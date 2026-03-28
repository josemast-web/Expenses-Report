"""
main.py  –  QuickBooks Expense Report Bot

Fetches the latest QB Excel report from a Google Drive folder,
processes expense data, generates an HTML summary + TXT attachments,
and delivers the report via Gmail SMTP.

Required environment variables  (see .env.example):
  DRIVE_FOLDER_ID        Google Drive folder containing QB reports
  GCP_CREDENTIALS_FILE   Path to the service account JSON (default: credentials.json)
  EMAIL_SENDER           Gmail address used to send
  EMAIL_PASSWORD         Gmail App Password
  EMAIL_RECIPIENTS       Comma-separated recipient addresses
  HIGH_VALUE_THRESHOLD   Alert threshold in USD (default: 3000)
  REPORT_LABEL           Optional label shown in the email footer
"""

import os
import io
import re
import smtplib
import logging
import pandas as pd
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from dotenv import load_dotenv

# Load .env for local development (no-op in CI/CD where vars are injected directly)
load_dotenv()


# ---------------------------------------------------------------------------
# 1. CONFIGURATION
# ---------------------------------------------------------------------------
class Config:
    # Google Drive folder ID – loaded from env, never hard-coded
    DRIVE_FOLDER_ID   = os.getenv("DRIVE_FOLDER_ID")
    CREDENTIALS_FILE  = os.getenv("GCP_CREDENTIALS_FILE", "credentials.json")
    TEMP_FILE_NAME    = "temp_qb_report.xlsx"

    # Email settings – all from environment
    EMAIL_SENDER      = os.getenv("EMAIL_SENDER", "")
    EMAIL_PASSWORD    = os.getenv("EMAIL_PASSWORD")

    _recipients_raw   = os.getenv("EMAIL_RECIPIENTS", "")
    EMAIL_RECIPIENTS  = [r.strip() for r in _recipients_raw.split(",") if r.strip()]

    # Report display label (team or org name, no company name hard-coded)
    REPORT_LABEL      = os.getenv("REPORT_LABEL", "Automated Report Bot")

    # Excel column names
    COL_DATE    = "Transaction date"
    COL_MEMO    = "Memo/Description"
    COL_QTY     = "Quantity"
    COL_PRODUCT = "Product/Service"
    COL_VENDOR  = "Vendor"
    COL_CUSTOMER = "Customer"
    COL_AMOUNT  = "Amount"

    # Terms excluded from the expense analysis
    EXCLUDE_TERMS = [
        "Shipping Charged to Materials", "Coupon Discount", "Sales Tax",
        "TARIFFS FEES", "SURCHARGES", "Surcharge Fee", "Tariff Surcharge",
    ]

    # Transactions above this amount trigger a high-value alert
    HIGH_VALUE_THRESHOLD = float(os.getenv("HIGH_VALUE_THRESHOLD", "3000"))


# Startup validation
if not Config.EMAIL_PASSWORD:
    print("[WARNING] EMAIL_PASSWORD env var not set - email delivery will fail.")
if not Config.DRIVE_FOLDER_ID:
    print("[WARNING] DRIVE_FOLDER_ID env var not set - Drive access will fail.")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("ExpenseBot")


# ---------------------------------------------------------------------------
# 2. GOOGLE DRIVE MODULE
# ---------------------------------------------------------------------------
class DriveManager:
    def __init__(self):
        self.scopes  = ["https://www.googleapis.com/auth/drive.readonly"]
        self.service = self._authenticate()

    def _authenticate(self):
        try:
            if not os.path.exists(Config.CREDENTIALS_FILE):
                logger.error("[AUTH] Credentials file not found: %s", Config.CREDENTIALS_FILE)
                return None
            creds = service_account.Credentials.from_service_account_file(
                Config.CREDENTIALS_FILE, scopes=self.scopes
            )
            return build("drive", "v3", credentials=creds)
        except Exception as e:
            logger.error("[AUTH] Authentication error: %s", e)
            return None

    def download_latest_report(self):
        """Find and download the most recent QB report file from the configured folder."""
        if not self.service:
            return None

        query = (
            f"'{Config.DRIVE_FOLDER_ID}' in parents "
            "and name contains 'QB' "
            "and name contains 'report' "
            "and trashed = false"
        )
        results = self.service.files().list(q=query, fields="files(id, name)").execute()
        files   = results.get("files", [])

        if not files:
            logger.warning("[DRIVE] No matching QB report files found in folder.")
            return None

        logger.info("[DRIVE] Scanning %d candidate files...", len(files))

        latest_file = None
        latest_date = datetime.min

        for file in files:
            name  = file["name"]
            match = re.search(r"QB.*?report.*?(\d{2})(\d{2})(\d{4})", name, re.IGNORECASE)

            if match:
                month, day, year = match.groups()
                try:
                    file_date = datetime(int(year), int(month), int(day))
                    if file_date > latest_date:
                        latest_date = file_date
                        latest_file = file
                except ValueError:
                    continue

        if latest_file:
            logger.info("[DRIVE] Selected file: %s", latest_file["name"])
            return self._download_file(latest_file["id"])

        logger.warning("[DRIVE] No file matched the expected date-based naming pattern.")
        return None

    def _download_file(self, file_id: str) -> str:
        """Download a Drive file by ID and save it to a local temp path."""
        request    = self.service.files().get_media(fileId=file_id)
        fh         = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done       = False

        while not done:
            _, done = downloader.next_chunk()

        fh.seek(0)
        with open(Config.TEMP_FILE_NAME, "wb") as f:
            f.write(fh.read())

        logger.info("[DRIVE] File downloaded to: %s", Config.TEMP_FILE_NAME)
        return Config.TEMP_FILE_NAME


# ---------------------------------------------------------------------------
# 3. DATA PROCESSING MODULE
# ---------------------------------------------------------------------------
class DataProcessor:
    @staticmethod
    def load_and_clean(file_path: str) -> pd.DataFrame:
        """Read the QB Excel report and apply initial cleaning rules."""
        try:
            logger.info("[DATA] Reading Excel file: %s", file_path)
            df = pd.read_excel(file_path, header=4)

            if df.empty or Config.COL_DATE not in df.columns:
                logger.warning("[DATA] File appears empty or has unexpected format.")
                return pd.DataFrame()

            df.columns = df.columns.str.strip()

            # Remove excluded memo terms
            mask = ~df[Config.COL_MEMO].astype(str).str.contains(
                "|".join(Config.EXCLUDE_TERMS), case=False, na=False
            )
            df = df[mask]

            # Type coercions
            df[Config.COL_DATE]    = pd.to_datetime(df[Config.COL_DATE], errors="coerce")
            df[Config.COL_AMOUNT]  = pd.to_numeric(df[Config.COL_AMOUNT], errors="coerce").fillna(0)
            df[Config.COL_VENDOR]  = df[Config.COL_VENDOR].fillna("Unknown Vendor")
            df[Config.COL_QTY]     = pd.to_numeric(df[Config.COL_QTY], errors="coerce").fillna(0)
            df[Config.COL_PRODUCT] = df[Config.COL_PRODUCT].fillna("")

            logger.info("[DATA] Loaded %d rows after exclusions.", len(df))
            return df

        except Exception as e:
            logger.error("[DATA] Error processing file: %s", e)
            return pd.DataFrame()

    @staticmethod
    def get_period_data(
        df: pd.DataFrame,
        days_lookback: int = None,
        current_week: bool = False,
    ) -> pd.DataFrame:
        """Filter DataFrame to the requested time window."""
        now       = datetime.now()
        today_end = now.replace(hour=23, minute=59, second=59, microsecond=999999)

        if current_week:
            monday     = now - timedelta(days=now.weekday())
            start_date = monday.replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            past_date  = now - timedelta(days=days_lookback)
            start_date = past_date.replace(hour=0, minute=0, second=0, microsecond=0)

        return df[
            (df[Config.COL_DATE] >= start_date) & (df[Config.COL_DATE] <= today_end)
        ].copy()


# ---------------------------------------------------------------------------
# 4. REPORT RENDERING MODULE
# ---------------------------------------------------------------------------
class ReportRenderer:
    @staticmethod
    def generate_html_section(df: pd.DataFrame, title: str, is_alert_active: bool = True) -> str:
        """Build a self-contained HTML section for a given time window."""
        if df.empty:
            return f"""
            <div style="background-color:#f8f9fa; border:1px solid #ddd; padding:20px; border-radius:8px; margin-bottom:25px;">
                <h3 style="color:#6c757d; margin-top:0;">{title}</h3>
                <p style="color:#6c757d;">No transactions found for this period.</p>
            </div>
            """

        total_spend  = df[Config.COL_AMOUNT].sum()
        count_tx     = len(df)
        count_vendor = df[Config.COL_VENDOR].nunique()
        count_proj   = df[Config.COL_CUSTOMER].nunique()

        # Top-10 aggregations
        vendor_stats = (
            df.groupby(Config.COL_VENDOR)
            .agg({Config.COL_AMOUNT: "sum", Config.COL_DATE: "count"})
            .sort_values(by=Config.COL_AMOUNT, ascending=False)
            .head(10)
        )
        project_stats = (
            df.groupby(Config.COL_CUSTOMER)
            .agg({Config.COL_AMOUNT: "sum", Config.COL_DATE: "count"})
            .sort_values(by=Config.COL_AMOUNT, ascending=False)
            .head(10)
        )

        # Unassigned transactions alert
        unassigned = df[df[Config.COL_CUSTOMER].isna() | (df[Config.COL_CUSTOMER] == "")]
        alert_html = ""
        if not unassigned.empty and is_alert_active:
            alert_html = f"""
            <div style="background-color:#fff3cd; border-left:5px solid #ffc107; padding:10px; margin:15px 0; color:#856404;">
                <strong>ATTENTION:</strong> {len(unassigned)} transactions are missing a Project/Customer assignment.
            </div>
            """

        # High-value purchases block
        high_value_df  = df[df[Config.COL_AMOUNT] > Config.HIGH_VALUE_THRESHOLD].sort_values(
            by=Config.COL_AMOUNT, ascending=False
        )
        high_value_html = ""
        if not high_value_df.empty:
            list_items = ""
            for _, row in high_value_df.iterrows():
                date_str = row[Config.COL_DATE].strftime("%m/%d")
                vendor   = row[Config.COL_VENDOR]
                amount   = row[Config.COL_AMOUNT]
                memo     = str(row[Config.COL_PRODUCT]) if pd.notna(row[Config.COL_PRODUCT]) else "No Desc"
                if len(memo) > 50:
                    memo = memo[:47] + "..."

                list_items += f"""
                <li style="margin-bottom:8px; border-bottom:1px dashed #eee; padding-bottom:5px;">
                    <span style="font-weight:bold; color:#d9534f;">${amount:,.2f}</span>
                    <span style="color:#555;"> | {date_str} | <strong>{vendor}</strong></span>
                    <br><span style="font-size:11px; color:#888; font-style:italic;">{memo}</span>
                </li>
                """

            high_value_html = f"""
            <div style="margin-top:25px; background-color:#fff0f0; border:1px solid #ffcccc; border-radius:5px; padding:15px;">
                <h4 style="margin-top:0; color:#c9302c; border-bottom:1px solid #e6b3b3; padding-bottom:5px;">
                    High Value Purchases (&gt; ${Config.HIGH_VALUE_THRESHOLD:,.0f})
                </h4>
                <ul style="list-style-type:none; padding-left:0; margin:0; font-size:13px;">
                    {list_items}
                </ul>
            </div>
            """

        def build_table(df_agg, col_title):
            rows = ""
            for name, row_data in df_agg.iterrows():
                amount = row_data[Config.COL_AMOUNT]
                count  = int(row_data[Config.COL_DATE])
                rows += f"""
                <tr>
                    <td style="padding:6px; border-bottom:1px solid #eee; font-size:11px;">{name}</td>
                    <td style="padding:6px; border-bottom:1px solid #eee; text-align:center; font-size:11px;">{count}</td>
                    <td style="padding:6px; border-bottom:1px solid #eee; text-align:right; font-size:11px;">${amount:,.2f}</td>
                </tr>
                """
            return f"""
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead>
                    <tr style="background-color:#e9ecef;">
                        <th style="padding:8px; text-align:left;">{col_title}</th>
                        <th style="padding:8px; text-align:center;"># Items</th>
                        <th style="padding:8px; text-align:right;">Amount</th>
                    </tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
            """

        return f"""
        <div style="background-color:#ffffff; border:1px solid #e0e0e0; border-radius:8px; padding:20px; margin-bottom:30px; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
            <h2 style="color:#2c3e50; margin-top:0; border-bottom:2px solid #0056b3; padding-bottom:10px; font-size:18px;">{title}</h2>

            <table width="100%" cellpadding="10" style="margin-bottom:20px;">
                <tr>
                    <td width="25%" style="background:#f8f9fa; text-align:center; border-radius:5px;">
                        <div style="font-size:10px; color:#666; text-transform:uppercase;">Total Spend</div>
                        <div style="font-size:18px; color:#d9534f; font-weight:bold;">${total_spend:,.0f}</div>
                    </td>
                    <td width="25%" style="text-align:center;">
                        <div style="font-size:10px; color:#666;">Transactions</div>
                        <div style="font-size:16px; font-weight:bold;">{count_tx}</div>
                    </td>
                    <td width="25%" style="text-align:center;">
                        <div style="font-size:10px; color:#666;">Vendors</div>
                        <div style="font-size:16px; font-weight:bold;">{count_vendor}</div>
                    </td>
                    <td width="25%" style="text-align:center;">
                        <div style="font-size:10px; color:#666;">Active Projects</div>
                        <div style="font-size:16px; font-weight:bold;">{count_proj}</div>
                    </td>
                </tr>
            </table>

            {alert_html}

            <table width="100%" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="48%" valign="top">{build_table(vendor_stats,  "Top 10 Vendors")}</td>
                    <td width="4%"></td>
                    <td width="48%" valign="top">{build_table(project_stats, "Top 10 Projects")}</td>
                </tr>
            </table>

            {high_value_html}
        </div>
        """

    @staticmethod
    def create_txt_attachments(df_week: pd.DataFrame, df_month: pd.DataFrame) -> list:
        """Write weekly and monthly plain-text reports to disk and return their paths."""
        files = []

        # --- Weekly transaction list ---
        f1_name = "Weekly_Transactions_List.txt"
        with open(f1_name, "w", encoding="utf-8") as f:
            f.write(f"WEEKLY REPORT - Generated: {datetime.now()}\n")
            f.write("Format: [Date] [Qty] [Product] | Amount | Vendor | Project\n")
            f.write("=" * 80 + "\n\n")

            if not df_week.empty:
                for _, row in df_week.sort_values(Config.COL_DATE).iterrows():
                    date_str = row[Config.COL_DATE].strftime("%Y-%m-%d")
                    qty      = int(row[Config.COL_QTY]) if pd.notna(row[Config.COL_QTY]) else 0
                    prod     = str(row[Config.COL_PRODUCT])[:25]
                    amt      = row[Config.COL_AMOUNT]
                    vend     = str(row[Config.COL_VENDOR])[:20]
                    cust     = str(row[Config.COL_CUSTOMER]) if pd.notna(row[Config.COL_CUSTOMER]) else "N/A"

                    f.write(f"[{date_str}] [{qty:<3}] {prod:<25} | ${amt:<8.2f} | {vend:<20} | {cust}\n")
            else:
                f.write("No transactions found.\n")
        files.append(f1_name)

        # --- Monthly project breakdown ---
        f2_name = "Monthly_By_Project.txt"
        with open(f2_name, "w", encoding="utf-8") as f:
            f.write(f"MONTHLY PROJECT REPORT - Generated: {datetime.now()}\n")
            f.write("Includes: Product & Quantity info\n")
            f.write("=" * 80 + "\n")

            if not df_month.empty:
                df_m = df_month.copy()
                df_m["Week"] = df_m[Config.COL_DATE].dt.isocalendar().week
                df_m[Config.COL_CUSTOMER] = df_m[Config.COL_CUSTOMER].fillna("UNASSIGNED")
                df_m = df_m.sort_values(by=["Week", Config.COL_CUSTOMER])

                for (cust, wk), group in df_m.groupby([Config.COL_CUSTOMER, "Week"]):
                    f.write(f"\nPROJECT: {cust} (Week {wk}) [Items: {len(group)}]\n")
                    f.write("-" * 80 + "\n")

                    for _, row in group.iterrows():
                        date_str = row[Config.COL_DATE].strftime("%m-%d")
                        qty      = int(row[Config.COL_QTY])
                        f.write(
                            f"   * {date_str} (Qty:{qty}) {row[Config.COL_PRODUCT]}"
                            f" | ${row[Config.COL_AMOUNT]:.2f} - {row[Config.COL_VENDOR]}\n"
                        )

                    f.write(f"   >>> SUBTOTAL: ${group[Config.COL_AMOUNT].sum():.2f}\n")

        files.append(f2_name)
        return files


# ---------------------------------------------------------------------------
# 5. EMAIL MODULE
# ---------------------------------------------------------------------------
class EmailService:
    @staticmethod
    def send_report(html_content: str, attachments: list) -> None:
        """Send the HTML report email with file attachments via Gmail SMTP."""
        msg = MIMEMultipart()
        msg["From"]    = Config.EMAIL_SENDER
        msg["To"]      = ", ".join(Config.EMAIL_RECIPIENTS)
        msg["Subject"] = f"Expense Report: {datetime.now().strftime('%Y-%m-%d')}"

        msg.attach(MIMEText(html_content, "html"))

        for filepath in attachments:
            if os.path.exists(filepath):
                with open(filepath, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(filepath)}",
                )
                msg.attach(part)

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(Config.EMAIL_SENDER, Config.EMAIL_PASSWORD)
            server.sendmail(Config.EMAIL_SENDER, Config.EMAIL_RECIPIENTS, msg.as_string())
            server.quit()
            logger.info("[EMAIL] Report sent successfully.")
        except Exception as e:
            logger.error("[EMAIL] Failed to send email: %s", e)


# ---------------------------------------------------------------------------
# 6. MAIN ORCHESTRATOR
# ---------------------------------------------------------------------------
def main() -> None:
    logger.info("[MAIN] Starting ExpenseBot...")

    # Step 1: Download latest QB report from Drive
    drive     = DriveManager()
    file_path = drive.download_latest_report()
    if not file_path:
        logger.error("[MAIN] No file downloaded. Aborting.")
        return

    # Step 2: Load and clean data
    df = DataProcessor.load_and_clean(file_path)
    if df.empty:
        logger.error("[MAIN] DataFrame is empty after processing. Aborting.")
        return

    # Step 3: Slice data by time window
    df_week  = DataProcessor.get_period_data(df, current_week=True)
    df_month = DataProcessor.get_period_data(df, days_lookback=30)

    # Step 4: Build HTML report
    html_header = f"""
    <div style="font-family:Arial, sans-serif; color:#333; max-width:900px; margin:auto;">
        <div style="background-color:#004085; color:white; padding:15px; border-radius:8px 8px 0 0; text-align:center;">
            <h1 style="margin:0; font-size:24px;">Expense Report</h1>
            <p style="margin:5px 0 0 0; font-size:14px; opacity:0.9;">Automated Financial Summary</p>
        </div>
        <div style="padding:20px; background-color:#f4f6f9;">
            <p><strong>Good morning Team,</strong><br>Please find below the updated expenses analysis.</p>
    """

    html_footer = f"""
            <hr style="border:0; border-top:1px solid #ddd; margin:20px 0;">
            <p style="font-size:11px; color:#888; text-align:center;">
                Generated automatically by {Config.REPORT_LABEL}.<br>
                Source: Google Drive QB Report.
            </p>
        </div>
    </div>
    """

    full_html = (
        html_header
        + ReportRenderer.generate_html_section(df_week,  "Current Week Activity")
        + ReportRenderer.generate_html_section(df_month, "Last 30 Days Overview")
        + html_footer
    )

    # Step 5: Generate TXT attachments
    attachments = ReportRenderer.create_txt_attachments(df_week, df_month)

    # Step 6: Send email
    EmailService.send_report(full_html, attachments)

    # Cleanup temp files
    try:
        os.remove(file_path)
        for att in attachments:
            os.remove(att)
    except OSError:
        pass

    logger.info("[MAIN] Process completed successfully.")


if __name__ == "__main__":
    main()
