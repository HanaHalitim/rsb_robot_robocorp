import os
import logging
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from dotenv import load_dotenv
from tenacity import retry, stop_after_attempt, wait_fixed

# Load credentials from .env file
load_dotenv('credentials.env')
USERNAME = os.getenv("ROBOT_USERNAME")
PASSWORD = os.getenv("ROBOT_PASSWORD")

# Configure logging
LOG_FILE = "robot.log"
logging.basicConfig(
    filename=LOG_FILE,
    filemode="w",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger()

# Ensure the output directory exists
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Custom exception for process errors
class ProcessError(Exception):
    pass

@task
def robot_spare_bin_python():
    """Main task to process sales data and export it as a PDF."""
    try:
        logger.info("Starting the robot process...")
        initialize_bot()
        process_sales_data()
        logger.info("Process completed successfully.")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")
    finally:
        end_process()

def initialize_bot():
    """Initialize browser and log in."""
    try:
        logger.info("Initializing browser...")

        # Check if credentials are loaded
        if not USERNAME or not PASSWORD:
            logger.error("Credentials are missing in the .env file.")
            raise ProcessError("Environment variables USERNAME and PASSWORD are required.")

        browser.configure(slowmo=100)
        browser.goto("https://robotsparebinindustries.com/")
        log_in()
        logger.info("Initialization complete.")
    except Exception as e:
        logger.error(f"Initialization failed: {e}")
        raise ProcessError("Initialization failed.")

def log_in():
    """Logs into the website using credentials."""
    try:
        logger.info("Attempting to log in...")
        page = browser.page()
        page.fill("#username", USERNAME)
        page.fill("#password", PASSWORD)
        page.click("button:text('Log in')")
        logger.info("Login successful.")
    except Exception as e:
        logger.error("Login failed.")
        raise ProcessError("Login failed.")

@retry(stop=stop_after_attempt(3), wait=wait_fixed(2))
def download_excel_file():
    """Download Excel file with retry logic."""
    try:
        logger.info("Downloading Excel file...")
        http = HTTP()
        http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)
        logger.info("Excel file downloaded successfully.")
    except Exception as e:
        logger.error("Failed to download Excel file.")
        raise ProcessError("Failed to download Excel file.")

def process_sales_data():
    """Read Excel data and fill forms."""
    try:
        download_excel_file()
        excel = Files()
        excel.open_workbook("SalesData.xlsx")
        data_rows = excel.read_worksheet_as_table("data", header=True)
        excel.close_workbook()

        logger.info(f"Processing {len(data_rows)} rows of sales data...")
        for row in data_rows:
            try:
                logger.info(f"Processing data for {row['First Name']} {row['Last Name']}")
                fill_and_submit_sales_form(row)
            except Exception as e:
                logger.error(f"Error processing row {row}: {e}")
                continue  # Log the error but continue processing other rows
        logger.info("All sales data processed.")
    except Exception as e:
        logger.error(f"Error during data processing: {e}")
        raise ProcessError("Failed to process sales data.")

def fill_and_submit_sales_form(sales_rep):
    """Fill and submit sales form for a single representative."""
    try:
        page = browser.page()
        page.fill("#firstname", sales_rep["First Name"])
        page.fill("#lastname", sales_rep["Last Name"])
        page.select_option("#salestarget", str(sales_rep["Sales Target"]))
        page.fill("#salesresult", str(sales_rep["Sales"]))
        page.click("text=Submit")
        logger.info(f"Form submitted for {sales_rep['First Name']} {sales_rep['Last Name']}")
    except Exception as e:
        logger.error(f"Failed to submit form for {sales_rep}: {e}")
        raise

def end_process():
    """Take a screenshot, export results to PDF, and log out."""
    try:
        logger.info("Taking screenshot and exporting results to PDF...")
        page = browser.page()

        # Save screenshot to output directory
        screenshot_path = os.path.join(OUTPUT_DIR, "sales_summary.png")
        page.screenshot(path=screenshot_path)
        logger.info(f"Screenshot saved to {screenshot_path}")

        # Export sales results to PDF in output directory
        sales_results_html = page.locator("#sales-results").inner_html()
        pdf_path = os.path.join(OUTPUT_DIR, "sales_results.pdf")
        pdf = PDF()
        pdf.html_to_pdf(sales_results_html, pdf_path)
        logger.info(f"PDF export completed and saved to {pdf_path}")

        page.click("text=Log out")
        logger.info("Logged out successfully.")
    except Exception as e:
        logger.error(f"End process failed: {e}")
