from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
import logging
import shutil

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    logger = logging.getLogger(__name__)
    open_robot_order_website()
    orders = get_orders()
    for order in orders:
        logger.info(order)
        close_modal()
        fill_form(order)
    zip_receipts()
    

def open_robot_order_website():
    """Configure browser and open website."""
    browser.configure(slowmo=100,)
    page = browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """
    Downloads a csv file from a specified link
    Returns the content of the csv file as a table.
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", target_file="orders.csv")
    tableReader = Tables()
    tableData = tableReader.read_table_from_csv(path="orders.csv", header=True)
    return tableData

def close_modal():
    '''Close the popup modal.'''
    page = browser.page()
    page.click("text=OK")

def fill_form(data):
    '''Fill the robot order form with data from the argument. Handle errors with submit button.'''
    page = browser.page()
    page.select_option(selector="id=head", value=data['Head'])
    page.click("id=id-body-" + data['Body'])
    page.get_by_placeholder("Enter the part number for the legs").fill(data['Legs'])
    page.fill("id=address", data['Address'])
    page.click("text=Preview")
    failed = True
    while (failed):
        page.click("id=order")
        if page.get_by_text("Receipt").is_visible():
            failed = False
    pdf_path = store_receipt_as_pdf(data['Order number'])
    screenshot_path = screenshot_robot(data['Order number'])
    embed_screenshot(screenshot_path, pdf_path)
    page.click("id=order-another")

def store_receipt_as_pdf(order_number):
    '''Find the receipt info and store it as a pdf.'''
    pdf = PDF()
    page = browser.page()
    receipt = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(receipt, "output/receipts/receipt" + order_number + ".pdf")
    return "output/receipts/receipt" + order_number + ".pdf"

def screenshot_robot(order_number):
    """Take a screenshot of the robot"""
    page = browser.page()
    page.locator('#robot-preview-image').screenshot(path="output/screenshot_robot" + order_number + ".png")
    return "output/screenshot_robot" + order_number + ".png"

def embed_screenshot(screenshot, pdf_file):
    '''Adds the screenshot of the robot to the corresponding pdf file'''
    pdf = PDF()
    pdf.add_files_to_pdf([screenshot], pdf_file, True)

def zip_receipts():
    '''Zips the created receipt files'''
    shutil.make_archive('output/receipts', 'zip', 'output/receipts')
