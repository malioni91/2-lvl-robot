from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive

import time


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=300,
    )
    page = browser.page()
    open_robot_order_website()
    download_csv()
    close_annoying_modal()
    orders = get_orders()
    for order in orders:
        fill_the_form(order) 
    
    archive_receipts()
            
    
def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv():
    """ Downloads excel file from given URL """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """ Read data from csv """
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return orders

def close_annoying_modal():
    """ Closes the website modal by clicking on OK"""
    page = browser.page()
    page.click("button:text('OK')")
    
    
def fill_the_form(order_info):
    """ Fills the form with data taken from the csv file"""
    page = browser.page()
    page.select_option("#head", str(order_info["Head"]))
    page.click('id=id-body-{0}'.format(order_info["Body"]))  
    legs_id = page.get_attribute('xpath=//label[text()="3. Legs:"]', 'for')
    page.fill("input[placeholder='Enter the part number for the legs']", str(order_info["Legs"]))
    page.fill("id=address", order_info["Address"])
    page.click("text=Preview")
    page.click("#order")
    
    if not page.locator("//div[contains(@class, 'alert-danger')]").is_visible():
        screenshot = screenshot_robot(str(order_info["Order number"]))
        pdf = store_receipt_as_pdf(str(order_info["Order number"]))  
        #embed_screenshot_toPDF(screenshot, pdf)
    
    while page.locator("//div[contains(@class, 'alert-danger')]").is_visible():
        try:
            page.click("#order", timeout=1000)
            screenshot = screenshot_robot(str(order_info["Order number"]))
            pdf = store_receipt_as_pdf(str(order_info["Order number"]))  
            #embed_screenshot_toPDF(screenshot, pdf)
        except Exception:
            break
    if page.locator("#order-another").is_visible():
        page.click("#order-another") 
        if page.locator("button:text('OK')").is_visible():
            close_annoying_modal()
         
    
    
def store_receipt_as_pdf(order_number):
    """Stores the order receipt as a PDF file and returns the file path"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    
    pdf = PDF()
    pdf_path = f"output/receipts/order_{order_number}_receipt.pdf"
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """ Takes a screenshot of the ordered robot and return file path """
    page = browser.page()
    screenshot_path = f"output/screenshots/order_{order_number}_screenshot.png"
    locator = page.locator("#robot-preview-image")
    page.screenshot(path=screenshot_path)
    return screenshot_path


def embed_screenshot_toPDF(pdf_path, screenshot_path):
    pdf = PDF()
    pdf.add_files_to_pdf(screenshot_path, target_document=pdf_path, append=True)
    

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'output/receipts.zip')
