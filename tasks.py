from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Browser.Selenium import Selenium
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100,)
    open_robot_order_website()
    close_annoying_modal()
    get_orders()
    archive_receipts()

def open_robot_order_website():
    """
    Opens robot order website.
    """

    page = browser.page()
    page.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    
    """Selects a button and closes the modal pop-up."""
    page = browser.page()
    page.click('//*[@class="alert-buttons"]/button[1]')


def get_orders():
    """
    Downloads the Excel file and reads it as a table.
    """
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite= True)

    library = Tables()
    orders = library.read_table_from_csv("orders.csv", columns= ["Order number", "Head", "Body", "Legs", "Address"])
    order = library.group_table_by_column(orders, "Order number")    
    
    for order in orders:
        fill_the_form(order)
        page = browser.page()

        while True:
            page.click("#order")
            order_another_btn = page.query_selector("#order-another")
            if order_another_btn:
                pdf_file = store_receipt_as_pdf(int(order["Order number"]))
                screenshot_path = screenshot_robot(int(order["Order number"]))
                embed_screenshot_to_receipt(screenshot_path, pdf_file)                
                order_another_robot()
                close_annoying_modal()
                break

def fill_the_form(order):
    """
    Fills order form and submits.
    """
    page = browser.page()
    page.select_option ("#head", order['Head'])
    page.click('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{0}]/label'.format(order['Body']))
    page.fill ("input[type='number']", order['Legs'])
    page.fill ("#address", order['Address'])
    page.click("#preview")    

def store_receipt_as_pdf(order_number):
    """
    Stores receipt as PDF.
    """
    page = browser.page()
    pdf = PDF()
    pdf_file = "output/receipts/{0}.pdf".format(order_number)
    receipt_info = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(receipt_info, pdf_file)
    return pdf_file        

def screenshot_robot(order_number):
    """
    Takes a screenshot of the page.
    """
    page = browser.page()
    screenshot_path = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path = screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot_path, pdf_file):
    """
    Embeds screenshot.
    """
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path = screenshot_path, source_path = pdf_file, output_path = pdf_file)

def order_another_robot():
    """
    Orders another robot.
    """
    page = browser.page()
    page.click("#order-another")  

def archive_receipts():
    """
    Combines receipts into one .zip file.
    """
    zip_file = Archive()
    zip_file.archive_folder_with_zip("./output/receipts/", "./output/receipts.zip")