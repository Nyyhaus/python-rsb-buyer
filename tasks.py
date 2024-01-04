from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive

page = browser.page()
http = HTTP()
pdf = PDF()
lib = Archive()

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    download_orders_file()
    process_orders_file()
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def download_orders_file():
    """Downloads csv from given URL"""
    
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def process_orders_file():
    """Reads values from csv as a table and loops them"""
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"],
        header=True,
    )
    
    for order in orders:
        fill_and_submit_orders_form(order)
        
def fill_and_submit_orders_form(order):
    """Close cookies and fill form"""
    close_annoying_modal()
    fill_the_form(order)
    
def close_annoying_modal():
    """Close the cookie prompt"""
    page.click("button:text('OK')")
    
def fill_the_form(order):
    """Fill the form fields"""
    
    page.select_option("#head", order["Head"])
    page.click(f'#id-body-{order["Body"]}') 
    page.fill(".form-control", order["Legs"])
    page.fill("#address", order["Address"])
    
    preview_and_order(order)

def preview_and_order(order):
    """Preview robot and order it"""
    
    page.click("button:text('Preview')")
    page.click("button:text('Order')")
    
    success = False
    while not success:
        try:
            store_receipt_as_pdf(order["Order number"])
            page.click("button:text('Order another robot')")
            success = True
        except Exception as e:
            print(e)
            page.click("button:text('Order')")
    
def store_receipt_as_pdf(order_number):
    """Converts the receipt into pdf with a screenshot"""  
    order_results_html = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(order_results_html, f"output/receipts/receipt{order_number}.pdf")
    screenshot_robot_and_embed_to_pdf(order_number)

def screenshot_robot_and_embed_to_pdf(order_number):
    """Takes a screenshot of a receipt and appends it to pdf""" 
    page.screenshot(path=f"output/order_summary{order_number}.png")
    list_of_files = [
        f"output/order_summary{order_number}.png",
    ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=f"output/receipts/receipt{order_number}.pdf",
        append=True,
    )
    
def archive_receipts():
    """Zips the receipt pdfs because Control Room does not support subdirectories"""
    lib.archive_folder_with_zip('./output/receipts', 'orders.zip', recursive=True)
