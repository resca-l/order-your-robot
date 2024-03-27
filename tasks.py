from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


@task
def oreder_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100,)
    open_the_intranet_website()
    orders = get_orders()
    insert_orders(orders)
    archive_receipts()

def archive_receipts():
    """Zip the directory"""
    zip = Archive()
    zip.archive_folder_with_zip(folder="output/receipts", archive_name="all_the_receipts.zip")

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def get_orders():
    """download orders file and return it as a table"""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)
    csv = Tables()
    workbook = csv.read_table_from_csv("orders.csv", header= True, delimiters=",", dialect="excel")
    return workbook

def insert_orders(table):
    """insert orders from a table"""
    page = browser.page()
    for elements in table:
        fill_the_form(elements)

        store_receipt_as_pdf(elements["Order number"])
        screenshot_robot(elements["Order number"])
        embed_screenshot_to_receipt(elements["Order number"])
       
        page.click("#order-another")

def embed_screenshot_to_receipt(number):
    """Merge receipt and preview"""
    img = [f"output/receipts/{number}_robot_preview.png"]
    target_pdf = f"output/receipts/{number}_receipt.pdf"
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=img,
        target_document=target_pdf,
        append = True
    )

def fill_the_form(order):
    """Fills with data from a row"""
    page = browser.page()
    close_annoying_modal()
    page.select_option("#head", str(order["Head"]))
    body = "#id-body-"+str(order["Body"])
    page.click(body)
    page.fill(".form-control", str(order["Legs"]))
    page.fill("#address", order["Address"])
    page.click("button:text('preview')")
    page.click("button:text('order')")

    while not page.query_selector("#order-another"):
        page.click("button:text('order')")
    

def screenshot_robot(order_number):
    """Take a screenshot of the robot image"""
    page = browser.page().locator("#robot-preview-image")
    page.screenshot(path=f"output/receipts/{order_number}_robot_preview.png")
     
def store_receipt_as_pdf(order_number):
    """Print an order receipt"""
    page = browser.page()
    pdf = PDF()
    receipt_html = page.locator("#order-completion").inner_html()
    pdf.html_to_pdf(receipt_html, f"output/receipts/{order_number}_receipt.pdf")

def close_annoying_modal():
    """Closing cookies and more"""
    page = browser.page()
    page.click("button:text('OK')")

