from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import os, shutil

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
        slowmo=100,
    )
    page = browser.page() # Use this as an argument
    
    open_robot_order_website()
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
        store_receipt_as_pdf(order["Order number"])
        page.click("#order-another")
    archive_receipts()

    # Cleanup function
    if os.path.isdir("output/receipts/"):
        shutil.rmtree("output/receipts/")

def open_robot_order_website():
    """
    Opens the Robot Order Website
    """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """
    Download Orders CSV
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return orders

def close_annoying_modal():
    """
    Close pop up
    """
    page = browser.page()
    page.click("button:text('OK')")

def fill_the_form(order):
    """
    Fill form with order data
    """
    page = browser.page()

    page.select_option("#head", str(order["Head"]))
    page.check("#id-body-" + str(order["Body"]))
    page.fill("xpath=//input[contains(@placeholder, 'legs')]", str(order["Legs"]))
    page.fill("#address", str(order["Address"]))
    page.click("#preview")
    
    page.click("#order")
    error = page.locator("xpath=//div[@role='alert'][contains(@class, 'danger')]").is_visible()
    
    if error:
        error_message = page.locator("xpath=//div[@role='alert'][contains(@class, 'danger')]").inner_text()
        print(error_message)
        while error:
            page.click("#order")
            error = page.locator("xpath=//div[@role='alert'][contains(@class, 'danger')]").is_visible()

def take_screenshot(image_locator, order_number):
    """
    Take a screenshot of the element
    """
    image_locator.screenshot(path="output/receipts/" + str(order_number) + ".png")

def store_receipt_as_pdf(order_number):
    """
    Save receipt and picture as pdf
    """
    page = browser.page()
    receipt = page.locator("#order-completion").inner_html()

    take_screenshot(page.locator("#robot-preview-image"), order_number)

    pdf = PDF()
    pdf.html_to_pdf(receipt, "output/receipts/" + str(order_number) + ".pdf")
    pdf.add_files_to_pdf(
        files=[
            "output/receipts/" + str(order_number) + ".pdf",
            "output/receipts/" + str(order_number) + ".png"
            ],
        target_document="output/receipts/" + str(order_number) + ".pdf"
    )

def archive_receipts():
    """
    Archive receipts into a zip file.
    """
    lib = Archive()
    lib.archive_folder_with_zip(include="*.pdf", folder="output/receipts/", archive_name="output/receipts.zip")