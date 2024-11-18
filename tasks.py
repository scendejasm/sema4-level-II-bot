import os
import zipfile

from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem

@task
def order_robots_from_RobotSpareBin():
    """Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        #slowmo=200,
        screenshot="on"
    )
    open_the_robot_order_website()
    #print(get_orders())
    get_orders()
    place_orders()
    archive_receipts()
   

def open_the_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page=browser.page()
    page.click("button:text('OK')")

def get_orders():
    """Downloads csv file from the given URL"""
    http = HTTP()
    url = "https://robotsparebinindustries.com/orders.csv"
    http.download(url=url, overwrite=True)

def place_orders():
    orders = Tables().read_table_from_csv("orders.csv")
    for order in orders:
        fill_the_form(order)
        submit_the_order()
        store_receipt_as_pdf(order['Order number'])
        screenshot_robot(order['Order number'])
        embed_screenshot_to_receipt(
            order['Order number'],
            order['Order number']
        )
        goto_next_order()

def fill_the_form(order):
    """fill data from the csv into the orders page"""
    page = browser.page()
    print("DEBUG: fill_the_form for order", order['Order number'])
    page.select_option("#head", order['Head'])
    page.click(f"input[name='body'][value='{order['Body']}']")
    page.fill("input[placeholder*='legs']", order['Legs'])
    page.fill("#address", order['Address'])

def submit_the_order():
    """Click the order button to finalize the order and handle possible errors"""
    page = browser.page()
    max_retries = 5 # Limit the number of retries
    retries = 0

    while retries < max_retries:
        try:
            print(f"Attempting to place the order (Attempt {retries + 1})...")
            page.click("button:text('Order')")

            # Wait for potential error or success
            page.wait_for_timeout(1000) # 1-second delay 
            if page.locator(".alert.alert-danger").count() > 0:
                print("An error was encountered: Retrying order")
                retries += 1
                continue
            else:
                print("Order has been placed successfully!")
                break
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            retries += 1
    else:
        print("Failed to place the order after maximum retries")

def store_receipt_as_pdf(order_number):
    """creates a pdf of the order reciept
        and places inside 
        orders/receipts/receipt_{order_number}.pdf
    """
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(order_receipt_html, f"output/receipts/receipt_{order_number}.pdf")

def screenshot_robot(order_number):
    """Takes a screenshot of the robot image displayed after
        an order is placed
    """
    page = browser.page()
    preview_image = page.locator("#robot-preview-image")
    clip_area = preview_image.bounding_box() #{'x': 100, 'y': 150, 'width':300, 'height': 200 }
    page.screenshot(clip=clip_area ,path=f"output/previews/preview_image_{order_number}.png")

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Appends the robot preview image to the related 
        receipt pdf file
    """
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[f"output/previews/preview_image_{screenshot}.png"],
        target_document=f"output/receipts/receipt_{pdf_file}.pdf",
        append=True
    )

def goto_next_order():
    """After the order has been placed and records of receipt
        and preview have been taken, click the button to go to 
        the next order page."""
    page = browser.page()
    page.click("button:text('Order another robot')")
    page.click("button:text('OK')")

def archive_receipts():
    source_folder = "output/receipts/" #adjust to your subdirectory
    output_zip = "archive_pdfs.zip"
    output_folder = "output/"

    pdf_files = [f for f in os.listdir(source_folder) if f.endswith(".pdf")]
    # Find all .pdf files in the subdirectory 

    if not pdf_files:
        print("No PDF files found in the specified direcory")
        return
    
    # Define the full path to the output zip file
    output_zip = os.path.join(output_folder, output_zip)
    
    # Create the zip archive from the found PDF files
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in pdf_files:
            file_path = os.path.join(source_folder, file)
            zipf.write(file_path, arcname=file) 

    print(f"Zip file created: {output_zip}")
