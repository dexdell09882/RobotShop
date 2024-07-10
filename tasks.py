from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
import time
from RPA.PDF import PDF
from RPA.Archive import Archive

order_number = "Order number"

head = "Head"

body = "Body"

legs = "Legs"

address = "Address"

@task
def order_robots_from_RobotSpareBin():

    browser.configure(
        slowmo=10,
    )

    open_robot_order_website()

    download_csv_orders_file()

    orders = get_orders()

    for row in orders:

        close_annoying_modal()

        fill_the_form(row)

        item_order_number = row[order_number]

        receipt = store_receipt_as_pdf(item_order_number)

        screenshot = screenshot_robot(item_order_number)

        embed_screenshot_to_receipt(screenshot, receipt)

    archive_receipts()
    

def open_robot_order_website():
    """Navigates to the given URL"""

    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def download_csv_orders_file():

    """Downloads excel file from the given URL"""

    http = HTTP()

    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def get_orders():
    """Read data from excel and fill in the sales form"""

    library = Tables()

    orders = library.read_table_from_csv(

    "orders.csv", columns=["Order number","Head","Body","Legs","Address"]

    )
    return orders

def close_annoying_modal():

    page = browser.page()

    page.click("button:text('OK')")


def fill_the_form(order):

    page = browser.page()
   
    page.select_option("//select[contains(@class, 'custom-select')]", order[head])

    page.click("input[type='radio'][value='" + order[body] + "']")

    page.fill("input.form-control[placeholder='Enter the part number for the legs']", order[legs])

    page.fill("input[type='text'][id='address']", order[address])

    page.click("#order")

    browser.configure(
        slowmo=5,
    )
  

def store_receipt_as_pdf(order_number):

    """Export the data to a pdf file"""

    browser.configure(
        slowmo=5,
    )

    page = browser.page()

    pdf = PDF()

    file_name = "output/receipt/order_number_" + order_number + ".pdf"

    #sales_results_html = page.locator("#receipt").inner_html()
    counter = 0
    while counter < 5:
        try:
            sales_results_html = page.locator("#receipt").inner_html(timeout=5000)
            counter = 6  
        except:
            try:
                page.click("#order",timeout=5000)
                browser.configure(
                    slowmo=5,
                )
                sales_results_html = page.locator("#receipt").inner_html(timeout=5000)
                counter = 6
            except:
                counter = counter + 1

    screenshot_file = screenshot_robot(order_number)

    #sales_results_html = sales_results_html+"<img src='screenshot_" + order_number + ".png'>"

    pdf.html_to_pdf(sales_results_html, file_name)
    
    #pdf.html_to_pdf(sales_results_html, file_name, working_directory="output/screenshot/")
    

    return file_name

def screenshot_robot(order_number):

    browser.configure(
        slowmo=10,
    )

    page = browser.page().locator("#robot-preview-image")

    file_name = "output/screenshot/screenshot_" + order_number + ".png"

    page.screenshot(path=file_name)

    return file_name
    
def embed_screenshot_to_receipt(screenshot, pdf_file):

    page = browser.page()

    #sales_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()

    #pdf.add_files_to_pdf([screenshot],pdf_file)
    
    #below is the recent one
    #pdf.add_files_to_pdf([pdf_file,screenshot,],pdf_file)
    pdf.add_watermark_image_to_pdf(image_path=screenshot, source_path=pdf_file,  output_path=pdf_file)

    #pdf.html_to_pdf(sales_results_html+"<img src='"+screenshot.replace('output/screenshot/','')+"'>", pdf_file, working_directory="output/screenshot/")
    
    

    page = browser.page()

    browser.configure(
        slowmo=5,
    )

    page.click("#order-another")
   


def archive_receipts():

    lib = Archive()

    lib.archive_folder_with_zip('output/receipt', 'output/archive_receipts.zip')