from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF 
from RPA.Archive import Archive

@task
def order_bot():
    """Use the orders file and fill in the orders"""
    browser.configure(
        slowmo=500,
    )
    download_csv_file()
    open_the_order_form()
    fill_in_order_form_with_csv()
    archive_receipts()
    
    

def download_csv_file():
    """downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def open_the_order_form():
    """opens the order form"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    click_pop_up()

def fill_in_order_form_with_csv():
    """takes the csv data and loops it into the form"""
    libary = Tables()
    orders = libary.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
        )
   
    for row in orders:
        fill_and_submit_form(row)
       
        

def fill_and_submit_form(order):
    """fills in the form and submits it"""
    page = browser.page()

    page.select_option("#head", order["Head"])
    page.click('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{0}]/label'.format(order["Body"]))
    page.fill("input[placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", order["Address"])
    if True:
        page.click("button:text('ORDER')")
        order_another = page.query_selector("#order-another")
        if order_another:
             pdf_path = pdf_receipt(str(order["Order number"]))
             screenshot_path = screenshot_robot(str(order["Order number"]))
             embed_screenshot_to_receipt(screenshot_path, pdf_path)
             page.click("button:text('ORDER ANOTHER ROBOT')")
             click_pop_up()
            
    

def click_pop_up():
        """clicks the ok on the pop up"""
        page = browser.page()

        page.click("text=OK")    

def pdf_receipt (order_number):
    """makes a pdf of the receipt html"""
    page = browser.page()
    receipt_html = page.locator("#order-completion").inner_html()

    pdf = PDF()
    pdf_path = "output/pdf/{0}.pdf".format(order_number)
    pdf.html_to_pdf(receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
     """screenshots the robot and adds the screenshot to the pdf"""
     page = browser.page()

     screenshot_path = "output/screenshot/{0}.png".format(order_number)
     page.locator("#robot-preview-image").screenshot(path=screenshot_path)
     return screenshot_path
   
def embed_screenshot_to_receipt(screenshot_path, pdf_path):
     pdf = PDF()

     pdf.add_watermark_image_to_pdf(image_path=screenshot_path,
                                    source_path=pdf_path,
                                    output_path=pdf_path)
     
def archive_receipts():
     lib = Archive()
     lib.archive_folder_with_zip("./output/pdf", "./output/pdf.zip")