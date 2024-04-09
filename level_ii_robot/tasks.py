from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """"
    Order robots from RobotSpareBin Industries Inc.
    Saves the order HTML reciept as a PDF file
    Saves the screenshot of the ordered robot
    Embeds the screenshot of the robot to the PDF receipt
    Creates ZIP archive of the reciepts and the images
    """
    browser.configure(slowmo=100)
    open_website()
    download_file()
    get_orders()
    orders=get_orders()
    for order in orders:
        close_warning()
        fill_form(order)
        screenshot_robot(order)
        store_receipt_as_pdf(order)
        embed_screenshot_to_receipt(order)
    archive_zip()

    #order_robot_order_website

def open_website():
    """Navigate to the given URL"""
    browser.goto('https://robotsparebinindustries.com/#/robot-order')

def download_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url='https://robotsparebinindustries.com/orders.csv', overwrite = True)

def get_orders():
    """Read data from csv and coverts to a data table"""
    library = Tables()
    orders=library.read_table_from_csv("orders.csv", columns=['Order number', 'Head', 
                                                              'Body', 'Legs','Address'])
    return orders

def close_warning():
    """Closes website pop-up window"""
    page=browser.page()
    page.click("button:text('OK')")

def fill_form(order_row):
    """Fills order form on website"""
    page = browser.page()
    
    body_list={ 1: 'Roll-a-thor body', 2: 'Peanut crusher body', 3:'D.A.V.E body', 
               4: 'Spanner mate body', 5: 'Andy Roid body', 6: 'Drillbit 2000 body'}
    
    page.select_option('#head', order_row['Head'])
    page.check('text={}'.format(body_list[int(order_row['Body'])]))
    page.fill('text=3. Legs:',order_row['Legs'])
    page.fill('#address',order_row['Address'])
    page.click('text=Preview')

def screenshot_robot(order_row):
    """Take a screenshot of the page"""
    page = browser.page()
    page.screenshot(path='output/screenshots/robot_{}.png'.format(str(order_row['Order number'])))

    page.on('dialog', lambda dialog: dialog.accept())
    page.click("button:text('Order')")
    while page.locator('css=div.alert-danger').is_visible() == True:
        page.click("button:text('Order')")

def store_receipt_as_pdf(order_row):
    """Export the order-completion to a pdf file"""
    page=browser.page()
    receipt_html=page.locator('#receipt').inner_html()
    
    pdf=PDF()
    pdf.html_to_pdf(receipt_html, 'output/receipts/receipt_{}.pdf'.format(str(order_row['Order number'])))

def embed_screenshot_to_receipt(order_row):
    """Adds screenshot of robot to reciept"""
    page=browser.page()
    pdf=PDF()
    
    order_num=order_row['Order number']
    file_list=["output/screenshots/robot_{}.png".format(order_num)]
    
    pdf.add_files_to_pdf(files=file_list, 
                         target_document="output/receipts/receipt_{}.pdf".format(order_num),
                         append=True)
    page.click('text=Order another robot')

def archive_zip():
    """zips file from output"""
    zip=Archive()
    zip.archive_folder_with_zip('./output/receipts','receipts.zip')
