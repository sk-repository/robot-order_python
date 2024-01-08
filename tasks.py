import time
from datetime import datetime, timedelta

from robocorp.tasks import task
from robocorp import vault

from RPA.HTTP import HTTP
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive
from RPA.Tables import Tables


receipt_dir='output/receipt/'
img_dir='output/img/'
zip_dir='output/archives/'

browser = Selenium()

@task
def order_robots_from_RobotSpareBin():
    get_order_list()
    open_robot_order_website()
    make_orders_from_csv()
    create_zip_archive()
    cleanup_temporary_files()
    # browser.close_browser()

def get_order_list():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", target_file="/output/orders.csv", overwrite=True)

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.open_browser("https://robotsparebinindustries.com/#/robot-order")

def click_ok_on_popup():
    browser.wait_until_page_contains_element("class:alert-buttons")
    browser.click_button("OK")

def click_order_confirm():
    browser.click_button("id:order")
    browser.wait_until_element_is_visible("id:receipt")
    
def return_to_order_form():
    browser.wait_until_element_is_visible("id:order-another", timeout=10)
    browser.click_button("id:order-another")

def make_singe_order(order):
    click_ok_on_popup()
    browser.wait_until_page_contains_element("tag:form") 
    browser.select_from_list_by_index("head", order["Head"])
    browser.select_radio_button("body", order["Body"])
    browser.input_text("xpath://html/body/div/div/div[1]/div/div[1]/form/div[3]/input", order["Legs"])
    browser.input_text("address", order["Address"])
    browser.click_button("Preview")
    click_order_confirm()

def make_pdf_for_single_receipt(receipt_filename, image_filename):
    pdf = PDF()
    pdf.open_pdf(receipt_filename)
    file_list = [receipt_filename, image_filename]
    pdf.add_files_to_pdf(file_list, receipt_filename, append=False)
    pdf.close_pdf()

def make_order_receipt():
    fs = FileSystem()
    fs.create_directory(receipt_dir)
    fs.create_directory(img_dir)
    browser.wait_until_element_is_visible("id:receipt")
    order_id = browser.get_text('//*[@id="receipt"]/p[1]')
    receipt_filename = f"{receipt_dir}receipt_{order_id}.pdf"
    receipt_html = browser.get_element_attribute('//*[@id="receipt"]', 'outerHTML')
    fs.create_file(receipt_filename, receipt_html)

    pdf = PDF()
    pdf.html_to_pdf(content=receipt_html, output_path=receipt_filename)

    browser.wait_until_element_is_visible("id:robot-preview-image")
    image_filename = f"{img_dir}robot_{order_id}.png"
    browser.screenshot("id:robot-preview-image", image_filename)
    make_pdf_for_single_receipt(receipt_filename, image_filename)

def create_zip_archive():
    fs = FileSystem()
    fs.create_directory(zip_dir)
    date = datetime.now()
    timestamp = date.strftime("%d%m%Y%H%M%S")
    zip_file_name = f"{zip_dir}/{timestamp}.zip"
    archive = Archive()
    archive.archive_folder_with_zip(receipt_dir, zip_file_name)

def cleanup_temporary_files():
    fs = FileSystem()
    fs.remove_directory(receipt_dir, recursive=True)
    fs.remove_directory(img_dir, recursive=True)

def make_orders_from_csv():
    tables = Tables()
    orders = tables.read_table_from_csv("/output/orders.csv")

    for order in orders:
        while True:
            try:
                make_singe_order(order)
                make_order_receipt()
                return_to_order_form()
            except Exception as e:
                browser.reload_page()
                continue
            break

if __name__ == "__main__":
    order_robots_from_RobotSpareBin()
