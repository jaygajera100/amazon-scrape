# pyinstaller --onefile --hidden-import selenium --add-binary "./drivers/chromedriver;./drivers/" amazon.py
import csv
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
import os
import time
# Project Start From here.


def get_chrome_web_driver(options):
    return webdriver.Chrome("./drivers/chromedriver.exe", chrome_options=options)


def get_web_driver_options():
    chrome_options = webdriver.ChromeOptions()
    return chrome_options


def set_ignore_certificate_error(options):
    options.add_argument('--ignore-certificate-errors')


def set_browser_as_incognito(options):
    options.add_argument('--incognito')


class Amazon_API:
    def __init__(self):
        options = get_web_driver_options()
        set_browser_as_incognito(options)
        # set_automation_as_head_less(options)
        set_ignore_certificate_error(options)
        self.driver = get_chrome_web_driver(options)

    def GPU(self):
        gui = Tk()
        gui.geometry("200x200")
        gui.title("Application")

        c = ttk.Button(gui, text="Select File",
                       command=self.run()).pack(pady=20)

        gui.mainloop()
        gui.quit()

    def run(self):
        file_path = self.file_path()
        print("Start Scripting...")

        # Creating New Excel File
        fieldnames = ['Vendor Name', 'Title', 'SKU', 'Vary By', 'Option 1 Name', 'Option 1 Value', 'Option 2 Name', 'Option 2 Value', 'Option 3 Name', 'Option 3 Value',
                      'Brand', "HTML Description", 'Description', "HTML Feature", 'Feature', 'Condition', 'Condition Grading', 'Warranty', 'Packing', 'What you get',
                      'Estimated Shipping Cost', 'Quantity', 'Cost', 'Suggested Sale Price', 'MSRP',  'Image_link_1', 'Image_link_2', 'Image_link_3', 'Image_link_4', 'Image_link_5', 'Image_link_6', 'Image_link_7', 'Image_link_8', 'Image_link_9', 'Image_link_10', 'Image_link_11', 'Image_link_12']
        with open(f'report.csv', "w", newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()

        # Open Excel file to get the links
        with open(file_path, "r") as f:
            reader = csv.reader(f, delimiter=',')
            links = []
            for i, row in enumerate(reader):
                # if i == 0:
                #     links.append(row[0])
                links.append(row[0])

        number = 1
        # Iterate Each link to get the details of the product
        for link in links:
            print(number)
            try:
                self.driver.get(link)
                self.driver.execute_script("window.stop();")
                title = self.get_title()
                asin = self.get_asin(link)
                brand = "Generic"
                condition = "New"
                UPC = None
                description = self.get_description()
                price = self.get_price()
                sell_price = None
                cost = None
                quantity = 500
                color = self.get_color(link)
                size = self.get_size(link)
                feature = self.get_feature()
                model = None
                packaging = None
                warranty = "30 Days"
                weight = self.get_weight()
                in_box_details = self.get_in_box_details()
                sizes_asin, image_links = self.get_images_links(color)
                if len(size["size"]) > 1:
                    vary = "Multivariate"
                else:
                    vary = "Standard"

                # Making Columns for new excel file
                product_info = {
                    'Vendor Name': 'ETC BUYS INC.',
                    'Title': title,
                    'SKU': asin,
                    'Vary By': vary,
                    'Option 1 Name': "Color",
                    'Option 1 Value': color,
                    'Option 2 Name': "Size",
                    'Option 2 Value': size,
                    'Option 3 Name': "Weight",
                    'Option 3 Value': weight,
                    'Brand': brand,
                    'HTML Description': "",
                    'Description': description,
                    'HTML Feature': "",
                    'Feature': description,
                    'Condition': condition,
                    'Condition Grading': "",
                    'Warranty': warranty,
                    'Packing': 'Opp',
                    'What you get': in_box_details,
                    'Estimated Shipping Cost': "",
                    'Quantity': quantity,
                    'Cost': None,
                    'Suggested Sale Price': None,
                    'MSRP': price,
                    'Image_link_1': "",
                    'Image_link_2': "",
                    'Image_link_3': "",
                    'Image_link_4': "",
                    'Image_link_5': "",
                    'Image_link_6': "",
                    'Image_link_7': "",
                    'Image_link_8': "",
                    'Image_link_9': "",
                    'Image_link_10': "",
                    'Image_link_11': "",
                    'Image_link_12': "",

                }
                with open(f'report.csv', "a", newline='', encoding="utf-8") as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    for m, c in enumerate(size['size']):
                        product_info['Option 2 Value'] = c
                        # product_info['ASIN'] = [m]
                        for n, s_a in enumerate(zip(color['color'], image_links, sizes_asin)):
                            product_info['SKU'] = s_a[2][m]
                            product_info['Option 1 Value'] = s_a[0]

                            image_link_1, image_link_2, image_link_3, image_link_4, image_link_5, image_link_6, image_link_7, image_link_8, image_link_9, image_link_10, image_link_11, image_link_12 = None, None, None, None, None, None, None, None, None, None, None, None
                            for i, image in enumerate(s_a[1]):
                                if i == 0:
                                    image_link_1 = image
                                if i == 1:
                                    image_link_2 = image
                                if i == 2:
                                    image_link_3 = image
                                if i == 3:
                                    image_link_4 = image
                                if i == 4:
                                    image_link_5 = image
                                if i == 5:
                                    image_link_6 = image
                                if i == 6:
                                    image_link_7 = image
                                if i == 7:
                                    image_link_8 = image
                                if i == 8:
                                    image_link_9 = image
                                if i == 9:
                                    image_link_10 = image
                                if i == 10:
                                    image_link_11 = image
                                if i == 11:
                                    image_link_12 = image

                            product_info['Image_link_1'] = image_link_1
                            product_info['Image_link_2'] = image_link_2
                            product_info['Image_link_3'] = image_link_3
                            product_info['Image_link_4'] = image_link_4
                            product_info['Image_link_5'] = image_link_5
                            product_info['Image_link_6'] = image_link_6
                            product_info['Image_link_7'] = image_link_7
                            product_info['Image_link_8'] = image_link_8
                            product_info['Image_link_9'] = image_link_9
                            product_info['Image_link_10'] = image_link_10
                            product_info['Image_link_11'] = image_link_11
                            product_info['Image_link_12'] = image_link_12

                            writer.writerow(product_info)

            except Exception as e:
                # print(e)
                if number != 0:
                    print("Link Doesn't Work")
                pass
            number = number + 1

        self.driver.quit()
        print("Done")

    def file_path(self):
        file = filedialog.askopenfile(mode='r')
        if file:
            filepath = os.path.abspath(file.name)
        else:
            filepath = ""
            print("Error in Selection of File!! Select the file Again !!")
        return filepath

    def get_title(self):
        try:
            return self.driver.find_element("id", 'productTitle').text
        except Exception as e:
            # print("Didn't get the title !!")
            return None

    def get_asin(self, link):
        try:
            asin = str(link).split('dp/')[1].split('/')[0]
            asin = f"{asin}-1"
            return asin
        except Exception as e:
            # print("Didn't get the ASIN !!")
            return None

    def get_description(self):
        try:
            desc = self.driver.find_element("id", 'feature-bullets').text
            return desc
        except Exception as e:
            # print("Didn't get the Description !!")
            return None

    def get_price(self):
        try:
            price = self.driver.find_element("xpath",
                                             "//span[contains(@class, 'a-price') and contains(@class, 'aok-align-center') and contains(@class, ' reinventPricePriceToPayMargin') and contains(@class, 'priceToPay')]").text
            price = ".".join(price.split('\n'))
            return price
        except Exception as e:
            try:
                price = self.driver.find_element("xpath",
                                                 "//span[contains(@class, 'a-price') and contains(@class, 'a-text-price') and contains(@class, 'a-size-medium') and contains(@class, 'apexPriceToPay')]").text
                return price
            except:
                # print("You didn't Catch the Price")
                return None

    def get_color(self, link):
        asin_color = f"{str(link).split('dp/')[1].split('/')[0]}-1"
        try:
            temp = self.driver.find_element('id', 'variation_color_name')
            temp = temp.find_elements('tag name', 'li')
            color = [x.get_attribute('title').split(
                "Click to select ")[1] for x in temp]
            asin_color = []
            for x in temp:
                try:
                    asin_c = f"{x.get_attribute('data-defaultasin')}-1"
                    if asin_c == '-1':
                        asin_c = f"{x.get_attribute('data-dp-url').split('/')[2]}-1"
                    asin_color.append(asin_c)
                except:
                    asin_color.append("")
            colors = {
                "color": color,
                "asin_color": asin_color
            }
            return colors
        except:
            try:
                temp = self.driver.find_element("xpath",
                                                "//tr[contains(@class, 'a-spacing-small') and contains(@class, 'po-color')]").text
                color = temp.split('Color ')[1]
                colors = {
                    "color": [color],
                    "asin_color": [asin_color]
                }
                return colors
            except Exception as e:
                colors = {
                    "color": [" "],
                    "asin_color": [asin_color]
                }
                return colors

    def get_size(self, link):
        asin_size = f"{str(link).split('dp/')[1].split('/')[0]}-1"
        try:
            temp = self.driver.find_element('id', 'variation_size_name')
            temp_options = temp.find_elements('tag name', 'option')
            size = [x.text for x in temp_options[1:]]
            asin_size = [
                f"{x.get_attribute('value').split(',')[1]}-1" for x in temp_options[1:]]

            if size == []:
                size = {
                    "size": ["4.00\"D x 4.00\"W x 4.00\"H"],
                    "asin_size": [f"{str(link).split('dp/')[1].split('/')[0]}-1"]
                }
            else:
                size = {
                    "size": size,
                    "asin_size": asin_size
                }
            return size
        except Exception as e:
            try:
                temp = self.driver.find_element("id",
                                                'prodDetails').text
                size = temp.split('Item Package Dimensions L x W x H ')[
                    1].split('\n')[0]

                if size == []:
                    size = {
                        "size": ["4.00\"D x 4.00\"W x 4.00\"H"],
                        "asin_size": [f"{str(link).split('dp/')[1].split('/')[0]}-1"]
                    }
                else:
                    size = {
                        "size": [size],
                        "asin_size": [asin_size]
                    }
                return size
            except:
                try:
                    temp = self.driver.find_element("xpath",
                                                    "//tr[contains(@class, 'a-spacing-small') and contains(@class, 'po-item_depth_width_height')]").text
                    size = temp.split('Product Dimensions ')[1]

                    if size == []:
                        size = {
                            "size": ["4.00\"D x 4.00\"W x 4.00\"H"],
                            "asin_size": [f"{str(link).split('dp/')[1].split('/')[0]}-1"]
                        }
                    else:
                        size = {
                            "size": [size],
                            "asin_size": [asin_size]
                        }
                    return size
                except:
                    size = {
                        "size": ["4.00\"D x 4.00\"W x 4.00\"H"],
                        "asin_size": [asin_size]
                    }
                    # print("You didn't Catch the Size")
                    return size

    def get_feature(self):
        try:
            temp = self.driver.find_element("xpath",
                                            "//tr[contains(@class, 'a-spacing-small') and contains(@class, 'po-special_feature')]").text
            feature = temp.split('Special Feature')[1]
            return feature
        except Exception as e:
            return None

    def get_weight(self):
        try:
            temp = self.driver.find_element("id",
                                            'productDetails_techSpec_section_1').text
            weight = temp.split('Item Weight ')[1].split('\n')[0]

            return weight
        except Exception as e:
            weight = "1 Pound"
            # print("You didn't Catch the Weight")
            return weight

    def get_in_box_details(self):
        try:
            temp = self.driver.find_element("id",
                                            'whatsInTheBoxDeck').text
            details = ",".join(temp.split("\n")[1:])
            return details
        except Exception as e:
            try:
                temp = self.driver.find_element("id",
                                                'prodDetails').text
                details = temp.split('Included Components')[
                    1].split('\n')[0]
                return details
            except:
                return None

    def get_images_links(self, color):
        try:
            # For Scrolling Images
            image_links = []
            sizes = []
            asin_list = [x.split('-')[0] for x in color['asin_color']]
            for asin in asin_list[0:]:
                self.driver.get(f"https://www.amazon.com/dp/{asin}")
                try:
                    temp = self.driver.find_element(
                        "class name", 'unrolledScrollBox')
                    temp1 = temp.find_elements("tag name", 'img')
                    im_links = [x.get_attribute('src') for x in temp1]

                    img_links = []
                    for link in im_links:
                        tt = [x for i, x in enumerate(
                            link.split(".")) if i != len(link.split('.'))-2]
                        tt.insert(-1, '_AC_UL1500_')
                        ttt = '.'.join(tt)
                        if "jpg" in ttt:
                            img_links.append(ttt)
                    size = self.get_size(f"https://www.amazon.com/dp/{asin}")
                    sizes.append(size['asin_size'])
                    image_links.append(img_links)
                except:
                    samp = self.driver.find_element(
                        'id', 'imageBlock_feature_div')
                    samp = samp.find_elements(
                        'tag name', 'script')[2].get_attribute('innerHTML')
                    temp_dd = samp.split("\"")
                    im_links = []
                    for i, y in enumerate(temp_dd):
                        if y == "hiRes":
                            im_links.append(temp_dd[i+2])

                    img_links = []
                    for link in im_links:
                        tt = [x for i, x in enumerate(
                            link.split(".")) if i != len(link.split('.'))-2]
                        tt.insert(-1, '_AC_UL1500_')
                        ttt = '.'.join(tt)
                        if "jpg" in ttt:
                            img_links.append(ttt)

                    size = self.get_size(f"https://www.amazon.com/dp/{asin}")
                    sizes.append(size['asin_size'])
                    image_links.append(img_links)
                # print(img_links)
            return sizes, image_links
        except Exception as e:
            return None


if __name__ == "__main__":
    am = Amazon_API()
    am.GPU()
