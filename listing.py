from bs4 import BeautifulSoup
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.remote.errorhandler import ErrorCode
from selenium.webdriver.remote.utils import format_json
import xlwt
from xlwt import Workbook
from openpyxl import load_workbook
from xlrd import open_workbook
from xlutils.copy import copy as xl_copy
import pandas as pd

N = 52  # number of pagination
INTERST_RATE = 1.5
LISTING_URL = "https://www.propertyguru.com.sg/listing/{}"
# URL = 'https://www.propertyguru.com.sg/property-for-sale?district_code%5B0%5D=D25&floor_level%5B0%5D=HIGH&maxprice=900000&property_type=N&property_type_code%5B0%5D=APT&property_type_code%5B1%5D=CLUS&property_type_code%5B2%5D=CONDO&property_type_code%5B3%5D=EXCON&property_type_code%5B4%5D=WALK&unselected=DISTRICT%7CD25&freetext=D25%20Admiralty%20/%20Woodlands&sort=size&order=desc'
# URL = "https://www.propertyguru.com.sg/property-for-sale/{}?MRT_STATIONS%5B0%5D=CC1&MRT_STATIONS%5B1%5D=CC10&MRT_STATIONS%5B10%5D=CE2&MRT_STATIONS%5B11%5D=CP1&MRT_STATIONS%5B12%5D=CR11&MRT_STATIONS%5B13%5D=CR13&MRT_STATIONS%5B14%5D=CR5&MRT_STATIONS%5B15%5D=CR8&MRT_STATIONS%5B16%5D=DT10&MRT_STATIONS%5B17%5D=DT11&MRT_STATIONS%5B18%5D=DT16&MRT_STATIONS%5B19%5D=DT17&MRT_STATIONS%5B2%5D=CC17&MRT_STATIONS%5B20%5D=DT18&MRT_STATIONS%5B21%5D=DT19&MRT_STATIONS%5B22%5D=DT22&MRT_STATIONS%5B23%5D=DT26&MRT_STATIONS%5B24%5D=EW1&MRT_STATIONS%5B25%5D=EW14&MRT_STATIONS%5B26%5D=EW8&MRT_STATIONS%5B27%5D=NE14&MRT_STATIONS%5B28%5D=NE4&MRT_STATIONS%5B29%5D=NE6&MRT_STATIONS%5B3%5D=CC30&MRT_STATIONS%5B30%5D=NS11&MRT_STATIONS%5B31%5D=NS12&MRT_STATIONS%5B32%5D=NS16&MRT_STATIONS%5B33%5D=NS21&MRT_STATIONS%5B34%5D=NS22&MRT_STATIONS%5B35%5D=NS24&MRT_STATIONS%5B36%5D=NS26&MRT_STATIONS%5B37%5D=NS27&MRT_STATIONS%5B38%5D=NS9&MRT_STATIONS%5B39%5D=TE10&MRT_STATIONS%5B4%5D=CC31&MRT_STATIONS%5B40%5D=TE11&MRT_STATIONS%5B41%5D=TE12&MRT_STATIONS%5B42%5D=TE13&MRT_STATIONS%5B43%5D=TE14&MRT_STATIONS%5B44%5D=TE15&MRT_STATIONS%5B45%5D=TE18&MRT_STATIONS%5B46%5D=TE2&MRT_STATIONS%5B47%5D=TE20&MRT_STATIONS%5B48%5D=TE21&MRT_STATIONS%5B49%5D=TE22&MRT_STATIONS%5B5%5D=CC32&MRT_STATIONS%5B50%5D=TE23&MRT_STATIONS%5B51%5D=TE24&MRT_STATIONS%5B52%5D=TE25&MRT_STATIONS%5B53%5D=TE26&MRT_STATIONS%5B54%5D=TE27&MRT_STATIONS%5B55%5D=TE3&MRT_STATIONS%5B56%5D=TE4&MRT_STATIONS%5B57%5D=TE5&MRT_STATIONS%5B58%5D=TE6&MRT_STATIONS%5B59%5D=TE7&MRT_STATIONS%5B6%5D=CC7&MRT_STATIONS%5B60%5D=TE9&MRT_STATIONS%5B7%5D=CC8&MRT_STATIONS%5B8%5D=CC9&MRT_STATIONS%5B9%5D=CE1&_freetextDisplay=TE12+Napier+MRT%2C+TE15+Great+World+MRT%2C+TE18+Maxwell+MRT%2C+CC31+Cantonment+MRT%2C+CC30+Keppel+MRT%2C+CC32+Prince+Edward+Road+MRT%2C+TE21+Marina+South+MRT%2C+TE22+Gardens+by+the+Bay+MRT%2C+TE23+Tanjong+Rhu+MRT%2C+TE24+Katong+Park+MRT%2C+TE25+Tanjong+Katong+MRT%2C+TE27+Marine+Terrace+MRT%2C+TE26+Marine+Parade+MRT%2C+CC7+Mountbatten+MRT%2C+CC8+Dakota+MRT%2C+NS12+Canberra+MRT%2C+NS11+Sembawang+MRT%2C+NS9%2FTE2+Woodlands+MRT%2C+TE3+Woodlands+South+MRT%2C+TE4+Springleaf+MRT%2C+TE6+Mayflower+MRT%2C+TE5+Lentor+MRT%2C+CR13%2FTE7+Bright+Hill+MRT%2C+CC17%2FTE9+Caldecott+MRT%2C+TE10+Mount+Pleasant+MRT%2C+DT10%2FTE11+Stevens+MRT%2C+TE13+Orchard+Boulevard+MRT%2C+NS22%2FTE14+Orchard+MRT%2C+CC1%2FNS24%2FNE6+Dhoby+Ghaut+MRT%2C+DT19%2FNE4+Chinatown+MRT%2C+DT18+Telok+Ayer+MRT%2C+EW14%2FNS26+Raffles+Place+MRT%2C+CE1%2FDT16+Bayfront+MRT%2C+DT17+Downtown+MRT%2C+CE2%2FNS27%2FTE20+Marina+Bay+MRT%2C+CC9%2FEW8+Paya+Lebar+MRT%2C+DT22+Jalan+Besar+MRT%2C+DT11%2FNS21+Newton+MRT%2C+CC10%2FDT26+MacPherson+MRT%2C+CR8%2FNE14+Hougang+MRT%2C+CP1%2FEW1%2FCR5+Pasir+Ris+MRT%2C+CR5%2C+CR11%2FNS16+Ang+Mo+Kio+MRT&freetext=TE12+Napier+MRT%2C+TE15+Great+World+MRT%2C+TE18+Maxwell+MRT%2C+CC31+Cantonment+MRT%2C+CC30+Keppel+MRT%2C+CC32+Prince+Edward+Road+MRT%2C+TE21+Marina+South+MRT%2C+TE22+Gardens+by+the+Bay+MRT%2C+TE23+Tanjong+Rhu+MRT%2C+TE24+Katong+Park+MRT%2C+TE25+Tanjong+Katong+MRT%2C+TE27+Marine+Terrace+MRT%2C+TE26+Marine+Parade+MRT%2C+CC7+Mountbatten+MRT%2C+CC8+Dakota+MRT%2C+NS12+Canberra+MRT%2C+NS11+Sembawang+MRT%2C+NS9%2FTE2+Woodlands+MRT%2C+TE3+Woodlands+South+MRT%2C+TE4+Springleaf+MRT%2C+TE6+Mayflower+MRT%2C+TE5+Lentor+MRT%2C+CR13%2FTE7+Bright+Hill+MRT%2C+CC17%2FTE9+Caldecott+MRT%2C+TE10+Mount+Pleasant+MRT%2C+DT10%2FTE11+Stevens+MRT%2C+TE13+Orchard+Boulevard+MRT%2C+NS22%2FTE14+Orchard+MRT%2C+CC1%2FNS24%2FNE6+Dhoby+Ghaut+MRT%2C+DT19%2FNE4+Chinatown+MRT%2C+DT18+Telok+Ayer+MRT%2C+EW14%2FNS26+Raffles+Place+MRT%2C+CE1%2FDT16+Bayfront+MRT%2C+DT17+Downtown+MRT%2C+CE2%2FNS27%2FTE20+Marina+Bay+MRT%2C+CC9%2FEW8+Paya+Lebar+MRT%2C+DT22+Jalan+Besar+MRT%2C+DT11%2FNS21+Newton+MRT%2C+CC10%2FDT26+MacPherson+MRT%2C+CR8%2FNE14+Hougang+MRT%2C+CP1%2FEW1%2FCR5+Pasir+Ris+MRT%2C+CR5%2C+CR11%2FNS16+Ang+Mo+Kio+MRT&maxprice=900000&mintop=2001&mrt_stations=CC1%2CCC10%2CCC17%2CCC30%2CCC31%2CCC32%2CCC7%2CCC8%2CCC9%2CCE1%2CCE2%2CCP1%2CCR11%2CCR13%2CCR5%2CCR8%2CDT10%2CDT11%2CDT16%2CDT17%2CDT18%2CDT19%2CDT22%2CDT26%2CEW1%2CEW14%2CEW8%2CNE14%2CNE4%2CNE6%2CNS11%2CNS12%2CNS16%2CNS21%2CNS22%2CNS24%2CNS26%2CNS27%2CNS9%2CTE10%2CTE11%2CTE12%2CTE13%2CTE14%2CTE15%2CTE18%2CTE2%2CTE20%2CTE21%2CTE22%2CTE23%2CTE24%2CTE25%2CTE26%2CTE27%2CTE3%2CTE4%2CTE5%2CTE6%2CTE7%2CTE9&order=desc&property_type=N&property_type_code%5B0%5D=APT&property_type_code%5B1%5D=CLUS&property_type_code%5B2%5D=CONDO&property_type_code%5B3%5D=EXCON&property_type_code%5B4%5D=WALK&sort=size"
# URL = 'https://www.propertyguru.com.sg/property-for-sale?market=residential&sort=size&order=desc&district_code%5B%5D=D25&freetext=D25+Admiralty+%2F+Woodlands&newProject=all&property_type=N&property_type_code%5B%5D=CONDO&property_type_code%5B%5D=APT&property_type_code%5B%5D=WALK&property_type_code%5B%5D=CLUS&property_type_code%5B%5D=EXCON&maxprice=900000&beds%5B%5D=1&floor_level%5B%5D=HIGH'
URL = 'https://www.propertyguru.com.sg/property-for-sale/{}?market=residential&sort=price&order=asc&freetext=BP6%2FDT1+Bukit+Panjang+LRT%2C+CC10%2FDT26+MacPherson+MRT%2C+CC13%2FNE12+Serangoon+MRT%2C+CC17%2FTE9+Caldecott+MRT%2C+CC19%2FDT9+Botanic+Gardens+MRT%2C+CG1%2FDT35+Expo+MRT%2C+CP1%2FEW1%2FCR5+Pasir+Ris+MRT%2C+CP2+Elias+MRT%2C+CP3%2FPE4+Riviera+MRT%2C+CP4%2FNE17%2FPTC+Punggol+MRT%2C+CR13%2FTE7+Bright+Hill+MRT%2C+CR3+Loyang+MRT%2C+CR4+Pasir+Ris+East+MRT%2C+CR5%2C+CR6+Tampines+North+MRT%2C+DT1%2FBP6+Bukit+Panjang+MRT%2C+DT10%2FTE11+Stevens+MRT%2C+DT11%2FNS21+Newton+MRT%2C+DT14%2FEW12+Bugis+MRT%2C+DT2+Cashew+MRT%2C+DT29+Bedok+North+MRT%2C+DT3+Hillview+MRT%2C+DT30+Bedok+Reservoir+MRT%2C+DT31+Tampines+West+MRT%2C+DT32%2FEW2+Tampines+MRT%2C+DT33+Tampines+East+MRT%2C+EW3+Simei+MRT%2C+EW4%2FCG+Tanah+Merah+MRT%2C+EW5+Bedok+MRT%2C+NE16%2FSTC+Sengkang+MRT%2C+NE18+Punggol+Coast+MRT%2C+NS11+Sembawang+MRT%2C+NS12+Canberra+MRT%2C+NS13+Yishun+MRT%2C+NS15+Yio+Chu+Kang+MRT%2C+NS22%2FTE14+Orchard+MRT%2C+NS9%2FTE2+Woodlands+MRT%2C+PE4%2FCP3+Riviera+LRT%2C+PTC%2FNE17%2FCP4+Punggol+LRT%2C+STC%2FNE16+Sengkang+LRT%2C+TE23+Tanjong+Rhu+MRT%2C+TE24+Katong+Park+MRT%2C+TE25+Tanjong+Katong+MRT%2C+TE27+Marine+Terrace+MRT%2C+TE28+Siglap+MRT%2C+TE3+Woodlands+South+MRT%2C+TE4+Springleaf+MRT%2C+TE5+Lentor+MRT%2C+TE6+Mayflower+MRT%2C+TE8+Upper+Thomson+MRT%2C+NS10+Admiralty+MRT%2C+DT12%2FNE7+Little+India+MRT%2C+NS20+Novena+MRT%2C+NS19+Toa+Payoh+MRT&property_type=N&property_type_code%5B%5D=CONDO&property_type_code%5B%5D=APT&property_type_code%5B%5D=WALK&property_type_code%5B%5D=CLUS&property_type_code%5B%5D=EXCON&maxprice=900001&beds%5B%5D=1&beds%5B%5D=2&beds%5B%5D=3&beds%5B%5D=4&beds%5B%5D=5&newProject=all&mintop=2000&maxtop=2020&floor_level%5B%5D=LOW&floor_level%5B%5D=MID&floor_level%5B%5D=HIGH&MRT_STATIONS%5B%5D=BP6&MRT_STATIONS%5B%5D=CC10&MRT_STATIONS%5B%5D=CC13&MRT_STATIONS%5B%5D=CC17&MRT_STATIONS%5B%5D=CC19&MRT_STATIONS%5B%5D=CG&MRT_STATIONS%5B%5D=CP1&MRT_STATIONS%5B%5D=CP2&MRT_STATIONS%5B%5D=CP3&MRT_STATIONS%5B%5D=CP4&MRT_STATIONS%5B%5D=CR13&MRT_STATIONS%5B%5D=CR3&MRT_STATIONS%5B%5D=CR4&MRT_STATIONS%5B%5D=CR5&MRT_STATIONS%5B%5D=CR6&MRT_STATIONS%5B%5D=DT1&MRT_STATIONS%5B%5D=DT10&MRT_STATIONS%5B%5D=DT11&MRT_STATIONS%5B%5D=DT14&MRT_STATIONS%5B%5D=DT2&MRT_STATIONS%5B%5D=DT26&MRT_STATIONS%5B%5D=DT29&MRT_STATIONS%5B%5D=DT3&MRT_STATIONS%5B%5D=DT30&MRT_STATIONS%5B%5D=DT31&MRT_STATIONS%5B%5D=DT32&MRT_STATIONS%5B%5D=DT33&MRT_STATIONS%5B%5D=DT9&MRT_STATIONS%5B%5D=EW1&MRT_STATIONS%5B%5D=EW12&MRT_STATIONS%5B%5D=EW2&MRT_STATIONS%5B%5D=EW3&MRT_STATIONS%5B%5D=EW4&MRT_STATIONS%5B%5D=EW5&MRT_STATIONS%5B%5D=NE12&MRT_STATIONS%5B%5D=NE16&MRT_STATIONS%5B%5D=NE17&MRT_STATIONS%5B%5D=NE18&MRT_STATIONS%5B%5D=NS11&MRT_STATIONS%5B%5D=NS12&MRT_STATIONS%5B%5D=NS13&MRT_STATIONS%5B%5D=NS15&MRT_STATIONS%5B%5D=NS21&MRT_STATIONS%5B%5D=NS22&MRT_STATIONS%5B%5D=NS9&MRT_STATIONS%5B%5D=PE4&MRT_STATIONS%5B%5D=PTC&MRT_STATIONS%5B%5D=STC&MRT_STATIONS%5B%5D=TE11&MRT_STATIONS%5B%5D=TE14&MRT_STATIONS%5B%5D=TE2&MRT_STATIONS%5B%5D=TE23&MRT_STATIONS%5B%5D=TE24&MRT_STATIONS%5B%5D=TE25&MRT_STATIONS%5B%5D=TE27&MRT_STATIONS%5B%5D=TE28&MRT_STATIONS%5B%5D=TE3&MRT_STATIONS%5B%5D=TE4&MRT_STATIONS%5B%5D=TE5&MRT_STATIONS%5B%5D=TE6&MRT_STATIONS%5B%5D=TE7&MRT_STATIONS%5B%5D=TE8&MRT_STATIONS%5B%5D=TE9&MRT_STATIONS%5B%5D=NS10&MRT_STATIONS%5B%5D=DT12&MRT_STATIONS%5B%5D=NE7&MRT_STATIONS%5B%5D=NS20&MRT_STATIONS%5B%5D=NS19'
HEADER = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
GOOGLE_MAPS_URL = 'https://www.google.com/maps/search/{}'
STYLE_ONE = xlwt.easyxf('pattern: pattern solid, fore_colour dark_blue;''font: colour white, bold True; align: horiz center;')
STYLE_TWO = xlwt.easyxf('font: colour black, bold False; align: horiz center;')


def get_driver(url, header=HEADER):
    options = Options()
    options.add_argument('--user-agent={}'.format(header))
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(executable_path="./chromedriver", options=options)
    driver.implicitly_wait(3)
    time.sleep(3)
    isRunning = 0
    try:
        driver.set_page_load_timeout(30)
        driver.get(url)
        isRunning = 1
    except:
        print("TimeOut Exception has been thrown: {}".format(url))
        driver.quit()
        driver = None
    return isRunning, driver


def format_float_value(val, index=1):
    return round(float(val.split(' ')[index].replace(',', '')), 2)


def get_avg_rental(lst):
    counter = 0
    total = 0
    for val in lst:
        if val:
            counter += 1
            total += format_float_value(val)
    return round(total / counter, 2)


def get_details_from_listing_page(data):
    isRunning, driver = get_driver(data['Website'])
    if not isRunning:
        return data

    # get name
    data['Name'] = driver.find_element_by_css_selector('div.property-info-element.location-info').find_element_by_class_name('listing-title').text

    #get price
    data['Price'] = driver.find_element_by_css_selector('div.price-overview-row.listing-detail-summary-bar-price[itemprop="offers"]').find_element_by_css_selector('span.element-label.price[itemprop="price"]').text

    # get price type tag
    try:
        data['Price Type Tag'] = driver.find_element_by_css_selector('div.price-overview-row.listing-detail-summary-bar-price[itemprop="offers"]').find_element_by_css_selector('div.element-label.price-type-tag').text
    except:
        data['Price Type Tag'] = "NA"

    data['Bed'] = int(driver.find_element_by_css_selector('div.property-info-element.beds').find_element_by_css_selector('span.element-label[itemprop="numberOfRooms"]').text)
    data['Bath'] = int(driver.find_element_by_css_selector('div.property-info-element.baths').find_element_by_css_selector('span.element-label').text)
    data['SQFT'] = format_float_value(driver.find_element_by_css_selector('div.property-info-element.area').find_element_by_css_selector('span.element-label').text, 0)
    data['PSF ($)'] = format_float_value(driver.find_element_by_css_selector('div.property-info-element.psf').find_element_by_css_selector('span.price-value').text, 0)

    # get address information
    data['Address'] = driver.find_element_by_css_selector('div.listing-address[itemprop="address"]').text
    data['Google Map Link'] = GOOGLE_MAPS_URL.format('+'.join(data['Address'].split(' ')))
    data['Area Code'] = data['Address'].split('(')[-1][:-1]

    # get stations near it
    try:
        data['Nearby Stations'] = ' / '.join(list(map(lambda x: x.text, driver.find_element_by_class_name('price-overview-nearby-poi').find_elements_by_tag_name('div')[1:])))
    except:
        data['Nearby Stations'] = ' / '.join([])

    details_section = driver.find_element_by_class_name('listing-details-primary')
    details_list = ['Furnishing', 'TOP', 'Floor Level', 'Currently Tenanted', 'Maintenance Fee']
    try:
        divs = details_section.find_elements_by_class_name('property-attr')
        label_found = []
        for div in divs:
            key = div.find_element_by_class_name('label-block').text
            value = div.find_element_by_class_name('value-block').text
            if key in details_list:
                data[key] = value
                label_found.append(key)
        for key in details_list:
            if key not in label_found:
                data[key] = 'NA'
    except:
        for key in details_list:
            data[key] = 'NA'

    house_features = []
    try:
        house_features = list(map(lambda x: x.text, driver.find_element_by_id('facilities').find_elements_by_class_name('expansible')[0].find_element_by_tag_name('ul').find_elements_by_tag_name('span')))
    except:
        house_features = []

    data['House features'] = ' / '.join(house_features)

    home_finance = driver.find_element_by_class_name('home-finance-widget__container')
    driver.execute_script('arguments[0].scrollIntoView(true);', home_finance)
    data['Price'] = format_float_value(home_finance.find_element_by_css_selector('input.form-control[name="propertyPrice"]').get_attribute('value'), 0)
    data['Loan Amount'] = format_float_value(home_finance.find_element_by_css_selector('input.form-control[name="loanAmount"]').get_attribute('value'), 0)
    data['Upfront Cost'] = round(data['Price'] - data['Loan Amount'], 2)
    home_finance.find_element_by_css_selector('input.form-control[name="interestRate"]').clear()
    time.sleep(2)
    home_finance.find_element_by_css_selector('input.form-control[name="interestRate"]').send_keys(str(INTERST_RATE))
    data['Interst Rate Applied'] = '{} %'.format(home_finance.find_element_by_css_selector('input.form-control[name="interestRate"]').get_attribute('value'))
    data['Monthly Mortgage (2%)'] = format_float_value(home_finance.find_element_by_class_name('home-finance-chart__card-context').find_elements_by_tag_name('div')[-1].text)
    home_finance.find_element_by_css_selector('button.btn.btn-default.home-finance-calc__btn').click()
    time.sleep(1)

    data['Monthly Mortgage'] = format_float_value(home_finance.find_element_by_class_name('home-finance-chart__card-context').find_elements_by_tag_name('div')[-1].text)
    data['Monthly Principal'] = format_float_value(home_finance.find_element_by_class_name('home-finance-chart__legend-container').find_element_by_class_name('home-finance-chart__legend').text)
    data['Monthly Interest'] = format_float_value(home_finance.find_elements_by_class_name('home-finance-chart__legend-container')[1].find_element_by_class_name('home-finance-chart__legend').text)

    price_insight = driver.find_element_by_id('price-insights-widget')
    driver.execute_script('arguments[0].scrollIntoView(true);', price_insight)
    price_insight.find_element_by_css_selector('button.btn.btn-default-outline[name="rent"]').click()
    time.sleep(1)

    try:
        last_transctions = list(map(lambda x: x.text, driver.find_element_by_id('priceInsightLastTransactionTab').find_element_by_class_name('price_widget_transaction_tab_rent__body').find_element_by_class_name('price_widget_transaction_tab__body_right').find_element_by_class_name('active').find_elements_by_class_name('price_widget_transaction_tab__price')))
        data['Avg Rental'] = get_avg_rental(last_transctions[: 10])  # based on recent 10 data
        data['Rental Yield'] = round((data['Avg Rental'] * 12) / data['Price'], 2)
    except:
        data['Avg Rental'] = 0.0
        data['Rental Yield'] = 0

    data['Cashflow'] = data['Avg Rental'] - data['Monthly Mortgage']
    data['Profitable'] = True if data['Cashflow'] > 0 else False

    driver.quit()

    return data


def get_listing_information(wb, URL, sheet_name, data):
    if sheet_exists(wb, sheet_name):
        return

    isRunning, driver = get_driver(URL)
    if not isRunning:
        return data

    time.sleep(1)
    listing_div = driver.find_element_by_id('listings-container')

    for i in listing_div.find_elements_by_xpath('./div'):

        if 'listing-card' in i.get_attribute('class'):
            listing_id = i.get_attribute('data-listing-id')
            dic = {}
            dic['Listing Id'] = listing_id
            dic['Website'] = LISTING_URL.format(listing_id)

            data[listing_id] = get_details_from_listing_page(dic)

    driver.quit()
    write_to_excel(wb, sheet_name, data)


def sheet_exists(wb, sheet_name):
    try:
        sheet = wb.get_sheet(sheet_name)
        return True
    except:
        return False


def add_or_get_sheet_by_name(wb, sheet_name, data):
    try:
        return wb.get_sheet(sheet_name)
    except:
        ws = wb.add_sheet(sheet_name)
        for listing_data in data.values():
            for id, key in enumerate(listing_data.keys()):
                ws.write(0, id, key, STYLE_ONE)
            break
        return ws


def write_to_excel(wb, sheet_name, data):
    ws = add_or_get_sheet_by_name(wb, sheet_name, data)

    row_counter = 1

    for listing_data in data.values():
        col_counter = 0
        for value in listing_data.values():
            ws.write(row_counter, col_counter, value, STYLE_TWO)
            col_counter += 1

        row_counter += 1

    wb.save("Listings.xls")


try:
    rb = open_workbook('Listings.xls', formatting_info=True)
    wb = xl_copy(rb)
except:
    rb = None
    wb = Workbook(encoding='utf-8')

thread_list = list()
for i in range(1, N + 1):  # 1 , N+1
    t = threading.Thread(name='Test {}'.format(i), target=get_listing_information(wb, URL.format(i), str(i), {}))
    t.start()
    thread_list.append(t)

for thread in thread_list:
    thread.join()

wb.save("Listings.xls")

# d = pd.ExcelFile("Listings.xls")
# all_dfs = pd.read_excel('Listings.xls', sheet_name=None)
# df = pd.concat(all_dfs, ignore_index=True)

# df.to_csv('Listings.csv', encoding='utf-8', index=False)
# rows_affected = df.loc[df['Interst Rate Applied'] == '21.5 %']
# print(len(rows_affected))
# print(rows_affected)
# print(listing_ids_affected[99])
# cols_to_update = ['Interst Rate Applied', 'Monthly Mortgage (2%)', 'Monthly Mortgage', 'Monthly Principal', 'Monthly Interest', 'Cashflow', 'Profitable']
# for index in rows_affected.index:
#     website = df.iloc[index]['Website']

#     isRunning, driver = get_driver(website)

#     vals_to_update = []

#     #get price
#     home_finance = driver.find_element_by_class_name('home-finance-widget__container')
#     driver.execute_script('arguments[0].scrollIntoView(true);', home_finance)

#     home_finance.find_element_by_css_selector('input.form-control[name="interestRate"]').clear()
#     time.sleep(3)
#     home_finance.find_element_by_css_selector('input.form-control[name="interestRate"]').send_keys(str(INTERST_RATE))
#     interst_rate_applied = '{} %'.format(home_finance.find_element_by_css_selector('input.form-control[name="interestRate"]').get_attribute('value'))
#     vals_to_update.append(interst_rate_applied)
#     old_monthly_mortgage = format_float_value(home_finance.find_element_by_class_name('home-finance-chart__card-context').find_elements_by_tag_name('div')[-1].text)
#     vals_to_update.append(old_monthly_mortgage)
#     home_finance.find_element_by_css_selector('button.btn.btn-default.home-finance-calc__btn').click()
#     time.sleep(1)

#     monthly_mortgage = format_float_value(home_finance.find_element_by_class_name('home-finance-chart__card-context').find_elements_by_tag_name('div')[-1].text)
#     vals_to_update.append(monthly_mortgage)
#     monthly_principle = format_float_value(home_finance.find_element_by_class_name('home-finance-chart__legend-container').find_element_by_class_name('home-finance-chart__legend').text)
#     vals_to_update.append(monthly_principle)
#     monthly_interest = format_float_value(home_finance.find_elements_by_class_name('home-finance-chart__legend-container')[1].find_element_by_class_name('home-finance-chart__legend').text)
#     vals_to_update.append(monthly_interest)
#     avg_rental = df.iloc[index]['Avg Rental']
#     cashflow = avg_rental - monthly_mortgage
#     vals_to_update.append(cashflow)
#     profitable = True if cashflow > 0 else False
#     vals_to_update.append(profitable)

#     driver.quit()

#     for i in range(len(cols_to_update)):
#         df.at[index, cols_to_update[i]] = vals_to_update[i]

# df.to_csv('Listings.csv', encoding='utf-8', index=False)
