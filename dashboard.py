import logging, datetime, traceback, os, time, re, pdftotext


try:
    application_path = os.path.dirname(os.path.abspath(__file__))

    import bs4, openpyxl
    from pyvirtualdisplay import Display
    from selenium import webdriver
    from selenium.webdriver.support.ui import WebDriverWait as WDW
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.expected_conditions import presence_of_element_located as pres
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import Select
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains

    startTime = datetime.datetime.now()

    def getTableData(columns):
        WDW(browser, 20).until(pres((By.LINK_TEXT, 'Print this Report'))).click()
        browser.switch_to.window(browser.window_handles[1])  # switch to the window with report
        soup = bs4.BeautifulSoup(browser.page_source, 'html.parser')
        tableData = []
        for i, row in enumerate(soup.select('tr')):
            cells = row.select('td')
            if not cells:
                continue
            cellsData = []
            for x, cell in enumerate(cells):
                # get only data from certain column indexes
                if x not in columns:
                    continue
                cellsData.append(cell.text.strip())
            if cellsData:
                tableData.append(cellsData)
        browser.close()
        browser.switch_to.window(browser.window_handles[0])  # switch back to the main window
        return tableData


    UPSlogin, UPSpassword, FedExlogin, FedExpassword, CPlogin, CPpassword = open('config.txt').read().strip().split(' ')

    disp = Display()  # virtual display
    disp.start()

    options = Options()
    driver_path = os.path.join(application_path, 'chromedriver.exe')
    options.add_argument("--start-normal")
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('prefs',  {
        "download.default_directory": application_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
        }
    )
    browser = webdriver.Chrome(options=options, executable_path=driver_path)
    browser.set_window_size(1800, 1000)

    # UPS part
    # log in the UPS billing center
    browser.get('https://www.apps.ups.com/ebilling/')
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#email"))).send_keys(UPSlogin)
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#pwd"))).send_keys(UPSpassword)
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#submitBtn"))).click()

    # get the list of all invoice numbers from Invoice Summary page from the whole time
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#navigationMenu > li:nth-child(5) > a"))).click()
    WDW(browser, 20).until(pres((By.LINK_TEXT, "Invoice Summary"))).click()
    WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromStatementDate"]'))).clear()
    WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromStatementDate"]'))).send_keys('01/01/2018')
    WDW(browser, 20).until(pres((By.XPATH, '//input[@name="toStatementDate"]'))).clear()
    WDW(browser, 20).until(pres((By.XPATH, '//input[@name="toStatementDate"]'))).send_keys(datetime.datetime.now().strftime('%m/%d/%Y'))
    WDW(browser, 20).until(pres((By.XPATH, '//input[@title="Search" and @type="submit"]'))).click()

    UPSData = []
    data = eval(open('UPS1.txt').read())
    cutOffDate = datetime.datetime(2021, 2, 1)
    for invoice in data:
        if invoice['invoice date'] < cutOffDate:
            UPSData.append(invoice)
    i = 1
    while True:
        lines = WDW(browser, 20).until(pres((By.XPATH, '//*[@id="ez"]/div[5]/form/div/div[5]/table/tbody'))).find_elements_by_tag_name('tr')[:-1]  # without total
        for l, line in enumerate(lines):
            if datetime.datetime.strptime(line.find_elements_by_tag_name('td')[2].text, '%m/%d/%Y') < cutOffDate:
                continue

            details = {
                'invoice number': line.find_elements_by_tag_name('td')[1].text,
                'invoice date': datetime.datetime.strptime(line.find_elements_by_tag_name('td')[2].text, '%m/%d/%Y')
            }
            # open in a new tab
            elem = line.find_elements_by_tag_name('td')[1]
            ActionChains(browser).move_to_element(elem).key_down(Keys.CONTROL).click(elem).key_up(Keys.CONTROL).perform()
            browser.switch_to.window(browser.window_handles[1])
            time.sleep(1)
            # get the data from the table
            soup = bs4.BeautifulSoup(browser.page_source, 'html.parser')
            tableLines = soup.select('table[id="Table1"] > tbody > tr')
            for row in tableLines[1:]:
                cells = row.select('td')
                if cells[1].text.strip() == '':
                    continue
                details[cells[0].text.strip()] = cells[1].text.strip()
            UPSData.append(details)
            print()
            # close the tab and switch to the first tab
            browser.close()
            browser.switch_to.window(browser.window_handles[0])
            print(f'Processing line {str(l)} out of {str(len(lines))}, page {str(i)}')
            time.sleep(1)

        if len(browser.find_elements_by_xpath('''//input[@onclick="document.tr1_pagingForm1.pagingAction.value='next'"]''')) == 0:
            break

        WDW(browser, 20).until(pres((By.XPATH, '//input[@title="Search" and @type="submit"]'))).click()
        for x in range(i):
            WDW(browser, 20).until(
                pres((By.XPATH, '''//input[@onclick="document.tr1_pagingForm1.pagingAction.value='next'"]'''))).click()
            time.sleep(0.5)
        i += 1

    # loop each invoice number on the Invoice Details page
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#navigationMenu > li:nth-child(5) > a"))).click()
    WDW(browser, 20).until(pres((By.LINK_TEXT, "Invoice Detail"))).click()
    for a, invoice in enumerate(UPSData):
        if invoice['invoice date'] < cutOffDate:
            continue
        print(f'Processing invoice {str(a)} out of {str(len(UPSData))}')
        WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromOtherDate"]'))).clear()
        WDW(browser, 20).until(pres((By.XPATH, '//input[@name="toOtherDate"]'))).clear()
        Select(browser.find_element_by_id('otherSelectionPickup')).select_by_visible_text('Select one')
        WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromStatementDate"]'))).clear()
        WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromStatementDate"]'))).send_keys('01/01/2018')
        WDW(browser, 20).until(pres((By.XPATH, '//input[@name="invNumber"]'))).clear()
        WDW(browser, 20).until(pres((By.XPATH, '//input[@name="invNumber"]'))).send_keys(invoice['invoice number'])
        WDW(browser, 20).until(pres((By.XPATH, '//input[@title="Search" and @type="submit"]'))).click()

        data = getTableData([6, 7, 8, 9, 12, 13, 14])
        shipments = data

        over100results = 'of 100' in browser.page_source
        if over100results:
            # if search by invoice number returns 100 shipments, get the returned shipments loop every 2 days interval of pickup date
            # get the start date for the loop
            startSearchDate = datetime.datetime.strptime(data[0][0], '%m/%d/%Y')
            endSearchDate = datetime.datetime.strptime(data[-1][0], '%m/%d/%Y')
            for item in data:
                if datetime.datetime.strptime(item[0], '%m/%d/%Y') < startSearchDate:
                    startSearchDate = datetime.datetime.strptime(item[0], '%m/%d/%Y')
            Select(browser.find_element_by_id('otherSelectionPickup')).select_by_visible_text('Pickup Date')
            # get the fist shipment for this invoice: up to the first known date
            WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromOtherDate"]'))).clear()
            WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromOtherDate"]'))).send_keys('01/01/2018')
            WDW(browser, 20).until(pres((By.XPATH, '//input[@name="toOtherDate"]'))).clear()
            WDW(browser, 20).until(pres((By.XPATH, '//input[@name="toOtherDate"]'))).send_keys(startSearchDate.strftime('%m/%d/%Y'))
            WDW(browser, 20).until(pres((By.XPATH, '//input[@title="Search" and @type="submit"]'))).click()
            shipments += getTableData([6, 7, 8, 9, 12, 13, 14])
            i = 0
            while True:
                fromDate = (startSearchDate + datetime.timedelta(days=i)).strftime('%m/%d/%Y')
                if startSearchDate + datetime.timedelta(days=i+1) > datetime.datetime.now():
                    toDate = datetime.datetime.now().strftime('%m/%d/%Y')
                # the next if clause looks really weird, but this is workaround for the bug in UPS system
                if datetime.datetime(2021, 1, 31, 0, 0) < startSearchDate + datetime.timedelta(days=i+1) < datetime.datetime(2021, 2, 3, 0, 0):
                    toDate = '02/03/2021'
                    i += 2
                else:
                    toDate = (startSearchDate + datetime.timedelta(days=i+1)).strftime('%m/%d/%Y')
                WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromOtherDate"]'))).clear()
                WDW(browser, 20).until(pres((By.XPATH, '//input[@name="fromOtherDate"]'))).send_keys(fromDate)
                WDW(browser, 20).until(pres((By.XPATH, '//input[@name="toOtherDate"]'))).clear()
                WDW(browser, 20).until(pres((By.XPATH, '//input[@name="toOtherDate"]'))).send_keys(toDate)
                WDW(browser, 20).until(pres((By.XPATH, '//input[@title="Search" and @type="submit"]'))).click()
                time.sleep(0.5)
                if 'No Information to Display' in browser.page_source and (startSearchDate + datetime.timedelta(days=i+1)) > endSearchDate:
                    break
                if 'No Information to Display' not in browser.page_source:
                    shipments += getTableData([6, 7, 8, 9, 12, 13, 14])
                if (startSearchDate + datetime.timedelta(days=i+1)) > datetime.datetime.now():
                    break
                i += 2


        # deduplicate shipments and remove special characters in numbers
        invoice['shipments'] = []
        for item in shipments:
            item[4] = item[4].strip('()$')
            item[5] = item[5].strip('()$')
            item[6] = item[6].strip('()$')
            if item not in invoice['shipments']:
                invoice['shipments'].append(item)

        file = open('UPS1.txt', 'w')
        file.write(str(UPSData))
        file.close()

    # FedEx part
    FedExData = []

    browser.get('https://www.fedex.com/fedexbillingonline/pages/accountsummary/accountSummaryFBO.xhtml')
    i = 0
    while True:
        time.sleep(2)
        i += 1
        if i == 10:
            logging.error('cant load fedex website after 10 attempts')
            break
        try:
            WDW(browser, 20).until(pres((By.XPATH, '//input[@name="username"]')))
            break
        except:
            print('cant load fedex website')
            browser.refresh()
    WDW(browser, 20).until(pres((By.XPATH, '//input[@name="username"]'))).send_keys(FedExlogin)
    WDW(browser, 20).until(pres((By.XPATH, '//input[@name="password"]'))).send_keys(FedExpassword)
    WDW(browser, 20).until(pres((By.XPATH, '//input[@name="login"]'))).click()
    time.sleep(3)

    # get information per invoice first
    WDW(browser, 20).until(pres((By.LINK_TEXT, 'Search/Download'))).click()
    WDW(browser, 20).until(pres((By.XPATH, '//*[@id="mainContentId:fromDate"]'))).clear()
    WDW(browser, 20).until(pres((By.XPATH, '//*[@id="mainContentId:fromDate"]'))).send_keys('04/01/2021')
    Select(WDW(browser, 20).until(pres((By.CSS_SELECTOR, '#mainContentId\:advTempl')))).select_by_visible_text('Invoices')
    time.sleep(2)
    WDW(browser, 20).until(pres((By.XPATH, '//*[@id="mainContentId:advancedSearchwldDataNewSrch"]'))).click()
    time.sleep(2)
    Select(WDW(browser, 20).until(pres((By.CSS_SELECTOR, '#mainContentId\:invAdvSrchRsltpaginator')))).select_by_visible_text('50')
    time.sleep(2)

    # open each invoice and save the details from it
    numOfTableLines = len(WDW(browser, 20).until(pres((By.XPATH, '//*[@id="mainContentId:invAdvSrchRsltTable:tbody"]'))).find_elements_by_tag_name('tr'))
    for x in range(numOfTableLines):
        print(f'Processing invoice {str(x)} out of {str(numOfTableLines)}')
        tableLine = WDW(browser, 20).until(pres((By.XPATH, '//*[@id="mainContentId:invAdvSrchRsltTable:tbody"]'))).find_elements_by_tag_name('tr')[x]
        if tableLine.text == '':
            continue
        # save the first invoice data
        details = {
            'invoice number': tableLine.find_elements_by_tag_name('td')[2].text.strip(),
            'invoice date': datetime.datetime.strptime(tableLine.find_elements_by_tag_name('td')[5].text.strip(), '%b %d, %Y')
        }
        # open invoice page
        tableLine.find_elements_by_tag_name('td')[2].find_element_by_tag_name('a').click()
        i = 0
        while True:
            time.sleep(1)
            if 'Invoice Detail View' in browser.page_source:
                break
        # open Details and save the rest of invoice data
        browser.find_element_by_link_text('View Details').click()
        while True:
            time.sleep(2)
            detailsElems = WDW(browser, 20).until(pres((By.XPATH,'//td[contains(@id, "mainContentId:j_id") and contains(@id, "-3-3") and @class="icePnlGrdCol2"]/div/table[2]'))).find_elements_by_tag_name('tr')
            if detailsElems:
                break
        for detailElem in detailsElems:
            if len(detailElem.find_elements_by_tag_name('td')) != 2:
                continue
            details[detailElem.find_elements_by_tag_name('td')[0].text] = detailElem.find_elements_by_tag_name('td')[1].text
        # get the data of each shipment on the invoice page, loop through pagination
        FedExShipmentsData = []
        while True:
            soup = bs4.BeautifulSoup(browser.page_source, 'html.parser')
            invoiceTableLines = soup.select('tbody[id="mainContentId:usInvoiceTable:tbody"]')[0].select('tr')
            for row in invoiceTableLines:
                if len(row.select('td')) < 3:
                    continue
                FedExShipmentsData.append([
                    invoiceTableLines[0].select('td')[2].text.strip(),  # tracking number
                    datetime.datetime.strptime(invoiceTableLines[0].select('td')[3].text.strip(), '%b %d, %Y'),  # invoice date
                    invoiceTableLines[0].select('td')[11].text.strip()  # total billed
                ])
            # if Next button is unclickable, break the loop
            try:
                if WDW(browser, 20).until(pres((By.XPATH, '//*[contains(@id, "mainContentId:j_id") and contains(@id, "next")]'))).get_attribute('class') == 'iceCmdLnk iceCmdLnk-dis':
                    break
            except:
                break
            # click Next
            WDW(browser, 20).until(pres((By.XPATH, '//*[contains(@id, "mainContentId:j_id") and contains(@id, "next")]'))).click()
            while True:
                time.sleep(2)
                # if the last tracking number is not on the page, the page has changed
                if FedExShipmentsData[-1][0] not in browser.page_source:
                    break

        details['shipments'] = FedExShipmentsData
        FedExData.append(details)

        file = open('FedEx1.txt', 'w')
        file.write(str(FedExData))
        file.close()

        # click back and check if it went back to the list of invoices
        WDW(browser, 20).until(pres((By.XPATH, '//*[@id="mainContentId:topbackAccSmmy"]'))).click()
        while True:
            time.sleep(2)
            if 'Search Criteria' in browser.page_source:
                break

    # Canada Past part
    CPData = []
    data = eval(open('CP1.txt').read())
    for invoice in CPData:
        if invoice['invoice number'] in ['9772129804', '9771224642', '9770308370', '9769404697', '9768463379', '9767373957', '9766476148', '9765507089', '9764477577', '9763321096', '9762313543', '9761298011', '9760217582', '9759641530', '9758365618', '9757133393', '9755805904', '9754382717', '9753159255', '9751901150', '9751019518', '9750099743', '9749116562', '9749163147', '9745522797']:
            CPData.append(invoice)

    # Open the web portal, log in
    browser.get('https://www.canadapost-postescanada.ca/dash-tableau/en/allinvoices')
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#sso_username"))).send_keys(CPlogin)
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#sso_password"))).send_keys(CPpassword)
    WDW(browser, 20).until(pres((By.CSS_SELECTOR, "#sso_action"))).click()

    # Select All Invoices
    WDW(browser, 20).until(pres((By.XPATH, '//div[@class="selected-option icon-dropdown-arrow" and @aria-owns="popup-dropdown-invoices-category"]'))).send_keys(Keys.PAGE_UP)
    time.sleep(1)
    while True:
        try:
            WDW(browser, 20).until(pres((By.XPATH, '//div[@class="selected-option icon-dropdown-arrow" and @aria-owns="popup-dropdown-invoices-category"]'))).click()
            break
        except:
            browser.find_element_by_tag_name('html').send_keys(Keys.ARROW_DOWN,Keys.ARROW_DOWN,Keys.ARROW_DOWN)
    time.sleep(1)
    while True:
        try:
            WDW(browser, 20).until(pres((By.XPATH, '//li[@id="searchingByAll"]'))).click()
            break
        except:
            browser.find_element_by_tag_name('html').send_keys(Keys.ARROW_DOWN,Keys.ARROW_DOWN,Keys.ARROW_DOWN)
    time.sleep(5)

    # get all invoice numbers, loop through each of them
    invoiceNumbers = []
    for line in WDW(browser, 20).until(pres((By.TAG_NAME, 'table-rows'))).find_elements_by_tag_name('li'):
        invoiceNumbers.append(line.find_element_by_tag_name('a').text.strip())

    # open each invoice PDF and parse the data
    for x, invoiceNumber in enumerate(invoiceNumbers):
        if invoiceNumber in ['9772129804', '9771224642', '9770308370', '9769404697', '9768463379',
                                         '9767373957', '9766476148', '9765507089', '9764477577', '9763321096',
                                         '9762313543', '9761298011', '9760217582', '9759641530', '9758365618',
                                         '9757133393', '9755805904', '9754382717', '9753159255', '9751901150',
                                         '9751019518', '9750099743', '9749116562', '9749163147', '9745522797']:
            continue

        print(f'Processing invoice {str(x)} out of {str(len(invoiceNumbers))}')
        data = {}
        browser.get(f'https://www.canadapost-postescanada.ca/dash-tableau/rs/0004008877/invoices/{invoiceNumber}?invoiceType=Details')
        while True:
            try:
                time.sleep(0.5)
                pdf = pdftotext.PDF(open(os.path.join(application_path, invoiceNumber + '.pdf'), 'rb'))
                break
            except:
                continue
        text = ''
        for page in pdf:
            text += page
        text = text.split('Total items shipped')[0]
        os.remove(os.path.join(application_path, invoiceNumber + '.pdf'))
        # get all parcel dates to use them to split the document
        shipments = []
        dates = []
        for item in re.findall(r'\n([12]\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01]))', text):
            dates.append(item[0])
        for i in range(len(dates)):
            if i == len(dates) - 1:
                textPerDate = text.split('\n' + dates[i])[1]
            else:
                textPerDate = text.split('\n' + dates[i])[1].split('\n' + dates[i])[0]
            trackNums = []
            for shipmentItem in re.findall(r'((\s+)(4008877(\d{9})))', textPerDate):
                trackNums.append(shipmentItem[2])
            for j in range(len(trackNums)):
                shipment = [trackNums[j]]
                shipment.append(datetime.datetime.strptime(dates[i], '%Y-%m-%d'))
                if j == len(trackNums) - 1:
                    textPerShipment = textPerDate.split(trackNums[j])[1]
                else:
                    textPerShipment = textPerDate.split(trackNums[j])[1].split(trackNums[j])[0]

                if 'Fuel Surcharge' in textPerShipment:
                    shipment.append(re.search(r'\d+\.\d{2}', textPerShipment.split('Fuel Surcharge')[1]).group())
                else:
                    shipment.append('0.0')

                if 'Volumetric Equivalent' in textPerShipment:
                    shipment.append(re.search(r'\d+\.\d{2}', textPerShipment.split('Volumetric Equivalent')[1]).group())
                else:
                    shipment.append('0.0')

                if 'Total' in textPerShipment:
                    shipment.append(re.search(r'\d+\.\d{2}', textPerShipment.split('Total')[1]).group())
                else:
                    shipment.append('0.0')

                if 'HST (ON)' in textPerShipment:
                    shipment.append(re.search(r'\d+\.\d{2}', textPerShipment.split('HST (ON)')[1]).group())
                else:
                    shipment.append('0.0')

                if 'GST' in textPerShipment:
                    shipment.append(re.search(r'\d+\.\d{2}', textPerShipment.split('GST')[1]).group())
                else:
                    shipment.append('0.0')

                shipments.append(shipment)

        CPData.append({
            'invoice number': invoiceNumber,
            'shipments': shipments
        })

        file = open('CP1.txt', 'w')
        file.write(str(CPData))
        file.close()

    browser.quit()
    disp.stop()

    outputData = []
    for invoice in UPSData:
        surcharges = [
            'UPS Returns-Transportation',
            'Packages Delivered but not Previously Billed',
            'Service Charge',
            'Address Corrections',
            'Shipping Charge Corrections',
            'Residential Adjustments-Shipping API',
            'Undeliverable Returns',
            'Fees',
            'Delivery Confirmation Signature Charges Not Billed-Shipping API',
            'Shipping API Voids',
            'Miscellaneous'
        ]
        surchargesTotal = 0
        for surcharge in surcharges:
            surchargesTotal += float(invoice.get(surcharge, '0').replace(',', '').strip('()$'))
        for shipment in invoice['shipments']:
            if float(shipment[6].replace(',', '')) == 0:
                continue
            try:
                date = datetime.datetime.strptime(shipment[0], '%m/%d/%Y')
            except:
                continue
            outputData.append([
                'UPS',
                date,
                shipment[6],  # total cost
                float(surchargesTotal) / len(invoice['shipments']),  # average Total Surcharges
                float(invoice.get('Address Corrections', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Address Corrections
                float(invoice.get('Packages Delivered but not Previously Billed', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Delivered but not Billed
                float(invoice.get('Residential Adjustments-Shipping API', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Residential Adjustments
                float(invoice.get('Shipping Charge Corrections', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Shipping Charge Corrections
                float(invoice.get('Fees', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Fees
                (float(invoice.get('HST', '0').replace(',', '').strip('()$')) + float(invoice.get('GST', '0').replace(',', '').strip('()$'))) / len(invoice['shipments'])  # average GST and HST
            ])

    for invoice in FedExData:
        surcharges = [
            'Fuel Surcharge',
            'Extended Delivery Area',
            'Residential Delivery',
            'Additional Handling Surcharge',
            'Address Correction',
            'Indirect Signature Req'
        ]
        surchargesTotal = 0
        for surcharge in surcharges:
            surchargesTotal += float(invoice.get(surcharge, '0').replace(',', '').strip('()$'))
        for shipment in invoice['shipments']:
            outputData.append([
                'FedEx',
                shipment[1],  # date
                shipment[2],  # total cost
                float(surchargesTotal) / len(invoice['shipments']),  # average Total Surcharges
                float(invoice.get('Address Correction', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Address Corrections
                float(invoice.get('Delivered but not Billed', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Delivered but not Billed - no such category in FedEx but I'll leave this call
                float(invoice.get('Residential Delivery', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Residential Adjustments
                float(invoice.get('Shipping Charge Corrections', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Shipping Charge Corrections
                float(invoice.get('Fees', '0').replace(',', '').strip('()$')) / len(invoice['shipments']),  # average Fees - no such category in FedEx but I'll leave this call
                (float(invoice.get('Canada GST', '0').replace(',', '').strip('()$')) + float(invoice.get('Canada HST', '0').replace(',', '').strip('()$'))) / len(invoice['shipments'])  # average GST and HST
            ])

    for invoice in CPData:
        for shipment in invoice['shipments']:
            outputData.append([
                'CP',
                shipment[1],  # date
                shipment[4],  # total cost
                float(shipment[2]) + float(shipment[3]),  # Total Surcharges
                0.0,  # average Address Corrections - no such category in Canada Post
                0.0,  # average Delivered but not Billed - no such category in Canada Post
                0.0,  # average Residential Adjustments- no such category in Canada Post
                0.0,  # average Shipping Charge Corrections- no such category in Canada Post
                0.0,  # average Fees - no such category in Canada Post
                float(shipment[5]) + float(shipment[6])
            ])

    sheets = []
    thisDay = datetime.datetime.now().day
    thisMonth = datetime.datetime.now().month
    for year in [2021, 2020, 2019]:
        shipmentsThisYear = []
        for a in outputData:
            if a[1].year == year:
                shipmentsThisYear.append(a)
        rows = []
        rows.append(['', 'Orders Shipped', 'Avg orders per day', 'Avg cost per package', 'HST + GST', 'Total Shipping Spend', 'Total Shipping Surcharges', 'Address Corrections', 'Packages Delivered but not billed', 'Residential Adj.', 'Shipping Charge Correction', 'Fees (i.e. Late fees)'])
        rows.append(['AGGREGATE'])
        for month in range(1, 13):
            shipmentsThisMonthAgr = []
            for x in shipmentsThisYear:
                if x[1].month == month:
                    shipmentsThisMonthAgr.append(x)
            if len(shipmentsThisMonthAgr) == 0:
                rows.append([datetime.datetime(year, month, 20, 0, 0, 0).strftime('%b'), 0, '', '', '', '', '', '', '', ''])
                continue
            daysInThisMonth = (shipmentsThisMonthAgr[0][1].replace(month=shipmentsThisMonthAgr[0][1].month % 12 +1, day=1)-datetime.timedelta(days=1)).day
            rows.append([
                shipmentsThisMonthAgr[0][1].strftime('%b'),
                len(shipmentsThisMonthAgr),
                round(len(shipmentsThisMonthAgr) / daysInThisMonth, 2),
                round(sum([float(item[2]) for item in shipmentsThisMonthAgr]) / len(shipmentsThisMonthAgr), 2),
                round(sum([float(item[9]) for item in shipmentsThisMonthAgr]), 2),
                round(sum([float(item[2]) for item in shipmentsThisMonthAgr]), 2),
                round(sum([item[3] for item in shipmentsThisMonthAgr]), 2),
                round(sum([item[4] for item in shipmentsThisMonthAgr]), 2),
                round(sum([item[5] for item in shipmentsThisMonthAgr]), 2),
                round(sum([item[6] for item in shipmentsThisMonthAgr]), 2),
                round(sum([item[7] for item in shipmentsThisMonthAgr]), 2),
                round(sum([item[8] for item in shipmentsThisMonthAgr]), 2)
            ])
        shipmentsYTD = []
        latestDate = datetime.datetime(2010, 1, 1)
        for y in shipmentsThisYear:
            if y[1] < datetime.datetime(year, thisMonth, thisDay, 0, 0):
                shipmentsYTD.append(y)
            if y[1] > latestDate:
                latestDate = y[1]
        if len(shipmentsYTD) == 0:
            rows.append(['YTD', 0, '', '', '', '', '', '', '', ''])
            continue
        rows.append([
            'YTD',
            len(shipmentsYTD),
            round(len(shipmentsYTD) / (latestDate - latestDate.replace(month=1, day=1)).days, 2),
            round(sum([float(item[2]) for item in shipmentsYTD]) / len(shipmentsYTD), 2),
            round(sum([float(item[9]) for item in shipmentsYTD]), 2),
            round(sum([float(item[2]) for item in shipmentsYTD]), 2),
            round(sum([item[3] for item in shipmentsYTD]), 2),
            round(sum([item[4] for item in shipmentsYTD]), 2),
            round(sum([item[5] for item in shipmentsYTD]), 2),
            round(sum([item[6] for item in shipmentsYTD]), 2),
            round(sum([item[7] for item in shipmentsYTD]), 2),
            round(sum([item[8] for item in shipmentsYTD]), 2)
        ])

        for carrier in ['CP', 'UPS', 'FedEx']:
            rows.append([])
            rows.append([carrier])
            shipmentsOfCarrier = []
            for b in shipmentsThisYear:
                if b[0] == carrier:
                    shipmentsOfCarrier.append(b)

            for month in range(1,13):
                shipmentsThisMonth = []
                for c in shipmentsOfCarrier:
                    if c[1].month == month:
                        shipmentsThisMonth.append(c)
                if len(shipmentsThisMonth) == 0:
                    rows.append(
                        [datetime.datetime(year, month, 20, 0, 0, 0).strftime('%b'), 0, '', '', '', '', '', '', '', ''])
                    continue
                daysInThisMonth = (shipmentsThisMonth[0][1].replace(month=shipmentsThisMonth[0][1].month % 12 + 1, day=1) - datetime.timedelta(days=1)).day
                rows.append([
                    shipmentsThisMonth[0][1].strftime('%b'),
                    len(shipmentsThisMonth),
                    round(len(shipmentsThisMonth) / daysInThisMonth, 2),
                    round(sum([float(item[2]) for item in shipmentsThisMonth]) / len(shipmentsThisMonth), 2),
                    round(sum([float(item[9]) for item in shipmentsThisMonth]), 2),
                    round(sum([float(item[2]) for item in shipmentsThisMonth]), 2),
                    round(sum([item[3] for item in shipmentsThisMonth]), 2),
                    round(sum([item[4] for item in shipmentsThisMonth]), 2),
                    round(sum([item[5] for item in shipmentsThisMonth]), 2),
                    round(sum([item[6] for item in shipmentsThisMonth]), 2),
                    round(sum([item[7] for item in shipmentsThisMonth]), 2),
                    round(sum([item[8] for item in shipmentsThisMonth]), 2)
                ])

            shipmentsYTDcarrier = []
            latestDate = datetime.datetime(2010, 1, 1)
            for y in shipmentsOfCarrier:
                if y[1] < datetime.datetime(year, thisMonth, thisDay, 0, 0):
                    shipmentsYTDcarrier.append(y)
                if y[1] > latestDate:
                    latestDate = y[1]
            if len(shipmentsYTDcarrier) == 0:
                rows.append(['YTD', 0, '', '', '', '', '', '', '', ''])
                continue
            rows.append([
                'YTD',
                len(shipmentsYTDcarrier),
                round(len(shipmentsYTD) / (latestDate - latestDate.replace(month=1, day=1)).days, 2),
                round(sum([float(item[2]) for item in shipmentsYTDcarrier]) / len(shipmentsYTDcarrier), 2),
                round(sum([float(item[9]) for item in shipmentsYTDcarrier]), 2),
                round(sum([float(item[2]) for item in shipmentsYTDcarrier]), 2),
                round(sum([item[3] for item in shipmentsYTDcarrier]), 2),
                round(sum([item[4] for item in shipmentsYTDcarrier]), 2),
                round(sum([item[5] for item in shipmentsYTDcarrier]), 2),
                round(sum([item[6] for item in shipmentsYTDcarrier]), 2),
                round(sum([item[7] for item in shipmentsYTDcarrier]), 2),
                round(sum([item[8] for item in shipmentsYTDcarrier]), 2)
            ])

        sheets.append(rows)

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = '2021'
    wb.create_sheet(title='2020')
    wb.create_sheet(title='2019')

    for i in range(len(sheets)):
        sheet = wb[wb.sheetnames[i]]
        for r, row in enumerate(sheets[i]):
            for c, cell in enumerate(row):
                if type(cell) == type(0.1):
                    value = round(cell, 2)
                else:
                    value = cell
                sheet.cell(row=r+1, column=c+1).value = cell

    wb.save('output.xlsx')

except:
    traceback.print_exc()
    reportName = os.path.join(application_path, 'log (shipping dashboard) ' + datetime.datetime.now().strftime('%b %d %Y %H-%M') + '.txt')
    logging.basicConfig(filename=reportName, level=logging.INFO, format=' %(asctime)s -  %(levelname)s -  %(message)s')
    logging.error(traceback.format_exc())