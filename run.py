from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import StaleElementReferenceException
import logging
import os
import getopt
import sys
import csv
from shutil import copyfile
from datetime import datetime, time


import invoice_xlsx_handler
import jobs_sheet_xlsx_handler
import email_handler

logger = logging.getLogger('label')

# config file
cfg = {}
with open('label.conf') as hlr:
    for line in hlr:
        split_line = line.split('::')
        cfg[split_line[0].strip()] = split_line[1].strip()


chromeOptions = webdriver.ChromeOptions()
prefs = {}
prefs["credentials_enable_service"] = False
prefs["password_manager_enabled"] = False
chromeOptions.add_experimental_option("prefs", prefs)

def start_driver():

    # driver
    if cfg['browser'] == 'Chrome-OSX':
        driver = webdriver.Chrome(chrome_options=chromeOptions)
    elif cfg['browser'] == 'Chrome':
        driver_path = os.path.join(os.path.dirname(__file__), 'chromedriver.exe')
        driver = webdriver.Chrome(driver_path, chrome_options=chromeOptions)
    driver.maximize_window()
    return driver


def label(driver, technician):
    import time

    driver.get(cfg['site_url'])

    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'userName'))).send_keys(technician[1])
        driver.find_element_by_css_selector('#password').send_keys(technician[2])
        driver.find_element_by_css_selector('#submitBtn').click()

        _wait_for_data_loaded(driver)
        _disable_location_services(driver)
        _wait_for_data_loaded(driver)
        _out_time_mask(driver)

        # schedule
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((
            By.CSS_SELECTOR, '#indexButtonsPanel>div:nth-of-type(1)'))).click()
        logger.debug('Schedule clicked upon')

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'listContainer')))

        _wait_for_assignments(driver)
        time.sleep(5)

        jobs_to_process = driver.find_elements_by_css_selector('#listContainer>div')

        if len(jobs_to_process) == 0:
            logger.info('No jobs to process for technician: {}'.format(technician[0]))
            return 'NOJOBS', 'NOJOBS', 'NOJOBS'

        logger.info('Initially found {} jobs to process for technician: {}.'.format(len(jobs_to_process), technician[0]))

        _abort_warnings(driver)

        # csv
        jobs = []

        # jobs sheet excel
        jobs_sheet = []

        # invoice sheet excel
        invoices_sheet = []

        i_ = 0

        while True:

            jobs_el = WebDriverWait(driver, 2).until(EC.presence_of_all_elements_located((
                By.CSS_SELECTOR, '#listContainer>div td #info')))

            if len(jobs_el) == i_:
                logger.info('Scraped {} jobs for technician: {}.'.format(len(jobs_el), technician[0]))
                break
            else:
                logger.debug('Job length: {} Job number: {}.'.format(len(jobs_el), i_))

            jobs_el[i_].click()
            i_ += 1

            # csv
            job = []
            _headings = []

            # job sheet excel
            job_sheet = [None] * 15

            # invoice sheet excel
            invoice_sheet = [None] * 13

            # ---------------------------
            job_sheet[0] = time.strftime('%d-%m-%Y', time.localtime())

            invoice_sheet[0] = time.strftime('%d-%m-%Y', time.localtime())
            invoice_sheet[1] = technician[0]

            scrape_time = time.strftime('%H:%M %d-%m-%Y', time.localtime())
            job.append(scrape_time)

            job.append(technician[0])

            WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.W6CMPropPanel_ValuePanel')))

            # ToW
            _headings.append('ToW')
            selector = '[objname="W6CMPropPanelText[CallID:String]"] #textBoxText'
            el = driver.find_element_by_css_selector(selector)
            tow = el.get_attribute('value')

            job.append(tow)
            job_sheet[8] = tow
            invoice_sheet[2] = tow

            # TaskType
            _headings.append('TaskType')
            selector = '[objname="W6CMPropPanelCombo[TaskType:foreign_Key]"] [objname="ComboBoxControl.Value"]'
            select_el = driver.find_element_by_css_selector(selector)
            select = Select(select_el)
            selected = select.first_selected_option
            task = selected.text

            job.append(task)

            # Contact Name
            _headings.append('Contact Name')
            selector = '[objname="W6CMPropPanelText[ContactName:String]"] #textBoxText'
            el = driver.find_element_by_css_selector(selector)
            contact_name = el.get_attribute('value')

            job.append(contact_name)
            job_sheet[3] = contact_name

            # Contact Phone Number
            _headings.append('Contact Phone Number')
            selector = '[objname="W6CMPropPanelText[ContactPhoneNumber:String]"] #textBoxText'
            el = driver.find_element_by_css_selector(selector)
            phone_numer = el.get_attribute('value')

            job.append(phone_numer)
            job_sheet[2] = phone_numer

            # Street
            _headings.append('Street')
            selector = '[objname="W6CMPropPanelText[Street:String]"] #textBoxText'
            el = driver.find_element_by_css_selector(selector)
            street = el.get_attribute('value')

            job.append(street)

            # City
            _headings.append('City')
            selector = '[objname="W6CMPropPanelText[City:String]"] #textBoxText'
            el = driver.find_element_by_css_selector(selector)
            city = el.get_attribute('value')

            job.append(city)

            # State
            _headings.append('State')
            selector = '[objname="W6CMPropPanelText[State:String]"] #textBoxText'
            el = driver.find_element_by_css_selector(selector)
            state = el.get_attribute('value')

            job.append(state)
            job_sheet[1] = street + ' ' + city + ' ' + state
            invoice_sheet[3] = street + ' ' + city + ' ' + state

            logger.info('Scraped first page of: ' + tow)

            _wait_for_icon_container_disappears(driver)

            # go right
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#navigateRight'))).click()

            selector = '[objname="W6CMPropPanelCombo[CommsCLRiskAssessment:foreign_Key]"]'
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector)))

            time.sleep(5)

            # go right
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#navigateRight'))).click()

            time.sleep(2)

            # click on tow number
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'itemText'))).click()

            time.sleep(2)

            # Appt Start (Date)
            _headings.append('Appt Start (Date)')
            selector = '[objname="W6CMPropPanelDate[WorkQueueStart:date]"] input.W6CMPropPanel_Date'
            el = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
            app_start_date = el.get_attribute('value')

            job.append(app_start_date)

            # Appt Start (Time)
            _headings.append('Appt Start (Time)')
            selector = '[objname="W6CMPropPanelDate[WorkQueueStart:date]"] input#timeInput'
            el = driver.find_element_by_css_selector(selector)
            app_start_time = el.get_attribute('value')

            job.append(app_start_time)

            # Appt Finish (Date)
            _headings.append('Appt Finish (Date)')
            selector = '[objname="W6CMPropPanelDate[WorkQueueEnd:date]"] input.W6CMPropPanel_Date'
            el = driver.find_element_by_css_selector(selector)
            app_finish_date = el.get_attribute('value')

            job.append(app_finish_date)

            # Appt Finish (Time)
            _headings.append('Appt Finish (Time)')
            selector = '[objname="W6CMPropPanelDate[WorkQueueEnd:date]"] input#timeInput'
            el = driver.find_element_by_css_selector(selector)
            app_finish_time = el.get_attribute('value')

            job.append(app_finish_time)

            job_sheet[6] = app_start_date + ' ' + app_start_time + ' - ' + app_finish_date + ' ' + app_finish_time

            # Job Requirement
            _headings.append('Job Requirement')
            selector = '[objname="W6CMPropPanelCombo[CommsJobRequirementType:foreign_Key]"] [objname="ComboBoxControl.Value"]'
            select_el = driver.find_element_by_css_selector(selector)
            select = Select(select_el)
            selected = select.first_selected_option
            job_req = selected.text

            job.append(job_req)

            # PSU Type
            _headings.append('PSU Type')
            selector = '[objname="W6CMPropPanelCombo[PSUPowerSupplyType:foreign_Key]"] [objname="ComboBoxControl.Value"]'
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.CSS_SELECTOR, selector)))
            select_el = driver.find_element_by_css_selector(selector)
            select = Select(select_el)
            selected = select.first_selected_option
            psu = selected.text

            job.append(psu)

            job_sheet[12] = psu
            invoice_sheet[6] = psu

            # NTD ID
            _headings.append('NTD ID')
            selector = '[objname="W6CMPropPanelText[NTDID:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            ntd_id = el.get_attribute('value')

            job.append(ntd_id)

            job_sheet[9] = ntd_id

            # FDH Nearest Address
            _headings.append('FDH Nearest Address')
            selector = '[objname="W6CMPropPanelText[FDHNearestAddress:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            fdh = el.get_attribute('value')
            fdh = fdh.replace('\n', ' ')

            rfind = fdh.rfind(',')
            job.append(fdh[:rfind])

            job_sheet[5] = fdh[:rfind]

            # Multiport Name
            _headings.append('Multiport Name')
            selector = '[objname="W6CMPropPanelText[MultiportName:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            multi = el.get_attribute('value')
            rfind_multi = multi.rfind('-')
            job.append(multi[rfind_multi + 1:])

            job_sheet[13] = multi[rfind_multi + 1:]

            # Multiport Number
            _headings.append('Multiport Number')
            selector = '[objname="W6CMPropPanelText[MultiportPortNumber:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            port = el.get_attribute('value')

            job.append(port)

            job_sheet[14] = port

            # Splitter Port Number
            _headings.append('Splitter Port Number')
            selector = '[objname="W6CMPropPanelText[SplitterPortNumber:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            port_num = el.get_attribute('value')

            job.append(port_num)

            job_sheet[11] = port_num

            # Splitter Port Slot Name
            _headings.append('Splitter Port Slot Name')
            selector = '[objname="W6CMPropPanelText[SplitterPortSlotName:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            slot = el.get_attribute('value')
            rfind_slot = slot.rfind('-')
            job.append(slot[rfind_slot + 1:])

            job_sheet[10] = slot[rfind_slot + 1:]

            # Local Port Number
            _headings.append('Local Port Number')
            selector = '[objname="W6CMPropPanelText[SplitterPortTrackName:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            local = el.get_attribute('value')

            job.append(local)

            job_sheet[7] = local

            # Splitter Port Rack Name  ( in csv: FDH Name)
            _headings.append('Splitter Port Rack Name')
            selector = '[objname="W6CMPropPanelText[LocalPortNumber:String]"] textarea#textBoxText'
            el = driver.find_element_by_css_selector(selector)
            rack = el.get_attribute('value')

            job.append(rack)

            job_sheet[4] = rack

            jobs.append(job)
            jobs_sheet.append(job_sheet)
            invoices_sheet.append(invoice_sheet)

            logger.info('Scraped second page of: ' + tow)

            # click on icon to trigger saving
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '.PanelSideMenuItemIcon'))).click()

            _abort_saving(driver)

        return jobs, jobs_sheet, invoices_sheet

    except StaleElementReferenceException:
        logger.info('Failed to scrape data: StaleElementReferenceException. Will re-try.')
        return None, None, None
    except:
        logger.info('Failed to scrape data. Will re-try.')
        logger.debug('ERROR: \n', exc_info=True)
        return None, None, None


def gmaps(driver, fetched_jobs, address_pointer=None):

    driver.get(cfg['google_maps_url'])

    try:
        fetch_jobs_with_url = []

        for fetch_job in fetched_jobs:

            fetch_job_with_url = fetch_job

            search = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'searchboxinput')))
            if not address_pointer:
                address = 'Australia, {}, {}, {}'.format(fetch_job[6], fetch_job[7], fetch_job[8])
            else:
                address = 'Australia, {}, {}'.format(fetch_job[16], fetch_job[8])

            search.send_keys(address)
            driver.find_element_by_css_selector('#searchbox-searchbutton').click()

            try:
                share_button = '.maps-sprite-pane-action-ic-share'
                share = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, share_button)))
                share.click()
            except TimeoutException:
                logger.info('Failed to get addresses for address: {}. Google failed to find it.'.format(address))
                fetch_job_with_url.append("")
                fetch_jobs_with_url.append(fetch_job_with_url)

                clear_search = '[guidedhelpid="clear_search"]'
                WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, clear_search))).click()
                continue

            modal = '.modal-dialog-content'
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, modal)))

            share_link = '#last-focusable-in-modal'
            el_shared = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, share_link)))
            link = el_shared.get_attribute('value')

            fetch_job_with_url.append(link)
            logger.info('address: {} \n maps url: {}'.format(address, link))

            close_modal = '.close-button-white-circle'
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, close_modal))).click()
            modal = '.modal-dialog-content'
            WebDriverWait(driver, 5).until_not(EC.presence_of_element_located((By.CSS_SELECTOR, modal)))

            clear_search = '[guidedhelpid="clear_search"]'
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, clear_search))).click()

            fetch_jobs_with_url.append(fetch_job_with_url)

        return fetch_jobs_with_url

    except:
        logger.info('Failed to get addresses from google maps. Will re-try.')
        logger.debug('ERROR: \n', exc_info=True)


def _wait_for_loading_data(locator, driver):
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, locator)))
        logger.debug('El: {} appeared'.format(locator))
    except TimeoutException:
        logger.debug('Element has not appeared')


def _wait_for_icon_container_disappears(driver):
    import time
    try:
        for _ in range(30):
            el = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.iconContainer')))
            if el.is_displayed():
                time.sleep(2)
                logger.debug('Element: {} displays'.format('.iconContainer'))
            else:
                logger.debug('El: {} doesnt display (is_displayed)'.format('.iconContainer'))
                break
    except TimeoutException:
        logger.debug('Element: {} doesnt display'.format('.iconContainer'))


def _wait_for_spinner(driver):
    try:
        loading = "#spinnerContainer"
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, loading)))
    except TimeoutException:
        logger.debug('Spinner has not appeared')


def _wait_for_data_loaded(driver):
    import time

    loading = "#loadingPanel"
    _wait_for_loading_data(loading, driver)

    for _ in range(120):
        try:
            el = WebDriverWait(driver, 1).until(EC.presence_of_element_located((
                By.CSS_SELECTOR, loading)))
            if el.is_displayed():
                time.sleep(1)
                raise
            else:
                break
        except TimeoutException:
            logger.debug('loadingPanel did not appear')
            return
        except:
            logger.debug('loadingPanel being spun')

    spinner = "#maskContainer"
    _wait_for_loading_data(spinner, driver)

    for _ in range(60):
        try:
            el = WebDriverWait(driver, 1).until(EC.presence_of_element_located((
                By.CSS_SELECTOR, spinner)))
            if el.is_displayed():
                logger.debug('spinner displaying')
                style = el.get_attribute('style')
                logger.debug('spinner style: {}'.format(style))
                time.sleep(1)
                raise
            else:
                logger.debug('spinner not displaying')
                style = el.get_attribute('style')
                logger.debug('spinner style: {}'.format(style))
                return
        except TimeoutException:
            logger.debug('spinner did not appear')
            return
        except:
            logger.debug('spinner being spun')


def _wait_for_assignments(driver):

    assignments = "#listContainer>div"

    for x_ in range(15):
        try:
            WebDriverWait(driver, 4).until(EC.presence_of_element_located((
                By.CSS_SELECTOR, assignments)))
            logger.debug('assignments loaded')
            break
        except TimeoutException:
            logger.info('assignments not yet loaded: ({})'.format(x_))


def _out_time_mask(driver):

    mask = ".cm_ui_window_buttons_surface"

    try:
        WebDriverWait(driver, 10).until_not(EC.presence_of_element_located((
            By.CSS_SELECTOR, mask)))
        logger.debug('.cm_ui_window_buttons_surface not displaying')
    except TimeoutException:
        logger.debug('.cm_ui_window_buttons_surface timed out')


def _abort_warnings(driver):
    import time

    try:
        driver.find_element_by_css_selector('#listContainer>div td').click()

        warning = ".cm_ui_window_buttons_surface .button1"
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, warning)))
        logger.debug('warning aborted')

        jquery = '$(\'.cm_ui_window_buttons_surface .button1\').trigger(\'click\')'
        driver.execute_script(jquery)

        _out_time_mask(driver)

        time.sleep(2)
    except TimeoutException:
        logger.debug('warning not loaded')


def _abort_saving(driver):
    import time

    try:
        pop_up = '.cm_ui_window_table_structure .button1:nth-of-type(2)'
        el = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, pop_up)))
        time.sleep(2)
        el.click()

        pop_up = '.cm_ui_window_table_structure'
        WebDriverWait(driver, 2).until_not(EC.presence_of_element_located((By.CSS_SELECTOR, pop_up)))

        logger.debug('saving aborted')

        time.sleep(0.5)
    except TimeoutException:
        logger.debug('saving not loaded')


def _disable_location_services(driver):
    import time

    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((
            By.CSS_SELECTOR, '#checkBoxImg'))).click()

        warning = ".cm_ui_window_buttons_button"
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((
            By.CSS_SELECTOR, warning))).click()
        logger.debug('location service aborted')

        WebDriverWait(driver, 2).until_not(EC.presence_of_element_located((
            By.CSS_SELECTOR, '#checkBoxImg')))
        time.sleep(1)
    except TimeoutException:
        logger.debug('location service not loaded')


def _create_temp():

    heading = ['Scrape Time', 'Technician', 'ToW', 'TaskType', 'Contact Name', 'Contact Phone Number', 'Street', 'City',
               'State', 'Appt Start (Date)', 'Appt Start (Time)', 'Appt Finish (Date)', 'Appt Finish (Time)', 'Job Requirement',
               'PSU Type', 'NTD ID', 'FDH Nearest Address', 'Multiport Name', 'Multiport Number', 'Splitter Port Number',
               'Splitter Port Slot Name', 'Local Port Number', 'FDH Name']# 'Maps URL', 'Neareset Maps URL']

    with open('./sheets/scraped.csv', 'wb') as hlr:
        wrt = csv.writer(hlr, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        wrt.writerow(heading)
        logger.info('heading added to scraped.csv: {}'.format(heading))


def _write_row(row):
    with open('./sheets/scraped.csv', 'ab') as hlr:
        wrt = csv.writer(hlr, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        wrt.writerow(row)
        logger.info('added to scraped.csv file: {}'.format(row))


def _read_rows():
    with open('./sheets/scraped.csv', 'rb') as hlr:
        rd = csv.reader(hlr, delimiter=',', quotechar='"')
        return [row for row in rd]


def _get_timestamp(hours=False):
    import time

    if not hours:
        return time.strftime('%d%m%y', time.localtime())
    else:
        return time.strftime('%d%m%y%H%M', time.localtime())

if __name__ == '__main__':
    verbose = None
    test_mode = None
    log_file = None
    scrape_iter = 5

    argv = sys.argv[1:]
    opts, args = getopt.getopt(argv, "ltvr:", ["log-file", "test", "verbose","run-mode="])
    for opt, arg in opts:
        if opt in ("-t", "--test"):
            test_mode = True
        elif opt in ("-v", "--verbose"):
            verbose = True
        elif opt in ("-r", "--run-mode"):
            run_mode = arg
        elif opt in ("-l", "--log-file"):
            log_file = True

    timestamp = _get_timestamp()
    console = logging.StreamHandler(stream=sys.stdout)
    logger.addHandler(console)
    ch = logging.Formatter('[%(levelname)s] %(message)s')
    console.setFormatter(ch)

    if log_file:
        log_file = os.path.join(os.path.dirname(__file__), timestamp + ".log")
        file_hndlr = logging.FileHandler(log_file)
        logger.addHandler(file_hndlr)
        ch_file = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
        file_hndlr.setFormatter(ch_file)

    if verbose:
        logger.setLevel(logging.getLevelName('DEBUG'))
    else:
        logger.setLevel(logging.getLevelName('INFO'))
    logger.debug('CLI args: {}'.format(opts))

    # update sheets
    if run_mode == 'email':
        now = datetime.now().time()
        if time(7, 30) <= now <= time(19, 0):
            logger.info('Time in range: 7:30 - 19:00')
        else:
            logger.info('Time not in range: 7:30 - 19:00')
            sys.exit()

    # _create_temp()

    technicians = [[cfg['name_1'], cfg['username_1'], cfg['pass_1'], cfg['email_1']], [cfg['name_2'], cfg['username_2'],
                    cfg['pass_2'], cfg['email_2']], [cfg['name_3'], cfg['username_3'], cfg['pass_3'], cfg['email_3']]]

    # invoice file if exists
    dirs = os.listdir(os.path.join(os.path.dirname(__file__), 'sheets'))
    for file_ in dirs:
        if file_.count('invoice'):
            logger.info('Invoice file found: {}'.format(file_))
            invoice_file = './sheets/' + file_
            break
    else:
        # create empty invoice sheet
        invoice_file = './sheets/invoice_{}.xlsx'.format(timestamp)
        copyfile('./sheets/template.xlsx', invoice_file)
        # invoice_xlsx_handler.create_file_and_write_to([])
        logger.info('Invoice file created: {}'.format(invoice_file))

    try:
        os.remove('./sheets/jobs.xlsx')
    except OSError:
        pass
    jobs_sheet_xlsx_handler.create_file_and_write_to([], './sheets/jobs.xlsx')

    for technician in technicians:

        address_urls = []

        for i in range(1, scrape_iter):
            logger.info('Scrape jobs begins for technician: {}. ({})'.format(technician[0], i))
            driver_ = start_driver()
            fetch_jobs, fetched_jobs_sheet, fetched_invoice_sheet = label(driver_, technician)

            if fetch_jobs:
                break
            else:
                if i == scrape_iter - 1:
                    file_name = _get_timestamp() + '.png'
                    logger.info('Taking screenshot: ' + file_name)
                    driver_.save_screenshot('./screencaps/' + file_name)
                driver_.delete_all_cookies()
                driver_.close()
        else:
            logger.error('Did not scrape jobs')
            driver_.quit()
            sys.exit()

        driver_.delete_all_cookies()
        driver_.close()

        if fetch_jobs != "NOJOBS":

            for i in range(1, 4):
                logger.info('Get address from google maps for technician: {}. ({})'.format(technician[0], i))
                driver_ = start_driver()
                jobs_with_url = gmaps(driver_, fetch_jobs)
                if jobs_with_url:
                    break
                else:
                    driver_.delete_all_cookies()
                    driver_.close()
            else:
                logger.error('Did not get address from google maps')
                driver_.quit()
                sys.exit()

            driver_.delete_all_cookies()
            driver_.close()

            for i in range(1, 4):
                logger.info('Get FDH address from google maps for technician: {}. ({})'.format(technician[0], i))
                driver_ = start_driver()
                jobs_with_url = gmaps(driver_, fetch_jobs, address_pointer='FDH')
                if jobs_with_url:
                    break
                else:
                    driver_.delete_all_cookies()
                    driver_.close()
            else:
                logger.error('Did not get FDH address from google maps')
                driver_.quit()
                sys.exit()

            driver_.delete_all_cookies()
            driver_.close()

            existing_data = _read_rows()
            existing_data_no_timestamp = [data_[1:] for data_ in existing_data]

            logger.info('Writing scraped jobs for technician: {}'.format(technician[0]))
            # writing rows without maps url
            for job_counter, row in enumerate(jobs_with_url):
                if row[1:-2] not in existing_data_no_timestamp:
                    logger.info('add scraped row: {}'.format(row))
                    _write_row(row[:-2])

                    jobs_sheet_xlsx_handler.add_to_file([fetched_jobs_sheet[job_counter]], './sheets/jobs.xlsx')
                    logger.info('add jobs to jobs sheet: \n {}'.format([fetched_jobs_sheet[job_counter]]))

                    invoice_xlsx_handler.add_to_file([fetched_invoice_sheet[job_counter]], invoice_file)
                    logger.info('add jobs to invoice sheet: \n {}'.format([fetched_invoice_sheet[job_counter]]))

                    address_urls.append([row[-2], row[-1]])
                else:
                    logger.debug('exists: skip scraped row: {}'.format(row))
                    logger.info('skip job sheet: \n {}'.format([fetched_jobs_sheet[job_counter]]))
                    logger.info('skip invoice sheet: \n {}'.format([fetched_invoice_sheet[job_counter]]))

            # email per technician
            if run_mode == 'email':
                logger.info('Run mode: email')
                if len(address_urls) > 0:
                    if test_mode:
                        recipient = cfg['test_email']
                    else:
                        recipient = technician[3]
                    logger.info('Recipient: ' + recipient)
                    email_handler._send_email(recipient, logger, address_urls, cfg)

                    # create a new job file
                    try:
                        os.remove('./sheets/jobs.xlsx')
                    except OSError:
                        pass
                    jobs_sheet_xlsx_handler.create_file_and_write_to([], './sheets/jobs.xlsx')
        else:
            logger.info('No jobs --> No address for technician: {}. ({})'.format(technician[0], i))

    if run_mode == 'print':
        logger.info('Run mode: print')
        print_file = os.path.join(os.path.dirname(__file__), 'sheets', 'jobs.xlsx')
        os.startfile(print_file, "print")
        logger.info('Sent to a printer: ' + print_file)

    driver_.quit()