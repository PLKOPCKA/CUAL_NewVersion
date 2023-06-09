from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from Password import fico_pass, fico_user
from SQL_Run_Proc import *
import time
from datetime import datetime
# import os
from Parameters import bands, cuts, max_solve, min_gap, scenario_name, \
    model_name, resolve, ISC_path, fico_results_archive, win_user, fico_cual_folder
# import openingZippedFile

options = Options()
service = Service(executable_path=r"E:\00_SQL Queries\Kamil\msedgedriver.exe")
options.binary_location = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
driver = webdriver.Edge(options=options, service=service)


def open_folder(name, double_click=False):
    done = False
    i = 0
    time.sleep(3)
    for table in driver.find_elements(By.XPATH, "//div[contains(@class, 'scenario-manager-items-list-container')]"):
        for row in table.find_elements(By.XPATH, "//div[contains(@class, 'tabulator-table')]"):
            for col in row.find_elements(By.XPATH, "//div[contains(@class, 'tabulator-row tabulator-selectable tabulator-row-odd')]"):
                if col.text == name:
                    i += 1
                    if i == 2:
                        if double_click:
                            chains = ActionChains(driver)
                            chains.double_click(col).perform()
                        else:
                            col.click()
                        done = True
                        break
            if done:
                break
        if done:
            break
    return


def navigate_to_frame(name):
    time.sleep(5)
    for el2 in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH,"//a[contains(@class, 'tab-item view-select-item')]"))):
        if el2.text == name:
            el2.click()
            break


def go_to_app():
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//ul[contains(@class, 'nav nav-pills pull-left')]/li"))):
        if el.text == 'APP':
            el.click()
            break


def click_job():
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//ul[contains(@class, 'nav nav-pills pull-left')]/li"))):
        if el.text == 'JOBS':
            el.click()
            break


def select_right_panel(name, iframe=True):
    time.sleep(5)
    if iframe:
        WebDriverWait(driver, 120).until(EC.frame_to_be_available_and_switch_to_it("view-iframe"))
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//ul[contains(@class, 'sidebar compact-sidebar list-group')]/li/a"))):
        if el.text == name:
            el.click()
            break


def go_to_job(operation, max_wait=3600, check=5, get_time=False):
    time.sleep(5)
    click_job()

    i = 0
    t = 0
    done = False
    try:
        while i <= max_wait:
            i += check
            t += check
            if t >= 200:
                click_job()
                t = 0
            time.sleep(check)
            j = 0
            for row in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//tr[contains(@data-id, '8df3e616-4224-4aa9-8cac-f9dda59ef203')]"))):
                j = 0
                model = ''
                process = ''
                status = ''
                exec_time = 'timer error'
                for col in row.find_elements(By.TAG_NAME, 'td'):
                    j += 1
                    if j == 2: model = col.text
                    if j == 3: process = col.text
                    if j == 5: status = col.text
                    if j == 10 and col.text != '': exec_time = col.text
                    if model == 'CUAL' and process == operation and status == 'Completed':
                        done = True
                        break
                if done:
                    break
            if done:
                break
    except:
        pass

    go_to_app()

    if get_time:
        return exec_time


def click_button(name, loc=1):
    time.sleep(5)
    i = 0
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//button[contains(@class, 'btn btn-primary')]"))):
        if el.text == name:
            i += 1
            if i == loc:
                el.click()
                break

    if name == 'DOWNLOAD RESULTS':
        select_right_panel('Results Export', False)
    driver.switch_to.default_content()


def lp_relaxation(tick=False):
    time.sleep(10)
    i = 0
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//div[contains(@class, 'checkbox')]/label"))):
        i += 1
        if i == 4:
            if tick:
                if not el.find_element(By.TAG_NAME, "input").is_selected():
                    el.click()
            else:
                if el.find_element(By.TAG_NAME, "input").is_selected():
                    el.click()
            break


def unclick_messages():
    time.sleep(3)
    try:
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("view-iframe"))
        for el in WebDriverWait(driver, 10).until(EC.visibility_of_any_elements_located((By.XPATH, "//div[contains(@id, 'message')]/div"))):
            el.click()
    except:
        pass
    driver.switch_to.default_content()


def gap_split_band(cuted, banded, part, itter):
    time.sleep(10)

    if part == 2:
        execute_sql_proc(cuted, banded, itter+1, ISC_path)
    else:
        new_bands(itter, cuted, banded)

    navigate_to_frame('Setup')
    select_right_panel('Mass Data Import')
    click_button('IMPORT DATA FROM DATABASE', 2)

    unclick_messages()
    go_to_job('RUN_ALL_QUERIES', 600, 5)

    navigate_to_frame('Optimization')
    unclick_messages()
    try:
        select_right_panel('Optimization')
    except:
        time.sleep(30)
        driver.refresh()
        unclick_messages()
        select_right_panel('Optimization')

    lp_relaxation(False)
    click_button('OPTIMIZE')
    exec_time = go_to_job('OPTIMIZE', max_solve+3600, 60, True)
    navigate_to_frame('Optimization')
    time.sleep(20)
    navigate_to_frame('Optimization')
    unclick_messages()
    try:
        select_right_panel('Optimization')
    except:
        time.sleep(20)
        driver.refresh()
        unclick_messages()
        select_right_panel('Optimization')

    # collect the data about Objective Function value and gap
    z = 0
    for row in driver.find_elements(By.TAG_NAME, 'h5'):
        z += 1
        if z == 2:
            objective = row.text
        if z == 3:
            gap = row.text

    driver.switch_to.default_content()

    # datetime object containing current date and time
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

    insert_outputs(scenario_name, banded, dt_string, objective, gap, exec_time)

    time.sleep(5)
    navigate_to_frame('Results')
    driver.refresh()
    time.sleep(30)
    select_right_panel('Results Export')
    click_button('POPULATE RESULTS')
    go_to_job('POPULATE_RESULTS', 600, 3)
    driver.refresh()
    unclick_messages()
    select_right_panel('Results Export')
    time.sleep(5)
    click_button('DOWNLOAD RESULTS')
