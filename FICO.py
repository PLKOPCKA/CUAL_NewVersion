# from msedge.selenium_tools import Edge, EdgeOptions
# from selenium import webdriver
# from selenium.webdriver.edge.options import Options
# from selenium.webdriver.edge.service import Service
# from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
from Password import fico_pass, fico_user
# from SQL_Run_Proc import *
# import time
# from datetime import datetime
# import os
# from Parameters import bands, cuts, max_solve, min_gap, scenario_name, \
#     model_name, resolve, ISC_path, fico_results_archive
import openingZippedFile
from FICO_func import *

start_time = time.time()
transform_all()

# options = EdgeOptions()
# options.use_chromium = True
# options.binary_location = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
# driver = Edge(executable_path=r"E:\00_SQL Queries\Kamil\msedgedriver.exe", options=options)

# driver.get('http://scs-prd-as02:8860/insight/login.jsp')
driver.get('https://scs-prd-as02.global.internal.carlsberggroup.com/insight/login.jsp')
driver.maximize_window()

try:

    login = driver.find_element(By.ID, 'inputEmail')
    login.send_keys(fico_user)

    password = driver.find_element(By.ID, 'inputPassword')
    password.send_keys(fico_pass)

    connect = driver.find_element(By.ID, 'loginBtn')
    connect.click()
    time.sleep(5)

    app_button = driver.find_element(By.XPATH, "//div[contains(@class, 'project-clickable-area')]")
    app_button.click()

    time.sleep(3)
    app_button = driver.find_element(By.XPATH, "//ul[contains(@class, 'shelf-item-list col-xs-11 scenario-manager-launcher ui-sortable')]")
    app_button.click()


    open_folder('Network_Design_Models')
    open_folder('CustomerAllocation')
    open_folder(fico_cual_folder)

    # select the CUAL model from the folder
    time.sleep(3)
    done = False
    for row in driver.find_elements(By.XPATH, "//div[contains(@class, 'tabulator-row tabulator-selectable tabulator-row-even')]"):
        for col in row.find_elements(By.XPATH, "//div[contains(@class, 'tabulator-cell')]"):
            if col.text == model_name:
                chains = ActionChains(driver)
                chains.double_click(col).perform()
                done = True
                break
        if done:
            break

    app_button.click()

    go_to_app()
    navigate_to_frame('Setup')
    select_right_panel('Mass Data Import')
    click_button('IMPORT DATA FROM DATABASE', 2)

    go_to_job('RUN_ALL_QUERIES', 600, 5)

    navigate_to_frame('Optimization')
    unclick_messages()
    select_right_panel('Optimization')

    lp_relaxation(True)

    # # Adjust Solver settings max time solve and min gap value
    z = 0
    p = 0
    for panel in WebDriverWait(driver, 5).until(EC.visibility_of_any_elements_located((By.XPATH, "//div[contains(@class, 'checkbox')]"))):
        p += 1
        if p == 1:
            for el in panel.find_elements(By.XPATH, "//input[contains(@class, 'form-control')]"):
                z += 1
                if z == 1:
                    el.clear()
                    el.send_keys(max_solve)
                if z == 2:
                    el.clear()
                    el.send_keys(str(min_gap))

    click_button('OPTIMIZE')
    go_to_job('OPTIMIZE', max_solve+1200, 60)

    driver.refresh()
    unclick_messages()

    navigate_to_frame('Results')
    driver.refresh()
    unclick_messages()
    select_right_panel('Results Export')
    click_button('POPULATE RESULTS')
    go_to_job('POPULATE_RESULTS', 600, 3)

    unclick_messages()

    select_right_panel('Results Export')
    time.sleep(5)

    # here can be added code to check if the AllCSVExports.zip file exists in the download folder
    openingZippedFile.check_if_file_exist(delete_zip=True)
    click_button('DOWNLOAD RESULTS')
    # if openingZippedFile.check_if_file_exist(delete_zip=False):
    #     openingZippedFile.move_file_to_new_directory(new_directory=(fico_results_archive
    #                                                                 + rf"\AllCSVExports_{scenario_name}_FirstRun.zip"))

    for k, band in enumerate(bands):
        i = 0
        if k == 0:
            i = 'first_run'
        else:
            i = bands[k - 1]
        if openingZippedFile.check_if_file_exist(delete_zip=False):
            openingZippedFile.move_file_to_new_directory(new_directory=(fico_results_archive
                                                                        + rf"\AllCSVExports_{scenario_name}_{i}.zip"))

        # if k == 4: exit()

        time.sleep(60)
        cut = cuts[k]
        if k > 0:
            update_lanes_values(scenario_name, bands[k-1])

        execute_sql_proc(cut, band, k, ISC_path)

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

        # added line if gap is too big
        gap_value = str.replace(str.replace(gap, 'Last Gap Value: ', ''), '%', '')
        if 'e' in gap_value:
            gap_value = 0
        else:
            gap_value = float(gap_value)

        if gap_value > min_gap and resolve:
            if k == 0:
                band_low = 0
            else:
                band_low = bands[k-1]

            update_bands(band_low, band)
            band_level1 = float(band_low + (band - band_low)/2)

            gap_split_band(cut, band_level1, 1, k)
            gap_split_band(cut, band, 2, k)
            update_lanes_values(scenario_name, band_level1)

        else:
            # datetime object containing current date and time
            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

            insert_outputs(scenario_name, band, dt_string, objective, gap, exec_time)

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
            # openingZippedFile.check_if_file_exist(delete_zip=True)
            click_button('DOWNLOAD RESULTS')

    if openingZippedFile.check_if_file_exist(delete_zip=False):
        openingZippedFile.move_file_to_new_directory(new_directory=(fico_results_archive
                                                     + rf"\AllCSVExports_{scenario_name}_Final.zip"))

    print("------------------------------------------------------------------")
    print("--- %s seconds ---" % (time.time() - start_time))
except Exception as exp:
    print(exp)
finally:
    log_out = driver.find_element(By.ID, 'user-dropdown')
    log_out.click()
    log_out = driver.find_element(By.ID, 'logout-link')
    log_out.click()
    time.sleep(3)
    driver.close()
    driver.quit()


