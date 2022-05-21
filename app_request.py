from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
from openpyxl import load_workbook


def kick_start(u_name, p_word):
    starting_time = time.perf_counter()
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    options.add_argument("--incognito")

    driver = webdriver.Chrome(executable_path='C:\drivers\chromedriver_win32\chromedriver.exe', options=options)
    # driver = webdriver.Firefox(executable_path='C:\drivers\geckodriver-v0.29.0-win64/geckodriver.exe')
    driver.get('http://portal.ui.edu.ng')  # FETCHES THE URL
    # driver.maximize_window()

    driver.find_element_by_xpath(
        '/html/body/div[2]/div/div[1]/div[1]/div/div[2]/a/div/p[2]').click()  # Goes to login # as existing student

    username = driver.find_element_by_name('UserName')
    password = driver.find_element_by_name('UserPassword1')
    username.send_keys(u_name)
    password.send_keys(p_word)
    driver.find_element_by_xpath('//*[@id="loginForm"]/table/tbody/tr[3]/td[2]/input').click()  # clicks login

    # THIS TRY-EXCEPT BLOCK ACCEPT ALERT ON A NORMAL LOGIN SITUATION
    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
    except TimeoutException:
        pass

    #  THIS IS TO HANDLE THE UPDATE PASSWORD PAGE WHEN THE SITE REQUESTS IT
    update_url = 'http://portal.ui.edu.ng/cportal/gc?Event=login'
    update_title = 'Global Portal - Change of Password'
    page_url = driver.current_url
    page_title = driver.title
    if page_url == update_url and page_title == update_title:
        old_password = driver.find_element_by_xpath('//*[@id="OldPassword"]')
        new_password = driver.find_element_by_xpath('//*[@id="NewPassword"]')
        confirm_new = driver.find_element_by_xpath('//*[@id="ConfirmPassword"]')

        old_password.send_keys(p_word)
        new_password.send_keys(p_word)
        confirm_new.send_keys(p_word)
        driver.find_element_by_xpath(
            '/html/body/table/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[2]/td/table/tbody/tr[6]/td[2]/input').click()  # UPDATES THE PASSWORD
        driver.switch_to.alert.accept()  # ACCEPT THE POP-UP ALERT

    # THIS TRY-BLOCK HANDLES PASSWORD ERROR
    # IT RETRIES TILL IT RESETS AND LOGS ITS DATA
    reset_page_title = driver.title
    reset_page_url = driver.current_url
    reset_title_check = 'University of Ibadan Portal'
    reset_url_check = 'http://portal.ui.edu.ng/cportal/gc?Event=login'
    if reset_page_title == reset_title_check and reset_page_url == reset_url_check:
        try:
            reseted = False
            while not reseted:
                class_name = driver.find_element_by_class_name('cpErrorText')
                xpath_name = driver.find_element_by_xpath('//*[@id="loginForm"]/table/tbody/tr[4]/td')
                if class_name.text == 'Your student status is not up-to-date':
                    with open('requested/specials/specials.txt', 'a') as rlog:
                        rlog.write(f'{u_name} \n')
                        rlog.close()
                        driver.close()
                        return

                elif class_name == xpath_name:
                    username = driver.find_element_by_name('UserName')
                    password = driver.find_element_by_name('UserPassword1')

                    username.send_keys(u_name)
                    password.send_keys(p_word)
                    driver.find_element_by_xpath('//*[@id="loginForm"]/table/tbody/tr[3]/td[2]/input').click()
                    continue

                elif class_name.text == 'Your student status is not up-to-date':
                    with open('requested/specials/specials.txt', 'a') as rlog:
                        rlog.write(f'{u_name} \n')
                        rlog.close()
                        driver.close()
                        return
                reseted = True

        except Exception as e:
            pass

            with open('requested/matric_reset/folder.txt', 'a') as log:
                log.write(f'{u_name} \n')
                log.close()
            driver.close()
            return
        return 
            # exit()

    driver.find_element_by_xpath(
        '//*[@id="modulePage643"]/table[2]/tbody/tr[2]/td[2]/fieldset[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/a').click()  # click personal detail

    matric_no = driver.find_element_by_xpath(
        '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[1]/tbody/tr[1]/td[2]/input').get_attribute(
        'value')
    l_name = driver.find_element_by_xpath(
        '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[2]/td[2]/input').get_attribute('value')
    f_name = driver.find_element_by_xpath(
        '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[1]/td[2]/input').get_attribute('value')

    # THIS RIGHT HERE CREATES THE FILE TO BE WRITTEN
    with open(f'requested/details/{matric_no}_{l_name}_{f_name}.txt', 'w') as o:
        o.close()

    rows = len(driver.find_elements_by_xpath('//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr'))
    # print(rows)

    # THIS HOLDS THE VALUE FOR THE LIST TO BE APPENDED TO THE EXCEL FILE
    excel_list = [f'{matric_no}']

    def personal_details():
        flagged_numbers = [20, 23, 24, 25]
        # flagged_numbers = [19, 22, 23, 24]  # THESE ARE THE ONES THAT ARE USED FOR STYLING THE LAYOUTS
        for r in range(1, rows + 1):
            if r in flagged_numbers:
                continue
            if r == 18 or r == 30:
                heading = driver.find_element_by_xpath(
                    '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[' + str(r) + ']/td[1]').text
                in_put = driver.find_element_by_xpath(
                    '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[' + str(r) + ']/td[2]/textarea')
                value = in_put.text
                with open(f'requested/details/{matric_no}_{l_name}_{f_name}.txt', 'a') as o:
                    o.write(f'{heading} {value} \n')
                    o.close()
                excel_list.append(str(value))
                # print(heading, '     ', value, '    ')
                continue

            heading = driver.find_element_by_xpath(
                '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[' + str(r) + ']/td[1]').text
            in_put = driver.find_element_by_xpath(
                '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr[' + str(r) + ']/td[2]/input')
            value = in_put.get_attribute('value')
            with open(f'requested/details/{matric_no}_{l_name}_{f_name}.txt', 'a') as o:
                o.write(f'{heading} {value} \n')
                o.close()
            excel_list.append(str(value))
            # print(heading, '     ', value, '    ')

    #  SECOND PART
    special_numbers = [6, 7, 8]  # THESES ONES USE TEXTAREA FOR INPUT
    second_rows = len(driver.find_elements_by_xpath(
        '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[1]/tbody/tr'))

    # print(second_rows)

    def second_part():
        for second_r in range(1, second_rows + 1):
            if second_r in special_numbers:
                second_heading = driver.find_element_by_xpath(
                    '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[1]/tbody/tr[' + str(
                        second_r) + ']/td[1]').text
                second_input = driver.find_element_by_xpath(
                    '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[1]/tbody/tr[' + str(
                        second_r) + ']/td[2]/textarea')
                second_value = second_input.text
                with open(f'requested/details/{matric_no}_{l_name}_{f_name}.txt', 'a') as o:
                    o.write(f'{second_heading} {second_value} \n')
                    o.close()
                excel_list.append(str(second_value))
                # print(second_heading, '     ', second_value, '    ')
                continue

            second_heading = driver.find_element_by_xpath(
                '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[1]/tbody/tr[' + str(
                    second_r) + ']/td[1]').text
            second_input = driver.find_element_by_xpath(
                '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[1]/tbody/tr[' + str(
                    second_r) + ']/td[2]/input')
            second_value = second_input.get_attribute('value')
            with open(f'requested/details/{matric_no}_{l_name}_{f_name}.txt', 'a') as o:
                o.write(f'{second_heading} {second_value} \n')
                o.close()
            excel_list.append(str(second_value))
            # print(second_heading, '     ', second_value, '    ')

    # SPONSOR PART
    sponsor_rows = len(driver.find_elements_by_xpath(
        '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[2]/tbody/tr'))

    # print(sponsor_rows)

    def sponsored_part():
        for sponsor_r in range(4, sponsor_rows + 1):
            if sponsor_r == 7:
                sponsored_heading = driver.find_element_by_xpath(
                    '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[2]/tbody/tr[' + str(
                        sponsor_r) + ']/td[1]').text
                sponsored_input = driver.find_element_by_xpath(
                    '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[2]/tbody/tr[' + str(
                        sponsor_r) + ']/td[2]/textarea')
                sponsored_value = sponsored_input.text
                with open(f'requested/details/{matric_no}_{l_name}_{f_name}.txt', 'a') as o:
                    o.write(f'{sponsored_heading} {sponsored_value} \n')
                    o.close()
                excel_list.append(str(sponsored_value))
                # print(sponsored_heading, '     ', sponsored_value, '    ')
                continue

            sponsored_heading = driver.find_element_by_xpath(
                '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[2]/tbody/tr[' + str(
                    sponsor_r) + ']/td[1]').text
            sponsored_input = driver.find_element_by_xpath(
                '//*[@id="modulePage432"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table[2]/tbody/tr[' + str(
                    sponsor_r) + ']/td[2]/input')
            sponsored_value = sponsored_input.get_attribute('value')
            with open(f'requested/details/{matric_no}_{l_name}_{f_name}.txt', 'a') as o:
                o.write(f'{sponsored_heading} {sponsored_value} \n')
                o.close()
            excel_list.append(str(sponsored_value))
            # print(sponsored_heading, '     ', sponsored_value, '    ')

    def final():
        # HERE WE LOAD THE WORKBOOK THEN APPEND THE LIST TO IT AND THEN SAVE
        wb = load_workbook('requested/Excel/details.xlsx')
        ws = wb.active
        ws.append(excel_list)
        wb.save('requested/Excel/details.xlsx')
        ending_time = time.perf_counter()
        print(ending_time - starting_time)
        driver.quit()

    def main():
        personal_details()
        second_part()
        sponsored_part()
        final()

    main()


# if __name__ == '__main__':
#     while True:
#         mats = input("Which profile to fetch ? >> ")
#         kick_start(str(mats), str(mats))
#         print('---------------Done-------------------')
