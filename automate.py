import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


def login(browser, login_link, user_name, password):
    browser.get(login_link)
    browser.save_screenshot("login.png")
    username_box = browser.find_element_by_name("username")
    password_box = browser.find_element_by_name("password")
    username_box.send_keys(user_name)
    password_box.send_keys(password)
    browser.find_element_by_css_selector("input[type='submit']").submit()
    print('Successfully Logged in')
    return True


def automate(
        browser, request_url, dsm_orders_url,intermidiate_req,
        etopup_order_url,
        push_button,
        filter_button,
        completed_button_1,
        completed_button_2,
        search_button,
        auto_refill_button_1,
        auto_refill_button_2,
        rows_button_1,
        rows_button_2,
        driver
):
    browser.get(intermidiate_req)
    time.sleep(1)
    browser.get(dsm_orders_url)
    time.sleep(1)
    browser.get(etopup_order_url)
    time.sleep(5)
    browser.find_element_by_css_selector(push_button).click()
    print('Push')
    time.sleep(5)
    if browser.find_elements_by_xpath(filter_button):
        browser.find_element_by_xpath(filter_button).click()
    print('Filter')
    browser.find_element_by_xpath(completed_button_1).click()
    time.sleep(2)
    browser.find_element_by_xpath(completed_button_2).click()
    print('Filter:Completed')
    browser.find_element_by_xpath(search_button).click()
    print('Filter:Searched')
    time.sleep(5)
    browser.find_element_by_xpath("//html").click()

    WebDriverWait(driver, 300).until(EC.visibility_of_element_located((By.XPATH,"//*[@id='root']/div/div/div/main/div"
                                                                                "/main/div/div[2]/div/div/div["
                                                                                "2]/div/div/div/table/tbody/tr[3]/td["
                                                                                "6]")))
    browser.find_element_by_xpath(auto_refill_button_1).click()
    browser.find_element_by_xpath(auto_refill_button_2).click()
    print('AutoRefill Filtered')
    browser.find_element_by_xpath("//html").click()
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    browser.find_element_by_xpath(rows_button_1).click()
    browser.find_element_by_xpath(rows_button_2).click()
    print('Got The Final Table')
    return True
