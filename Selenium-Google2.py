from selenium import webdriver
import urllib.request
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import sys
import _socket
import unittest, time, re
import xlrd
import xlwt
import unittest, time
a = xlwt.Workbook()
sheet = a.add_sheet('sheet1',cell_overwrite_ok=True)
workbook = xlrd.open_workbook("Test_EXCEL_Ref.xlsx")
worksheet = workbook.sheet_by_index(0)
driver = webdriver.Chrome()
i = 0
while (i < 5):

    x = i
    xx = 1
    col = worksheet.col_values(x)
    driver.get("https://google.com/")
    input_element = driver.find_element_by_name("q")
    input_element.send_keys(col[0]+ " dividend 2016")
    input_element.submit()
    driver.find_element_by_name("btnG").click()
    time.sleep(3)
    sheet.write(0, i, col[0])
    RESULTS_LOCATOR1 = "(//div/h3/a)[1]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(xx,i,item1.text)
    xx = xx + 1
    #RESULTS_LOCATOR2 = "(//div/cite[contains(@class,'_Rm')])[1]"

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[1]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[1]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    xx = xx + 1

    RESULTS_LOCATOR1 = "(//div/h3/a)[1]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[2]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1


    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[2]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[2]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    xx = xx + 1

    RESULTS_LOCATOR1 = "(//div/h3/a)[2]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[3]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[3]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[3]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    RESULTS_LOCATOR1 = "(//div/h3/a)[3]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[4]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[4]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[4]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    RESULTS_LOCATOR1 = "(//div/h3/a)[4]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[5]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[5]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[5]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    RESULTS_LOCATOR1 = "(//div/h3/a)[5]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[6]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[6]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[6]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    RESULTS_LOCATOR1 = "(//div/h3/a)[6]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[7]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[7]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[7]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    RESULTS_LOCATOR1 = "(//div/h3/a)[7]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[8]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[8]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[8]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    RESULTS_LOCATOR1 = "(//div/h3/a)[8]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[9]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[9]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)

    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[9]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)

    RESULTS_LOCATOR1 = "(//div/h3/a)[9]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    RESULTS_LOCATOR1 = "(//div/h3/a)[10]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR1)))

    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item1 in page1_results1:

        print(item1.text)
        sheet.write(0, i, col[0])
        sheet.write(xx,i,item1.text)
    xx = xx + 1

    links = driver.find_elements_by_xpath("(//h3[@class='r']/a[@href])[10]")
    results = []
    for link in links:
        url = link.get_attribute('href')
        results.append(url)
    print(results)
    sheet.write(xx, i, results)


    #RESULTS_LOCATOR2 = "(//div/cite[contains(@class,'_Rm')])[10]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR2)))

    #page1_results2 = driver.find_elements(By.XPATH, RESULTS_LOCATOR2)

    #for item2 in page1_results2:

        #print(item2.text)
        #sheet.write(xx,i,item2.text)
    xx = xx + 1
    RESULTS_LOCATOR3 = "(//span[@class='st'])[10]"

    #WebDriverWait(driver, 10).until(
        #EC.visibility_of_element_located((By.XPATH, RESULTS_LOCATOR3)))

    page1_results3 = driver.find_elements(By.XPATH, RESULTS_LOCATOR3)

    for item3 in page1_results3:

        print(item3.text)
        sheet.write(xx,i,item3.text)


    RESULTS_LOCATOR1 = "(//div/h3/a)[10]"
    page1_results1 = driver.find_elements(By.XPATH, RESULTS_LOCATOR1)

    for item in page1_results1:
        item.click()
        time.sleep(2)

    try:
        data = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(data,'html.parser')
        texts = soup.find_all('p')
        textss = str(texts)
    except:
        print('No Results')
        sheet.write(xx, i, 'No Results')
        driver.execute_script("window.history.go(-1)")
        time.sleep(2)
        a.save("google_test.xls")
        xx = xx + 2
        pass
    else:
        try:
            print(textss)
            sheet.write(xx, i, textss)
            driver.execute_script("window.history.go(-1)")
            time.sleep(2)
            a.save("google_test.xls")
        except:
            print('File Too Large')
            xx = xx + 2
            pass
        else:
            xx = xx + 2

    time.sleep(1)
    i = i + 1
    a.save("google_test.xls")
    if (i > 3):
        break


