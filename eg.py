inputExcelFile = "D:\pythonProject\UN_NGO.xlsx"
newWorkbook = openpyxl.load_workbook(inputExcelFile)
driver = webdriver.Chrome('chromedriver')
driver.get('https://esango.un.org/civilsociety/displayConsultativeStatusSearch.do?method=list&show=5059')
driver.maximize_window()
sleep(2)
state = driver.find_element("xpath","/html/body/div[9]/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr[2]/td/table/tbody/tr[5]/td/ol/li[29]/a")
state.click()
sleep(2)
for pages in range(60,116):
    newWorkbook.create_sheet(f'Sheet{pages}')
    newWorkbook.save(inputExcelFile)
    state_name = f"https://ngodarpan.gov.in/index.php/home/statewise_ngo/11524/19/{pages}?per_page=100"
    driver.get(state_name)
    count_names = Select(driver.find_element("xpath", "/html/body/div[9]/div[1]/div[3]/div/div/div[1]/form/div[2]/select"))
    count_names.select_by_index(3)
    sleep(2)
    firstWorksheet = newWorkbook[f"Sheet{pages}"]
    for names in range(1,101):
        name = driver.find_element("xpath",f"/html/body/div[9]/div[1]/div[3]/div/div/div[2]/table/tbody/tr[{names}]/td[2]/a")
        driver.execute_script("arguments[0].scrollIntoView();", name)
        firstWorksheet.cell(row=names+1, column=2, value=name.text)
        regi_num = driver.find_element("xpath",f"/html/body/div[9]/div[1]/div[3]/div/div/div[2]/table/tbody/tr[{names}]/td[3]")
        driver.execute_script("arguments[0].scrollIntoView();", regi_num)
        firstWorksheet.cell(row=names+1, column=3, value=regi_num.text)
        name.click()
        sleep(5)
        mobile_num = driver.find_element("xpath","/html/body/div[9]/div[3]/div[2]/div/div[2]/div/div/table[8]/tbody/tr[5]/td[2]")
        driver.execute_script("arguments[0].scrollIntoView();", mobile_num)
        firstWorksheet.cell(row=names+1, column=4, value=mobile_num.text)
        email_id = driver.find_element("xpath","/html/body/div[9]/div[3]/div[2]/div/div[2]/div/div/table[8]/tbody/tr[7]/td[2]")
        driver.execute_script("arguments[0].scrollIntoView();", email_id)
        firstWorksheet.cell(row=names+1, column=5, value=email_id.text)
        web_URL = driver.find_element("xpath","/html/body/div[9]/div[3]/div[2]/div/div[2]/div/div/table[8]/tbody/tr[6]/td[2]")
        driver.execute_script("arguments[0].scrollIntoView();", web_URL)
        firstWorksheet.cell(row=names+1, column=6, value=web_URL.text)
        print(f"Page {pages} in {names}. Completed! ({regi_num.text})")
        driver.refresh()
        sleep(2)
        newWorkbook.save(inputExcelFile)
        sleep(2)
print("West Bengal State Completed!")
driver.close()