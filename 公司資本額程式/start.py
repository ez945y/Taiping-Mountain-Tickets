from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from time import sleep

header = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/12.0.3 Safari/605.1.15"
ser = Service('.\chromedriver.exe')
op = webdriver.ChromeOptions()
op.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
op.add_experimental_option('useAutomationExtension', False)
op.add_experimental_option("prefs", {"profile.password_manager_enabled": False, "credentials_enable_service": False})
driver=webdriver.Chrome(service=ser,options=op)
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
  "source": """
    Object.defineProperty(navigator, 'webdriver', {
      get: () => undefined
    })
  """
})
driver.maximize_window()


with open("./number.csv",encoding="utf-8") as f:
    df=pd.read_csv(f)

list1 = []
list2 = []
list3 = []
n_all = 0
n_fault = 0
length = df.shape[0]
print()
print('      開始作業')
print('------',n_all + 1,"/",length,'------') 
for number in df['統編號碼']:
    try:
      driver.get('https://findbiz.nat.gov.tw/fts/query/QueryList/queryList.do') #開啟網站
      sleep(2)
      if n_all == 0:
        driver.find_element("xpath",'//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[3]').click()
        driver.find_element("xpath",'//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[5]').click()
        driver.find_element("xpath",'//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[9]').click()
      driver.find_element("xpath",'//*[@id="qryCond"]').send_keys(number)
      driver.find_element("xpath",'//*[@id="qryBtn"]').click()
      sleep(1)
      
      title = driver.find_element("xpath",'//*[@id="vParagraph"]/div/div[1]/a')
      print(title.text)
      list1.append(title.text)
      title.click()
      sleep(0.5)
      try:
        cap = driver.find_element("xpath",'//*[@id="tabCmpyContent"]/div/table/tbody/tr[5]/td[1]')
        if cap.text == '資本總額(元)':
          cap_val = driver.find_element("xpath",'//*[@id="tabCmpyContent"]/div/table/tbody/tr[5]/td[2]')
          list2.append(cap_val.text)
          print('資本額(元) : ', cap_val.text)
        elif cap.text == '章程所訂外文公司名稱':
          cap_val = driver.find_element("xpath",'//*[@id="tabCmpyContent"]/div/table/tbody/tr[6]/td[2]')
          list2.append(cap_val.text)
          print('資本額(元) : ', cap_val.text)
        else:
          list2.append("查無")
          print('資本額(元) : ', "查無")
      except:
        cap = driver.find_element("xpath",'//*[@id="tabBusmContent"]/div/table/tbody/tr[8]/td[1]')
        if cap.text == '資本額(元)':
          cap_val = driver.find_element("xpath",'//*[@id="tabBusmContent"]/div/table/tbody/tr[8]/td[2]')
          list2.append(cap_val.text)
          print('資本額(元) : ', cap_val.text)
        else:
          list2.append("查無")
          print('資本額(元) : ', "查無")

      try:
        real = driver.find_element("xpath",'//*[@id="tabCmpyContent"]/div/table/tbody/tr[6]/td[1]')
        if real.text == '實收資本額(元)':
          real_val = driver.find_element("xpath",'//*[@id="tabCmpyContent"]/div/table/tbody/tr[6]/td[2]')
          list3.append(real_val.text)
          print('實收資本額(元) : ', real_val.text)
        elif cap.text == '章程所訂外文公司名稱':
          real_val = driver.find_element("xpath",'//*[@id="tabCmpyContent"]/div/table/tbody/tr[7]/td[2]')
          list3.append(real_val.text)
          print('實收資本額(元) : ', real_val.text)
        else:
          list3.append("查無")
          print('實收資本額(元) : ', "查無")

      except:
        cap = driver.find_element("xpath",'//*[@id="tabBusmContent"]/div/table/tbody/tr[9]/td[1]')
        if cap.text == '實收資本額(元)':
          cap_val = driver.find_element("xpath",'//*[@id="tabBusmContent"]/div/table/tbody/tr[9]/td[2]')
          list3.append(real_val.text)
          print('實收資本額(元) : ', real_val.text)
        else:
          list3.append("查無")
          print('實收資本額(元) : ', "查無")

      n_all += 1
      if n_all == 10:
        print('------------')
        break 

      print('------',n_all + 1,"/",length,'------') 
      sleep(2)
      
    except:
      n_all += 1
      n_fault += 1

      print("      查無資料")
      try:
        if driver.find_element("xpath",'//*[@id="tabFactContent"]/h3').text == '工廠基本資料':
          list2.append("查無")
          list3.append("查無")
      except:
        list1.append("查無")
        list2.append("查無")
        list3.append("查無")

      sleep(2)
      if n_all == 10:
        print('------------')
        break 

      print('------',n_all + 1,"/",length,'------') 
      continue

print('已完成作業, Total : ', n_all, ' 查無數: ', n_fault)
df['公司名稱'] = pd.Series(list1)
df['資本總額(元)'] = pd.Series(list2)
df['實收資本額(元)'] = pd.Series(list3)

df.to_csv("work_coumplete.csv", index = False)

