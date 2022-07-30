from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep

with open('info.txt','r',encoding="utf-8") as f:
  list=[]
  line ='firstline'
  while line:
        line = f.readline()
        line = line.strip()
        if line != '':
          list.append(line)
          
phone = list[0]
id =list[1]
visa = list[2]
name = list[3]
birth = list[4]
email = list[5]
mode = int(list[6])

calendar = int(list[7])
row =int(list[8])
col=int(list[9])

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

driver.get('https://tpsr.forest.gov.tw/TPSOrder/wSite/index.do?action=indexPage#')
locator = (By.CSS_SELECTOR,'#show')
WebDriverWait(driver, 1).until(EC.presence_of_all_elements_located(locator))
driver.find_element_by_xpath('//*[@id="show"]/div[1]/a/img').click()#關掉公告


img_base64 = browser.execute_script("""
    var ele = arguments[0];
    var cnv = document.createElement('canvas');
    cnv.width = ele.width; cnv.height = ele.height;
    cnv.getContext('2d').drawImage(ele, 0, 0);
    return cnv.toDataURL('image/jpeg').substring(22);    
    """, driver.find_element_by_xpath("//*[@id="captchaImg"]"))
 
with open("captcha_login.png", 'wb') as image:
    image.write(base64.b64decode(img_base64))

file = {'file': open('captcha_login.png', 'rb')}  #下載下來的一般驗證碼(Normal Captcha)圖片
 
api_key = '37eff46c56aad756a746e751942346a6'
data = {
    'key': api_key,
    'method': 'post'
}
 
response = requests.post('http://2captcha.com/in.php', files=file, data=data)
print(f'response:{response.text}')
 
if response.ok and response.text.find('OK') > -1:
 
    captcha_id = response.text.split('|')[1]  #擷取驗證碼ID
 
    for i in range(10):
        
        response = requests.get(
            f'http://2captcha.com/res.php?key={api_key}&action=get&id={captcha_id}')
        print(f'response:{response.text}')
 
        if response.text.find('CAPCHA_NOT_READY') > -1:  #尚未辨識完成
            time.sleep(3)
 
        elif response.text.find('OK') > -1:
            captcha_text = response.text.split('|')[1]  #擷取辨識結果
            
            captcha = driver.find_element_by_xpath('//*[@id="captcha"]')
            login = driver.find_element_by_xpath('//*[@id="signForm"]/div[2]/input[1]')

            captcha.send_keys(captcha_text)  #輸入剛剛擷取的一般驗證碼辨識結果
            login.click()  #點擊登入按鈕
 
            break
 
        else:
            print('取得驗證碼發生錯誤!')
else:
    print('提交驗證碼發生錯誤!')

locator = (By.CSS_SELECTOR,'body > div > div.content > div:nth-child(3) > ul > li:nth-child(1) > form > div > img')
WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(locator))
driver.execute_script(f'window.scrollTo(0,200)')
driver.find_element_by_xpath(f'/html/body/div/div[2]/div[3]/ul/li[{mode}]/form/input[1]').click()#訂房
'''/html/body/div/div[2]/div[3]/ul/li[2]/form/input[1]'''
locator = (By.CSS_SELECTOR,'body > div > div.content > div.list01 > ul')
WebDriverWait(driver, 1).until(EC.presence_of_all_elements_located(locator))

driver.execute_script(f'window.scrollTo(0,800)')
locator = (By.CSS_SELECTOR,'#calendar3 > tbody > tr:nth-child(4) > td:nth-child(2) > span')
'''//*[@id="calendar3"]/tbody/tr[3]/td[7]/span'''
WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(locator))
driver.find_element_by_xpath(f'//*[@id="calendar{calendar}"]/tbody/tr[{row}]/td[{col}]/span/a').click()
locator = (By.CSS_SELECTOR,'body > div > div.content > div.listTb5')
WebDriverWait(driver, 1).until(EC.presence_of_all_elements_located(locator))
driver.find_element_by_css_selector('#htx_iphone').send_keys(phone)
driver.find_element_by_xpath('//*[@id="agree"]').click()
driver.find_element_by_xpath('//*[@id="form1"]/div[2]/input[1]').click()


locator = (By.CSS_SELECTOR,'#form1 > p')
WebDriverWait(driver, 1).until(EC.presence_of_all_elements_located(locator))
driver.find_element_by_css_selector('#htx_idnumber').send_keys(id)
driver.find_element_by_css_selector('#htx_passport').send_keys(visa)
driver.find_element_by_css_selector('#htx_name').send_keys(name)
driver.find_element_by_css_selector('#htx_email').send_keys(email)

js = 'document.getElementById("htx_birthday_display").removeAttribute("readonly")'
driver.execute_script(js)
driver.find_element_by_css_selector('#htx_birthday_display').clear()
driver.find_element_by_css_selector('#htx_birthday_display').send_keys(birth)

driver.execute_script(f'window.scrollTo(0,200)')
#driver.find_element_by_xpath('//*[@id="form1"]/div[2]/input[1]').click()