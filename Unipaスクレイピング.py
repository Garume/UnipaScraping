from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from xlwt import Workbook
import time
import chromedriver_binary

URL = 'https://portal.sa.dendai.ac.jp/up/faces/login/Com00505A.jsp'
ID = input('ID')
PASS = input('PASS')
YOBI = input("何曜日? 0:月 1:火 2:水 3:木 4:金 5:土 6:日")
JIGEN = str(int(input("何時限目? 1~8")))


#ブラウザ起動
options = Options()
options.binary_location = "C:\Program Files\Google\Chrome\Application\chrome.exe"
#options.add_argument("--headless")
driver = webdriver.Chrome(options=options)
driver.get(URL)
wait = WebDriverWait(driver, 10)

#Excel初期化
wb = Workbook()
ws = wb.add_sheet("YOBI="+YOBI+" JIGEM="+JIGEN)
ws.write(0,0,"YOBI="+YOBI+"JIGEM="+JIGEN)
ws.write(0,1,"授業コード")
ws.write(0,2,"授業名")
ws.write(0,3,"英文名")
ws.write(0,4,"開講年度学期")
ws.write(0,5,"単位")
ws.write(0,6,"教室")
ws.write(0,7,"担当教員")
ws.write(0,8,"授業形式")
ws.write(0,9,"遠隔授業方法")
ws.write(0,10,"目的概要")
ws.write(0,11,"達成目標")
ws.write(0,12,"関連科目")
ws.write(0,13,"履修条件")
ws.write(0,14,"教科書名")
ws.write(0,15,"参考署名")
ws.write(0,16,"評価方法")
ws.write(0,17,"学習教育目標との対応")
ws.write(0,18,"DPとのの対応")
ws.write(0,19,"アクティブラーニングの実施")
ws.write(0,20,"ICTの活用")
ws.write(0,21,"実践的活用")
ws.write(0,22,"自由記載欄")
ws.write(0,23,"自由記載欄")
ws.write(0,24,"第1回")
ws.write(0,25,"第2回")
ws.write(0,26,"第3回")
ws.write(0,27,"第4回")
ws.write(0,28,"第5回")
ws.write(0,29,"第6回")
ws.write(0,30,"第7回")
ws.write(0,31,"第8回")
ws.write(0,32,"第9回")
ws.write(0,33,"第10回")
ws.write(0,34,"第11回")
ws.write(0,35,"第12回")
ws.write(0,36,"第13回")
ws.write(0,37,"第14回")
ws.write(0,38,"Email")
ws.write(0,39,"質問への対応")
ws.write(0,40,"履修上の注意事項 クラス分け")
ws.write(0,41,"履修上の注意事項　ガイダンス")
ws.write(0,42,"学習上の助言")



#ログイン
driver.find_element_by_id('loginForm:userId').send_keys(ID)
driver.find_element_by_id('loginForm:password').send_keys(PASS)
driver.find_element_by_id('loginForm:loginButton').click()

print('ログイン成功')
#actions = ActionChains(driver)
#actions.move_to_element(driver.find_element_by_class_name("ui-menuitem-text")).perform()
# men
#u = driver.find_element_by_id('menu5')
# driver.getMouse().mouseMove(menu)
wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="menuForm:mainMenu"]/ul/li[4]'))).click()
driver.find_element_by_xpath('//*[@id="menuForm:mainMenu"]/ul/li[4]/ul/table/tbody/tr/td[1]/ul/li[2]').click()

print('シラバス検索ページ')
Syobi = str(2*int(YOBI)+1)
Sjigen = str(2*(int(JIGEN)-1)+1)
Xyobi = '//*[@id="funcForm:yobiList"]/tbody/tr/td['+Syobi+']/div/div[2]'
Xjigen = '//*[@id="funcForm:jigenList"]/tbody/tr/td['+Sjigen+']/div/div[2]'
print(Xyobi)
print(Xjigen)
driver.find_element_by_xpath(Xyobi).click()
driver.find_element_by_xpath(Xjigen).click()
#selector1 = Select(driver.find_element_by_id('funcForm:yobiList'))
#selector2 = Select(driver.find_element_by_id('funcForm:jigenList'))
#selector1.select_by_value(YOBI)
#selector2.select_by_value(JIGEN)
driver.find_element_by_xpath('//*[@id="funcForm:search"]').click()

print('検索完了ページ')
xpath = '//*[@id="funcForm:table_paginator_bottom"]/span[1]'
trs = wait.until(EC.visibility_of_element_located((By.XPATH,xpath)))
print(trs.text)
Strs = input("何件抽出しますか？")
clasnum=0
for i in range(int(Strs)):
    try:
        clasnum+=1;
        print("funcForm:"+str(i)+":jugyoKmkName")
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="funcForm:table:'+str(i)+':jugyoKmkName"]'))).click()
        print("nowCaptureing"+str(i))
        #//*[@id="pkx02301:ch:table"]/div[3]/div[2]/div 授業コード
        #//*[@id="pkx02301:ch:table"]/div[4]/div[2]
        #//*[@id="pkx02301:ch:table"]/div[3]/div[4]
        #//*[@id="pkx02301:ch:table"]/div[29]/div[2]/div/div/div目的概要
        #//*[@id="pkx02301:ch:table"]/div[11]/div[2]/div
        #//*[@id="pkx02301:ch:table"]/div[43]/div[2]/div
        #//*[@id="pkx02301:ch:table"]/div[47]/div[2]/div
        timewait = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="pkx02301:ch:table"]/div[3]/div[2]/div')))
        print(timewait.text)
        number = 0
        for i in range(3,50,1):
            number+=1
            try:
                atd = driver.find_element_by_xpath('//*[@id="pkx02301:ch:table"]/div[' + str(i) +']/div[2]/div')
                print(atd.text)
                ws.write(clasnum,number,atd.text)
            except:
                if number == 10 or 11:
                    number -= 1
                    pass
                else:
                    ws.write(clasnum,number,"")
        driver.find_element_by_xpath('//*[@id="pkx02301:dialog"]/div[1]/a[1]/span').click()
    except:
        pass
wb.save("YOBI="+YOBI+" JIGEM="+JIGEN+".xls")
driver.quit()