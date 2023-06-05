## 상품 크롤링 ##
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import pyautogui

options = Options() # 크롬 옵션 객체 생성
options.add_experimental_option("detach", True) # 크롬 창을 종료해도 프로세스가 종료되지 않도록 설정
service = Service(ChromeDriverManager().install())  # 크롬 드라이버 경로 설정
driver = webdriver.Chrome(service=service, options=options) # 크롬 드라이버 객체 생성

url = 'https://shopping.naver.com/home'
driver.maximize_window()
driver.get(url) # 사이트 연결하기
print(driver.title)     # 타이틀 읽어오기
driver.implicitly_wait(2) # 2초 대기하기

input_product = pyautogui.prompt("검색할 상품명 입력")

driver.find_element(By.CLASS_NAME, '_searchInput_search_text_3CUDs').send_keys(input_product)
driver.find_element(By.CLASS_NAME, '_searchInput_button_search_1n1aw').click()

driver.implicitly_wait(2) # 2초 대기하기
# 스크롤 전 높이
b_height = driver.execute_script("return window.scrollY")

# 무한 스크롤 처리
while True:
    try:
        #스크롤 내리기
        more = driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.END)
        time.sleep(1)

        after_height = driver.execute_script("return window.scrollY")

        if b_height == after_height:
            break

        b_height = after_height

    except:
        break

product = driver.find_elements(By.CLASS_NAME, 'product_item__MDtDF')

import pandas as pd

data = []

for i, t in enumerate(product):
  name = t.find_element(By.CSS_SELECTOR, 'div.product_title__Mmw2K').text
  price = t.find_element(By.CSS_SELECTOR, 'span.price_num__S2p_v').text 
  reviews = t.find_element(By.CSS_SELECTOR, 'em.product_num__fafe5').text + "개"
  link = t.find_element(By.CSS_SELECTOR, 'div.product_title__Mmw2K > a').get_attribute("href")
  print("크롤링중...")
  data.append([name, price, reviews, link])
  print("엑셀에 데이터 추가중...")

print( )
print("크롤링 완료!")
df = pd.DataFrame(data, columns=['제품명', '가격','리뷰', '링크'])
df.to_excel(input_product+'.xlsx', index=False)
print("엑셀 데이터 추가 완료!", "\n")

driver.quit()

## 엑셀 크기 조정 ##
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# 엑셀 파일 생성
workbook = Workbook()
worksheet = workbook.active

# 데이터프레임 내용을 엑셀에 저장
header = df.columns.tolist()
worksheet.append(header)
for _, row in df.iterrows():
    worksheet.append(row.tolist())

# 열의 폭 자동 조정
for column_cells in worksheet.columns:
    max_length = 0
    column = [cell for cell in column_cells]
    column_name = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2) * 1.2
    worksheet.column_dimensions[column_name].width = adjusted_width

# 엑셀 파일 저장
workbook.save(input_product+'.xlsx')
print("엑셀파일 저장!","\n")

## 메일 전송 ## 
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email import charset

# 한글 인코딩 설정
charset.add_charset('utf-8', charset.QP, charset.QP, 'utf-8')

smtp = smtplib.SMTP("smtp.naver.com", 587)
smtp.ehlo()
smtp.starttls()

input_ID = pyautogui.prompt("네이버 ID 입력")
input_PW = pyautogui.prompt("네이버 PW 입력")
smtp.login(input_ID, input_PW) # 네이버 아이디, 비밀번호

input_email = pyautogui.prompt("받는 사람 이메일")

me = 'quan0808@naver.com' # 보내는 이메일
you = input_email # 받는 이메일
subject = input_product + " 추천 목록"

input_message = pyautogui.prompt("본문 내용 입력")

message = input_message

msg = MIMEMultipart()
msg['Subject'] = Header(subject, 'utf-8')
msg['From'] = me
msg['To'] = you

text = MIMEText(message, 'plain', 'utf-8')
msg.attach(text)

# 엑셀 파일 첨부
file_name = input_product + ".xlsx"

with open(file_name, 'rb') as excel_file:
    attachment = MIMEApplication(excel_file.read())
    attachment.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', file_name))
    msg.attach(attachment)

# 메일 보내기
smtp.sendmail(me, you, msg.as_string())
smtp.quit()
print("이메일 전송 완료!!","\n")

## 메일 전송 확인 ##
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
url = 'http://naver.com'
driver.get(url)
driver.implicitly_wait(1)

elem = driver.find_element(by=By.CLASS_NAME, value="link_login")
elem.click()

