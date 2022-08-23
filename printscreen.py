import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import win32com.client

outlook = win32com.client.Dispatch('outlook.application')

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get('http:xxxx.com')
driver.save_screenshot('screenshot.png')
driver.find_element("name", "#login-username").send_keys('someusername')
time.sleep(5)
driver.find_element("name", "#login-password").send_keys('somepassword')
time.sleep(10)

driver.find_element("name", "form-control").submit()
driver.save_screenshot('signIn.png')

mail = outlook.CreateItem(0)
mail.To = 'somemail@gmail.com'
mail.Subject = 'Test'
mail.HTMLBody = '<h3>This is HTML Body</h3>'
mail.Body = "This is the normal body"
mail.Attachments.Add('C:\\location\\screenshot.png')
mail.Attachments.Add('C:\\location.png')
mail.Send()












