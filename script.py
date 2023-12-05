from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import os

os.chdir('C:/Users/LENOVO/Documents/learn 000x1-stage/githubDocs/bulk-msg-whatsapp-selenium')

# change the name of file and sheet_name

excel_data = pandas.read_excel('testtt.xlsx', sheet_name='Sheet1')

count = 0

msg = "Vous envisagez des études à l'étranger ? Explorez des opportunités passionnantes avec EPS Education. 🌍✨ Basés à Marrakech, Agadir, Rabat et Casablanca. 💼 Pour une consultation gratuite, contactez-nous pour plus d'informations. 📚🎓 Nous sommes là pour vous aider !   Considering studying abroad? Explore exciting opportunities with EPS Education. 🌍✨ Based in Marrakech, Agadir, Rabat, and Casablanca. 💼 For a free consultation, contact us for more information. 📚🎓 We're here to help! 🌟        هل تفكر في الدراسة في الخارج؟ استكشف الفرص المثيرة مع EPS Education. 🌍✨ مقرنا في مراكش، أكادير، الرباط والدار البيضاء. 💼 للحصول على استشارة مجانية، اتصل بنا للمزيد من المعلومات. 📚🎓 نحن هنا للمساعدة !🌟                     ---------------- 𝗖𝗼𝗻𝘁𝗮𝗰𝘁𝗲𝘇 𝗻𝗼𝘂𝘀 𝗱𝗲́𝘀 𝗺𝗮𝗶𝗻𝘁𝗲𝗻𝗮𝗻𝘁 ! 📲 WhatsApp : +212 702-099988 🌐 Ou inscrivez-vous directement sur : https://www.epseducation.ma/inscription/ 💻"

# you can change the URL of the image 
image_path = 'C:/Users/LENOVO/Documents/learn 000x1-stage/stage01/test3/pic.jpeg'


chrome_options = webdriver.ChromeOptions()
chrome_driver_path = ChromeDriverManager().install()

driver = webdriver.Chrome(options=chrome_options)

driver.get('https://web.whatsapp.com')
input("Press ENTER after login into Whatsapp Web and your chats are visiable.")
for column in excel_data['contact'].tolist():
    try:
        # you can change the the title of column and the message you wanna send

        url = 'https://web.whatsapp.com/send?phone={}&text={}'.format(excel_data['contact'][count], msg)
        sent = False
        print('Navigating to URL:', url)
        driver.get(url)

        attachment_box = WebDriverWait(driver, 50).until(
        EC.element_to_be_clickable((By.XPATH, '//span[contains(@data-icon, "attach-menu-plus")]')))
        attachment_box.click()
        image_box = driver.find_element(By.XPATH,'//input[contains(@accept, "image/*,video/mp4,video/3gpp,video/quicktime")]')


        image_box.send_keys(image_path)
        

        try:
            send_button = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.XPATH,'//span[contains(@data-icon, "send")]')))
        except Exception as e:
            # change the column's name
            print("Sorry message could not sent to " + str(excel_data['contact'][count]))
        else:
            sleep(2)
            actions = ActionChains(driver)
            actions.move_to_element(send_button).click().perform()
            sent = True
            sleep(5)
            print('Message sent to: ' + str(excel_data['contact'][count]))


    except Exception as e:
        # change the column's name
        print('Failed to send message to ' + str(excel_data['contact'][count]) + " " + str(e))

    finally:
        count = count + 1
driver.quit()
print("The script executed successfully.")