from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

class GmailAutomation:
    def __init__(self, driver_path):
        self.driver = webdriver.Chrome(service=Service(driver_path))

    def login(self, email, password):
        self.driver.get("https://mail.google.com")
        WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.ID, "identifierId")))
        self.driver.find_element(By.ID, "identifierId").send_keys(email)
        self.driver.find_element(By.ID, "identifierNext").click()
        WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.NAME, "Passwd")))
        self.driver.find_element(By.NAME, "Passwd").send_keys(password)
        self.driver.find_element(By.ID, "passwordNext").click()
        WebDriverWait(self.driver, 15).until(EC.url_contains("inbox"))
        return "inbox" in self.driver.current_url.lower()

    def send_email(self, recipient, subject, body, attachment_path=None):
        try:
            WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[gh='cm']"))).click()
            to_field = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='To recipients']")))
            to_field.send_keys(recipient)

            try:
                dropdown = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.aXS[data-name='" + recipient + "']")))
                dropdown.click()
            except:
                print("No dropdown found for recipient")

            subject_field = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.NAME, "subjectbox")))
            subject_field.click() 
            subject_field.clear()  
            subject_field.send_keys(subject)

            body_field = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[aria-label='Message Body']")))
            body_field.click()  
            body_field.send_keys(body)

            if attachment_path:
                attachment_input = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']")))
                attachment_input.send_keys(attachment_path)
                time.sleep(2)

            send_button = WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Send ‪(Ctrl-Enter)‬']")))

            self.driver.execute_script("arguments[0].click();", send_button)
            time.sleep(2)
        except Exception as e:
            print(f"Error while sending email: {e}")

    def logout_and_close(self):
        try:
            profile_icon = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[aria-label*='Google Account:']"))
            )
            profile_icon.click()

            for attempt in range(3):  
                try:
                    sign_out_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "div.SedFmc"))
                    )
                    self.driver.execute_script("arguments[0].click();", sign_out_button)
                    break
                except:
                    time.sleep(2)
        except Exception as e:
            print(f"Error during logout: {e}")
        finally:
            self.driver.quit()
