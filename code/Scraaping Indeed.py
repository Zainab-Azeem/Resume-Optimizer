from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import logging

def setup_driver():
    options = Options()
    options.headless = True
    service = Service('C:/Users/Sher Ali/Downloads/chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    logging.info("Initialized ChromeDriver successfully.")
    return driver

def scrape_job_details(url):
    driver = setup_driver()
    
    try:
        driver.get(url)
        logging.info(f"Opened URL: {url}")
        time.sleep(5)

        job_details = {}
        
        try:
            job_details['Job Title'] = driver.find_element(By.XPATH, '//span[contains(@class, "jobsearch-JobInfoHeader-title")]').text
            logging.info(f"Job Title: {job_details['Job Title']}")
            
            job_details['Company Name'] = driver.find_element(By.ID, 'companyLink').text
            logging.info(f"Company Name: {job_details['Company Name']}")
            
            job_details['Location'] = driver.find_element(By.ID, 'location-collapsed-header').text
            logging.info(f"Location: {job_details['Location']}")
            
            job_details['Salary'] = driver.find_element(By.XPATH, '//span[contains(@class, "css-19j1a75")]').text
            logging.info(f"Salary: {job_details['Salary']}")
            
            job_details['Job Type'] = driver.find_element(By.XPATH, '//p[contains(text(), "Job Type")]').text.split(': ')[1]
            logging.info(f"Job Type: {job_details['Job Type']}")
            
            job_description_element = driver.find_element(By.XPATH, '//h2[contains(@id, "jobDescriptionTitleHeading")]/following-sibling::div')
            job_details['Job Description'] = job_description_element.text if job_description_element else "Not Available"
            logging.info("Job Description retrieved.")
        
        except Exception as e:
            logging.error(f"Error occurred while scraping job details: {e}")
            job_details = None

    except Exception as e:
        logging.error(f"Error occurred while opening the URL: {e}")
        job_details = None
    
    finally:
        driver.quit()
        logging.info("ChromeDriver closed.")

    return job_details

def save_to_excel(job_details):
    if job_details:
        df = pd.DataFrame([job_details])
        output_file = "job_details.xlsx"
        
        try:
            df.to_excel(output_file, index=False)
            logging.info(f"Job details saved to {output_file}")
            
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                for cell in worksheet["1:1"]:
                    cell.font = cell.font.copy(bold=True)
                    cell.alignment = cell.alignment.copy(horizontal='center')

                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except Exception as e:
                            logging.error(f"Error adjusting column width: {e}")
                    adjusted_width = max_length + 2
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
                
            logging.info("Excel file formatted successfully.")
        except Exception as e:
            logging.error(f"Failed to save job details to Excel: {e}")
    else:
        logging.warning("No job details to save.")

if __name__ == "__main__":
    job_url = "https://pk.indeed.com/jobs?q=web+developer&l=Pakistan&from=searchOnDesktopSerp&vjk=454324cde9c2e255"
    job_details = scrape_job_details(job_url)
    save_to_excel(job_details)
