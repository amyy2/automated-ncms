from selenium.webdriver.common.by import By
from selenium import webdriver
 
def extract_info(driver, dispo, AC):
    '''
    Extracts an HTML file containing all relevant NCMS information from the NCMS extract website
 
        Args:
            driver (WebDriver): pointer to Selenium WebDriver object defined outside the function
            dispo (bool): True for extracting disposition text, False otherwise
            AC (int): Number corresponding to current aircraft
    '''
 
    url = "http://eclipse.ca.aero.bombardier.net/quality/ncms/NCMSExtTool.asp"
    driver.get(url)
   
    # check boxes in output field selection
    NCR_display_status = driver.find_element(By.ID, 'Ncr_Display_StatusX')
    NCR_display_status.click()
    work_center = driver.find_element(By.ID, 'WCX')
    work_center.click()
    aircraft_number = driver.find_element(By.ID, 'ACX')
    aircraft_number.click()
    part_description = driver.find_element(By.ID, 'PdescX')
    part_description.click()
    actual_group_name = driver.find_element(By.ID, 'ActualGroupNameX')
    actual_group_name.click()
    NCR_type = driver.find_element(By.ID, 'NCR_TypeX')
    NCR_type.click()
    component_number = driver.find_element(By.ID, 'PartNumberX')
    component_number.click()
    discrepency_text = driver.find_element(By.ID, 'DiscrepTextX')
    discrepency_text.click()
    actual_queue_aging = driver.find_element(By.ID, 'InQueueAgingX')
    actual_queue_aging.click()
    link_disc_to = driver.find_element(By.ID, 'LinkDiscT0X')
    link_disc_to.click()
    discrepency_number = driver.find_element(By.ID, 'DiscreNumX')
    discrepency_number.click()
    preliminary_cause_code = driver.find_element(By.ID, 'PrelimCauseCodeX')
    preliminary_cause_code.click()
    discrepency_status = driver.find_element(By.ID, 'DiscrepStatusX')
    discrepency_status.click()
    part_number_affected = driver.find_element(By.ID, 'PN_affectedX')
    part_number_affected.click()
 
    if dispo == True:
        disposition_text = driver.find_element(By.ID, 'DispotextX')
        disposition_text.click()
 
    # link disc to
    link_disc_to_textbox = driver.find_element(By.ID, 'LinkDiscT0')
    link_disc_to_textbox.click()
    aircraft = driver.find_element(By.XPATH, '/html/body/form/div/table/tbody/tr[2]/td[2]/input')
    aircraft.click()
    assembly = driver.find_element(By.XPATH, '/html/body/form/div/table/tbody/tr[3]/td[2]/input')
    assembly.click()
 
    # discrepency status
    discrepency_status_textbox = driver.find_element(By.ID, 'DiscrepStatus')
    discrepency_status_textbox.click()
    open = driver.find_element(By.XPATH, '/html/body/form/div/table/tbody/tr[4]/td[2]/input')
    open.click()
    working = driver.find_element(By.XPATH, '/html/body/form/div/table/tbody/tr[5]/td[2]/input')
    working.click()
 
    # aircraft numbers
    range = driver.find_element(By.ID, 'AcSelR')
    range.click()
    ac_from = driver.find_element(By.ID, 'AcFrom')
    ac_from.send_keys(str(AC))
    ac_to = driver.find_element(By.ID, 'AcTo')
    ac_to.send_keys(str(AC))
   
    # submit
    submit = driver.find_element(By.ID, 'button')
    submit.click()
 
    print('Finished processing', ('with dispo' if dispo == True else 'without dispo'))
 
if __name__ == "__main__":
    driver = webdriver.Chrome('C:/Users/B0640469/Documents/webdriver/chromedriver.exe')
    extract_info(driver, True)
    extract_info(driver, False)