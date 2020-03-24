from selenium import webdriver
import xlsxwriter
import time

workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

driver = webdriver.Chrome()

driver.get("https://www.apgo.org/students/residency-directory/search-residency-directory/")

time.sleep(5)

elems = driver.find_elements_by_xpath("//*[@id='Content']/div/div[2]/div/div[2]/div[1]/div")

elems_len = len(elems)

a = 2
for i in range(elems_len-1):

    elem = driver.find_element_by_xpath("//*[@id='Content']/div/div[2]/div/div[2]/div[1]/div["+str(a)+"]/div[1]/span[2]/a")
    
    driver.get(elem.get_attribute("href"))

    time.sleep(5)

    state = driver.find_element_by_xpath("//*[@id='Content']/div/div[2]/div[1]/div[2]/div[2]/div[2]/div[3]/strong").text
    worksheet.write('A'+str(a), state)

    program = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[3]/div[1]/div[3]/strong').text
    worksheet.write('B'+str(a), program)
    print(program)
    
    city = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[2]/div[1]/div[3]/strong').text
    worksheet.write('C'+str(a), city)

    program_director = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[4]/div[5]/div[3]/strong').text
    worksheet.write('D'+str(a), program_director)

    pd_email = ""
    worksheet.write('E'+str(a), pd_email)
    
    program_coord = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[4]/div[8]/div[3]/strong').text
    worksheet.write('F'+str(a), program_coord)

    program_coord_mail = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[4]/div[10]/div[3]/a/strong').text
    worksheet.write('G'+str(a), program_coord_mail)

    program_coord = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[4]/div[8]/div[3]/strong').text
    worksheet.write('H'+str(a), program_coord)

    cleckship_direct = ""
    worksheet.write('I'+str(a), cleckship_direct)

    cleckship_direct_mail = ""
    worksheet.write('A'+str(a), cleckship_direct_mail)

    cleckship_coord = ""
    worksheet.write('J'+str(a), cleckship_coord)

    cleckship_coord_mail = ""
    worksheet.write('K'+str(a), cleckship_coord_mail)

    website = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[3]/div[3]/div[3]/a/strong').text
    worksheet.write('L'+str(a), website)

    phone = driver.find_element_by_xpath('//*[@id="Content"]/div/div[2]/div[1]/div[2]/div[4]/div[6]/div[3]/strong').text
    worksheet.write('M'+str(a), phone)

    driver.get("https://www.apgo.org/students/residency-directory/search-residency-directory/")

    time.sleep(5)

    a+=1

workbook.close()
    
