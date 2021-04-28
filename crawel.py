# -*- coding: utf-8 -*-
"""
Created on Thu Apr 22 13:44:33 2021

@author: li651
"""
import time
import datetime
from selenium import webdriver
from bs4 import BeautifulSoup
import openpyxl
from pathlib import Path


browser = webdriver.Chrome('chromedriver.exe')
filename = 'Book1.xlsx'
xlsx_file = Path(filename)
wb_obj = openpyxl.load_workbook(xlsx_file) 
# Read the active sheet:
sheet = wb_obj.active

def AutoLogin():
    browser.get('https://www.linkedin.com/login?fromSignIn=true&trk=guest_homepage-basic_nav-header-signin')
    file = open('config.txt')
    lines = file.readlines()
    username = lines[0]
    password = lines[1]
    elementID = browser.find_element_by_id('username')
    elementID.send_keys(username)
    elementID = browser.find_element_by_id('password')
    elementID.send_keys(password)
    elementID.submit()

#Retrieve the data from excel sheet    
def getData(company,location):
    data = {} #for saving the data retrieve from excel in map format 
    ent=[] # company name list
    loc= [] #company location list
    data[company] = ent
    data[location] = loc
    loc_index = -1
    entity_index = -1
    rows = 0
    for row in sheet.iter_rows():
        count = 0
        for cell in row:
            #title name
            if rows == 0:
                #Search for the company and location columns in file and get their index
                if cell.value == company:
                    entity_index = count
                elif cell.value == location:
                    loc_index = count
            else:
                #if row currently equal to company or location index we should save the cell value in data
                if count == entity_index:
                    data[company].append(cell.value)
                elif count == loc_index:
                    if cell.value == 'SH':
                        data[location].append('shanghai')
                    elif cell.value == 'BJ':
                        data[location].append('beijing')
                    elif cell.value == 'SZ':
                        data[location].append('shenzhen')
                    elif cell.value == 'JS':
                        data[location].append('jiangsu')
                    elif cell.value == 'GZ':
                        data[location].append('guangzhou')
                    else:
                        data[location].append(cell.value)       
            count+=1
        rows+=1
    return data
#type in the company name,hr,location as keyword to search the people on Linkedin
def SearchTalent(Business_Entity,Location,title):
    browser.get('https://www.linkedin.com/search/results/all/?keywords='+Business_Entity+'%20'+title+'%20'+Location+'&origin=GLOBAL_SEARCH_HEADER')

#Retrieve the people's profile links 
def getProfile():
    time.sleep(5)
    links=[]
    page_source = browser.page_source
    soup = BeautifulSoup(page_source, 'html5lib')
    name_a = soup.find_all('a',{'class':'app-aware-link'},href=True)
    for a in name_a:
        links.append(a['href'])
    links = list(dict.fromkeys(links)) #filter out duplicated links
    return links

#Retrieve the data from the profile link and write in the excel file
def getAndWriteInfo(links,row,col,Business_Entity,position):
    org = col
    for link in links:
        browser.get(link)
        time.sleep(3)
        #get soruce code and the information we want
        page_source = browser.page_source
        soup = BeautifulSoup(page_source, 'html5lib')
        #get people's name and write in
        name_name = soup.find('li',{'class':'inline t-24 t-black t-normal break-words'})
        name = name_name.get_text().strip()
        sheet[col+str(row)] = name
        #move to next cell
        temp = ord(col[0])
        col = chr(temp+1)
        #get people's job-title,company and write in
        name_title = soup.find_all('h3',{'class':'t-16 t-black t-bold'})
        companys = soup.find_all('p',{'class':'pv-entity__secondary-title t-14 t-black t-normal'})
        t_size = len(name_title)
        c_size = len(companys)
        for i in range(min(t_size,c_size)):
            title = name_title[i].get_text().strip()
            company = companys[i].get_text().strip()
            #Search by the key word 'HR' and company name
            if position in title and Business_Entity in company:
                sheet[col+str(row)] = title
                temp = ord(col[0])
                col = chr(temp+1)
                sheet[col+str(row)] = company
                temp = ord(col[0])
                col = chr(temp+1)
        #Get the contact information
        browser.get(link+'/detail/contact-info/')
        time.sleep(3)
        page_source = browser.page_source
        soup = BeautifulSoup(page_source, 'html5lib')
        contacts = soup.find_all('div',{'class':'pv-contact-info__ci-container t-14'})
        con_size = len(contacts)
        for i in range(con_size):
            contact = contacts[i].get_text().strip()
            sheet[col+str(row)] = contact
            temp = ord(col[0])
            col = chr(temp+1)
        wb_obj.save(filename) #save once one person's information being written
    #intitize the col
    col = org

#Find and write all HR position for each company on the list in relative location
def ImplementSheet(data,row,col,company,location,title):
    data_size = len(data[company])
    for i in range(row-2,data_size):
        SearchTalent(data[company][i],data[location][i],title)
        links = getProfile()
        try:
            getAndWriteInfo(links,row,col,data[company][i],title)
        except AttributeError as error:
            # Output expected AttributeErrors.
            print(error)
        except Exception as exception:
            # Output unexpected Exceptions.
            print(exception)
        row+=1
#apply the searching and writing process
def application():
    company = input ("Enter the company respresentation title: ")
    location = input ("Enter the location reprsentation title: ")
    title = input ("Enter the title for people you want crawl on Linkedin: ")
    start_row = int(input("Enter the rows index you want to start: "))
    start_col = input("Enter the rows column index you want to start: ")
    print("start runing: ",datetime.datetime.now())
    AutoLogin()
    data = getData(company,location)
    print("start crawling: ",datetime.datetime.now())
    ImplementSheet(data, start_row, start_col,company,location,title)
    print('Done: ',datetime.datetime.now())

#execute the program
application()
