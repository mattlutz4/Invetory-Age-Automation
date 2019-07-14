'''
dirtydozen.py compiles the dirty dozen report, containing the 12 oldest vehicles that are not excluded. 
This script converts the tables into dataframes to carry out the necessary operations, then writes back to the original xlsx file.

dirtydozen.py takes the given excel file as input; the only difference is ‘Dirty Dozen Challenge.xlsx’ was renamed ‘DDC.xlsx’. 
The script uses the original sheet names, therefore any sheet name changes will result in an error. 
Make sure that this script and the DDC.xlsx file are in the same folder before running. 

Arguments:
	inventories from different dealerships in the DDC.xlsx file

Output:
	dirty dozen, excluded and all stores tables, written to the DDC.xlsx file

'''

import pandas as pd
from bs4 import BeautifulSoup
from requests import get
from time import sleep
from openpyxl import load_workbook
import datetime

#This function checks if the car is listed online by searching for the vin on the appropriate dealer site
def is_online(car):
    vin = car.VIN
    url = store_url[car.Store]
    r = get(url+'?search='+vin)
    html_soup = BeautifulSoup(r.text, 'html.parser')
    sleep(0.5) #I dont want to request more that 1/0.5 second
    vehicle_count=html_soup.find_all('span', class_="vehicle-count") #If the html span 'vehicle-count' exists then the car is listed online
    if not vehicle_count:
        return False
    else:
        return True


def find_dd(df):
	bad = []
	for i in range(25): #make sure we get to twelve after moving some vehicles to excluded
	    if is_online(df.iloc[i]):
	        if df.iloc[i].Miles <=1000:
	                dirtydozen.loc[i] = df.iloc[i]
	                if len(dirtydozen) == 12: # once dirty dozen is at 12, break out of for loop
	                    break
	                else:
	                	continue
	                
	        else: #if miles greater than 1000, send to excluded
	            excluded.loc[i] = df.iloc[i]
	            bad.append(i)
	    else: #if car is not listed online, send to excluded
	        excluded.loc[i] = df.iloc[i]
	        bad.append(i)

	allstores.drop(bad, inplace = True) #remove the excluded vehicles from allstores

#path of input file
DDSpath = 'DDC.xlsx'
#Reading input inventories
ch = pd.read_excel(DDSpath, sheet_name = 'CH')
hh = pd.read_excel(DDSpath, sheet_name = 'HH')
ano = pd.read_excel(DDSpath, sheet_name = 'ANO')
cm = pd.read_excel(DDSpath, sheet_name = 'CM')
hy = pd.read_excel(DDSpath, sheet_name = 'HY')
ml = pd.read_excel(DDSpath, sheet_name = 'ML')

#Creating allstores dataframe from 6 dealership inventories
allstores = pd.concat([ch,hh,ano,cm,hy,ml],ignore_index = True)
allstores.sort_values('Age', ascending = False, inplace = True)
allstores.reset_index(drop=True, inplace = True)
#This adds the days past 6/27/2019 to the age column
t0 = datetime.date(2019,6,27)
tdelta = datetime.date.today()- t0
allstores.Age = allstores.Age + tdelta.days
#Dealership websites
cmweb = 'https://www.classicmazda.com/new-inventory/index.htm'
chweb = 'https://www.classichonda.com/new-inventory/index.htm'
hhweb =  'https://www.hollerhonda.com/new-inventory/index.htm'
anoweb = 'https://www.audinorthorlando.com/new-inventory/index.htm'
hyweb =  'https://www.hollerhyundai.com/new-inventory/index.htm'
mlweb = 'https://www.mazdalakeland.com/new-inventory/index.htm'

#Dictionary relating dealership code to the correst website
store_url = {'CM':cmweb,'CH':chweb,'HH':hhweb, 'ANO':anoweb,'HY':hyweb,'ML':mlweb}

#Create Empty dirty dozen and Excluded data frames
columns = ['Stock', 'VIN', 'Vehicle', 'Age', 'Miles', 'Store']
dirtydozen = pd.DataFrame(columns=columns)
excluded = pd.DataFrame(columns=columns)


def main():
	
	find_dd(allstores)

	#Writing back to DDC.xlsx
	book = load_workbook(DDSpath)
	writer = pd.ExcelWriter(DDSpath, engine='openpyxl') 
	writer.book = book
	writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

	dirtydozen.to_excel(writer, "Dirty Dozen", index = False)
	excluded.to_excel(writer, "Excluded", index = False)
	allstores.to_excel(writer, "All Stores", index = False)
	writer.save()

if __name__ == '__main__':
	main()