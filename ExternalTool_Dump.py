#MIT Open Source License
#Copyright 2021 Michael Boncoddo
#Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import requests
import xlwt

# This tool will dump all external tools to an excel spreadsheet.

# CONFIGURE API ENDPOINT AND TOKEN BELOW

# Example API_URL https://yourschool.instructure.com/api/v1/"
API_URL = ""
ACCESS_TOKEN = ""

# CONFIGURE API ENDPOINT AND TOKEN ABOVE

# FUNCTIONS BELOW

def generateAPICall(aUrl, apiCall, uToken):
    return aUrl + apiCall + "/?per_page=50&access_token=" + uToken

def getPrimaryAccounts():
    rUrl = generateAPICall(API_URL, "accounts", ACCESS_TOKEN)
    r = requests.get(url = rUrl)
    
    jsonObject = r.json()
    
    accountList = []
    for element in jsonObject:
        if element['id'] != "None":
            accountInfo = []
            accountInfo.append(element['id'])
            accountInfo.append(element['name'])
            accountList.append(accountInfo)

    #Pagination
    while r.links['current']['url'] != r.links['last']['url']:
        rUrl = r.links['next']['url'] + "&access_token=" + ACCESS_TOKEN
        r = requests.get(url = rUrl)
        jsonObject = r.json()
        for element in jsonObject:
            if element['id'] != "None":
                accountInfo = []
                accountInfo.append(element['id'])
                accountInfo.append(element['name'])
                accountList.append(accountInfo)
    
    return accountList
    
def getSubAccounts(primaryAccount):
    queryString = "accounts/" + str(primaryAccount) + "/sub_accounts"
    rUrl = generateAPICall(API_URL, queryString, ACCESS_TOKEN)
    r = requests.get(url = rUrl)
    
    jsonObject = r.json()
    
    subAccountList = []
    for element in jsonObject:
        if element['id'] != "None":
            accountInfo = []
            accountInfo.append(element['id'])
            accountInfo.append(element['name'])
            subAccountList.append(accountInfo)
    
    #Pagination
    while r.links['current']['url'] != r.links['last']['url']:
        rUrl = r.links['next']['url'] + "&access_token=" + ACCESS_TOKEN
        r = requests.get(url = rUrl)
        jsonObject = r.json()
        for element in jsonObject:
            if element['id'] != "None":
                accountInfo = []
                accountInfo.append(element['id'])
                accountInfo.append(element['name'])
                subAccountList.append(accountInfo)
            
    return subAccountList
    
def getExternalToolsList(accountId):
    queryString = "accounts/" + str(accountId) + "/external_tools"
    rUrl = generateAPICall(API_URL, queryString, ACCESS_TOKEN)
    r = requests.get(url = rUrl)
    
    jsonObject = r.json()
    
    externalToolsList = []
    for element in jsonObject:
        if element['id'] != "None":
            externalToolsList.append(element['name'])
    
    #Pagination
    while r.links['current']['url'] != r.links['last']['url']:
        rUrl = r.links['next']['url'] + "&access_token=" + ACCESS_TOKEN
        r = requests.get(url = rUrl)
        jsonObject = r.json()
        
        for element in jsonObject:
            if element['id'] != "None":
                externalToolsList.append(element['name'])
            
    return externalToolsList
    
def getExcelHeader():
    header_font = xlwt.Font()
    header_font.name = 'Arial'
    header_font.bold = True
    header_style = xlwt.XFStyle()
    header_style.font = header_font
    return header_style

# FUNCTIONS ABOVE

# Spreadsheet for Exporting Data
book = xlwt.Workbook(encoding='utf-8', style_compression = 0)

# Get Accounts List
accountList = getPrimaryAccounts()
currentIndex = 1
for currentAccount in accountList:
    subAccounts = getSubAccounts(currentAccount[0])
    externalToolsList = getExternalToolsList(currentAccount[0])
    currentName = str(currentIndex) + " - " + currentAccount[1][:25]
    currentSheet = book.add_sheet(currentName, cell_overwrite_ok = True)
    currentSheet.write(0,0,currentAccount[1] + " - External Tools", getExcelHeader())
    currentRow = 1
    for currentTool in externalToolsList:
        currentSheet.write(currentRow,0,currentTool)
        currentRow += 1
    
    currentIndex += 1
    
    for currentSubAccount in subAccounts:
        subExternalToolsList = getExternalToolsList(currentSubAccount[0])
        currentSubName = str(currentIndex) + " - " + currentSubAccount[1][:25]
        currentSubSheet = book.add_sheet(currentSubName, cell_overwrite_ok = True)
        currentSubSheet.write(0,0,currentSubAccount[1] + " - External Tools", getExcelHeader())
        currentRow = 1
        for currentTool in externalToolsList:
            currentSubSheet.write(currentRow,0,currentTool)
            currentRow += 1
        for currentTool in subExternalToolsList:
            currentSubSheet.write(currentRow,0,currentTool)
            currentRow += 1
            
        currentIndex += 1
        
book.save("External_Tools.xls")