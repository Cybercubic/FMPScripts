import requests
import openpyxl
import time

API_Key = 'YOUR_API_KEY'

url1 = 'https://financialmodelingprep.com/api/v3/profile/'
url2 = 'https://financialmodelingprep.com/api/v3/balance-sheet-statement/'
url3 = 'https://financialmodelingprep.com/api/v3/income-statement/'
url4 = 'https://financialmodelingprep.com/api/v3/historical-price-full/'

#Dataload

def makeBase (filename):
    
    wb = openpyxl.load_workbook(filename = filename)

    sheet = wb['List1']

    def dataLoad(url1, url2, url3, url4):
    
        response1 = requests.request("GET", url1)
        response2 = requests.request("GET", url2)
        response3 = requests.request("GET", url3)
        response4 = requests.request("GET", url4)
    
        companyProfile = response1.json()
        companyProfile = companyProfile[0]
        companyName = companyProfile['companyName']
        currentPrice = companyProfile['price']
        companyMC = companyProfile['mktCap'] / 1000000
        companyIndustry = companyProfile['industry']
        companyDescription = companyProfile['description']
    
        companyBalance = response2.json()
        companyBalance = companyBalance[0]
        companyEV = (companyMC * 1000000 + companyBalance['shortTermDebt'] + companyBalance['longTermDebt'] 
                     + companyBalance['minorityInterest'] + companyBalance['minorityInterest'] - companyBalance['cashAndShortTermInvestments'])
        companyEV = companyEV / 1000000
        
        companyIncome = response3.json()
        
        companyRevenueQ1 = companyIncome[0]['revenue']
        companyRevenueQ2 = companyIncome[1]['revenue']
        companyRevenueQ3 = companyIncome[2]['revenue']
        companyRevenueQ4 = companyIncome[3]['revenue']
        
        companyIncomeQ1 = companyIncome[0]['netIncome']
        companyIncomeQ2 = companyIncome[1]['netIncome']
        companyIncomeQ3 = companyIncome[2]['netIncome']
        companyIncomeQ4 = companyIncome[3]['netIncome']
    
        companyRevenue = (companyRevenueQ1 + companyRevenueQ2 + companyRevenueQ3 + companyRevenueQ4) / 1000000
        companyNetIncome = (companyIncomeQ1 + companyIncomeQ2 + companyIncomeQ3 + companyIncomeQ4) / 1000000
    
        companyPrices = response4.json()
        companyPrices = companyPrices['historical']
        companyPrice1d = companyPrices[0]['close']
        companyPrice1w = companyPrices[5]['close']
        companyPrice1m = companyPrices[21]['close']
        companyPrice3m = companyPrices[65]['close']
        print(companyPrice3m)
    
        changeDaily = currentPrice / companyPrice1d - 1
        changeWeekly = currentPrice / companyPrice1w - 1
        changeMonthly = currentPrice / companyPrice1m - 1
        change3m = currentPrice / companyPrice3m - 1
    
        data_list = [companyName, currentPrice, changeDaily, changeWeekly, changeMonthly, change3m,
                    companyMC, companyEV, companyRevenue, companyNetIncome, companyIndustry, companyDescription]
    
        return data_list

    for i in range (1, 145):
        time.sleep(1)
        try:
            print('Everything OK')
            tickerCell = 'B' + str(i + 1)
            ticker = sheet[tickerCell].value
            print(ticker)
        
            url5 = url1 + ticker + '?apikey=' + API_Key
            url6 = url2 + ticker + '?apikey=' + API_Key + '&limit=4'
            url7 = url3 + ticker + '?period=quarter&apikey=' + API_Key + '&limit=4'
            url8 = url4 + ticker + '?apikey=' + API_Key + '&serietype=line'
        
            print(url5, url6, url7, url8)

            data_list = dataLoad(url5, url6, url7, url8)
            print(data_list)

            my_alphabet = ['C', 'D', 'E', 'F', 'G', 'H', 'I',
                          'J', 'K', 'L', 'M', 'N']

            for j in range (len(data_list)):
                sheet[my_alphabet[j] + str(i + 1)] = data_list[j]

        except Exception:
            print('Not Like This')
            sheet['D' + str(i + 1)] = 'error'

    wb.save(filename)
    
makeBase('YOUR_EXCEL_FILE_DIR.xlsx')