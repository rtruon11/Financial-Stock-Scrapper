from lxml \
import html #
import requests #import librarie
import xlrd, xlwt #import librarie

def Search(): # definiton of seach function 
    print('Stock ticket: ', ticker)
    url = ("https://www.marketwatch.com/investing/stock/%s" % ticker)
    urlopen = requests.get(url)
    tree = html.fromstring(urlopen.content)
    price = tree.xpath('//span[@class="last-value"]/text()')
    description = tree.xpath('//p[@class="description__text"]/text()')
    low_high = tree.xpath('//span[@class="kv__value kv__primary "]/text()')
    print('Description: ', description[0])
    print('Current stock price: $',price[0])
    print('Open stock price: $', low_high[0])
    print('Todays high and low prices: Low', low_high[1], 'High')


while 1==1 :
    usersearch = str(input("\nWhat is the name of the company? ")).title()
    workbook = xlrd.open_workbook('companylist.csv.xlsx', on_demand = True)
    worksheet = workbook.sheet_by_index(0)
    i = 0
    g = 0
    try:
        for i in range(worksheet.nrows):
            if usersearch in worksheet.cell(i,1).value:
                try:
                    ticker = (worksheet.cell(i,0).value)
                    i += 1
                    g = 1
                    Search()
                except:
                    print('An error occured')
                    continue
        if g != 1:
            print("Company not found")
    except:
        print('Company not found')
        continue
