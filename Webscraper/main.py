from multiprocessing.connection import wait
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import time
import datetime
lex = 1
stop = False

def main(nex):
    req = requests.get("https://www.google.com/finance/quote/UBSG:SWX")
    soup = BeautifulSoup(req.content, "html.parser" )

    res = soup.find(jsname="ip75Cb", class_="kf1m0")

    c_output = res.get_text()
    output = c_output.replace("CHF", "")

    wb = load_workbook('quotes.xlsx')
    ws = wb.active
    ws['A'+nex].value = output
    wb.save('quotes.xlsx')

    
while stop == False :
    rn = str(datetime.datetime.now().time())
    
    if rn >= "09:00:00.000000" and rn <= "09:00:02.000000":
        time.sleep(2)
        lex = lex + 1
        main(str(lex))
