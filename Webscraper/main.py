import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import time
import datetime
#In case the programm stops running, you should change the "Lex" value to the Number of the last box.
#Ex. the programm wrote untill the box A20 in excel, change "Lex" to 20 and rerun the program. 

stop = False

def main(nex):
    #Imports data from Google Finance
    req = requests.get("https://www.google.com/finance/quote/UBSG:SWX")
    soup = BeautifulSoup(req.content, "html.parser" )

    res = soup.find(jsname="ip75Cb", class_="kf1m0")

    c_output = res.get_text()
    output = c_output.replace("CHF", "")

    #Exports the data to the Excel spreadsheet
    wb = load_workbook('quotes.xlsx')
    ws = wb.active
    ws['A'+nex].value = output
    wb.save('quotes.xlsx')
    
    textfile = open("A-num.txt", "r")
    txtplus1 = int(textfile.read())+1
    textfile.close
    textfile= open("A-num.txt", "w")
    textfile.write(str(txtplus1))
    textfile.close

#Programm loop  
while stop == False :
    rn = str(datetime.datetime.now().time())
    
    if rn >= "20:07:00.000000" and rn <= "20:07:02.000000":
        time.sleep(2)
        textfile = open("A-num.txt", "r")
        lex = textfile.read()
        textfile.close
        main(str(lex))
