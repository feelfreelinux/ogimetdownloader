#!/usr/bin/python
from bs4 import BeautifulSoup
import time,os,calendar
from tqdm import *
import urllib,xlwt,sys
from dateutil.rrule import rrule, MONTHLY
from datetime import datetime
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Arkusz 1")
written=0
iterations=0
def month_iter(start_month, start_year, end_month, end_year):
    start = datetime(start_year, start_month, 1)
    end = datetime(end_year, end_month, 1)

    return ((d.month, d.year) for d in rrule(MONTHLY, dtstart=start, until=end))

print("Podaj zakres dat w formacie M/RRRR-M/RRRR np.: 5/2014-6/2016")
zakres = input(" :")
os.system("clear")
print("Chcesz zmienić domyślny nr stacji(Katowice)?(t/n)")
wybor = input(" :")
os.system("clear")
if(wybor=="t"):
	print("Podaj nr stacji: ")
	nrstacji = input(" :")
	os.system("clear")
elif(wybor=="n"):
	nrstacji="12560"
else:
	print("Zły format!")
	sys.exit()
print("Podaj nazwe pliku do zapisu(.xls)")
nazwapliku = input(" :")
date1, date2 = zakres.split('-')
date1month, date1year = date1.split('/')
date2month, date2year = date2.split('/')
written=1;
def rokPrzestepny(rok):
	if rok % 4 == 0 and (rok % 100 != 0 or rok % 400 == 0):
		return True
	else:
		return False
for m in month_iter(int(date1month), int(date1year), int(date2month), int(date2year)):
	iterations=iterations+1
pbar=tqdm(total=iterations)
for m in month_iter(int(date1month), int(date1year), int(date2month), int(date2year)):
	blokada = False
	w, h =19999, 19999
	dane = [[0 for x in range(w)] for y in range(h)]
	daysof = calendar.monthrange(m[1], m[0])[1]
	
	with urllib.request.urlopen(("http://www.ogimet.com/cgi-bin/gsynres?ind="+str(nrstacji)+"&lang=en&decoded=yes&ndays="+str(daysof)+"&ano="+str(m[1])+"&mes="+str(m[0])+"&day="+str(daysof)+"&hora=23")) as response:
		html = response.read()   
	soup = BeautifulSoup(html, 'html.parser')
	tablica = soup.find('table', attrs={'bgcolor':'#d0d0d0'})
	x=False
	for tr in range(1,((daysof*24))):
		x=False
		for td in range(0,15):
			print(tr,td)
			dane[tr-1][td]=tablica.findAll('tr')[tr].findAll('td')[td].get_text()
			
	for x in range(0,(daysof*24)-1):
		for y in range(0,15):
			sheet1.write(written+x,y, dane[x][y] )
	written = written+(daysof*24)-1
	pbar.update(1)

pbar.close()
book.save(nazwapliku)
