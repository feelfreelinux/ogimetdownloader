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
sheet1.write(0,0,"Data")
sheet1.write(0,1,"Temp(Max)")
sheet1.write(0,2,"Temp(Min)")
sheet1.write(0,3,"Temp(Śr)")
sheet1.write(0,4,"Temp(Noc)")
sheet1.write(0,5,"Wilg(%)")
sheet1.write(0,6,"Wiatr(Kier)")
sheet1.write(0,7,"Wiatr(?)")
sheet1.write(0,8,"Wiatr?)")
sheet1.write(0,9,"Ciś(Hp)")
sheet1.write(0,10,"?")
sheet1.write(0,11,"?")
sheet1.write(0,12,"?")
sheet1.write(0,13,"Słonce")
sheet1.write(0,14,"Widoczność")
sheet1.write(0,15,"Godzina")
written=1;
def rokPrzestepny(rok):
	if rok % 4 == 0 and (rok % 100 != 0 or rok % 400 == 0):
		return True
	else:
		return False
for m in month_iter(int(date1month), int(date1year), int(date2month), int(date2year)):
	for xxx in range(0,24):
		iterations=iterations+1
pbar=tqdm(total=iterations)
for m in month_iter(int(date1month), int(date1year), int(date2month), int(date2year)):
	blokada = False
	for xx in range(1,25):
		w, h = 15, 31
		dane = [[0 for x in range(w)] for y in range(h)]
		daysof = calendar.monthrange(m[1], m[0])[1]
		with urllib.request.urlopen(("http://www.ogimet.com/cgi-bin/gsynres?lang=en&ord=REV&ndays="+str(daysof)+"&ano="+str(m[1])+"&mes="+str(m[0])+"&day="+str(daysof)+"&hora="+str(xx)+"&ind="+str(nrstacji))) as response:
			html = response.read()
		soup = BeautifulSoup(html, 'html.parser')
		tablica = soup.find('table', attrs={'bgcolor':'#d0d0d0'})
		if(blokada==False):
			if(xx==24):
				daysof = daysof-1
				blokada=True
		x=False
		for tr in range(1,(daysof+1)):
			x=False
			for td in range(0,15):
				if(x==False):
					ddatee = tablica.findAll('tr')[tr+1].findAll('td')[td].get_text()
					msc, day = ddatee.split('/')
					dane[tr-1][td] = str(m[1])+"/"+msc+"/"+day
					x=True
				else:
					dane[tr-1][td]=tablica.findAll('tr')[tr+1].findAll('td')[td].get_text()
				

		for x in range(0,daysof):
			for y in range(0,16):
				if(y<15):
					sheet1.write(written+x,y, dane[x][y] )
				elif(y==15):
					sheet1.write(written+x,y, str(xx)+":00")
		written = written+daysof
		pbar.update(1)
		
	
pbar.close()
book.save(nazwapliku)
