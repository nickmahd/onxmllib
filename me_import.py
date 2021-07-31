import fnmatch, os, zipfile, clr, shutil, System, sys, calendar, time
sys.path.append(r'C:\EPPlus')
clr.AddReferenceToFile('EPPlus.dll')
from OfficeOpenXml import ExcelPackage
clr.AddReference('System.Xml')
clr.AddReference('WindowsBase')
from System.Xml import *
from datetime import datetime, timedelta,date
from System.Threading import Thread
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#basePath = 'C:/Users/rasrinivasan.ONCEPT/Desktop/files/'
basePath = 'G:/Electricity/ME/MISO/Financials/'
logpath = 'C:/logs_me/'
path = basePath + 'to_import/'
workpath = basePath + 'Summaries/'
def log(message):
	print message
	now = datetime.now()
	logfilename = str(now.month) + "-" + str(now.day) + "-" + str(now.year) + ".txt"
	logfile = logpath + logfilename
	if not os.path.isfile(logfile):
		open(logfile, 'w').close()
	with open(logfile, "a") as myfile:
		myfile.write(now.strftime('%d/%m/%y %H:%M %p') + ": ")
		myfile.write(str(message) + "\n")
def email(Exception, e):
	print Exception
	print e
	body = "error Exception! " + e.message + "\n"+str(Exception)
	msg = MIMEText(body)
	msg['Subject'] = 'ME_watch Error'
	msg['From'] = "Me_watch@oncept.net"
	rec = ['venkat@oncept.net']
	msg['To'] = "venkat@oncept.net"
	s = smtplib.SMTP('10.10.0.222')
	s.sendmail(msg['From'],msg['To'], msg.as_string())
	s.quit()
def daily(type,sdate):
	#me update

	date = sdate - timedelta(days=7)
	filename = type + ' Summary ' + str(date.year) + '.xlsx'
	
	log(filename)
	fs = System.IO.FileInfo(workpath+filename)
	ps = ExcelPackage(fs)
	ws = ps.Workbook
	s = ws.Worksheets[calendar.month_name[date.month]]
	col = date.day + 1
	
	trone = s.Cells[3,col].Value;
	if trone == None:
		trone = 0.0;
	trtwo = s.Cells[4,col].Value;
	if trtwo == None:
		trtwo = 0.0;
	
	trade_results = trone + trtwo
	trade_results = -1*trade_results
	totalsum = 0.0

	for i in range(s.Dimension.End.Row-3):
		if s.Cells[i+3,col].Value != None:
			totalsum = totalsum + s.Cells[i+3,col].Value
	ps.Stream.Close()
		
	print type, date.year
	filename = type + ' FTR ' + str(date.year) + '.xlsx'
	log(filename)
	fs = System.IO.FileInfo(workpath + filename)
	ps = ExcelPackage(fs)
	ws = ps.Workbook
	s = ws.Worksheets[calendar.month_name[date.month]]
	col = date.day + 1
	
	totalftr = 0.0

	for i in range(s.Dimension.End.Row-3):
		if s.Cells[i+3,col].Value != None:
			totalftr = totalftr + s.Cells[i+3,col].Value
	ps.Stream.Close()
	
	total = totalftr - totalsum
	fees = total - trade_results
	
	send(type, trade_results, fees, total, date)
def colTotal(s,col):
	total = 0.0
	#lastrow=""
	for i in range(s.Dimension.End.Row-2):

		if s.Cells[i+3,col].Value != None:
			total = total + s.Cells[i+3,col].Value
		#lastrow=s.Cells[i+3,0].Value
	#if lastrow!="Other Amount":
	#	total = total + s.Cells[s.Dimension.End.Row,col].Value
	return total
def send(type, trade, fees, total, date):
	body = "Trading Results: " + str(trade) + '\r\n' + "Fees and Charges: " + '{0:,.2f}'.format(fees) + '\r\n' + "Total: " + '{0:,.2f}'.format(total)
	msg = MIMEText(body)
	msg['Subject'] = type + ' ' + date.strftime("%m/%d/%Y") + " Daily Summary"
	msg['From'] = type + "_summary@oncept.net"
	msg['To'] = "jp@oncept.net"
	
	msg.add_header('To','venkat@oncept.net')
	if type == "ME2":
		msg.add_header('To','shal@oncept.net')
	if type == "ME":
		msg.add_header('To','nzhang@oncept.net')
	s = smtplib.SMTP('10.10.0.222')
	s.sendmail(msg['From'],msg.get_all('To'), msg.as_string())
	s.quit()
def weekly(type,sdate):
	date = sdate - timedelta(days=7)
	wdate = date - timedelta(days=6)
	log("weekly report starting - beginning date: " + wdate.strftime("%m/%d/%Y"))
	filename = type + ' FTR ' + str(wdate.year) + '.xlsx'
	log(filename)
	fs = System.IO.FileInfo(workpath+filename)
	ps = ExcelPackage(fs)
	ws = ps.Workbook
	s = ws.Worksheets[calendar.month_name[wdate.month]]
	row = [];
	for i in range(s.Dimension.End.Row-2):
		row.append(s.Cells[i+2,1].Value + " Sum: ")
	amt = [];
	for i in range(len(row)):
		amt.append(0.0)
	ps.Stream.Close()
	while wdate != date + timedelta(days=1):
		filename = type + ' FTR ' + str(wdate.year) + '.xlsx'
		fs = System.IO.FileInfo(workpath+filename)
		ps = ExcelPackage(fs)
		ws = ps.Workbook
		s = ws.Worksheets[calendar.month_name[wdate.month]]
		col = wdate.day + 1
		for i in range(len(row)):
			if i == 0:
				cel = colTotal(s,col)
			else:
				cel = s.Cells[i+2,col].Value
			if cel == None:
				cel = 0;
			amt[i] = amt[i] + cel
		wdate = wdate + timedelta(days=1)
		ps.Stream.Close()
	ps.Stream.Close()
	sendweekly(type + " FTR",row,amt,date)
def weeklySummary(type,sdate):
	date = sdate - timedelta(days=7)
	wdate = date - timedelta(days=6)
	log("weekly report starting - beggining date: " + wdate.strftime("%m/%d/%Y"))
	filename = type + ' Summary ' + str(wdate.year) + '.xlsx'
	log(filename)
	fs = System.IO.FileInfo(workpath+filename)
	ps = ExcelPackage(fs)
	ws = ps.Workbook
	s = ws.Worksheets[calendar.month_name[wdate.month]]
	row = [];
	lastCell=""
	for i in range(s.Dimension.End.Row-2):
		row.append(s.Cells[i+2,1].Value + " Sum: ")
		lastCell=s.Cells[i+2,1].Value
	if lastCell!="Other Amount":
		row.append("Other Amount")
	amt = [];
	for i in range(len(row)):
		amt.append(0.0)
	ps.Stream.Close()
	while wdate != date + timedelta(days=1):
		filename = type + ' Summary ' + str(wdate.year) + '.xlsx'
		fs = System.IO.FileInfo(workpath+filename)
		ps = ExcelPackage(fs)
		ws = ps.Workbook
		s = ws.Worksheets[calendar.month_name[wdate.month]]
		col = wdate.day + 1
		for i in range(len(row)):
			if i == 0:
				cel = colTotal(s,col)
			else:
				cel = s.Cells[i+2,col].Value
			if cel == None:
				cel = 0;
			amt[i] = amt[i] + cel
		wdate = wdate + timedelta(days=1)
		ps.Stream.Close()
	ps.Stream.Close()
	sendweekly(type + " Summary",row,amt,date)
#def sendweekly_old(type,row,amt,date):
#	odate = date - timedelta(days=6)
#	body = "Week of: " + odate.strftime("%m/%d/%Y") + " - " + date.strftime("%m/%d/%Y")
#	for i in range(len(row)):
#		body = body + "\r\n" + row[i] + '{0:,.2f}'.format(amt[i])
#	msg = MIMEText(body)
#	msg['Subject'] = type + ' ' + odate.strftime("%m/%d/%Y") + " - " + date.strftime("%m/%d/%Y") + " Weekly Summary"
#	msg['From'] = type + "_summary@oncept.net"
#	msg['To'] = "jp@oncept.net"
#	if type == "ME Summary" or type == "ME2 Summary":
#		msg.add_header('To','laura@oncept.net')
#	s = smtplib.SMTP('10.10.0.222')
#	s.sendmail(msg['From'],msg.get_all('To'), msg.as_string())
#	s.quit()
def sendweekly(type,row,amt,date):
	odate = date - timedelta(days=6)
	body = ( "<html><head></head><body>\r\n"
		    +"Week of: " + odate.strftime("%m/%d/%Y") + " - " + date.strftime("%m/%d/%Y") + "<br/>\r\n"
			+"<table border=\"1\">\r\n")
	for i in range(len(row)):
		body = body + "<tr><td> "+ row[i] + "</td><td> "+ '{0:,.2f}'.format(amt[i]) +" </td></tr>\r\n"
	body = body + "</table></body></html>"
	#msg = MIMEText(body)
	msg = MIMEMultipart('alternative')
	msg['Subject'] = type + ' ' + odate.strftime("%m/%d/%Y") + " - " + date.strftime("%m/%d/%Y") + " Weekly Summary"
	msg['From'] = type + "_summary@oncept.net"
	msg['To'] = "jp@oncept.net"
	#if type == "ME Summary" or type == "ME2 Summary":
	msg.add_header('To','laura@oncept.net')
	msg.add_header('To','venkat@oncept.net')

	text = "You need a html capable mail reader to read this message"
	# Record the MIME types of both parts - text/plain and text/html.
	part1 = MIMEText(text, 'plain')
	part2 = MIMEText(body, 'html')	
	# Attach parts into message container.
	# According to RFC 2046, the last part of a multipart message, in this case
	# the HTML message, is best and preferred.
	msg.attach(part1)
	msg.attach(part2)

	s = smtplib.SMTP('10.10.0.222')
	s.sendmail(msg['From'],msg.get_all('To'), msg.as_string())
	s.quit()
def writeAO(expath, d,sdate,otheramt):
	fixedTop = ['Day Ahead Virtual Energy Amount','Real Time Virtual Energy Amount','Real Time Market Administration Amount','Day Ahead Market Administration Amount','Financial Transmission Rights Market Administration Amount','Day Ahead Schedule 24 Allocation Amount','Real Time Schedule 24 Allocation Amount','Real Time Schedule 24 Distribution Amount','Day Ahead Revenue Sufficiency Guarantee Distribution Amount','Day Ahead Revenue Sufficiency Guarantee Make Whole Payment Amt','Real Time Revenue Sufficiency Guarantee First Pass Dist Amount','Real Time Revenue Sufficiency Guarantee Make Whole Payment Amt'];
	rest = ['Day Ahead Regulation Amount','Financial Transmission Guarantee Uplift Amount','Real Time Excessive Deficient Energy Deployment Charge Amount','Day Ahead Spinning Reserve Amount','Day Ahead Supplemental Reserve Amount','Day Ahead Asset Energy Amount','Day Ahead Financial Bilateral Transaction Congestion Amount','Day Ahead Financial Bilateral Transaction Loss Amount','Day Ahead Congestion Rebate on Carve-Out Grandfathered Agrmnts','Day Ahead Losses Rebate on Carve-Out Grandfathered Agrmnts','Day Ahead Congestion Rebate on Option B Grandfathered Agrmnts','Day Ahead Losses Rebate on Option B Grandfathered Agrmnts','Day Ahead Non-Asset Energy Amount','Auction Revenue Rights Transaction Amount','Financial Transmission Rights Annual Transaction Amount','Auction Revenue Rights Infeasible Uplift Amount','Auction Revenue Rights Stage 2 Distribution Amount','Financial Transmission Rights Full Funding Guarantee Amount','Financial Transmission Rights Hourly Allocation Amount','Financial Transmission Rights Monthly Allocation Amount','Financial Transmission Rights Monthly Transaction Amount','Financial Transmission Rights Transaction Amount','Financial Transmission Rights Yearly Allocation Amount','Contingency Reserve Deployment Failure Charge Amount','Excessive Energy Amount','Net Regulation Adjustment Amount','Non-Excessive Energy Amount','Real Time Regulation Amount','Regulation Cost Distribution Amount','Real Time Spinning Reserve Amount','Spinning Reserve Cost Distribution Amount','Real Time Supplemental Reserve Amount','Supplemental Reserve Cost Distribution Amount','Real Time Asset Energy Amount','Real Time Demand Response Allocation Uplift Charge','Real Time Financial Bilateral Transaction Congestion Amount','Real Time Financial Bilateral Transaction Loss Amount','Real Time Congestion Rebate on Carve-Out Grandfathered Agrmnts','Real Time Losses Rebate on Carve-Out Grandfathered Agrmnts','Real Time Distribution of Losses Amount','Real Time Miscellaneous Amount','Real Time Non-Asset Energy Amount','Real Time Net Inadvertent Distribution Amount','Real Time Price Volatility Make Whole Payment Amt','Real Time Revenue Neutrality Uplift Amount','Real Time Uninstructed Deviation Amount'];
	total = fixedTop + rest;
	date = sdate.strftime("%m/%d/%Y");

	f = System.IO.FileInfo(expath)
	p = ExcelPackage(f)
	w = p.Workbook
	activesheet = w.Worksheets[calendar.month_name[sdate.month]]

	row = 1;
	col = sdate.day + 1;
	
	activesheet.Cells[row,col].Value = date
	
	row = 3
	keys = d.keys()
	for i in total:
		if i in keys:
			activesheet.Cells[row,col].Value = float(d[i])
		row = row + 1
	activesheet.Cells[row,1].Value = "Other Amount"
	activesheet.Cells[row,col].Value = otheramt
	p.Save()
	p.Stream.Close()
	

	p = ExcelPackage(f)
	w = p.Workbook
	activesheet = w.Worksheets[calendar.month_name[sdate.month]]
	#make active view of correct area
	if sdate.month > 1:
		nt = NameTable()
		nsmgr = XmlNamespaceManager(nt)
		nsmgr.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" )
		v2=w.WorkbookXml.SelectNodes("/", nsmgr).Item(0)
		v=v2.GetElementsByTagName("workbookView", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" ).Item(0)
		v.SetAttribute("activeTab", str(sdate.month-1))


	#hide all zero rows
	hideRow = 1
	row = 3 + len(fixedTop)
	while row <= activesheet.Dimension.End.Row-1:
		activesheet.Row(row).Hidden = 1;
		for i in range(activesheet.Dimension.End.Column-2):
			if activesheet.Cells[row,i+2].Value != "None" and activesheet.Cells[row,i+2].Value != None:
				if activesheet.Cells[row, i+2].Value != 0.0:
					activesheet.Row(row).Hidden = 0;
					continue;
		row = row + 1;		
	
	p.Save()
	p.Stream.Close()
def writeFTR(expath,d,sdate):
	fixedTop = ['Financial Transmission Rights Market Administration Amount','Financial Transmission Rights Hourly Allocation Amount','Auction Revenue Rights Infeasible Uplift Amount','Financial Transmission Guarantee Uplift Amount'];
	rest = ['Auction Revenue Rights Transaction Amount','Financial Transmission Rights Annual Transaction Amount','Auction Revenue Rights Stage 2 Distribution Amount','Financial Transmission Rights Full Funding Guarantee Amount','Financial Transmission Rights Monthly Allocation Amount','Financial Transmission Rights Monthly Transaction Amount','Financial Transmission Rights Transaction Amount','Financial Transmission Rights Yearly Allocation Amount'];
	total = fixedTop + rest;
	date = sdate.strftime("%m/%d/%Y");
	
	f = System.IO.FileInfo(expath)
	p = ExcelPackage(f)
	w = p.Workbook
	activesheet = w.Worksheets[calendar.month_name[sdate.month]]
	
	row = 1;
	col = sdate.day + 1;
	
	activesheet.Cells[row,col].Value = date
	
	row = 3
	for i in total:
		activesheet.Cells[row,col].Value = float(d[i])
		row = row + 1
	
	p.Save()
	p.Stream.Close()
	p = ExcelPackage(f)
	w = p.Workbook
	activesheet = w.Worksheets[calendar.month_name[sdate.month]]
	#make active view of correct area
	#activesheet.Select(activesheet.Cells[1,col].Address)

	#hide all zero rows
	hideRow = 1
	row = 3 + len(fixedTop)
	while row <= activesheet.Dimension.End.Row:
		activesheet.Row(row).Hidden = 1;
		for i in range(activesheet.Dimension.End.Column-2):
			if activesheet.Cells[row,i+2].Value != "None" and activesheet.Cells[row,i+2].Value != None:
				if activesheet.Cells[row, i+2].Value != 0.0:
					activesheet.Row(row).Hidden = 0;
					continue;
		row = row + 1;	
	
	p.Save()
	p.Stream.Close()
def importFTR(file, type):
	importFile = path + file
	z = zipfile.ZipFile(importFile)
	l = z.namelist();
	idx = [i for i, item in enumerate(l) if fnmatch.fnmatch(item,'FTR*S7.xml')]
	imp = z.read(l[idx[0]])
	
	doc = XmlDocument()
	doc.LoadXml(imp)
	
	n = doc.SelectNodes("//SCHEDULED_DATE")
	time = n[0].InnerText
	odate = datetime.strptime(time,"%m/%d/%Y");
	sdate = odate - timedelta(days=7);
	
	#using date determine workbook/worksheet
	exportfile = type + " FTR " + str(sdate.year) + ".xlsx"

	#ensure workbook exists or make workbook
	if not os.path.isfile(workpath+exportfile):
		shutil.copyfile(workpath+r"\templates\FTR_template.xlsx",workpath+exportfile);
	
	#generate field:amt dictionary
	n = doc.SelectNodes("//CHG_TYP/CHG_TYP_NM")
	names = [];
	for c in n:
		names.append(c.InnerText)
	
	n = doc.SelectNodes("//CHG_TYP/STLMT_TYP/AMT")
	amts = [];
	for c in n:
		amts.append(c.InnerText)
	
	d = {};
	for i in range(len(names)):
		d[names[i]]=amts[i];
	
	z.close()	
	
	writeFTR(workpath + exportfile, d, sdate)
def importAO(file, type):
	importFile = path + file
	log(importFile)
	z = zipfile.ZipFile(importFile)
	l = z.namelist();
	print("==0==")	   
	idx = [i for i, item in enumerate(l) if fnmatch.fnmatch(item,'AO-*.xml')]
	print(str(idx))
	zi = z.getinfo(l[idx[0]])
	print("00")
	if(zi.compress_size<=0):
		raise Exception('Empty/Invalid AO-*.xml file in ' + importFile)
	print("==1==")
	imp = z.read(l[idx[0]])
	
	doc = XmlDocument()
	doc.LoadXml(imp)
	
	n = doc.SelectNodes("//SCHEDULED_DATE")
	time = n[0].InnerText
	odate = datetime.strptime(time,"%m/%d/%Y");
	sdate = odate - timedelta(days=7);
	
	#using date determine workbook/worksheet
	exportfile = type + " Summary " + str(sdate.year) + ".xlsx"

	#ensure workbook exists or make workbook
	if not os.path.isfile(workpath+exportfile):
		shutil.copyfile(workpath+r"\templates\AO_template.xlsx",workpath+exportfile);
	
	#generate field:amt dictionary
	n = doc.SelectNodes("//CHG_TYP/CHG_TYP_NM")
	names = [];
	for c in n:
		names.append(c.InnerText);
	
	n = doc.SelectNodes("//CHG_TYP/STLMT_TYP[1]/AMT")
	other = doc.SelectNodes("//CHG_TYP/STLMT_TYP/AMT")

	amts = []
	otheramt=0.00
	s7amt=0.00
	for c in n:
		amts.append(c.InnerText);
		s7amt+=float(c.InnerText)
	for c in other:
		otheramt+=float(c.InnerText)
	otheramt-=s7amt
	d = {};
	for i in range(len(names)):
		d[names[i]]=amts[i];
	z.close()	
	
	writeAO(workpath + exportfile, d, sdate,otheramt)
def update(files):
	try:
		fileType = 'ME'
		for file in files:	
			if fnmatch.fnmatch(file,'*ME_*.zip'):
				fileType='ME'
			elif fnmatch.fnmatch(file,'*ME2_*.zip'):
				fileType='ME2'
			elif fnmatch.fnmatch(file,'*QTRO_*.zip'):
				fileType='QTRO';
			elif fnmatch.fnmatch(file,'*GOE_*.zip'):
				fileType = 'GOE'
			elif fnmatch.fnmatch(file,'*QTRO2_*.zip'):
				fileType='QTRO2'
			elif fnmatch.fnmatch(file,'*ALTRO_*.zip'):
				fileType='ALTRO'
			elif fnmatch.fnmatch(file,'*ALTR_*.zip'):
				fileType='ALTR'


			exportpath = basePath + 'Settlements.'+fileType+'/'
			importAO(file, fileType)
			importFTR(file, fileType)
			#send the daily email for the file just processed
			date = file.split('_')[1].split('.')[0]
			odate = datetime.strptime(date,"%Y%m%d")
			daily(fileType,odate)
			#send the weekly email for the ME file just processed IF the last date is a Friday
			if odate.weekday() == 4:
				weekly(fileType,odate)
				weeklySummary(fileType,odate)


			#if fnmatch.fnmatch(file,'*ME_*.zip'):
			#	exportpath = basePath + 'Settlements.ME/'
			#	importAO(file, 'ME')
			#	importFTR(file, 'ME')
			#	#send the daily email for the ME file just processed
			#	date = file.split('_')[1].split('.')[0]
			#	odate = datetime.strptime(date,"%Y%m%d")
			#	daily("ME",odate)
			#	#send the weekly email for the ME file just processed IF the last date is a Friday
			#	if odate.weekday() == 4:
			#		weekly("ME",odate)
			#		weeklySummary("ME",odate)
			#elif fnmatch.fnmatch(file,'*ME2_*.zip'):
			#	exportpath = basePath + 'Settlements.ME2/'
			#	importAO(file,'ME2')
			#	importFTR(file, 'ME2')
			#	#send the daily email for the ME2 file just processed
			#	date = file.split('_')[1].split('.')[0]
			#	odate = datetime.strptime(date,"%Y%m%d")
			#	daily("ME2",odate)
			#	#send the weekly email for the ME file just processed IF the last date is a Friday
			#	if odate.weekday() == 4:
			#		weekly("ME2",odate)
			#		weeklySummary("ME2",odate)

				
			stamp = file.split('.')[0].split('_')[1];
			year = stamp[0:4]
			month = stamp[4:6]
			day = stamp[6:]
			exportpath = exportpath + year + '/' + month + '/'
			if not os.path.exists(exportpath):
				os.makedirs(exportpath)
			log(path+file)
			log(exportpath)
			log("=============================")
			try:
				shutil.copy(path+file, exportpath)
				os.remove(path+file)
			except Exception,e:
				try:
					Thread.CurrentThread.Join(1000 * 15)
					shutil.copy(path+file, exportpath)
					os.remove(path+file)
				except Exception,e:
					log("failure moving")
	except Exception,e:
		log("Exception:")
		log(e.errno)
		log(e.strerror)
		
#today = date.today()
#odate = today-timedelta(days=3)
#weeklySummary("ME",odate)						 