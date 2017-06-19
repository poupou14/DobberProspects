#!/usr/bin/python 
from HTMLParser import HTMLParser
import DPDataFormat
import os,string, sys
import urllib
import copy
import chardet
#### SPECIFIC IMPORT #####
sys.path.append("/home/ugoos/Developpement/HockeysFuture/Import/xlrd-0.7.1")
sys.path.append("/home/ugoos/Developpement/HockeysFuture/Import/xlwt-0.7.2")
sys.path.append("/home/ugoos/Developpement/HockeysFuture/Import/pyexcelerator-0.6.4.1")

from pyExcelerator import *
import xlwt
from xlrd import open_workbook
from xlwt import Workbook,easyxf,Formula,Style
#from lxml import etree
import xlrd

currentPlayer_g = ''

def onlyascii(char):
    if ord(char) <= 0 or ord(char) > 127: 
	return ''
    else: 
	return char

def isnumber(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

class DPParser():

	def __init__(self): 
		self.__status = False
		self.__counter = 0
		self.__rootAdd = "http://www.dobberprospects.com/"

	def readDP(self, inputFile_p, outpuFile_p):
		# On va lire le fichier Input
		global currentPlayer_g
		self.getNames(inputFile_p)
		myGoogleForDPParser = GoogleForDPParser()
		len_l = len(DPDataFormat.listPlayer)
		for playerName_l in DPDataFormat.listPlayer.keys() :
			self.__counter+=1
			print "%d" % self.__counter 
			print "%s" % playerName_l 
			myGoogleForDPParser.reset()
			hfPlayerUrl_l = myGoogleForDPParser.getHfPlayerUrl(playerName_l)
			try :
				myDPParserSummary = DPParserSummary()
				myDPParserSummary.reset()
				print "try to read %s" % hfPlayerUrl_l
				url = urllib.urlopen(hfPlayerUrl_l)
				myDPParserSummary.html = url.read()
				url.close()
				currentPlayer_g = DPDataFormat.listPlayer[playerName_l]['Name']
				myDPParserSummary.html = filter(onlyascii, myDPParserSummary.html)
				myDPParserSummary.feed(myDPParserSummary.html)	
				DPDataFormat.listPlayer[playerName_l]['html'] = hfPlayerUrl_l
			except IOError:
				print "%s not readable" % hfPlayerUrl_l 
			
		self.__status = True
		self.writeOuput(outpuFile_p)	
		print DPDataFormat.listPlayer



	def getNames(self, inputFileName) :

		try:
			inputBook_l = open_workbook(inputFileName, encoding_override="cp1252")
		except:            
			print "Error: Impossible to open file %s" % inputFileName                       
			sys.exit()
    
		#read the sheet
		try:
			worksheet_l = inputBook_l.sheet_by_index(0)
		except:
			print  "Error: no sheet 0 found in %s" % xlsfile
			sys.exit()
		
		numRows_l = worksheet_l.nrows - 1
		numCells_l = worksheet_l.ncols - 1
		currRow_l = -1
		while currRow_l < numRows_l:
			currRow_l += 1
			row_l = worksheet_l.row(currRow_l)
			# Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
			name_l = worksheet_l.cell_value(currRow_l, 0)
			print name_l
			DPDataFormat.currentPlayer = dict(DPDataFormat.emptyPlayer)
			DPDataFormat.currentPlayer['Name'] = name_l
			DPDataFormat.listPlayer[name_l] = DPDataFormat.currentPlayer
			teamLNHV2_l = worksheet_l.cell_value(currRow_l, 1)
			DPDataFormat.listPlayer[name_l]['TeamLNHV2'] = teamLNHV2_l
			print "Team : %s" % DPDataFormat.listPlayer[name_l]['TeamLNHV2'] 


	def writeOuput(self, outputFile_p):
#		Book to read xls file (output of main_CSVtoXLS)
    		workbook1_l = Workbook()
    		newPlayerSheet_l = workbook1_l.add_sheet("Prospects", cell_overwrite_ok=True)

		newPlayerIndex_l = 1
		title = False

		for playerName_l in DPDataFormat.listPlayer.keys() :
			player_l = DPDataFormat.listPlayer[playerName_l]
			if not(title) :
				self.addClmnTitle(newPlayerSheet_l, player_l)
				title = True
				self.addPlayer(newPlayerSheet_l, player_l, newPlayerIndex_l) 
				newPlayerIndex_l += 1
			else:
				self.addPlayer(newPlayerSheet_l, player_l, newPlayerIndex_l) 
				newPlayerIndex_l += 1
		workbook1_l.save(outputFile_p)

	def addClmnTitle(self, sheet, player) :
		indexClmn_l = 0
		for key in DPDataFormat.keyListPlayer :
		#for key in player.keys() :
			value_l = filter(onlyascii, key)
			sheet.write(0, indexClmn_l, value_l)
			indexClmn_l += 1
 	
	def addPlayer(self, sheet, player, index):
		indexClmn_l = 0
		style = easyxf('font: underline single')
#		for key in player.keys() :
		for key in DPDataFormat.keyListPlayer :
			if isnumber(player[key]) :
				try :
					if player[key].find(".") != -1 :
						value_l = float(player[key])
					else:
						value_l = player[key].replace("\n", " ").replace("  "," ")
					sheet.write(index, indexClmn_l, value_l)
				except :
					value_l = player[key]
					sheet.write(index, indexClmn_l, value_l)
					
			else :
				if key == "html" :
					link = 'HYPERLINK("%s";"%s")' % (player[key], player['Name'])
					value_l = Formula(link)
					sheet.write(index, indexClmn_l, value_l, style)
				else :
					value_l = filter(onlyascii, player[key])
					sheet.write(index, indexClmn_l, value_l)
			indexClmn_l += 1
				
			
class GoogleForDPParser(HTMLParser): 

	def __init__(self): 
		self.__playerUrlFound = False
		self.__playerUrl = ""

	def getHfPlayerUrl(self, playerName_p) :
		self.__playerUrlFound = False
		self.__playerUrl = ""
		query_l = ''.join((playerName_p, " dobber+prospects"))
    		try:
			filename = 'http://www.google.com/search?' + urllib.urlencode({'q': query_l })
    		except :
			filename = 'http://www.google.com/search?'
		cmd = os.popen("lynx -dump %s | grep www.dobberprospects.com\/" % filename)
		output = cmd.read()
		print "Output=\n%s" % filename
		lineResult_l = output.splitlines()
		cmd.close()
		print "html page : %s" % output
		for line_l in lineResult_l :
			wordsResult_l = line_l.split(" ")
			for word_l in wordsResult_l :
				if word_l.find("www.dobberprospects.com") != -1 :
					self.__playerUrlFound = True
					self.__playerUrl = word_l
					break
			if self.__playerUrlFound : 
				break
		if self.__playerUrl.find("http") == -1 :
			self.__playerUrl = ''.join(("http://", self.__playerUrl))
		return self.__playerUrl	
		

class DPParserSummary(HTMLParser): 

	def __init__(self): 
		print "************ INIT DPParserSummary *************************"
		self.__readOK = False 
		self.__getTbody = False 
		self.__betweenTag = False
		self.__title = False 
		self.__playerUrl = False
		self.__nextOutlook = 0
		self.__nextObservations = -1
		self.__nextFootage = 0
		self.__nextFantasyOutlook = 0
		self.__nextNameAndPos = 0
		self.__nextShoot = 0
		self.__nextDrafted = 0
		self.__nextBorn = 0
		self.__nextHeight = 0
		self.__nextWeight = 0

	def setReadNext(self, val) :
		self.__readNext = val

	def readNext(self) :
		return self.__readNext	

	def getNextPage(self) :
		return self.__nextPage	

	def handle_starttag(self, tag, attrs):
		if tag == "tbody" :
			self.__getTbody = True
		elif tag == "p" and len(attrs) >= 1 and self.__nextNameAndPos < 2 :
			if attrs[0][0].lower() == "align" and attrs[0][1].lower()=="center" and self.__nextNameAndPos == 0: 
				self.__nextNameAndPos = 1
		elif tag == "p" and self.__nextNameAndPos == 2 :
			self.__nextShoot = 1
			self.__nextNameAndPos = 3
		elif tag == "p" and self.__nextShoot == 2 :
			self.__nextHeight = 1
			self.__nextShoot = 3
		elif tag == "p" and self.__nextHeight == 2 :
			self.__nextWeight = 1
			self.__nextHeight = 3
		elif tag == "p" and self.__nextWeight == 2 :
			print "nextBorn =1"
			self.__nextBorn = 1
			self.__nextWeight = 3
		elif tag == "p" and self.__nextBorn == 2 :
			self.__nextDrafted = 1
			self.__nextObservations = 0
			self.__nextBorn = 3
		#elif tag == "strong" and self.__nextDrafted == 2 and self.__nextObservations == 0:
		elif tag == "strong" and self.__nextObservations == 0:
			self.__nextObservations = 1
			self.__nextDrafted = 3
		elif tag == "strong" and self.__nextObservations == 2:
			self.__nextOutlook = 1
			self.__nextObservations = 3
		elif tag == "strong" and self.__nextOutlook == 2:
			self.__nextOutlook = 3
			self.__nextFootage = 1
				
		self.__betweenTag = True

	def handle_data(self, data):
		if self.__nextNameAndPos == 1 :
			datas = data.split(",")
			if len(datas) > 1 :
				DPDataFormat.listPlayer[currentPlayer_g]['Pos'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print "Pos : %s" % datas[1]
			self.__nextNameAndPos = 2
		elif self.__nextShoot == 1 :#and data.find("Shoots") != -1 :
			#print "shoot pos =%s" % data
			datas = data.split(":")
			if len(datas) > 1 and ((data.lower().find("shoots") != -1) or (data.lower().find("catches") != -1)) :
				DPDataFormat.listPlayer[currentPlayer_g]['Shoot'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print "Shoot : %s" % datas[1]
				self.__nextShoot = 2
		elif self.__nextHeight == 1 and data.find("Height") != -1 :
			datas = data.split(":")
			if len(datas) > 1:
				DPDataFormat.listPlayer[currentPlayer_g]['Height'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print "Height : %s" % datas[1]
			self.__nextHeight = 2
		elif self.__nextWeight == 1 and data.find("Weight") != -1 :
			datas = data.split(":")
			if len(datas) > 1:
				DPDataFormat.listPlayer[currentPlayer_g]['Weight'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print "Weight : %s" % datas[1]
			self.__nextWeight = 2
		elif self.__nextBorn == 1 and data.find("Born") != -1:
			datas = data.split(":")
			if len(datas) > 1:
				DPDataFormat.listPlayer[currentPlayer_g]['Born'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print 	"Born : %s" % data
				print 	"Born-2 : %s" % DPDataFormat.listPlayer[currentPlayer_g]['Born']
			self.__nextBorn = 2
		elif self.__nextDrafted == 1 and data.find("Drafted") != -1 :
			datas = data.split(":")
			if len(datas) > 1:
				print "Drafted : %s" % datas[1]
				DPDataFormat.listPlayer[currentPlayer_g]['Drafted'] = datas[1]
			self.__nextDrafted = 2
		elif self.__nextObservations == 1 and data.lower().find("antasy outlook") == -1:
			DPDataFormat.listPlayer[currentPlayer_g]['Observations'] += data
		elif self.__nextObservations == 1 and data.lower().find("antasy outlook") != -1:
			datas = data.split(":")
			if len(datas) > 1:
				DPDataFormat.listPlayer[currentPlayer_g]['Score'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print "Score : %s" % data
				self.__nextFantasyOutlook = 2
				self.__nextOutlook = 1
			else:
				self.__nextFantasyOutlook = 1
			self.__nextObservations = 2
		elif self.__nextFantasyOutlook == 1 :
			datas = data.split(":")
			if data.find(":") != -1 :
				i = data.find(":")
				data = data[i+1:]
				DPDataFormat.listPlayer[currentPlayer_g]['Score'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print "Score : %s" % data
				self.__nextFantasyOutlook = 2
				self.__nextOutlook = 1
			else:
				data = data.replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				if len(data) <= 2:
					DPDataFormat.listPlayer[currentPlayer_g]['Score'] = data
					print "Score : %s" % data
					self.__nextFantasyOutlook = 2
					self.__nextOutlook = 1

		elif self.__nextFantasyOutlook == 2 :
			datas = data.split(":")
			if len(datas) > 1:
				DPDataFormat.listPlayer[currentPlayer_g]['Score'] = datas[1].replace(" ","").replace("\t","").replace("\r","").replace("\n","")
				print "Score : %s" % data
			self.__nextOutlook = 1
			self.__nextFantasyOutlook = 3
		elif self.__nextOutlook == 1 and data.find("Footage") != 0 :
			DPDataFormat.listPlayer[currentPlayer_g]['Outlook'] += data
		elif self.__nextOutlook == 1 and data.find("Footage") == 0 :
			self.__nextOutlook = 2
			self.__nextFootage = 1
		elif self.__nextFootage == 1 :
			DPDataFormat.listPlayer[currentPlayer_g]['Footage'] += data
			self.__nextFootage = 2
			
			

	def handle_endtag(self, tag):
		self.__betweenTag = False



class DPWriter() :

	def __init__(self):
		self.__workingSheet = ""


	def fileUpdate(self, inputFileName, outputFileName):
    		workbook1_l = Workbook()
    		worksheet1_l = workbook1_l.add_sheet("Sheet1", cell_overwrite_ok=True)


#		 Remplissage des champs pour chaque article
		indexRaw_l = 0
		for article_l in listeArticle :
#			print article_l
			indexClmn_l = 0
			if indexRaw_l == 0 :
#		 Remplissage de la premiere ligne
				for key_l in keyListeArticle:
					worksheet1_l.write(0, indexClmn_l, article_l[key_l])
					indexClmn_l+=1
			else:
				for key_l in keyListeArticle:
					if isinstance(article_l[key_l], str):
						article_l[key_l] = cleanString(article_l[key_l])
					worksheet1_l.write(indexRaw_l, indexClmn_l, article_l[key_l])
					indexClmn_l+=1
#				print article_l
			indexRaw_l += 1
    		workbook1_l.save(name)



def open_excel_sheet():
    """ Opens a reference to an Excel WorkBook and Worksheet objects """
    workbook = Workbook()
    worksheet = workbook.add_sheet("Sheet 1")
    return workbook, worksheet

def write_excel_header(worksheet, title_cols):
    """ Write the header line into the worksheet """
    cno = 0
    for title_col in title_cols:
        worksheet.write(0, cno, title_col)
        cno = cno + 1
    return

def write_excel_row(worksheet, rowNumber, columnNumber):
    """ Write a non-header row into the worksheet """
    cno = 0
    for column in columns:
        worksheet.write(lno, cno, column)
        cno = cno + 1
    return

def save_excel_sheet(workbook, output_file_name):
    """ Saves the in-memory WorkBook object into the specified file """
    workbook.save(output_file_name)
    return

