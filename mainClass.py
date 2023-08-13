import os, shutil, re
from datetime import datetime
import MyLibs.Scan_DirsFiles as Scan

class Letter():
	#Excel Table
	ShtName1 = 'Реестр'
	ShtName2 = 'Документы'
	ShtName3 = 'Другое'
	
	Col_Width = {
				ShtName1:{'A:A' : 15, 'B:B' : 6.22, 'C:C' : 13, 'D:D' : 9.33, 'E:E' : 30, 'F:F' : 8.11},
				ShtName2:{'A:A' : 8.11},
				ShtName3:{'A:A' : 8.11}
				}

	
	Colors = {}
	
	#ShtName1		
	Col_Adr = 0				#'Адресат' (контрагент)
	Col_InOut = 1			#'Вх/Исх' and hyperlink to folder
	Col_Num = 2				#'№'
	Col_Date = 3			#'Дата'
	Col_Subj = 4			#'Тема'
	
	Col_FolderLink = 1		# hyperlink to folder
	Col_FileLink = 2		# hyperlink to file
	Col_AnswLink = 7		# hyperlink to file with answer


	#Folder with letters
	pathToScan = r'D:\@Письма'
	FoldersToSkip = '@'
	FldIn = 'Входящие'
	FldOut = 'Исходящие'
	FldDocs = 'Документы'
	
	Reestr = pathToScan + '\\'+ 'РеестрПисем.xlsx'
	DocsToSearch = ['ТТ', 'ТЗ', 'ВИ'] 

	Today = datetime_to_str(now_local(), '%Y.%m.%d')
	SaveTo = rf'D:\@Письма\{Today}.xlsx'
	
	def __init__(self): pass

	def ReadXLS(self):
		print('Read Table...', end=' ')
		from openpyxl import load_workbook
		WBook = load_workbook(self.Reestr)
		self.Data = {}
		for name in WBook.sheetnames:
			ws = WBook[name]
			for row in ws.iter_rows(max_row=1, values_only = True): Title = row #read only cell values (values_only = True)
			rows, links = [], []
			for row in ws.iter_rows(min_row=2):	#read cell as Excel objects with all attributes (values_only = True)
				tmp = []
				for cell in row:
					try: 
						hyperlink = cell.hyperlink.target
					except: 
						hyperlink = None
										
					value = cell.value
					if type(value) == datetime: value = datetime.strftime(value, "%d.%m.%Y")
					tmp.append([value, hyperlink])
					
					
				if name == self.ShtName1: links.append(tmp[self.Col_FileLink][1])
				elif name == self.ShtName2: pass
				rows.append(tmp)
			self.Data[name] = [Title, rows, links]
		print('OK')	
		
	def ScanFiles(self):
		print('Scan Files...', end=' ')
		FilePaths, _ = Scan.SubdirScanPaths(self.pathToScan)
		self.PathsToLetters, self.PathsToOther, self.PathsToDocs = [], [], []
		
		for path in FilePaths:
			newpath = path.replace(self.pathToScan + '\\', '')
			tmp = newpath.split('\\')
			if len(tmp) < 3: continue #skip files in root and in Cooperation folders
			elif self.FoldersToSkip in tmp[0]: continue #Skip folders with '@'
				
			elif (self.FldIn in tmp[1]) or (self.FldOut in tmp[1]):
			
				if ('Приложени' in tmp[-1]): self.PathsToOther.append(path)#Приложения
				elif ('.pdf' in tmp[-1]) == False: self.PathsToOther.append(path) #non 'pdf' files
				else: self.PathsToLetters.append(path)
			elif self.FldDocs in tmp[1]: 
				self.PathsToDocs.append(path)
			else: print('Need to classify:', path)
		print('OK')

	def UpdateLetters(self):
		print('Update data...', end=' ')
		Rename = input('ReName files? (y/n)')
		Title, rows, LinksToFiles = self.Data[self.ShtName1]
		UpdatedData = []
		for FilePath in self.PathsToLetters:
			Adr, InOut, Number, Date, Topic, FilePath, Attach, Ext = self.Split(FilePath)
		
			if FilePath in LinksToFiles:
				N = LinksToFiles.index(FilePath)
				row = rows[N]
				
				if row[self.Col_Adr][0] == None: row[self.Col_Adr][0] = Adr
				if row[self.Col_InOut][0] == None: 
					if 'Исх' in InOut: row[self.Col_InOut][0] = 'Исх.'
					if 'Вх' in InOut: row[self.Col_InOut][0] = 'Вх.'
				if row[self.Col_Date][0] == None: row[self.Col_Date][0] = Date
				if row[self.Col_Num][0] == None: row[self.Col_Num][0] = Number
				if row[self.Col_Subj][0] == None: row[self.Col_Subj][0] = Topic
				
				#pathToFolder
				row[Col_FolderLink][1] = f'{self.pathToScan}\\{row[self.Col_Adr][0]}\\{row[self.Col_InOut][0]}' 
				#pathToFile
				if Rename == 'y':
					newpath = f'{self.pathToScan}\\{row[self.Col_Adr][0]}\\{row[self.Col_InOut][0]}\\{row[self.Col_Date][0]} {row[self.Col_Num][0]} {row[self.Col_Subj][0]}{Ext}'
					row[Col_FileLink][1] = newpath
					self.MoveFile(FilePath, newpath)
				else:
					row[Col_FileLink][1] = FilePath
				
				UpdatedData.append(row)
				
				#search for attach
				tmp = f'{self.pathToScan}\\{row[Adr][0]}\\{row[InOut][0]}\\{row[Date][0]} {row[Number][0]}'
				for AttPath in self.PathsToAttach:
					if tmp in AttPath:
						Adr, InOut, Number, Date, Topic, FilePath, Attach, Ext = self.Split(AttPath)
						newpath = f'{self.pathToScan}\\{row[self.Col_Adr][0]}\\{row[self.Col_InOut][0]}\\{row[self.Col_Date][0]} {row[self.Col_Num]} {Attach}{Ext}'
						self.MoveFile(AttPath, newpath)
			else:
				L = len(Title)
				newrow = [[None, None] for i in range(L)]
				newrow[self.Col_Adr][0] = Adr
				if 'Исх' in InOut: newrow[self.Col_InOut][0] = 'Исх.'
				if 'Вх' in InOut: newrow[self.Col_InOut][0] = 'Вх.'
				newrow[self.Col_Date][0] = Date
				newrow[self.Col_Num][0] = Number
				newrow[self.Col_Subj][0] = Topic
				
				newrow[self.Col_FolderLink][1] = f'{self.pathToScan}\\{Adr}\\{InOut}' 
				newrow[self.Col_FileLink][1] = FilePath
				
				UpdatedData.append(newrow)
		self.Data[self.ShtName1] = Title, UpdatedData, LinksToFiles 
		print('OK')
			
	def Split(self, FilePath):
		if self.pathToScan in FilePath: newpath = FilePath.lstrip(self.pathToScan + '\\')
		else: newpath = FilePath
		
		Adr, InOut, filename = newpath.split('\\')
		
		ExtRe = r'(.\w{3,4})$'
		match = re.compile('(.+)' + ExtRe).search(filename)
		tmp, Ext = match.groups()
		
		AttachRe = r'приложени.+'
		match = re.compile(AttachRe, re.IGNORECASE).search(tmp)
		if match == None: Attach = None
		else:
			match = re.compile('(.+)' + r'(Приложени.+)$').search(tmp)
			tmp, Attach = match.groups()
		
		tmp2 = tmp.split(' ')
		if len(tmp2) == 3: Date, Number, Topic = tmp2
		else: 
			DateRe = r'(\d{4}.\d{2}.\d{2})'
			match = re.compile(DateRe + '(.+)').search(tmp)
			Date, tmp = match.groups()
			
			NumRe = r'([0-9-_]+)'
			match = re.compile(NumRe + '(.+)').search(tmp)
			Number, Topic = match.groups()
			
		Number = Number.strip(' ')
		Number = Number.strip('_')
		Number = Number.replace('_', '/')
		Topic = Topic.strip(' ')
		Topic = Topic.strip('_')
		Topic = Topic.replace('_', ' ')
			
		return Adr, InOut, Number, Date, Topic, FilePath, Attach, Ext

	
	def MoveFile(self, oldPath, newPath):
		tmp = newPath.split('\\')
		newdir = ('\\').join(tmp[:-1])
		if not os.path.exists(newdir): os.makedirs(newdir)
		shutil.move(oldPath, newPath)

	def WriteXLS(self):
		print('Write to file...', end=' ')
		import xlsxwriter
		#Create new file
		workbook = xlsxwriter.Workbook(self.SaveTo)
		
		# Add a bold format to use to highlight cells.
		title_format = workbook.add_format({'bold': True, 'align': 'center'})
		
		cell_format = workbook.add_format({'align': 'center'})
		
		for ShtName in self.Data: 
			Title, rows, _ = self.Data[ShtName]
			ws = workbook.add_worksheet(ShtName)
			
			for col, width in self.Col_Width[ShtName].items():
				ws.set_column(col, width)
			
			
			ws.write_row ('A1', Title, title_format)
			
			for row_num, row in enumerate(rows):
				for col_num, data in enumerate(row):
					value, hLink = data
					if hLink == None: 
						if type(value) == 'datetime.datetime': 
							print(value)
							ws.write_datetime(row_num+1, col_num, cell_format, value)
						else: ws.write(row_num+1, col_num, value, cell_format)
					else: ws.write_url(row_num+1, col_num, hLink, cell_format, string = value)

		workbook.close()
		print('Ok')
	
if __name__ == '__main__':

	Reestr = Letter()
	Reestr.ReadXLS()
	Reestr.ScanFiles()
	Reestr.UpdateLetters()
	Reestr.WriteXLS()
	
		
	# There is no way to specify “AutoFit” for a column in the Excel file format. This feature is only
	# available at runtime from within Excel. It is possible to simulate “AutoFit” in your application by
	# tracking the maximum width of the data in the column as your write it and then adjusting the
	# column width at the end.
	