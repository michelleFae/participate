import gspread
from oauth2client.service_account import ServiceAccountCredentials
from colorama import Fore, Back, Style 
import operator, random

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',
"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope) #has been hidden on github for security reasons
client = gspread.authorize(creds)


##Google Form Answers Sheet ##
responseSheetName = 'Responses-Participate' #CHANGE ME
responseSheetId = "1i0m5W171l7JsnulqAZty4JZLzjWxI8KerSW_1d2zepQ" #CHANGE ME #check for way to automatically do this
responseSheet = client.open(responseSheetName).sheet1 


##point Sheet ##
pointSheetName = 'Discussion Points 61A'
pointSheetId = "1NBJOYy1PdExE_YOCBz2eRcZCHGkbLetuAiKhDaUd-2A" #CHANGE ME #check for way to automatically do this
pointSheet = client.open(pointSheetName).sheet1 


##column numbers##
responseIdCol = 2
responseOptionCol = 3

pointSheetEmailCol = 1
pointSheetNameCol = 2
pointSheetPointCol = 3
pointSheetIdCol = 4
pointSheetAnonNameCol = 5

#ID of person to get it correct first
firstCorrect = ""
firstCorrectId = 0

#list of options, ordered by timestamp
options = None

#number of people who answered the google form
numResponses = 0

#last filled in row in resopnseSheet
lastFilledInRow = 25 #UPDATE

#first row with valid data in pointSheet
offestPointSheet = 3

#first row with valid data in responseSheet
offestOptionSheet = 3

#total valid entries in pointSheet
totalStudents = lastFilledInRow - offestPointSheet

#equivalent to stopping the collection of information (but doesn't actually). 
#fills in data structures. Call before calling other functions
#simultaneously allocates points for the correct optionNum
#allocates points and resets sheet too
def stop(optionNum,pointVal=1):
	global numResponses, firstCorrect, options, firstCorrectId
	optionNum = str(optionNum)
	options = responseSheet.col_values(responseOptionCol)
	numResponses = len(options) - offestOptionSheet + 1
	print(numResponses)
	try:
		indexOfFirstCorrect = options.index(optionNum) + 1
		firstCorrectId = responseSheet.cell(indexOfFirstCorrect, responseOptionCol - 1).value
		firstCorrect = pointSheet.cell(int(firstCorrectId) % 100, pointSheetAnonNameCol).value

		checkUnique()

		privOptionPoints(optionNum, pointVal)
		reset()
	except ValueError:
		print("no one has answered option " + optionNum)



#get number of respondents
def getNumPpl():
	print(numResponses)
	return numResponses


#gets the ID with the fastest time and correct answer
def getFirstCorrect():
	print("Congratulations " + firstCorrect + "!")
	return firstCorrect

#delete row #use this
#remember to reset gform too
#dont forget to give points!
def reset(): 
	responseSheet.resize(rows=offestOptionSheet-1)
	responseSheet.resize(rows=70)


#makes sure all email ids are unique to prevent cheating
def checkUnique():
	ids = responseSheet.col_values(responseIdCol)
	if numResponses > len(set(ids)) - offestOptionSheet + 1:
		print("error: CHEATING IN PROGRESS")
   

#allocates points to all people who selected that option
def privOptionPoints(optionNum, pointVal=1):
	for i in range(offestOptionSheet - 1,  numResponses + offestOptionSheet - 1):
		if options[i] == optionNum:
			#get id num
			print("options[i]:" + str(options[i]))
			idNum = int(responseSheet.cell(i + 1, responseIdCol).value)
			row = privIdToRow(idNum)
			#verify 4 digit id
			correctIdNum = int(pointSheet.cell(row, pointSheetIdCol).value)
			if idNum != correctIdNum:
				print("incorrect id " + str(idNum) + str(correctIdNum))
			privUpdatePoints(row, pointSheetPointCol, pointVal)


#Updates points
def privUpdatePoints(row, col, pointVal=1):
	currPoints = int(pointSheet.cell(row, col).value)
	pointSheet.update_cell(row, col, currPoints + pointVal)


#finds points for a particular entry
#entry: either a particular name, email, anon. name, or id
#assumes no 2 people with the same name
def privGetPoints(entry, col):
	if privFinder(entry, col) != -1:
		return int(pointSheet.cell(i + 1, pointSheetPointCol).value)
	print(str(entry) + " does not exist. Points can't be retrieved.")
	return 0	

#finds points given name
def pointByName(name):
	return privGetPoints(name, pointSheetNameCol)

#finds points given email
def pointByEmail(email):
	return privGetPoints(email, pointSheetEmailCol)	

#finds points given idVal
def pointByID(idVal):
	idVal = str(idVal)
	return privGetPoints(idVal, pointSheetIdCol)

#finds points given anonymous userName
def pointByAnon(anonName):
	return privGetPoints(anonName, pointSheetAnonNameCol)

#sets points based on either name or idNum
def setPoints(name, idNum=-1, pointVal=1):
	if idNum != -1:
		#then search by id
		row = privIdToRow(idNum)
	else:
		#search by name
		row = privFinder(name, pointSheetNameCol)
		if row == -1:
			return print("person not found")

	privUpdatePoints(row, pointSheetPointCol, pointVal)	
	menu()	

#finds row number of entry entry in column col. Returns -1 if entry does not exist.
def privFinder(entry, col):
	entry = entry.lower()
	valsList = pointSheet.col_values(col)
	for i in range(offestPointSheet - 1, totalStudents + offestPointSheet):
		if valsList[i].lower() == entry:
			return i + 1
	return -1			


#converts ID to row number in pointSheet where entry is stored
#last 2 digits of person correspond to cell in point sheet
def privIdToRow(idNum):
	try:
		idNum = int(idNum)
	except TypeError:
		print("Invalid ID")
		return -1
	return idNum % 100

#Displays top <numberToDisplay> anon. names in the class overall.
def highestScore(numberToDisplay=3):
	allPoints = [int(i) for i in pointSheet.col_values(pointSheetPointCol)[offestPointSheet - 1:]] #1st valid row=0th index 
	
	while numberToDisplay > 0:

		maxValue = max(allPoints)
		highestPointRows = [i+offestPointSheet for i in range(0,len(allPoints)) if allPoints[i] == maxValue]
		highestPointAnonNames = [pointSheet.cell(row, pointSheetAnonNameCol).value for row in highestPointRows]
		print("With a score of " + str(maxValue) + " ...")
		for anonName in highestPointAnonNames:
			print(Fore.WHITE + Back.BLUE + Style.BRIGHT + "Congratulations " + anonName + " !")
			print(Style.RESET_ALL) 

		
		for row in highestPointRows:
			allPoints[row - offestPointSheet] = float("-inf")
			numberToDisplay -= 1




#displays possible commands
def menu():
	message = """
	_____________________________________________
	|	~ONLINE~								
	|	■ stop(optionNum,pointVal=1) 			
	|			#stops data collection and 		
	|				allocates points  			
	|	☺ getNumPpl()  							
	|			#number of respondants			
	|	☼ getFirstCorrect()						
	|			#returns anon. name of 			
	|				person with fastest			
	♥				time  						
	|	◄ reset() 								
	|			#discards data collected		
	|											
	|											
	|	~LIVE~									
	V	♫ setPoints(name, idNum=-1, pointVal=1)
	|			#sets points based on either	
	|				name or idNum  				
	|			#if idNum = -1, sets based 		
	|				on name, else based on 		
	|				idNum
	|  	🌝  randomizer()
	|			# displays random name
	|       					
	|											
	|	~POINT INFORMATION~
	|	░ highestScore(numberToDisplay=3)
	|			#displays top x people overall	
	|	☻ pointByName(name) 					
	|			#finds points given name  		
	|	⌂ pointByEmail(email)					
	♥			#finds points given email 		
	|	◘ pointByID(idVal):						
	|			#finds points given idVal		
	|	♀ pointByAnon(anonName)					
	V			#finds points given anonymous	
	|				userName					
	|	▲ menu()								
	|			#displays menu					
	|___________________________________________
	"""	
	print(Back.WHITE + Style.BRIGHT + message)
	print(Style.RESET_ALL) 

menu()	

#displays a random name
def randomizer():
	name = ''
	while name=='free' or name=='':
		rowNum = random.randint(offestPointSheet, lastFilledInRow + 1)
		name = pointSheet.cell(rowNum, pointSheetNameCol).value
	print(name)





