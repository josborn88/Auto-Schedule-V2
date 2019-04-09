class OpenLongShift:
	def __init__(self, name, clearance, color, colorTwo, database, rowKnow, start, end):
		self.name = name
		self.clearance = clearance
		self.backFill = PatternFill(start_color = color, end_color = color, fill_type = 'solid')
		self.backFillTwo = PatternFill(start_color = colorTwo, end_color = colorTwo, fill_type = 'solid')
		self.database = database
		self.rowKnow = rowKnow
		self.shiftStart = start
		self.shiftEnd = end
		
def buildOpenLong(shift):

	times = {'9:45', 30, '10a': 31, '10:15a': 32, '10:30a': 33, '10:45a': 34, '11a': 35, '11:15a': 36, '11:30a': 37, '11:45a': 38, '12p': 39, '12:15p': 40, '12:30p':41, '12:45p':42, '1p':43, '1:15p': 44, '1:30p': 45, '1:45p': 46, '2p': 47, '2:15p': 48, '2:30p': 49, '2:45p': 50, '3p': 51, '3:15p': 52, '3:30p': 53, '3:45p': 54, '4p': 55, '4:15p': 56, '4:30p': 57, '4:45p': 58, '5p': 59, '5:15p': 60, '5:30p': 61}           

	tried = 1000
	go = True
	meetingFill = 00FFFF00
	
	knows = checkDatabase.checkTraining(shift)
	knowsFirst = []

	for i in range(0, len(knows)):
		knowsFirst.append(knows[i].split()[0])

		#times.get(shift.start)
		
	while go == True:	
		assigned = False
		while assigned == False:
			who = random.randint(config.floorStart, config.floorEnd - 1)
			if (config.dpop.cell(row = times.get(shift.start), column = who).fill == config.blankFill or config.dpop.cell(row = times.get(shift.start), column = who).fill == config.meetingFill) and (config.dpop.cell(row = times.get(shift.start) + 3, column = who).fill == config.blankFill):			
				if ((config.dpop.cell(row = 21,column = who).value in knows) or (config.dpop.cell(row=21,column=who).value in knowsFirst)):
					config.dpop.cell(row = times.get(shift.start), column = who).value = time + shift.name.split()[0]
					if shift.name.split()[0] != shift.name.split()[-1]:
						config.dpop.cell(row = times.get(shift.start) + 1, column = who).value = time + shift.name.split()[1]
					config.dpop.cell(row = start, column = who).fill = shift.backFill
					for i in range(times.get(shift.start) + 1, times.get(shift.end)):
						config.dpop.cell(row = i, column = who).fill = shift.backFillTwo
					
					assigned = True
					
			
					if config.dpop.cell(row = times.get(shift.end), column = who).fill == shift.backFillTwo:
						tried = 0
						go = False
						break
					
				else:
					tried -= 1
					if tried <= 0:
						go = False
						break
			else:
				tried -= 1
				if tried <= 0:
					go = False
					break	