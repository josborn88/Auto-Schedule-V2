import defineShifts, config, random, checkDatabase
from config import borderGoneTB, borderGoneB, borderGoneT, borderGoneTBold, borderGoneBBold

def buildShift(shift):
	
	knows = checkDatabase.checkTraining(shift)
	knowsFirst = []
	for i in range(0, len(knows)):
		knowsFirst.append(knows[i].split()[0])
		
	
	tried = 1000
	go = True
	
	while go == True:
		for h in range(35,59, shift.long):
			assigned = False
			if shift.name == 'Cafe Host':
				if h > 51:
					go = False
					tried = 0
					assigned = True
			while assigned == False:
				who = random.randint(config.floorStart, config.floorEnd - 1)
				if (config.dpop.cell(row = h, column = who).fill == config.blankFill) and (config.dpop.cell(row = h + shift.long - 1, column = who).fill == config.blankFill):			
					if (config.dpop.cell(row = h - shift.long, column = who).value not in shift.restrictions) and ((config.dpop.cell(row=21,column=who).value in knows) or (config.dpop.cell(row=21,column=who).value in knowsFirst)):
						config.dpop.cell(row = h, column = who).value = shift.name.split()[0]
						if shift.name.split()[0] != shift.name.split()[-1]:
							config.dpop.cell(row = h + 1, column = who).value = shift.name.split()[1]
						config.dpop.cell(row = h, column = who).fill = shift.backFill
						config.dpop.cell(row = h + 1, column = who).fill = shift.backFill
						config.dpop.cell(row = h, column = who).border = config.borderGoneB
						config.dpop.cell(row = h + 1, column = who).border = config.borderGoneT
						assigned = True
						
						if shift.long == 4:
							config.dpop.cell(row = h + 2, column = who).fill = shift.backFill
							config.dpop.cell(row = h + 3, column = who).fill = shift.backFill
							config.dpop.cell(row = h + 1, column = who).border = config.borderGoneTB
							config.dpop.cell(row = h + 2, column = who).border = config.borderGoneTB
							config.dpop.cell(row = h + 3, column = who).border = config.borderGoneT
						for i in range(config.floorStart, config.floorEnd - 1):
							if config.dpop.cell(row = 58, column = i).fill == shift.backFill:
								tried = 0
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
		if tried <= 0:
			go = False
			break
			
def buildLksShift(shift):
	from openpyxl.styles import Color, colors, Font
	
	deskFont=Font(color=colors.WHITE)
	
	knows = checkDatabase.checkTraining(shift)
	knowsFirst = []
	for i in range(0, len(knows)):
		knowsFirst.append(knows[i].split()[0])
		
	
	tried = 1000
	go = True
	
	while go == True:
		for h in range(31, 59, shift.long):
			assigned = False
			while assigned == False:
				who = random.randint(config.lksStart, config.lksEnd - 1)
				if (config.dpop.cell(row = h, column = who).fill == config.blankFill) and (config.dpop.cell(row = h + shift.long - 1, column = who).fill == config.blankFill):			
					if (config.dpop.cell(row = h - shift.long, column = who).value not in shift.restrictions) and ((config.dpop.cell(row=21,column=who).value in knows) or (config.dpop.cell(row=21,column=who).value in knowsFirst)):
						config.dpop.cell(row = h, column = who).value = shift.name.split()[0]
						if shift.name.split()[0] != shift.name.split()[-1]:
							config.dpop.cell(row = h + 1, column = who).value = shift.name.split()[1]
						config.dpop.cell(row = h, column = who).fill = shift.backFill
						config.dpop.cell(row = h + 1, column = who).fill = shift.backFill
						config.dpop.cell(row = h + 2, column = who).fill = shift.backFill
						config.dpop.cell(row = h + 3, column = who).fill = shift.backFill
						if shift.name == 'LKS Desk':
							config.dpop.cell(row = i, column = who).font = deskFont
							config.dpop.cell(row = i + 1,column = who).font = deskFont
						
						assigned = True
						for i in range(config.lksStart, config.lksEnd - 1):
							if config.dpop.cell(row = 58, column = i).fill == shift.backFill:
								tried = 0
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
		if tried <= 0:
			go = False
			break

def buildCampinShiftAm(shift):
	
	knows = checkDatabase.checkTraining(shift)
	knowsFirst = []
	for i in range(0, len(knows)):
		knowsFirst.append(knows[i].split()[0])
		
	
	tried = 1000
	go = True
	
	while go == True:
		for h in range(35,59, shift.long):
			assigned = False
			if shift.name == 'Cafe Host':
				if h > 51:
					go = False
					tried = 0
					assigned = True
			while assigned == False:
				who = random.randint(config.floorStart, config.floorEnd - 1)
				if (config.dpop.cell(row = h, column = who).fill == config.blankFill) and (config.dpop.cell(row = h + shift.long - 1, column = who).fill == config.blankFill):			
					if (config.dpop.cell(row = h - shift.long, column = who).value not in shift.restrictions) and ((config.dpop.cell(row=21,column=who).value in knows) or (config.dpop.cell(row=21,column=who).value in knowsFirst)):
						config.dpop.cell(row = h, column = who).value = shift.name.split()[0]
						if shift.name.split()[0] != shift.name.split()[-1]:
							config.dpop.cell(row = h + 1, column = who).value = shift.name.split()[1]
						config.dpop.cell(row = h, column = who).fill = shift.backFill
						config.dpop.cell(row = h + 1, column = who).fill = shift.backFill
						config.dpop.cell(row = h, column = who).border = config.borderGoneB
						config.dpop.cell(row = h + 1, column = who).border = config.borderGoneT
						assigned = True
						
						if shift.long == 4:
							config.dpop.cell(row = h + 2, column = who).fill = shift.backFill
							config.dpop.cell(row = h + 3, column = who).fill = shift.backFill
							config.dpop.cell(row = h + 1, column = who).border = config.borderGoneTB
							config.dpop.cell(row = h + 2, column = who).border = config.borderGoneTB
							config.dpop.cell(row = h + 3, column = who).border = config.borderGoneT
						for i in range(config.floorStart, config.floorEnd - 1):
							if config.dpop.cell(row = 58, column = i).fill == shift.backFill:
								tried = 0
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
		if tried <= 0:
			go = False
			break	
		
def buildClose(shift):
	tried = 1000
	go = True
	
	knows = checkDatabase.checkTraining(shift)
	knowsFirst = []
	
	for i in range(0, len(knows)):
		knowsFirst.append(knows[i].split()[0])
		
	while go == True:	
		assigned = False
		while assigned == False:
			who = random.randint(config.floorStart, config.floorEnd - 1)
			if (config.dpop.cell(row=57, column=who).fill == config.blankFill) and (config.dpop.cell(row=58, column=who).fill == config.blankFill):			
				if ((config.dpop.cell(row=21,column=who).value in knows) or (config.dpop.cell(row=21,column=who).value in knowsFirst)):
					config.dpop.cell(row = 57, column = who).value = 'CL ' + shift.name.split()[0]
					if shift.name.split()[0] != shift.name.split()[-1]:
						config.dpop.cell(row = 58, column = who).value = 'CL ' + shift.name.split()[1]
					config.dpop.cell(row = 57, column = who).fill = shift.backFill
					config.dpop.cell(row = 58, column = who).fill = shift.backFillTwo
					assigned = True
					
					for i in range(config.floorStart, config.floorEnd - 1):
						if config.dpop.cell(row = 58, column = i).fill == shift.backFillTwo:
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
