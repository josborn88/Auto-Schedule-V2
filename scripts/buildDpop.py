import config, openpyxl, testForShifts, lksAnimalFunction, LksRatExploreFunction, storytimeFunction, openingScripts, os
from assignShift import buildShift as buildShift,buildLksShift as buildLksShift, buildClose as buildClose, buildShow as buildShow
from defineShifts import Shift as Shift, ClosingShift as ClosingShift, ShowShift as ShowShift


dbPath = os.path.dirname(os.path.dirname( __file__ ))

#define EPT shifts
amnh = Shift('Mythic Creatures', 4, ['HWU', 'Dinosaurs', 'Cafe'], '00339966', 'Temporary Programs', 9)
amnh2 = Shift('Mythic Creatures', 4, ['HWU', 'Dinosaurs', 'Cafe'], '00339966', 'Temporary Programs', 9)

hwu = Shift('HWU (Mezz)', 2, ['HWU', 'Dinosaurs', 'Cafe', amnh.name.split()[0]], '0000FF00', 'Attractions', 12)
dinos = Shift('Dinosaurs', 4, ['HWU', amnh.name.split()[0], 'Cafe', 'Dinosaurs'], 'D8E4BC', 'Dinos', 11)
travel = Shift('Crocs', 4, [], '0099CCFF', 'Temporary Programs', 8)
cafeHost = Shift('Cafe Host', 4, ['HWU', 'Dinosaurs', amnh.name.split()[0], 'Cafe'], 'E6B8B7', 'Gadgets Cafe', 8)

#define LKS shifts
desk = Shift('LKS Desk', 4, [], '60497A', 'LKS Duties', 4)
bkl = Shift('LKS BKL', 4, [], '00FFFF00', 'LKS Duties', 7)
studio = Shift('LKS Studio', 4, [], 'AEAAAA', 'LKS Duties', 7)

#define closing shifts
closeOcean = ClosingShift('Ocean', 'FF99CC', 'FF99CC', 'Floor Operations', 14)
closeLifeProgress = ClosingShift('Life Prog', 'ff6600', 'FFCC00', 'Floor Operations', 8)
closeGadgets = ClosingShift('Gadgets', '99cc00', '99cc00', 'Floor Operations', 6)
closeEnergy = ClosingShift('EE BSP', '7b7b7b', '7b7b7b', 'Floor Operations', 16)

#Define Show Shifts
chemLiveAm = ShowShift('Chem Live', '0000CCFF', '10:30a', '12p', 'Show Training', 8)
chemLivePm = ShowShift('Chem Live', '0000CCFF', '1:30p', '3p', 'Show Training', 8)

eg1040 = ShowShift('EG', '00CCFFFF', '10:30a', '11a', 'Show Training', 59)
eg140 = ShowShift('EG', '00CCFFFF', '1:30p', '2p', 'Show Training', 59)
eg440 = ShowShift('EG', '00CCFFFF', '4:30p', '5p', 'Show Training', 59)

rats1230 = ShowShift('Rat Ball', 'FFCC00', '12p', '1:15p', 'Show Training', 51)
rats330 = ShowShift('Rat Ball', 'FFCC00', '3p', '4:15p', 'Show Training', 51)
#define maintainance shifts
cafePrep = ShowShift('Cafe Prep', 'E6B8B7', '4p', '4:30p', 'Gadgets Cafe', 8)
#build the dpop


#Assign EPT Shifts
madeEPT = False
a = 0
while madeEPT == False:
	#Load DPOP to create
	
	config.getday = openpyxl.load_workbook(dbPath + '\\' + config.whatDay + '.xlsx')
	config.dpop = config.getday[config.getday.sheetnames[0]]
	#assign Shows
	#ratBallFunction.buildRatBall()

	buildShow(chemLiveAm)
	buildShow(chemLivePm)
	
	buildShow(eg1040)
	buildShow(eg140)
	buildShow(eg440)
	
	buildShow(rats1230)
	buildShow(rats330)
	
	#build opening shifts
	openingScripts.openAmnh()
	openingScripts.openAmnhTwo()
	openingScripts.openDinos()
	openingScripts.openHwu()
	openingScripts.openSpace()
	openingScripts.openOcean()
	openingScripts.openLife()
	openingScripts.openGadgets()
	
	#build the main shifts
	buildShift(amnh)
	buildShift(amnh2)
	buildShift(hwu)
	buildShift(dinos)
	buildShift(travel)
	buildShift(cafeHost)
	
	#Maintain Shifts
	buildShow(cafePrep)
	buildShow(cafePrep)
	#roveWindFunction.buildWindRove()
	
	#closing shifts
	buildClose(closeEnergy)
	buildClose(closeOcean)
	buildClose(closeLifeProgress)
	buildClose(closeGadgets)
	
	config.getday.save(dbPath + '\\process\\' + config.whatDay + ' EPT.xlsx')

	finishedEPT = testForShifts.eptTest()

	
	if finishedEPT == True:

		madeEPT = True
		print('EPT finished')
		break
	else:
		if a > 1000:
			madeEPT = True
			print('Human Needed')
			break
		else:
			print('Try '+str(a))
			a+=1

#Assign LKS Shifts			
madeLks = False
a = 0
while madeLks == False:


	#Load DPOP to create
	
	config.getday = openpyxl.load_workbook(dbPath + '\\process\\'+config.whatDay+' EPT.xlsx')
	config.dpop = config.getday[config.getday.sheetnames[0]]
	
	#assign Shows
	lksAnimalFunction.buildAnimalShow()
	LksRatExploreFunction.buildMorningShow()
	storytimeFunction.buildStorytime()
	
	#build the main shifts
	buildLksShift(desk)
	buildLksShift(bkl)
	buildLksShift(studio)

	
	config.getday.save(dbPath + '\\finished\\' + config.whatDay + '.xlsx')
	
	finishedLKS = testForShifts.lksTest()
	
	
	if finishedLKS == True:
    
		madeLks = True
		print('LKS finished')
		break
	else:
		if a > 500:
			madeLks = True
			print('Human Needed')
		else:
			print('Try '+str(a))
			a+=1			
			

print('DPOP Made')
