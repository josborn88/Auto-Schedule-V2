import config, openpyxl, testForShifts, lksAnimalFunction, LksRatExploreFunction, storytimeFunction, openingScripts, os
from assignShift import buildShift as buildShift,buildLksShift as buildLksShift, buildClose as close, buildShow as buildShow, buildOpen as open
from defineShifts import Shift as Shift, CloseShift as CloseShift, ShowShift as ShowShift, OpenShift as OpenShift


dbPath = os.path.dirname(os.path.dirname( __file__ ))

#Section 2: Defining Shifts

#define EPT shifts
amnh = Shift('Mythic Creatures', '', 4, ['HWU', 'Dinosaurs', 'Cafe'], '00339966', 'Temporary Programs', 9)
amnh2 = Shift('Mythic Creatures', 'WOSU', 4, ['HWU', 'Cafe'], '00339966', 'Temporary Programs', 9)
hwu = Shift('HWU', 'Mezz', 2, ['HWU', 'Dinosaurs', 'Cafe', amnh.name.split()[0]], '0000FF00', 'Attractions', 12)
dinos = Shift('Dinosaurs', '', 4, ['HWU', amnh.name.split()[0], 'Cafe', 'Dinosaurs'], 'D8E4BC', 'Dinos', 11)
travel = Shift('Crocs', 'G1 & O', 4, [], '0099CCFF', 'Temporary Programs', 8)
cafeHost = Shift('Cafe Host', 'Gad', 4, ['HWU', 'Dinosaurs', amnh.name.split()[0], 'Cafe'], 'E6B8B7', 'Gadgets Cafe', 8)


#define LKS shifts
desk = Shift('LKS Desk', '(North)', 4, ['Desk'], '60497A', 'LKS Duties', 4)
bkl = Shift('LKS BKL', '', 4, ['BKL'], '00FFFF00', 'LKS Duties', 7)
studio = Shift('LKS Studio', '(South)', 4, ['Studio'], 'AEAAAA', 'LKS Duties', 7)


#define Opening closing shifts

#define Opening Long Shifts
opGadgets = OpenShift('Gad Cafe', 'Prep', '99cc00', 'e6b8b7', 'Floor Operations', 6, '10a', '11a')
OpHWU = OpenShift('HWU', 'Mezz', '0000FF00', '0000FF00', 'Attractions', 3, '10a', '11a')
OpDinos = OpenShift('Dinos', '', 'D8E4BC', 'D8E4BC', 'Dinos', 11, '9:45a', '11a')
OpOcean = OpenShift('Ocean LP', '', 'FF99CC', 'FF99CC', 'Floor Operations', 14, '10a', '10:30a')
OpSpace = OpenShift('Space Mezz', '', 'cc99ff', 'cc99ff', 'Floor Operations', 11, '10a', '10:30a')
OpLife = OpenShift('Life Prog', '', 'ff6600', 'FFCC00', 'Floor Operations', 7, '10a', '10:30a')
OpEE = OpenShift('EE Prarie', '', '76933c', '76933c', 'Floor Operations', 16, '10a', '10:30a')
OpTravel = OpenShift('Crocs', '', '0099CCFF', '0099CCFF', 'Temporary Programs', 8, '10a', '11a')

OpAmnhOne = OpenShift(amnh.name, amnh.clearance, '00339966', '00339966', amnh.database, amnh.rowKnow, '9:45a', '11a')
OpAmnhTwo = OpenShift(amnh.name, amnh2.clearance, '00339966', '00339966', amnh.database, amnh.rowKnow, '9:45a', '11a')


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
windtubes = ShowShift('Windtubes', '0099CCFF', '4p', '4:30p', 'Floor Operations', 2)

#Build closing Shifts
clOcean = CloseShift('Ocean', 'FF99CC', 'FF99CC', 'Floor Operations', 14)
clLifeProg = CloseShift('Life Prog', 'ff6600', 'FFCC00', 'Floor Operations', 8)
clGadgets = CloseShift('Gadgets', '99cc00', '99cc00', 'Floor Operations', 6)
clEnergy = CloseShift('EE BSP', '76933c', '76933c', 'Floor Operations', 16)
clHWU = CloseShift( 'HWU Space', '0000FF00', 'cc99ff', 'Attractions', 12)
#build the dpop



#Section 3: Make the dpop

#Assign EPT Shifts
madeEPT = False
a = 0
while madeEPT == False:
	#Load DPOP to create
	
	config.getday = openpyxl.load_workbook(dbPath + '\\' + config.whatDay + '.xlsx')
	config.dpop = config.getday[config.getday.sheetnames[0]]
	#assign Shows
	#ratBallFunction.buildRatBall()
	if a < 500:
		buildShow(chemLiveAm)
		buildShow(chemLivePm)
	
		buildShow(eg1040)
		buildShow(eg140)
		buildShow(eg440)
	
		buildShow(rats1230)
		buildShow(rats330)
	
	#build opening shifts
	#openingScripts.openDinos()
	
	
	open(opGadgets)
	open(OpAmnhOne)
	open(OpAmnhTwo)
	open(OpHWU)
	open(OpDinos)
	open(OpOcean)
	open(OpSpace)
	open(OpLife)
	open(OpEE)
	open(OpTravel)

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
	buildShow(windtubes)
	buildShow(windtubes)
	#roveWindFunction.buildWindRove()
	
	#closing shifts
	
	close(clOcean)
	close(clLifeProg)
	close(clGadgets)
	close(clEnergy)
	close(clHWU)
	
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
