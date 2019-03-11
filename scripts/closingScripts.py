def closeOcean():

	from config import borderGoneTB, borderGoneB, borderGoneT, borderGoneTBold, borderGoneBBold
	
	dbEnd=0
	for i in range(1,100):
		if config.opsdb.cell(row=i,column=2).value==None:
			dbEnd=i
			break
			
	#fill patterns
	blankFill=config.PatternFill(start_color="FFFFFFFF", end_color = 'FFFFFFFF', fill_type='solid')
	oceanFill=config.PatternFill(start_color="FF99CC", end_color = 'FF99CC', fill_type='solid')
	#marked off in database
	canCloseOcean=[]
	for i in range(1,dbEnd):
		if config.opsdb.cell(row=i, column=14).value==True:
			canCloseOcean.append(config.opsdb.cell(row=i,column=2).value.split()[0] + ' ' + config.opsdb.cell(row=i,column=2).value.split()[1][0])
	
	canCloseOceanFirstName=[]
	for i in range(0,len(canCloseOcean)):
		canCloseOceanFirstName.append(canCloseOcean[i].split()[0])
			
	#assign closing ocean
	tried=1000
	go=True
	while go==True:
		assigned=False
		while assigned==False:
			who=random.randint(config.floorStart,config.floorEnd-1)
			if (config.dpop.cell(row=57, column=who).fill == config.blankFill) and (config.dpop.cell(row=58, column=who).fill == config.blankFill) and ((config.dpop.cell(row=21,column=who).value in canCloseOcean) or (config.dpop.cell(row=21,column=who).value in canCloseOceanFirstName)):
				config.dpop.cell(row=57, column=who).value='Ocean'
				config.dpop.cell(row=57, column=who).fill=oceanFill
				config.dpop.cell(row=58, column=who).value='Close'
				config.dpop.cell(row=58, column=who).fill=oceanFill
				tried=0
				assigned=True
				for i in range(config.floorStart,config.floorEnd):			
					if config.dpop.cell(row=57, column=i).value =='Ocean':
						triedOpen=0
					else:
						tried-=1
						#print('HWU'+str(tried))
						if tried <= 0:
							go=False
							break
			else:
				tried-=1
				#print('HWU'+str(tried))
				if tried <= 0:
						go=False
						break
		if tried <= 0:
			go=False
			break

def closeHwu():
	