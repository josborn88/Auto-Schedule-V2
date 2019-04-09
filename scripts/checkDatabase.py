import config, openpyxl

shows=config.data['Show Training']
attractions=config.data['Attractions']
cafe=config.data['Gadgets Cafe']
temp=config.data['Temporary Programs']
dinos=config.data['Dinos']
lks=config.data['LKS Duties']
ops=config.data['Floor Operations']




def checkTraining(shift):

	training = config.data[shift.database]
	
	knows = []
	
	for i in range(2, 55):
		if (training.cell(row = i, column = shift.rowKnow).value == True) or ((training.cell(row = i, column = shift.rowKnow).value != None) and (training.cell(row = i, column = shift.rowKnow).value != False)):
			knows.append(training.cell(row=i,column=2).value.split()[0] + ' ' + training.cell(row=i,column=2).value.split()[1][0])
	
	return knows

	