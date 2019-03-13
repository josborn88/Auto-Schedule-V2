#create a GUI
import openpyxl, PySimpleGUI as sg

layout = [
	[sg.Text('Lets make a DPOP shall we?')],
	[sg.Text('DPOP to Make', size=(15, 1)), sg.InputText('')],
	[sg.Checkbox('')],     
    [sg.Submit(), sg.Cancel()]
	]
window = sg.Window('AutoPops, gotta have \'em').Layout(layout)
button, values = window.Read()
getday = openpyxl.load_workbook('C:\\Users\\JOsbor01\\Desktop\\AutoPop\\'+values[0]+'.xlsx')


print(getday)
print(values[1])