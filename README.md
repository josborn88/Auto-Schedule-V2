# AutoShells V2, what it is
This is a small program I made to help create daily schedules for my workplace. It creates the shells, or the rough outlines for each day. The daily schedules are made in Excel and our weekly schedules are created with WhenIwork.com. This program scrapes the data from WhenIWork and adds it to the Excel documents. From there the shells are filled in based on who is trained on what activities we are doing that day. I am working on a program to fill in each day, you can find that project on my GitHub under AutoPops.

This is the second iteration I have made for this task. The first version was much longer at over 1000 lines of code. After learning more about Python and programming in general I went back and rewrote it, this being the result.

# Prerequisites
AutoShells runs in Pyhon 3.6 or higher. It also requires the Selenium and OpenPyXl packages. To manipulate the Excel files you need Microsoft Excel, though any program that reads/writes .xlsx files should work decently.
You also need to be using WhenIwork.com for scheduling you team.


# Installing

To install download the Python script and the file "AutoShellsMasterExample.xlsx" from this repository. change the file paths in the script to match wher ou wnat them on your system. 

# Running

To run: run the script in whichever way you prefer. It will launch a web browser and go to WhenIwork.com. Login to WhenIwork and navigate to the weekly schedule you want to make shells for. Once that page is loaded, return to the running Python script. there will be a prompt asking what you wish to save the resulting shells as. Enter your desired file name (E.G. Shells 2.7.19-2.13.19) and hit enter. 

At this point the script will run and save your shells. When the script is finished it will print out the message "Shells are done".

Make any changes you need to to the shells and relax in the joy that you just saved time and energy in making the shells!
