# AutoShellsV2
This is a small program I created to automatically create outlines for the daily schedules for a science museum.

It launches a browser to WhenIWork.com, scrapes the schedule information for the week, and then takes that data and fills in an Excel document that is used for the daily schedule.

This is the second iteration I have made for this task. The first version was muh longer (over 1000 lines of code) and so hard to maintain and adapt and our need changed.

To run the program you need Python with openpyxl and selenium installed. Run the Python script, it will launch a web browser and take you to WhenIWork.com, after you log in and navigate to the week you want to create schedules for. Return to Python and tell it what to save our schdule as (E.G. Schedule 2.1.19-2.6.19) and it will gather the date, parse it out, and will in the Excel file.

In this repository is a blank Excel scheudle and an example of one filled in by this program. After the data is filled in you can reove the extra columns and add any notes or shift specific needs to the day. 

While ther eare a few bugs to fix still an d I would like to add a few axtra features overall I am happy with this program.
