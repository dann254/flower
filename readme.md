# Flower.

This app demonstrates the functionality of xlrd and xlwt.

# usage
	This is an app that manages salary for employees in a flower company.
	The app reads the document flower.xlsx, calculates the salary totals following the required instructions and writing the result on a new document called Salaries.xls.

# installation
	1. install python 2.7
	2. install pip
	3. using pip, install requirements "pip install -r requirements.txt"
	4. run salary.py using python to excecute the app.


# what was needed
	1. All employees are payed ksh 10 for each flower cut.
	2. Badly cut flowers are not included in employee's salary or flower count
	3. All the employees that cut more than 5000 flowers are intled to ksh 5 bonus for each extra flower cut.
	4. All employees that have more than 100 badly cut flowers get ksh 5 deducted from their salary for each extra badly cut flower.
	5. Employees that have less than 10 badly cut flowers are entitled to 3% of raw salary bonus.

# contents
	1. salary.py - this is the main app that does the calculations.
	2. flowers.xlsx - this is the flower database.
# original document - flower.xlsx.

	The app reads this document as the flower database.
![flower.xlsx screenshot](screen_shots/flower.PNG?raw=true "")


# result document Salaries.xls.

	This is the result created after the app runs.
	
	## all the contents of This document are completely generated by the app. 
![result.xls screenshot](screen_shots/result.PNG?raw=true "")

# N/B
note that the employees in the included documents are completely fictional.
