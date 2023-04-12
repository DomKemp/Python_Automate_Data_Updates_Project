# Python_Automate_Data_Updates_Project

1. For this project, I have created a Python script to automate data updates. I am using an Excel file called "budget.xlsx" and a CSV file called "transactions.csv".
The budget.xlsx file has a transactions sheet that has expenses and income. It then adds up the different categories for the expenses and income and displays the values on the summary sheet. The script will extract the data from the "transactions.csv" file and insert this data into the "transactions" sheet of the "budget.xlsx" file, under correct columns.

"Transactions.csv" file:

![transactions csv file](https://user-images.githubusercontent.com/127214128/231582775-714f8e32-ab83-43ec-9ab7-257a5eddda71.PNG)

Transactions sheet of "budget.xlsx" file:

![transactions sheet](https://user-images.githubusercontent.com/127214128/231578411-4b74f918-09e2-4e3f-ae0f-5435690ca018.PNG)

Summary sheet of "budget.xlsx" file:

![Summary tab](https://user-images.githubusercontent.com/127214128/231575606-fe9845b6-49d5-46c1-a4e8-d0e32661b320.PNG)

Steps to run the script: 

1. Download the Project (budget.xlsx, transactions.csv and main-final.py files).
2. Create a folder in your computer and add all 3 files to the folder.
3. Open the terminal on your computer and install openpyxl module using the following command: pip install openpyxl
4. IMPORTANT! Make sure to change the directory in the script to match the directory that you are using.
5. Run the script
6. After running the script, the transactions will be filled in.

Before:
![budget file before](https://user-images.githubusercontent.com/127214128/231588428-ebe707e6-e0bd-4990-9295-888ad2052c5e.PNG)

After: ![transactions sheet](https://user-images.githubusercontent.com/127214128/231588573-926cdc6e-8b1d-4bc9-81d4-7e3622898e6f.PNG)
