# Azure_CostDatabase_BoQAutomation

## Aims and Objectives

The aim of this project is to develop a user interface for MEP (Mechanical, Electrical, and Plumbing) BoQ (Bill of Quantities) in Excel, providing predictive rates as engineers enter items and descriptions. The predictive rates are generated using information from a database of BoQs of previous engineering projects.

The objectives are:
1. To convert all MEP BoQ (Excel) into a standard format.
2. To store all BoQs on Azure.
3. To perform data cleaning on the data.
4. To use machine learning to interpret text and to predict actual rates using the database.

## Methodology

### Step 1: Converting BoQs to a Standard Format using VBA

I have all my BoQs stored in a folder on my Desktop. I used VBA to convert the BoQs into a standard format. The latter was stored in a different folder.

#### Original BoQ:
![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/c18a441a-0701-407a-ab54-b40913799458)

#### Modified BoQ:
![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/7c0169bd-14e1-40c0-b433-2cc0ed14bf38)

### Step 2: Storing BoQs on Azure

All BoQ files were then imported into a single table on Azure as shown below:

![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/719c71fc-77a5-4153-9152-ed7e1fc19677)

### Step 3: Data Cleaning with Python

A program was then written in Python to import the data from Azure and perform data cleaning, such as handling duplicate entries and empty rows. The cleaned data was then imported into a separate table on Azure.

### Step 4: Machine Learning with Python

Another program was written in Python to import the cleaned data from Azure to perform machine learning.

## Assistance Needed

I am seeking your valuable insights and opinions on the methodology I am using to achieve the aims and objectives of this project. Specifically, I would appreciate feedback on the following aspects:

1. The effectiveness of the VBA script for converting BoQs into a standard format.
2. Suggestions for best practices in storing BoQ data on Azure.
3. Alternative approaches or improvements to the machine learning model used.

Feel free to provide feedback through GitHub issues or directly by reaching out. Thank you for your guidance and contributions!
