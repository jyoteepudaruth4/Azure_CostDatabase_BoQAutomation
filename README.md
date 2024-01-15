# Azure_CostDatabase_BoQAutomation

## Aims and Objectives

The aim of this project is to develop a user interface for filling in MEP (Mechanical, Electrical, and Plumbing) BoQ (Bill of Quantities) in Excel, providing predictive rates as engineers enter items and descriptions. The predictive rates are generated in real-time using information from a database of BoQs from previous engineering projects.

The objectives are:
1. To convert all MEP BoQ (Excel) into a standard format.
2. To store all BoQs on Azure.
3. To perform data cleaning on the data.
4. To use machine learning to interpret text and to predict actual rates using the database.

## Methodology

### Step 1: Converting BoQs to a Standard Format using VBA

All BoQs are stored in a folder on my Desktop. I used VBA to convert the BoQs into a standard format and to store them in a different folder on my Desktop.

#### Snippet of a random BoQ excel file (original):
![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/c18a441a-0701-407a-ab54-b40913799458)

#### Snippet of the modified excel file:
![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/7c0169bd-14e1-40c0-b433-2cc0ed14bf38)

### Step 2: Storing BoQs on Azure

All BoQ files were then imported simultaneously into a single table on Azure as shown below:

![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/719c71fc-77a5-4153-9152-ed7e1fc19677)

### Step 3: Data Cleaning with Python

A program was then written in Python to import the data from Azure and perform data cleaning, such as handling duplicate entries and empty rows. The cleaned data was then imported into a separate table on Azure as shown below:

![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/d17accf3-ab78-457f-9c35-b2e271900442)

### Step 4: Machine Learning with Python

Another program was written in Python to import the cleaned data from Azure to perform machine learning.

## Methodology

The following is what I have achieved so far:

![image](https://github.com/jyoteepudaruth4/Azure_CostDatabase_BoQAutomation/assets/156639095/fce6dd04-efa2-4921-9f09-42186a7b2897)

## Assistance Needed

I am seeking your valuable insights and opinions on the methodology I am using to achieve the aims and objectives of this project. Specifically, I would appreciate feedback on the following aspects:

1. The effectiveness of the VBA script for converting BoQs into a standard format.
2. Suggestions for best practices in storing BoQ data on Azure.
3. Alternative approaches or improvements to the machine learning model used.

Feel free to provide feedback through GitHub issues or directly by reaching out. Thank you for your guidance and contributions!
