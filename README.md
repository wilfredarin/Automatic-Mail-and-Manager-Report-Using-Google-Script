# Automatic mail through Google Sheet using Google Script

### Problem Statement 
We have Three Designations District Manager (DM), Whole Sale Distributors (WDs) and Area Manager (AMs). WDs have to upload a file using Google form and so that it reaches the  District Manager, since there are  100s of WDs under a District Manager, it's hard for a DM to keep track of all the files. So the DM assigns the task of following up WD's file uploads to few of the AMs.<br>

__Following Documentation Explains the approach taken to solve this problem__



This is Google Script Code, it automaticaly sends a Customised mail to the assigned AM, when ever a  WD fills the form  upload the daily report.

IT also sends regular update on mail to all the Area Managers updating them about the current stautus, with information regarding number and names of WDs who are due to send their Report and who have not send send their report.

## How To Use IT ?


## Step 1 : Create a Google Form<br>
Start with creating a Google Form, With the following fields Name,WD Code (i.e Employee Code) and File Upload.<br> 
 
![Google Form](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Google%20Form.png?raw=true)


## Step 2 : Link the Google Form to Google Sheet for accepting responses

Create 4 sheets including, Form Response Sheet, i.e Google Sheet will have 4 Sheets with the following names:

  **Sheet 1 : "Form Responses 1"   <br>
  Sheet 2 : "AM Data"   <br>
  Sheet 3 : "WD Data"  <br>
  Sheet 4 : "WD AM Mapper"**  <br><br><br>



**Sheet 1 : Form Responses 1 ->  Add an extra column of Mail Status to the Response Sheet** <br><br>
![Google Form Response Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Form%20Response.png?raw=true)


<br><br>


**Sheet 2 : AM Data -> This Sheet Contains Area Manager Name, Area Manager Code and Area Manager Email** <br><br>
![AM Data Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/AM%20DATA.png?raw=true)

<br><br>
 
**Sheet 3 :  WD Data -> WD Name & WD Code** <br>   <br>
![WD Data Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/WD%20Data.png?raw=true)

<br><br>

**Sheet 4 : WD AM Mapper Data -> This Sheet has  WD Code mapped against Area Manager Code (AM Code)** <br><br>
![WD AM Mapper Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/WD%20AM%20Mapper.png?raw=true)


<br><br>
## Step 3: ADD Google Script Code, main.gs From this repository to the Script editor in Google Sheet
<br>

[Copy this Code into the Google Sheet Script Editor](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/main.gs)

<br>

<br>
The Google Script Code has 5 Functions to implement the required task.<br><br>
1. main()<br>
2. fetchManagerCode()<br>
3. updateManagers()<br>
4. fetchWdName()<br>
5. fetchAmNameEmail()<br>


### main() Function
<br>
This function gets triggered when ever a form is filled.<br>

It Checks the mail status column of the Response Sheet, and on finding a row whoose mail status is not Done,<br>
it sends a mail to the Employees' Manager. <br><br>

![Mail to Manager](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/mail.png?raw=true)
<br><br>
### fetchManagerCode() Function
<br>
This function takes Employee Code (WD Code in our case) as it's input and returns the Employee's Manager Code.
<br><br>
### updateManagers() Function
<br>

This function sends a detailed report to the Manager, which includes the Number and Name the of Employees who have sent their reports and also of those who have not sent. It uses a function sendManagerReport.<br><br>
![Mail to Manager](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Manager%20Report.png?raw=true)
<br><br>
### fetchWdName()
<br>
This function takes Employee ID (WD Code) and Returns the Employee Name (WD Name).
<br><br>

### fetchAmNameEmail()
<br>
This function takes Manager Code ( AM Code ) and returns the Manager Name and Manager Email address.
<br><br>

## Step 4 : Add Triggers<br><br>

### Trigger on Main() function to send Mail on every form submit

<br><br>
![Main Trigger](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Trigger%20Main.png?raw=true)
<br><br>

### Triger on updateManagers() function to send detail reports
<br><br>
![updateManagers Trigger](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/trigger%20update%20managers.png?raw=true)
<br><br>










