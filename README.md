# Automatic-Mail-and-Manager-Report-Using-Google-Script

This is Google Script Code, it automaticaly sends a Customised mail to the assigned Manager, when an employee fills a form.

IT also Sends Regular Update on Mail to all the Managers updating them about the current stautus, with information regarding number and names of Employees who are due to 
send their Report and who have not send send their report.

How To Use IT ?


Start with creating a Google Form, With Name,WD Code (i.e Employee Code) and File Upload Fields. 

![Google Form](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Google%20Form.png?raw=true)


The Google Sheet Will have 4 Sheets:

  Sheet 1 : Form Response Sheet   <br>
  Sheet 2 : WD  AM Mapper Sheet  <br>
  Sheet 3 : AM Data Sheet  <br>
  Sheet 4 : WD Data Sheet  <br><br><br>



Sheet 1 : Add an extra column of Email Stautus to the Response Sheet <br>
![Google Form Response Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Form%20Response.png?raw=true)


<br>


Sheet 2 : This Sheet Contains Data of Manager, Name, Code and Email <br>
![AM Data Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Google%20Form.png?raw=true)

<br>
 
Sheet 3 :  Data of Employee, Name & Employee Code   <br>
![WD Data Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/WD%20Data.png?raw=true)

<br>

Sheet 4 : This Sheet has Employee ID (WD Code) mapped against Manager ID (AM Code) <br>
![WD AM Mapper Sheet](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/WD%20AM%20Mapper.png?raw=true)


<br><br><br>
The Google Script Code has 5 Functions to implement the required task.<br><br>
1. main
2. fetchManagerCode
3. updateManagers
4. fetchWdName
5. fetchAmNameEmail
<br><br>

main Function 
<br>
This function gets triggered when ever a form is filled.<br>

It Checks the mail status column of the Response Sheet, and on finding a row whoose mail status is not Done,<br>
it sends a mail to the Employees' Manager. <br><br>
![Mail to Manager](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/mail.png?raw=true)
<br><br>
fetchManagerCode
<br>
This function takes Employee Code (WD Code in our case) as it's input and returns the Employee's Manager Code.
<br><br>
updateManagers
<br>

This function sends a detailed report to the Manager, which includes the Number and Name the of Employees who have sent their reports and also of those who have not sent. It uses a function sendManagerReport.<br><br>
![Mail to Manager](https://github.com/wilfredarin/Automatic-Mail-and-Manager-Report-Using-Google-Script/blob/main/Manager%20Report.png?raw=true)
<br><br>
fetchWdName
<br>
This function takes Employee ID (WD Code) and Returns the Employee Name (WD Name).
<br><br>
fetchAmNameEmail
<br>
This function takes Manager Code ( AM Code ) and returns the Manager Name and Manager Email address.
<br><br>







