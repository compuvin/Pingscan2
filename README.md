# Pingscan 2.0
Scans a range of IP addresses and reports to a MySQL database

The idea behind this VBScript is to have one script that takes care of all that is needed to ping a range IP addresses and report the data to a MySQL database.
The VBScript will prompt for all the information that is needed the first time that it is run. It will then create a Powershell script based on that data which will run and create a CSV file.
The CSV file with the active hosts will then be run against the database and important changes emailed to the administrator.

If any changes need to be made to the setup, the psapp.ini file that is created during the first run can be edited. The changes take place immediately the next time the script is run. It is *almost* a portable app.

So this is really version 2.0 of a program that I designed years ago to catch people hopping onto an open Wi-Fi. There is a custom ASP web interface that I designed for it which I hope to convert and make it public at some point.

*Can now be set to download the MAC address CSV file automatically*
*Now prompts to create database and table*
