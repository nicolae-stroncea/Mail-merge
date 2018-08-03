# Mail-merge
Mail Merger for Windows

### What is this repository for? ### 

Send customized emails to your email list 

### How do I get set up? ### 

Requirements: 

#### General ####
* Ensure you download all of the required libraries.
* Fill in your email address(i.e johnCena@gmail.com), your password(i.e hunter2), the host(i.e smtp.gmail.com), the port number of the host(i.e 587 for gmail) in  user_details.py 
* To launch the app, just launch script_gui and complete all necessary fields.

#### Setting up the excel sheet ####
* Have a xlsx file where you have the people you want to send your emails to, each on their own row.
* If you have a column which has the time, you must call it "Time", and the time has to be in the Excel Time format. 
* If you have a column which has the date, you must call it "Date". 
* Column with emails must be titled "email".

### Setting up the text file ###
* Include a "###" whenever you plan to fill in a word from your excel spreadsheet, and include the column name of that category(i.e name, date, food, etc) in the gui, in the right order.
Example:
Text: Hello, ###, please show up at ###, at ###. We're excited to see you there!
Categories to put in the gui(from excel spreadsheet column names): Name, Location, Time.





