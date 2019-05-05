# Set Instructions 
This Python script allows you to save only certain columns of an Excel spreadsheet into a different Excel spreadsheet in a specified order. This is useful when using excel spreadsheets in presentations or Powerpoints, so that an image of relevant data can fit on the screen.

### Background ###
This is a project I worked on for my father while I was on winter break. He often needs to keep only certain columns from an Excel spreadsheet for presentations or reports.

### How it works ###
The script is console based and needs to be run in the same directory as the Excel file you wish to read from. From the script you can either set a new `instruction` or run an Excel file through a new `instruction`. These `instruction`s are named (by the user), and saved in a file. Once an Excel file is run through an `instruction`, the resulting columns are saved in a new Excel file named [initial Excel file name]\_[instruction name].

### Note ###
Since this was written primarily for use by my father, there is decent, but not comprehensive input checking.
