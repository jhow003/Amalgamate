# Amalgamate
Powershell Script for combining multiple Excel documents and removing duplicates
>Dedicated to all Past, Present, and Future SANS students


## Required Module
- Import-Excel
  - In PowerShell run `Install-Module ImportExcel -scope CurrentUser` 
  
## Usage
- Create a single Excel spread sheet with all combined Data to be sorted
- Place the Excel file in a new folder (location of your choosing)
- Run PowerShell ISE
- Open Amalgate within ISE
- In the console window `CD` to `C:\FOLDER-YOU-CREATED>`
- Change the `./test.xlsx` on line 25 to ` ./Your-File-Name.xlsx`
- Sit back, relax, and watch the status bar
- Once complete your new Excel doc will be open and be saved within the folder you created above named `FinalMM-dd-yyyy.T.HH-mm-ss.xlsx`
- Enjoy


## Notes
Currently this script will not retain text formating i.e. Bolds, font size, font type, etc.. 

