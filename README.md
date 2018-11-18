# gsuite-serial-letter
GSuite Add-On for creating Serial Letters using Google Spreadsheets and Google Docs

## How it works
- Import your user data as csv in Google Spreadsheets
- You can create Letter Templates in Google Docs with the syntax {{ key }} where key corresponts to the column of the csv file
- The Add-On creates a folder with all the personalized Letters as pdfs' that you can download or share

## Getting Started
1. Create a new project at: https://script.google.com/, If it is asking for what document type, use Spreadsheet
2. Copy the code from the github and paste it into the Google Script project
3. Test if it works:	
    - Create a Letter Template (example: https://docs.google.com/document/d/1aRgYrXG1r-DY09EiRj60ZfwDSDY8W8FEJYXg6G-Jldg/)
    - Create a Spreadsheet with Data: (example: https://docs.google.com/spreadsheets/d/1USsd85E2VSqADYozrsKRh9B8a2uwDFFDbrd3NyLtpuE/
    - To test it in the documents, open the Google Script Editor with your project and click on Execute > Test as Addon and select the Spreadsheet you created.
    - Within the then opened Spreadsheet click on Add-ons > Letter Template > Start
    - Copy the Url of the Letter Document and paste it into the input field in the sidebar element.
    - Click on ‘PDFs erstellen’ and you get all PDFs’

## Screenshots
https://docs.google.com/document/d/16pyYak3txJCaTSwKxSQrOykGRXNZZbQaUT4xm_p_-0s/edit#

