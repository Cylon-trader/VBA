# VBA
This is MS Excel macro with "Transactions sorter by address" - similar than "One click Pivot table".

1. Activate the "Developer" tab on Excel
  1.1 Right click anywhere on the ribbon, and then click "Customize the Ribbon".
  1.2 Under Customize the Ribbon, on the right side of the dialog box, select Main tabs (if necessary).
  1.3 Check the "Developer" check box
  1.4 Click OK
  
2. Add macro to Excel document
  2.1 Open the "Developer" tab
  2.2 Click the "Visual Basic" button
  2.3 Right click anywhere on the Project explorer field, and then click "Import file" and open downloaded *.bas file
  2.4 Then you will see Module added to Project you can close the VBA window.

3. Copy data from the w8io to the Excel
  3.1 Open the https://w8io.ru and paste the /WavesAddress after the website link
  3.2 Filter transactions for interested currency (click on Waves, for example)
  3.3 The link would be like this: https://w8io.ru/WavesAddress/f/Waves
  3.4 Copy all data (ctrl+a) from the w8io and past to the Excel's A1 cell

4. Run the macro
  4.1 Open the "Developer" tab
  4.2 Click the "Macros" button, chose and Run the macro
