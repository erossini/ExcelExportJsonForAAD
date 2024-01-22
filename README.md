# Excel Export Json For Azure Active Directory
A VBA script that converts Excel tables to JSON format and exports the data to a file at the location of your choice, in particular for `Groups` and `AppRoles` for `Azure Active Directory`.

![image](https://github.com/erossini/ExcelExportJsonForAAD/assets/9497415/480b40c8-cc85-4c1e-92f2-fcd8f59d41fb)

### Installation
You can use this script by following these steps:

1. Open up Microsoft Excel
2. Go to the **Developer** tab (For more information on how to show the developer tab, go [here](https://support.office.com/en-us/article/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45?omkt=en-001&ui=en-US&rs=en-001&ad=US))
3. Click on **Visual Basic**, in the upper left corner of the window
4. In the toolbar at the top of the window that appears, click on **File** > **Import file...**
5. Select **ExcelToJSON.bas** and click on **Open**
6. Click on **File** > **Import file...** for a second time
7. Select **ExcelToJSONForm.frm** and click on **Open** (make sure that **ExcelToJSONForm.frx** is located in the same folder, or this step will not work)

### Usage
To use the script, you need an Excel file with at least one table in it. Once you do, follow these instructions:

1. Go to the **Developer** tab
2. Click on **Macros**
3. Select **yourfile.XLSB!ExcelToJSON.ExcelToJSON**
4. Click on **Run**
5. In the window that appears, select which tables that you would like to export, and then click on **Submit**
6. Finally select the name for the JSON file that will be selected as well as the location that you would like to save the file in
