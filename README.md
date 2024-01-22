# Excel Export JSON for Azure Active Directory
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

## Example
In an Excel file, you map the `Groups` for `Azure Active Directory` that you want to create or associate. For example, you have a table like that.

| GroupName          | AppReg         | AppRoles      |
| ------------------ | -------------- | ------------- |
| {ENV}_Contributors | {ENV}_API      | Designer      |
| {ENV}_Contributors | {ENV}_API      | Editor        |
| {ENV}_Contributors | {ENV}_API      | Team_Users    |
| {ENV}_Contributors | {ENV}_API      | Viewer        |
| {ENV}_Contributors | {ENV}_UI       | Designer      |
| {ENV}_Contributors | {ENV}_UI       | Editor        |
| {ENV}_Contributors | {ENV}_UI       | Team_Users    |
| {ENV}_Contributors | {ENV}_UI       | Viewer        |
| {ENV}_Contributors | {ENV}_UI       | AdminDesigner |
| {ENV}_Contributors | {ENV}_UI       | AdminEditor   |
| {ENV}_Contributors | {ENV}_UI       | AdminViewer   |
| {ENV}_Contributors | {ENV}_API2_API | Team_Users    |
| {ENV}_Contributors | {ENV}_API2_API | AdminDesigner |
| {ENV}_Contributors | {ENV}_API2_API | AdminEditor   |
| {ENV}_Contributors | {ENV}_API2_API | AdminViewer   |
| {ENV}_Dev_Leads    | {ENV}_API      | Admin         |
| {ENV}_Dev_Leads    | {ENV}_API      | Designer      |
| {ENV}_Dev_Leads    | {ENV}_API      | Editor        |
| {ENV}_Dev_Leads    | {ENV}_API      | Exporter      |
| {ENV}_Dev_Leads    | {ENV}_API      | Importer      |
| {ENV}_Dev_Leads    | {ENV}_API      | Team_Users    |
| {ENV}_Dev_Leads    | {ENV}_API      | Viewer        |
| {ENV}_Dev_Leads    | {ENV}_UI       | Admin         |
| {ENV}_Dev_Leads    | {ENV}_UI       | Designer      |
| {ENV}_Dev_Leads    | {ENV}_UI       | Editor        |
| {ENV}_Dev_Leads    | {ENV}_UI       | Exporter      |
| {ENV}_Dev_Leads    | {ENV}_UI       | Importer      |
| {ENV}_Dev_Leads    | {ENV}_UI       | Team_Users    |
| {ENV}_Dev_Leads    | {ENV}_UI       | Viewer        |
| {ENV}_Dev_Leads    | {ENV}_UI       | AdminAdmin    |
| {ENV}_Dev_Leads    | {ENV}_UI       | AdminDesigner |
| {ENV}_Dev_Leads    | {ENV}_UI       | AdminEditor   |
| {ENV}_Dev_Leads    | {ENV}_UI       | AdminExporter |
| {ENV}_Dev_Leads    | {ENV}_UI       | AdminImporter |
| {ENV}_Dev_Leads    | {ENV}_UI       | AdminViewer   |

Now, the issue is how to create a Json file for this table. There is an export in Excel that creates a Json but not in the format that is required for the Active Directory. By the way, the expected `json` is like the following one

```
{
  "Groups": [
    {
      "GroupName": "{ENV}_Contributors",
      "AppRegs": [
        {
          "AppRegName": "{ENV}_API",
          "AppRoles": [
            "Designer",
            "Editor",
            "Team_Users",
            "Viewer"
          ]
        },
        {
          "AppRegName": "{ENV}_UI",
          "AppRoles": [
            "Designer",
            "Editor",
            "Team_Users",
            "Viewer",
            "AdminDesigner",
            "AdminEditor",
            "AdminViewer"
          ]
        },
        {
          "AppRegName": "{ENV}_API2_API",
          "AppRoles": [
            "Team_Users",
            "AdminDesigner",
            "AdminEditor",
            "AdminViewer"
          ]
        }
      ]
    },
    {
      "GroupName": "{ENV}_Dev_Leads",
      "AppRegs": [
        {
          "AppRegName": "{ENV}_API",
          "AppRoles": [
            "Admin",
            "Designer",
            "Editor",
            "Exporter",
            "Importer",
            "Team_Users",
            "Viewer",
            "TRSCore"
          ]
        },
        {
          "AppRegName": "{ENV}_UI",
          "AppRoles": [
            "Admin",
            "Designer",
            "Editor",
            "Exporter",
            "Importer",
            "Team_Users",
            "Viewer",
            "AdminAdmin",
            "AdminDesigner",
            "AdminEditor",
            "AdminExporter",
            "AdminImporter",
            "AdminViewer",
            "TRSCore"
          ]
        },
        {
          "AppRegName": "{ENV}_API2_API",
          "AppRoles": [
            "Team_Users",
            "AdminAdmin",
            "AdminDesigner",
            "AdminEditor",
            "AdminExporter",
            "AdminImporter",
            "AdminViewer",
            "TRSCore"
          ]
        }
      ]
    },
    {
      "GroupName": "{ENV}_DevOps",
      "AppRegs": [
        {
          "AppRegName": "{ENV}_API",
          "AppRoles": [
            "Admin",
            "Designer",
            "Editor",
            "Exporter",
            "Importer",
            "Team_Users",
            "Viewer"
          ]
        },
        {
          "AppRegName": "{ENV}_UI",
          "AppRoles": [
            "Admin",
            "Designer",
            "Editor",
            "Exporter",
            "Importer",
            "Team_Users",
            "Viewer",
            "AdminAdmin",
            "AdminDesigner",
            "AdminEditor",
            "AdminExporter",
            "AdminImporter",
            "AdminViewer"
          ]
        },
        {
          "AppRegName": "{ENV}_API2_API",
          "AppRoles": [
            "Team_Users",
            "AdminAdmin",
            "AdminDesigner",
            "AdminEditor",
            "AdminExporter",
            "AdminImporter",
            "AdminViewer"
          ]
        }
      ]
    },
    {
      "GroupName": "{ENV}_Internal_Client_Support",
      "AppRegs": [
        {
          "AppRegName": "{ENV}_API",
          "AppRoles": [
            "Designer",
            "Editor",
            "Team_Users",
            "Viewer",
            "TRSCore"
          ]
        },
        {
          "AppRegName": "{ENV}_UI",
          "AppRoles": [
            "Designer",
            "Editor",
            "Team_Users",
            "Viewer",
            "AdminDesigner",
            "AdminEditor",
            "AdminViewer",
            "TRSCore"
          ]
        },
        {
          "AppRegName": "{ENV}_API2_API",
          "AppRoles": [
            "Team_Users",
            "AdminDesigner",
            "AdminEditor",
            "AdminViewer",
            "TRSCore"
          ]
        }
      ]
    },
    {
      "GroupName": "{ENV}_Users",
      "AppRegs": [
        {
          "AppRegName": "{ENV}_API",
          "AppRoles": [
            "Editor",
            "Team_Users",
            "Viewer"
          ]
        },
        {
          "AppRegName": "{ENV}_UI",
          "AppRoles": [
            "Editor",
            "Team_Users",
            "Viewer",
            "AdminEditor",
            "AdminViewer"
          ]
        },
        {
          "AppRegName": "{ENV}_API2_API",
          "AppRoles": [
            "Team_Users",
            "AdminEditor",
            "AdminViewer"
          ]
        }
      ]
    }
  ]
}
```

Because this structure is a little but complex, I have to create something my own export. With this code, when I run the `Macro`, I get a window with the list of the table in the spreadsheet.

![image](https://github.com/erossini/ExcelExportJsonForAAD/assets/9497415/2dda0ff4-40bf-429a-b6d1-306fbfb14b5e)

Then, I can select one or more tables I want to export. Remember this script generates only one `json` file. After that, I have to choose the location and the name of the file I want to create.
