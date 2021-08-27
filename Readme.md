<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/163293147/18.2.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830483)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/CustomExportDocumentExample/Form1.cs) (VB: [Form1.vb](./VB/CustomExportDocumentExample/Form1.vb))
<!-- default file list end -->
# Dashboard for WinForms - How to add custom information to the exported Excel document


The  [DashboardDesigner.CustomizeExportDocument](https://docs.devexpress.com/Dashboard/DevExpress.DashboardWin.DashboardDesigner.CustomizeExportDocument) event allows you to obtain the exported document's stream (the [e.Stream](https://docs.devexpress.com/Dashboard/DevExpress.DashboardCommon.CustomizeExportDocumentEventArgs.Stream) property) and change the document's layout.

In this example the [Workbook](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Workbook) component loads the [Excel document](https://docs.devexpress.com/Dashboard/15181) for further processing. The resulting document includes a custom header and highlighted text.


![screenshot](https://github.com/DevExpress-Examples/WinForms-DashboardDesigner-CustomizeExportDocument-example/blob/18.2.4%2B/images/screenshot.png)

## Documentation

- [Printing and Exporting](https://docs.devexpress.com/Dashboard/15181/common-features/printing-and-exporting)

## More Examples
- [Dashboard for WinForms - How to add custom information to the exported dashboard](https://supportcenter.devexpress.com/ticket/details/t466558)
