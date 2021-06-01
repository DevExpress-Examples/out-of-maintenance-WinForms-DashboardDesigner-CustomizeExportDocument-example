<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/CustomExportDocumentExample/Form1.cs) (VB: [Form1.vb](./VB/CustomExportDocumentExample/Form1.vb))
<!-- default file list end -->
# How to add custom information to the exported Excel document


The  [DashboardDesigner.CustomizeExportDocument](https://docs.devexpress.com/Dashboard/DevExpress.DashboardWin.DashboardDesigner.CustomizeExportDocument) event allows you to obtain the exported document's stream (the [e.Stream](https://docs.devexpress.com/Dashboard/DevExpress.DashboardCommon.CustomizeExportDocumentEventArgs.Stream) property) and change the document's layout.

In this example the [Workbook](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Workbook) component loads the [Excel document](https://docs.devexpress.com/Dashboard/15181) for further processing. The resulting document includes a custom header and highlighted text.


![screenshot](https://github.com/DevExpress-Examples/WinForms-DashboardDesigner-CustomizeExportDocument-example/blob/18.2.4%2B/images/screenshot.png)

## Documentation

- [Printing and Exporting](https://docs.devexpress.com/Dashboard/15181/common-features/printing-and-exporting)
