Imports DevExpress.DashboardCommon
Imports DevExpress.Spreadsheet
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Namespace CustomExportDocumentExample
    Partial Public Class Form1
        Inherits DevExpress.XtraEditors.XtraForm

        Public Sub New()
			InitializeComponent()
			dashboardDesigner1.CreateRibbon()
			AddHandler dashboardDesigner1.ConfigureDataConnection, AddressOf DashboardDesigner1_ConfigureDataConnection
			dashboardDesigner1.LoadDashboard("Dashboard.xml")
			dashboardDesigner1.ExportToExcel("test.xlsx")
			System.Diagnostics.Process.Start("test.xlsx")
		End Sub

		Private Sub DashboardDesigner1_ConfigureDataConnection(ByVal sender As Object, ByVal e As DashboardConfigureDataConnectionEventArgs)
			Dim parameters As ExtractDataSourceConnectionParameters = TryCast(e.ConnectionParameters, ExtractDataSourceConnectionParameters)
			If parameters IsNot Nothing Then
				parameters.FileName = Path.GetFileName(parameters.FileName)
			End If
		End Sub
		Private Sub dashboardDesigner1_CustomizeExportDocument(ByVal sender As Object, ByVal e As DevExpress.DashboardCommon.CustomizeExportDocumentEventArgs) Handles dashboardDesigner1.CustomizeExportDocument
			If e.ExportAction = DashboardExportAction.ExportToExcel Then
				Dim fileStream As FileStream = TryCast(e.Stream, FileStream)
				fileStream.Position = 0

				Dim workbook As New Workbook()
				If e.ExcelExportOptions.Format = ExcelFormat.Xlsx Then
					workbook.LoadDocument(fileStream, DocumentFormat.Xlsx)
				ElseIf e.ExcelExportOptions.Format = ExcelFormat.Xls Then
					workbook.LoadDocument(fileStream, DocumentFormat.Xls)
				ElseIf e.ExcelExportOptions.Format = ExcelFormat.Csv Then
					workbook.LoadDocument(fileStream, DocumentFormat.Csv)
				End If
				For Each sheet As Worksheet In workbook.Worksheets
					' Export to CSV without images, cell merging and coloring. 
					If e.ExcelExportOptions.Format = ExcelFormat.Csv Then
						sheet.Rows.Insert(0)
						Dim textCell As Cell = sheet.Cells(0, 0)
						textCell.Value = "Custom Document Header"
					' Export to the Excel workbook with images, cell merging and coloring.
					Else
						sheet.Rows.Insert(0, 3)
						sheet.Pictures.AddPicture(My.Resources.dxLogo, sheet.Range.FromLTRB(0, 0, 5, 2), True)
						Dim textCell As Cell = sheet.Cells(0, 5)
						textCell.Value = "Custom Document Header"

						sheet.MergeCells(sheet.Range.FromLTRB(5, 0, 8, 0))
						Dim formatting As Formatting = textCell.BeginUpdateFormatting()
						formatting.Fill.BackgroundColor = Color.LightBlue
						textCell.EndUpdateFormatting(formatting)
						textCell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
					End If
				Next sheet
				fileStream.Position = 0
				fileStream.SetLength(0)
				If e.ExcelExportOptions.Format = ExcelFormat.Xlsx Then
					workbook.SaveDocument(fileStream, DocumentFormat.Xlsx)
				ElseIf e.ExcelExportOptions.Format = ExcelFormat.Xls Then
					workbook.SaveDocument(fileStream, DocumentFormat.Xls)
				ElseIf e.ExcelExportOptions.Format = ExcelFormat.Csv Then
					workbook.SaveDocument(fileStream, DocumentFormat.Csv)
				End If
			End If
		End Sub
	End Class
End Namespace
