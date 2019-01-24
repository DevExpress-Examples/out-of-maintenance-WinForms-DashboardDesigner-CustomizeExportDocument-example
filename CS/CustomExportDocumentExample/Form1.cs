using DevExpress.DashboardCommon;
using DevExpress.Spreadsheet;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace CustomExportDocumentExample
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
            dashboardDesigner1.CreateRibbon();
            dashboardDesigner1.ConfigureDataConnection += DashboardDesigner1_ConfigureDataConnection;
            dashboardDesigner1.LoadDashboard("Dashboard.xml");
            dashboardDesigner1.ExportToExcel("test.xlsx");
            System.Diagnostics.Process.Start("test.xlsx");
        }

        private void DashboardDesigner1_ConfigureDataConnection(object sender, DashboardConfigureDataConnectionEventArgs e)
        {
            ExtractDataSourceConnectionParameters parameters = e.ConnectionParameters as ExtractDataSourceConnectionParameters;
            if (parameters != null)
                parameters.FileName = Path.GetFileName(parameters.FileName);
        }
        private void dashboardDesigner1_CustomizeExportDocument(object sender, DevExpress.DashboardCommon.CustomizeExportDocumentEventArgs e)
        {
            if (e.ExportAction == DashboardExportAction.ExportToExcel)
            {
                FileStream fileStream = e.Stream as FileStream;
                fileStream.Position = 0;

                Workbook workbook = new Workbook();
                if (e.ExcelExportOptions.Format == ExcelFormat.Xlsx)
                    workbook.LoadDocument(fileStream, DocumentFormat.Xlsx);
                else if (e.ExcelExportOptions.Format == ExcelFormat.Xls)
                    workbook.LoadDocument(fileStream, DocumentFormat.Xls);
                else if (e.ExcelExportOptions.Format == ExcelFormat.Csv)
                    workbook.LoadDocument(fileStream, DocumentFormat.Csv);
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Export to CSV without images, cell merging and coloring. 
                    if (e.ExcelExportOptions.Format == ExcelFormat.Csv)
                    {
                        sheet.Rows.Insert(0);
                        Cell textCell = sheet.Cells[0, 0];
                        textCell.Value = "Custom Document Header";
                    }
                    // Export to the Excel workbook with images, cell merging and coloring.
                    else
                    {
                        sheet.Rows.Insert(0, 3);
                        sheet.Pictures.AddPicture(Properties.Resources.dxLogo, sheet.Range.FromLTRB(0, 0, 5, 2), true);
                        Cell textCell = sheet.Cells[0, 5];
                        textCell.Value = "Custom Document Header";

                        sheet.MergeCells(sheet.Range.FromLTRB(5, 0, 8, 0));
                        Formatting formatting = textCell.BeginUpdateFormatting();
                        formatting.Fill.BackgroundColor = Color.LightBlue;
                        textCell.EndUpdateFormatting(formatting);
                        textCell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    }
                }
                fileStream.Position = 0;
                fileStream.SetLength(0);
                if (e.ExcelExportOptions.Format == ExcelFormat.Xlsx)
                    workbook.SaveDocument(fileStream, DocumentFormat.Xlsx);
                else if (e.ExcelExportOptions.Format == ExcelFormat.Xls)
                    workbook.SaveDocument(fileStream, DocumentFormat.Xls);
                else if (e.ExcelExportOptions.Format == ExcelFormat.Csv)
                    workbook.SaveDocument(fileStream, DocumentFormat.Csv);
            }
        }
    }
}
