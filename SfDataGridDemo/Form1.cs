using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Grid;
using Syncfusion.WinForms.DataGrid;
using Syncfusion.WinForms.DataGrid.Enums;
using Syncfusion.WinForms.DataGridConverter;
using Syncfusion.WinForms.DataGridConverter.Events;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataGrid_WF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            OrderInfoCollection orderInfoCollection = new OrderInfoCollection();
            this.sfDataGrid1.DataSource = orderInfoCollection.Orders;

            //Creating object for a stacked header row.
            var stackedHeaderRow1 = new StackedHeaderRow();

            //Adding stacked column to stacked columns collection available in stacked header row object.
            stackedHeaderRow1.StackedColumns.Add(new StackedColumn() { ChildColumns = "OrderID", HeaderText = "Order Details" });
            stackedHeaderRow1.StackedColumns.Add(new StackedColumn() { ChildColumns = "CustomerID,CustomerName,", HeaderText = "Customer Details" });
            stackedHeaderRow1.StackedColumns.Add(new StackedColumn() { ChildColumns = "Country,ShipCity", HeaderText = "City Details" });

            //Adding stacked header row object to stacked header row collection available in SfDataGrid.
            sfDataGrid1.StackedHeaderRows.Add(stackedHeaderRow1);
                      
            this.sfDataGrid1.DrawCell += SfDataGrid1_DrawCell;
        }

        private void SfDataGrid1_DrawCell(object sender, Syncfusion.WinForms.DataGrid.Events.DrawCellEventArgs e)
        {
            if ((e.DataRow as DataRowBase).RowType == RowType.StackedHeaderRow)
            {
                if (e.CellValue.ToString() == "Order Details")
                {
                    e.Style.BackColor = Color.DarkCyan;
                    e.Style.TextColor = Color.White;
                }
                if (e.CellValue.ToString() == "Customer Details")
                {
                    e.Style.BackColor = Color.LightCyan;
                }
                if (e.CellValue.ToString() == "City Details")
                {
                    e.Style.BackColor = Color.DarkGray;
                    e.Style.TextColor = Color.White;
                }
            }
        }

      


        private void Excel_Exporting(object sender, EventArgs e)
        {
            var options = new ExcelExportingOptions();
            var excelEngine = sfDataGrid1.ExportToExcel(sfDataGrid1.View, options);
            var workBook = excelEngine.Excel.Workbooks[0];
            workBook.Worksheets[0].Range["A1:A1"].CellStyle.Color = Color.DarkCyan;
            workBook.Worksheets[0].Range["A1:A1"].CellStyle.Font.Color = ExcelKnownColors.White;
            workBook.Worksheets[0].Range["B1:C1"].CellStyle.Color = Color.LightCyan;
            workBook.Worksheets[0].Range["B1:C1"].CellStyle.Font.Color = ExcelKnownColors.Black;
            workBook.Worksheets[0].Range["D1:E1"].CellStyle.Color = Color.DarkGray;
            workBook.Worksheets[0].Range["D1:E1"].CellStyle.Font.Color = ExcelKnownColors.White;
            workBook.SaveAs("SampleRange.xlsx");
          
        }

  
        private void PDF_Exporting(object sender, EventArgs e)
        {
            PdfExportingOptions options = new PdfExportingOptions();
            options.ExportStackedHeaders = true;
            options.CellExporting += options_CellExporting;
            var document = sfDataGrid1.ExportToPdf(options);
            document.Save("Sample.pdf");
        }

        private void options_CellExporting(object sender, DataGridCellPdfExportingEventArgs e)
        {
            if(e.CellType==ExportCellType.StackedHeaderCell)
            {
               if(e.CellValue.ToString()=="Order Details")
                {
                    
                    var cellStyle = new PdfGridCellStyle();
                    cellStyle.BackgroundBrush = PdfBrushes.DarkCyan;
                    cellStyle.TextBrush = PdfBrushes.White;
                    e.PdfGridCell.Style = cellStyle;
                }
               else if (e.CellValue.ToString() == "Customer Details")
                {

                        var cellStyle = new PdfGridCellStyle();
                        cellStyle.BackgroundBrush = PdfBrushes.LightCyan;
                        e.PdfGridCell.Style = cellStyle;
                }
                else if (e.CellValue.ToString() == "City Details")
                {
                        var cellStyle = new PdfGridCellStyle();
                        cellStyle.BackgroundBrush = PdfBrushes.DarkGray;
                        cellStyle.TextBrush = PdfBrushes.White;
                    e.PdfGridCell.Style = cellStyle;
                }
            }
          
        }
    }
}
