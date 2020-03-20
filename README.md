# How to export the WinForms DataGrid StackedHeaders with different back color to Excel and PDF?

## About the sample
This example illustrates how to export the WinForms DataGrid StackedHeaders with different back color to Excel and PDF

By default, actual value only will be exported to Pdf. You can customize StackedHeader cellstyle by using PdfGridCellStyle instance through CellExporting event in PdfExportingOptions.

```C#
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
```


Exporting SfDataGrid StackedHeaders with different back colors by using ExcelExportingOptions by defining the cell style for the StackedHeader cell range. 

```C#

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
```

## Requirements to run the demo
Visual Studio 2015 and above versions
