
namespace ExcelAddIn1
{
    using System;
    using Microsoft.Office.Tools.Ribbon;
    using Excel = Microsoft.Office.Interop.Excel;

    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonRandomNum_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range firstRow = activeSheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeSheet.get_Range("A1");
            newFirstRow.Value2 = (new Random()).Next(100).ToString();
        }
    }
}
