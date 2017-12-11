using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using myExcel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace myHomework
{
    public partial class Ribbon1
    {
        bool state = false;
        DateTime getTimeFirst, getTimeSecond;
        //Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnEncrypt_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng = Globals.ThisAddIn.Application.Selection;
            rng.encrypt(rng.Offset[1, 0]);
        }

        private void btnGetGontent_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng = Globals.ThisAddIn.Application.ActiveCell;
            string web = @"http://www.matrix67.com/blog/feed";
            rng.writhContent(web);
        }

        private void getTimeDiff_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng = Globals.ThisAddIn.Application.ActiveCell;
            if (state == false)
            {
                getTimeFirst = myExcelClass.getTimeFirstStr();
                rng.Value = getTimeFirst;
                state = !state;
            }
            else
            {
                getTimeSecond = myExcelClass.getTimeFirstStr();
                string timeDiff = myExcelClass.getTimeDiff(getTimeFirst);
                rng.Offset[1, 0].Value = getTimeSecond;
                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                workbook.Save();
                rng.Offset[2, 0].Value = getTimeDiff;
                state = !state;
            }
        }
    }
}

