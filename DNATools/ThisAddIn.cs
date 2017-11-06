using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace DNATools
{
    public class ThisAddIn:IExcelAddIn
    {
        public void AutoOpen()
        {
            try
            {
                object xlApp = ExcelDnaUtil.Application;
                Type xlType = xlApp.GetType();
                CommandBars customBars = (CommandBars)xlType.InvokeMember("Commandbars", BindingFlags.GetProperty, null, xlApp, null);
                CommandBar customBar = customBars["Excel-DNA"];
                CommandBarButton sutBtn1 = (CommandBarButton)customBar.Controls["Test Button 1"];
                if (sutBtn1 != null)
                {
                    sutBtn1.FaceId = 69;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
        public void AutoClose()
        {
            //MessageBox.Show("Close");
        }
    }
}
