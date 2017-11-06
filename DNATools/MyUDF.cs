using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DNATools
{
    public class MyUDF
    {
        [ExcelCommand(Description = "Test Button 1", MenuName = "Excel-DNA", MenuText = "Test Button 1")]
        public static void sutTest1()
        {
            MessageBox.Show("Test Button 1");
        }

        [ExcelCommand(Description = "Test Button 2", MenuName = "Excel-DNA", MenuText = "Test Button 2")]
        public static void sutTest2()
        {
            MessageBox.Show("Test Button 2");
        }

        private static string string_0 = "%$#@REWQK6543JHGF432*&";
    }
}
