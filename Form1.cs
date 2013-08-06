using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                String inputFile = @"D:\Excel\Input.xlsx";

                Excel.Application oXL = new Excel.Application();


#if DEBUG
                oXL.Visible = true;
                oXL.DisplayAlerts = true;
#else
                oXL.Visible = false; 
                oXL.DisplayAlerts = false;
#endif


                //Open the Excel File
                Excel.Workbook oWB = oXL.Workbooks.Open(inputFile);

                String SheetName = "ExperimentSheet";
                Excel._Worksheet oSheet = oWB.Sheets[SheetName];
                //oSheet = oWB.ActiveSheet;
                //oSheet = oWB.Sheets.[1];


                //We already know the Address Range of the cells

                String start_range = "A2";
                String end_range = "A11";


                Object[,] values = oSheet.get_Range(start_range, end_range).Value2;

                int t = values.GetLength(0);
                for (int i = 1; i <= values.GetLength(0); i++)
                {
                    String val = values[i, 1].ToString();
                }


                //    Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                //    int lastRow = last.Row;
                //    int lastColumn = last.Column;
                //    Excel.Range range = oSheet.get_Range("A1", last);


                oXL.Quit();

                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                oSheet = null;
                oWB = null;
                oXL = null;
                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
            }
            catch (Exception ex)
            {
                String errorMessage = "Error reading the Excel file : " + ex.Message;
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        // Return the column name for this column number.
        private string ColumnNumberToName(int col_num)
        {
            // See if it's out of bounds.
            if (col_num < 1) return "A";

            // Calculate the letters.
            string result = "";
            while (col_num > 0)
            {
                // Get the least significant digit.
                col_num -= 1;
                int digit = col_num % 26;

                // Convert the digit into a letter.
                result = (char)((int)'A' + digit) + result;

                col_num = (int)(col_num / 26);
            }
            return result;
        }
    }
}
