using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace WindowsFormsApplication14
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            string fileName = "records.xls";
            string fileName2 = "report.xls";
            string sourcePath = "c:/Users/dimitris/Desktop/first project";
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            int x;
           
            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oWB.SaveCopyAs("C:/Users/dimitris/Desktop/first project/report.xls");





            }

            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
            string targetPath = "c:/Users/dimitris/Desktop/first project";

            // Use Path class to manipulate file and directory paths. 
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName2);


            System.IO.File.Copy(sourceFile, destFile, true);

            /* oXL = new Excel.Application();
               oWB = oXL.Workbooks.Open(sourceFile);
               oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item(1);
               oWB.Close(true, null, null);
               oRng = oSheet.UsedRange;
               for (rCnt = 1; rCnt <= oRng.Rows.Count; rCnt++)
               {
                   for (cCnt = 1; cCnt <= oRng.Columns.Count; cCnt++)
                   {
                       str = (string)(oRng.Cells[rCnt, cCnt] as Excel.Range).Value2;
                       MessageBox.Show(str);
                   }
               }
               oWB.Close(true, null, null);
               oXL.Quit();
               releaseObject(oSheet);
               releaseObject(oWB);
               releaseObject(oXL);

           }
           private void releaseObject(object obj)
           {
               try
               {
                   System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                   obj = null;
               }
               catch (Exception ex)
               {
                   obj = null;
                   MessageBox.Show("Unable to release the Object " + ex.ToString());
               }
               finally
               {
                   GC.Collect();
               }
           } 
               */
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            String str;
            int rCnt = 0;
            int cCnt = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sourceFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            List<employment> list = new List<employment>(range.Rows.Count);
            employment k;
            k::id = 5;
            list.Add(k);
            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {




                    str = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    MessageBox.Show(str);
                }
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
        }
        class employment {
            int id;
            string name;
            string surname;
        
        
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {


      
           

        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
           
            string id;
            if (e.KeyCode == Keys.Enter)
            {
                
                e.Handled = true;
                
                id = textBox1.Text;
                if (id== "221")
                {
                    MessageBox.Show("vrethike o ipallilos");

                }
                textBox1.Text = String.Empty;
                employment x;
                
            }
            
        }
        
    }
    
}
