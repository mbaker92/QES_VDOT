/*
 * Author: Matthew Baker
 * Program: NewVdotProcessing
 * Date: April 1, 2017
 * 
 * Details: C# program to run Access and Excel Scripts on Databases for VDOT.
 * 
 * 
 * TODO : - Uncomment ShoulderMacro Stuff when Shoulder Scripts are ready
 *        - Test Program on Databases with Shoulder Scripts
*/
using System;
using System.Reflection;
using Microsoft.Office.Core;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Access = Microsoft.Office.Interop.Access;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewVdotProcessing
{
    public partial class Form1 : Form
    {
        private string pathfull;
        public string path { get { return pathfull; } set { pathfull = value; } }
        bool choseFolder = false;

        public Form1()
        {
            InitializeComponent();
        }


        /* Browser Click
         *  Will show a folder browse window. If folder was choosen, but Start was not clicked. Delete the Macros
         *  Store the Path
         *  Show path in the window
         *  Rename the old processed excel files if they exist in the folder
         *  Set the path of the excel and access scripts
         *  Combine the path and the user's name of new database
         *  Rename Database in folder if already exists
         *  Copy new database and excel macros to the chosen folder
         *  Set chosenFolder to true 
         */

        private void Browser_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                if(choseFolder == true)
                {
                    deleteFile("QAResultsACPMac.xlsm");
                    deleteFile("CRCMac.xlsm");
                    deleteFile("JCPMac.xlsm");
                    //  deleteFile("ShoulderComparisonMac.xlsm");
                }


                // Get the Path of Where the user wants the new Database
                path = folderBrowserDialog1.SelectedPath;


                // Set Label3 to display the path the the user chose
                label3.Text = path;


                // If there are old files from previous processing, Rename them with time in appended infront of name
                RenameOldFiles();


                // The File Path with the Excel Scripts
                string ExcelPath = " "; // Removed Actual Directory
                string AccessPath = " "; // Removed Actual Directory
                
                
                // Combine the path selected by user with the name of the database without extension
                string DestFile = System.IO.Path.Combine(path, textBox1.Text);
  
                // Add Extension
                DestFile += ".accdb";


                //Check if File Exists. If it does append time to front of old file
                if (System.IO.File.Exists(DestFile))
                {
                    System.IO.File.Move(DestFile, path + @"\" + DateTime.Now.ToString("HHmmss") + textBox1.Text + @".accdb");
                }


                // Copy Access Database with scripts over to Location specifed by user with the extension and changed name.
                System.IO.File.Copy(AccessPath, DestFile , false);
                
                
                // Copy Excel Files over
                System.IO.File.Copy(ExcelPath + "QAResultsACPMac.xlsm", path + @"\QAResultsACPMac.xlsm", true);
                System.IO.File.Copy(ExcelPath + "CRCMac.xlsm", path + @"\CRCMac.xlsm", true);
                System.IO.File.Copy(ExcelPath + "JCPMac.xlsm", path + @"\JCPMac.xlsm", true);
                // System.IO.File.Copy(ExcelPath + "ShoulderComparisonMac.xlsm", path + @"\ShoulderComparisonMac.xlsm", true);


                choseFolder = true;

            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        /* Start_Click Will start and run Access Macro.
         * Once Access is done, the function will run the excel macros,
         * delete the excel macro files from the directory, and let the user know its done.
         */

        private void Start_Click(object sender, EventArgs e)
        {
            if (choseFolder == true)
            {
                // Start Access with the new database with the Scripts
                Access.Application oAccess = new Access.Application();
                oAccess.Visible = true;
                oAccess.OpenCurrentDatabase(path + @"\" + textBox1.Text + ".accdb", false, "");

                // Run the Start Macro in the Database to execute the scripts
                oAccess.DoCmd.RunMacro("StartMacro");
                oAccess.DoCmd.Quit(Access.AcQuitOption.acQuitSaveAll);

                // Release Resources 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                oAccess = null;


                // Run the Excel Macros
                excelMacros(path + @"\QAResultsACPMac.xlsm", "QABeautify");
                excelMacros(path + @"\CRCMac.xlsm", "QABeautify");
                excelMacros(path + @"\JCPMac.xlsm", "QABeautify");
                //  excelMacros(path + @"\ShouderComparisonMac.xlsm", "Shoulderup");


                // Delete Macro Files from the folder
                deleteFile("QAResultsACPMac.xlsm");
                deleteFile("CRCMac.xlsm");
                deleteFile("JCPMac.xlsm");
                //  deleteFile("ShoulderComparisonMac.xlsm");
                deleteFile("*.bas");


                // Let User know that the program is done 
                DialogResult result = MessageBox.Show("Processing Has Been Completed", "VDOT Processing", MessageBoxButtons.OK);
                choseFolder = false; 
            }
        }

        /*
         * excelMacros will start and run the excel macros 
         */


        private void excelMacros(string excelMac, string macName)
        {
            // Open Excel
            object oMissing = System.Reflection.Missing.Value;
            Excel.Application oExcel = new Excel.Application();
            oExcel.Visible = true;
            Excel.Workbooks oBooks = oExcel.Workbooks;
            Excel._Workbook oBook = null;
            oBook = oBooks.Open(excelMac, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            // Run the Excel Macro
            oExcel.Run(macName);
            oBook.Close(false, oMissing, oMissing);
            oExcel.Quit();


            //Release Resources
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
            oBook = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
            oBooks = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
            oExcel = null;

        }


        /* RenameOldFiles()
         * Checks if old processed excel files exist in the folder and appends the time to the front of the files
         */ 

        private void RenameOldFiles()
        {
            if(System.IO.File.Exists(path + @"\ACP_QA_Results.xlsx"))
            {
                System.IO.File.Move(path + @"\ACP_QA_Results.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "ACP_QA_Results.xlsx");
            }

            if (System.IO.File.Exists(path + @"\CRCP_QA_Results.xlsx"))
            {
                System.IO.File.Move(path + @"\CRCP_QA_Results.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "CRCP_QA_Results.xlsx");
            }

            if (System.IO.File.Exists(path + @"\JRCP_QA_Results.xlsx"))
            {
                System.IO.File.Move(path + @"\JRCP_QA_Results.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "JRCP_QA_Results.xlsx");
            }

            if (System.IO.File.Exists(path + @"\TableForCompareACP.xlsx"))
            {
                System.IO.File.Move(path + @"\TableForCompareACP.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "TableForCompareACP.xlsx");
            }

            if (System.IO.File.Exists(path + @"\TableForCompareCRC.xlsx"))
            {
                System.IO.File.Move(path + @"\TableForCompareCRC.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "TableForCompareCRC.xlsx");
            }

            if (System.IO.File.Exists(path + @"\TableForCompareJCP.xlsx"))
            {
                System.IO.File.Move(path + @"\TableForCompareJCP.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "TableForCompareJCP.xlsx");
            }

            if (System.IO.File.Exists(path + @"\ShoulderComparison.xlsx"))
            {
                System.IO.File.Move(path + @"\ShoulderComparison.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "ShoulderComparison.xlsx");
            }

            if (System.IO.File.Exists(path + @"\QESInHouseComparisons.xlsx"))
            {
                System.IO.File.Move(path + @"\QESInHouseComparisons.xlsx", path + @"\" + DateTime.Now.ToString("HHmmss") + "QESInHouseComparisons.xlsx");
            }
        }


        /* deleteFile
         *  Will delete a file based on a string passed to it if it exists in the folder
         */ 

        private void deleteFile(string fileName)
        {
            // Get the FullPath of the File you want to delete
            String FullPath = path + @"\" + fileName;

            // If the file exists, delete that file
            if (System.IO.File.Exists(FullPath))
            {
                System.IO.File.Delete(FullPath);
            }
        }


    }
}
