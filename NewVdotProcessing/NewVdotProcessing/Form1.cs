/*
 * Author: Matthew Baker
 * Program: NewVdotProcessing
 * Date: April 1, 2017
 * Modified: October 27, 2017
 * Version : 3.0
 *      3.0 - Rewrite of code into better functions for easier understanding. 
 *            Place Access, Excel, and Instructions into Solution Folder to prevent having to change paths again.
 *            Some Try/Catch blocks to prevent program from quitting if Access and Excel macros mess up or are cancelled.
 *      2.0 - Change Access and Excel Paths to reflect new location of database and macros  
 *      1.0 - Initial Layout of Figuring Out Program Flow.
 * Details: C# program to run Access and Excel Scripts on Databases for VDOT.
 * 
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
using Word = Microsoft.Office.Interop.Word;

namespace NewVdotProcessing
{
    public partial class Form1 : Form
    {
        // Global Variables
        private string path;
        private bool choseFolder = false;
        private string ExecutingPath = "";
        private static string AccessFile = "GetAll.accdb";
        private static string Instruction = "Processing Instructions.docx";
        private List<string> OldFiles;
        private List<string> ReqFiles;

        /* Function: Form1
         * Purpose: Initialize form, set the Executing path to the path the program is executing in,
         *          Populate the oldfile list of previously outputted file, and populate the required
         *          file list.
         */

        public Form1()
        {
            InitializeComponent();
            ExecutingPath = Environment.CurrentDirectory;
            PopulateOldFileList();
            PopulateReqFileList();
        }


        /* Function: Browser_Click
         * Purpose: Have the user select the folder they want to place the database and final excel files in.
         *          DeleteMacros in another folder if another folder was chosen before. Get the path and display
         *          it on the form. Rename the old files if there are any from before and set choseFolder to true
         *          so that the start button can be clicked.
         */

        private void Browser_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                // If a folder was chosen previously and Startbutton was not clicked.
                if(choseFolder == true)
                {
                    // DeleteMacros from previous folder
                    DeleteMacros();
                }


                // Get the Path of Where the user wants the new Database
                path = folderBrowserDialog1.SelectedPath;


                // Set Label3 to display the path the the user chose
                label3.Text = path;


                // If there are old files from previous processing, Rename them with time in appended infront of name
                RenameOldFiles();

                // Set choseFolder to true so StartButton can be pressed.
                choseFolder = true;
            }
        }


        /* Function: Start_Click
         * Purpose: Start_click will copy the GetAll.accdb and rename it to the user selected name. Then it will call
         *          CopyRequiredFiles, RunAccess, RunExcel, and DeleteMacros functions. A messagebox is then shown to
         *          tell the user that the program is done processing the databases and that the excel files are ready
         *          for viewing.
         */

        private void Start_Click(object sender, EventArgs e)
        {
            if (choseFolder == true)
            {             
                // Combine the path selected by user with the name of the database without extension
                string DestFile = System.IO.Path.Combine(path, textBox1.Text);

                // Add Extension to user selected filename
                DestFile += ".accdb";


                //Check if chosen Access file name Exists. If it does append time to front of old file
                if (System.IO.File.Exists(DestFile))
                {
                    System.IO.File.Move(DestFile, path + @"\" + DateTime.Now.ToString("HHmmss") + textBox1.Text + @".accdb");
                }

                // Copied the required files to the user selected folder
                CopyRequiredFiles(DestFile);

                try
                {
                    // Run Access Macro
                    RunAccess(path + @"\" + textBox1.Text + ".accdb");

                    // Run the Excel Macros
                    RunExcel();
                }
                catch (System.Runtime.InteropServices.COMException )
                {
                    // If Access or Excel fails or is cancelled, Alert user, delete the macros from the folder, and set choseFolder to false

                    MessageBox.Show("Please choose the folder again to restart.", " An Error Occurred", MessageBoxButtons.OK);
                    DeleteMacros();
                    choseFolder = false;
                    return;
                }

                // Delete files that are not needed in the folder
                DeleteMacros();


                // Let User know that the program is done and set choseFolder to false
                DialogResult result = MessageBox.Show("Processing Has Been Completed", "VDOT Processing", MessageBoxButtons.OK);
                choseFolder = false; 
            }
        }


        /* Function : RunAccess
         * Purpose : RunAccess will create a new instance of Access and open the user named database that is passed into it.
         *           It will then run the StartMacro that is inside of that access database. Once the macro is done, the program
         *           quit access, release the resources and exit the function.
         */

         private void RunAccess(string database)
        {
            // Start Access with the new database with the Scripts
            Access.Application oAccess = new Access.Application();
            oAccess.Visible = true;
            oAccess.OpenCurrentDatabase(database, false, "");

            try
            {
                // Run the Start Macro in the Database to execute the scripts
                oAccess.DoCmd.RunMacro("StartMacro");
                oAccess.DoCmd.Quit(Access.AcQuitOption.acQuitSaveAll);
            }
            catch(System.Runtime.InteropServices.COMException ) 
            {
                // If the macro is cancelled or there is some type of failure, quit access, release resources and throw exception to calling function
                oAccess.DoCmd.Quit(Access.AcQuitOption.acQuitSaveNone);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                oAccess = null;
                throw;
            }

            // Release Resources 
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
            oAccess = null;
        }


        /* Function: RunExcel
         * Purpose: RunExcel will call the excelMacro function to run the required macros.
         *          The ACP and Shoulder macros will always run. The CRCP and JRCP macros
         *          will only run if there are excel files that were created by the Access
         *          macro ran from RunAccess.
         */

        private void RunExcel()
        {
            // ACP Macro
            excelMacros(path + @"\QAResultsACPMac.xlsm", "QABeautify");

            // CRCP Macro
            if (System.IO.File.Exists(path + @"\CRCP_QA_Results.xlsx"))
            {
                excelMacros(path + @"\CRCMac.xlsm", "QABeautify");
            }

            // JRCP Macro
            if (System.IO.File.Exists(path + @"\JRCP_QA_Results.xlsx"))
            {
                excelMacros(path + @"\JCPMac.xlsm", "QABeautify");
            }

            //Shoulder Macro
            excelMacros(path + @"\ShoulderComparisonMac.xlsm", "Shoulderup");
        }


        /* Function: excelMacros
         * Purpose: excelMacros will take the name of the macro file and the name of the macro. It will
         *          start an instance of excel and run the macro from that file. Once finished, it will
         *          close excel and release the resources excel required.
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

            try
            {
                // Run the Excel Macro
                oExcel.Run(macName);
                oBook.Close(false, oMissing, oMissing);
                oExcel.Quit();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // If there is a failure or the user cancelled, close and quit Excel. Then release resources and throw exception to calling function.
                oBook.Close(false, oMissing, oMissing);
                oExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
                oBook = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
                oBooks = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                oExcel = null;
                throw;
            }

            //Release Resources
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
            oBook = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
            oBooks = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
            oExcel = null;
        }


        /* Function: RenameOldFiles
         * Purpose: RenameOldFiles will rename the files that may be in the folder from a previous processing attempt
         *          so that no files are accidentally destroyed. Each file from the OldFiles list will be check. If
         *          the user has one of the files open, the function will show a message box alerting the user to close
         *          it and recall itself so that the file is not skipped in the renaming process.
         */ 

        private void RenameOldFiles()
        {
            // Run through OldFile list
            for(int i=0; i< OldFiles.Count; i++)
            {
                try
                {
                    // Rename file if it exists to have time appended to the front of the filename.
                    if (System.IO.File.Exists(path + @"\" + OldFiles[i]))
                    {
                        System.IO.File.Move(path + @"\" + OldFiles[i], path + @"\" + DateTime.Now.ToString("HHmmss") + OldFiles[i]);
                    }
                } catch(Exception)
                {

                    // if the file is opened, notify user and call RenameOldFiles again until the file is renamed.
                    if (MessageBox.Show("Close " + path + @"\" + OldFiles[i] + " To Continue Processing", "Error", MessageBoxButtons.OK) == DialogResult.OK)
                        {
                            RenameOldFiles();
                        }
                }
            }
        }


        /* Function: deleteFile
         * Purpose: deleteFile will delete any filename in the user selected path that is passed into the function. 
         */ 

        private void deleteFile(string fileName)
        {
            // Get the FullPath of the File you want to delete
            string FullPath = path + @"\" + fileName;

            // If the file exists, delete that file
            if (System.IO.File.Exists(FullPath))
            {
                System.IO.File.Delete(FullPath);
            }
        }


        /* Function: InstructButton_Click
         * Purpose: The InstructButton_Click function will open up a word document with instructions on how to use the
         *          program and for any troubleshooting that the user needs to do._
         */

        private void InstructButton_Click(object sender, EventArgs e)
        {
            Word.Application oWord = new Word.Application();
            oWord.Visible = true;
            Console.WriteLine(ExecutingPath + @"\ReqFiles\" + Instruction);
            oWord.Documents.Open(ExecutingPath + @"\ReqFiles\" + Instruction);
        }


        /* Function: button2_Click
         * Purpose: button2_Click will open up Form2 for a highlevel review
         * 
         * NOTE: button2_Click is not used anymore since the highlevel review is done from the
         *       Access Macro in the RunAccess function.
         */

        private void button2_Click(object sender, EventArgs e)
        {
           // this.Hide();
            Form2 Highlevel= new Form2();
            Highlevel.Show();
        }


        /* Function: PopulateOldFileList
         * Purpose: PopulateOldFileList will create the OldFiles list and add the filenames that are created
         *          by an old processing attempt in the same user selected folder.
         */

        private void PopulateOldFileList()
        {
            OldFiles = new List<string>();
            OldFiles.Add("ACP_QA_Results.xlsx");
            OldFiles.Add("CRCP_QA_Results.xlsx");
            OldFiles.Add("JRCP_QA_Results.xlsx");
            OldFiles.Add("ACP_FOR_COMPARE.xlsx");
            OldFiles.Add("CRCP_FOR_COMPARE.xlsx");
            OldFiles.Add("JRCP_FOR_COMPARE.xlsx");
            OldFiles.Add("ShoulderComparison.xlsx");
            OldFiles.Add("InHouseComparison_ACP.xlsx");
        }


        /* Function: PopulateReqFileList
         * Purpose: PopulateReqFileList will create the ReqFile list and add the files that will need to be copied
         *          from the ReqFiles folder of the executing directory to the user selected folder.
         */

        private void PopulateReqFileList()
        {
            ReqFiles = new List<string>();
            ReqFiles.Add(AccessFile);
            ReqFiles.Add("QAResultsACPMac.xlsm");
            ReqFiles.Add("CRCMac.xlsm");
            ReqFiles.Add("JCPMac.xlsm");
            ReqFiles.Add("ShoulderComparisonMac.xlsm");
        }


        /* Function: CopyRequiredFiles
         * Purpose: CopyRequiredFiles will copy the macro and access files from the ReqFiles folder of the executing
         *          directory to the user selected directory. The function will rename the GetAll.accdb file to the
         *          destFile string that is passed into the function during the copied.
         */

        private void CopyRequiredFiles(string destFile)
        {
            for (int i = 0; i< ReqFiles.Count; i++)
            {
                // if GetAll.accdb is being copied, then copy it to the directory as user selected name.
                if(ReqFiles[i] == AccessFile)
                {
                    System.IO.File.Copy(ExecutingPath + @"\ReqFiles\" + AccessFile, destFile, false);
                }
                else
                {
                    System.IO.File.Copy(ExecutingPath + @"\ReqFiles\" + ReqFiles[i], path + @"\" + ReqFiles[i], true);
                }
            }
        }


        /* Function: DeleteMacros
         * Purpose: The DeleteMacros function will delete the files that are not needed in the user selected directory
         *          once the program is done.
         */

        private void DeleteMacros()
        {
            // Delete Macro Files from the folder
            deleteFile("QAResultsACPMac.xlsm");
            deleteFile("CRCMac.xlsm");
            deleteFile("JCPMac.xlsm");
            deleteFile("ShoulderComparisonMac.xlsm");
            deleteFile("*.bas");
        }
    }
}
