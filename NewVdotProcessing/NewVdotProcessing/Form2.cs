/* Author: Matthew Baker
 * Purpose: High level review is done with form 2. It will run the Access macro in the GetAll database.
 * Date Created: April 1, 2017
 * Date Modified: October 27, 2017 -  Form 2 is no longer used. High Level is done at the same time as the other processing.
 */


using System;
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

    public partial class Form2 : Form
    {

        private string path;
        private string ExecutingPath = "";
        bool choseFolder = false;

        public Form2()
        {
            InitializeComponent();
            ExecutingPath = Environment.CurrentDirectory;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {

                // Get the Path of Where the user wants the new Database
                path = folderBrowserDialog1.SelectedPath;


                // Set Label3 to display the path the the user chose
                label3.Text = path;


                // The File Path with the Excel Scripts
                string AccessPath = ExecutingPath + @"\ReqFiles\GetAll.accdb";


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
                System.IO.File.Copy(AccessPath, DestFile, false);


                choseFolder = true;

            }

            // Do the Same thing as in Form1
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Start the High Level Process
            if (choseFolder == true)
            {
                // Start Access with the new database with the Scripts
                Access.Application oAccess = new Access.Application();
                oAccess.Visible = true;
                oAccess.OpenCurrentDatabase(path + @"\" + textBox1.Text + ".accdb", false, "");
                try
                {
                    // Run the Start Macro in the Database to execute the scripts
                    oAccess.DoCmd.RunMacro("HighLevelMacro");
                    oAccess.DoCmd.Quit(Access.AcQuitOption.acQuitSaveAll);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    MessageBox.Show("Choose Folder To Restart", "Error", MessageBoxButtons.OK);
                }
                // Release Resources 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                oAccess = null;

                // Let User know that the program is done 
                DialogResult result = MessageBox.Show("High-Level Completed", "High-Level", MessageBoxButtons.OK);
                choseFolder = false;
            }
        }
    }
}
