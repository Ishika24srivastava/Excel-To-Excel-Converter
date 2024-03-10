using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _006_HierarchyConverter_V4
{
    public partial class Form1 : Form
    {
        public static string inputPath = "";
        public static string validationPath = "";
        public static string selectedFolderPath = "";
        public static string jobCodePath = "";
        public static string maximoJobpath="";
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            SetterToInitialValues();

        }
        private void InputBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                inputPath = openFileDialog.FileName;
                textBox1.Text = inputPath;
                VerificationBtn.Enabled = true;
                RestartBtn.Enabled = true;
            }
        }
        private void OutputBtn_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                string selectpath = inputPath.Remove(inputPath.LastIndexOf("\\"));
                // folderDialog.SelectedPath = @"C:\";
                folderDialog.SelectedPath = selectpath;

                // Show the dialog and check if the user clicked OK
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    // The selected folder path is in folderDialog.SelectedPath
                    string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    selectedFolderPath = folderDialog.SelectedPath + "\\Full_Output_" + currentDateTime + ".xlsx";
                    textBox2.Text = folderDialog.SelectedPath;
                    CnvrtBtn.Enabled = true;
                }
            }
        }

        private void CnvrtBtn_Click(object sender, EventArgs e)
        {
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string excelFileName = "Jobs Sheet.xlsx";
            string JobSheetPath = Path.Combine(exeDirectory, excelFileName);
            if (File.Exists(JobSheetPath))
            {
                jobCodePath = JobSheetPath;
                InputBtn.Enabled = false;
                CnvrtBtn.Enabled = false;
                OutputBtn.Enabled = false;
                RestartBtn.Enabled = false;
                VerificationBtn.Enabled = false;
                progressBar.Visible = true;
                Controller2.Start(progressBar, label1);
            }
            else
            {
                MessageBox.Show($"There might be two possible reasons.\n1. The Job Sheet file might not be located in the same directory as the software is running.\n2. The file name could be different. In such case, consider renaming it to \"Jobs Sheet.xlsx\".", "Job Sheet File is missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            RestartBtn.Enabled = true;
        }
        private void VerificationBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                validationPath = openFileDialog.FileName;
                textBox3.Text = validationPath;
                OutputBtn.Enabled = true;
            }
        }

        private void ExitBtn_Click(object sender, EventArgs e)
        {
            //  ExcelInteropExampl1.ResourceRelease();
            System.Windows.Forms.Application.Exit();
        }

        private void RestartBtn_Click(object sender, EventArgs e)
        {
            RestartBtn.Enabled = false;
            SetterToInitialValues();
        }
        private void SetterToInitialValues()
        {
            CnvrtBtn.Enabled = false;
            OutputBtn.Enabled = false;
            VerificationBtn.Enabled = false;
            RestartBtn.Enabled = false;
            InputBtn.Enabled = true;
            progressBar.Visible = false;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            inputPath = "";
            validationPath = "";
            selectedFolderPath = "";
            label1.Text = "";
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
        }

        private void MaximoJob_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog3=new OpenFileDialog();
            openFileDialog3.Filter = "Excel Files|*.xlsx;*.xls;";
            if (openFileDialog3.ShowDialog() == DialogResult.OK)
            {
                validationPath = openFileDialog3.FileName;
                textBox4.Text = validationPath;
                OutputBtn.Enabled = true;
            }

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}