using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace _006_HierarchyConverter_V4




{
    internal class Controller2
    {
        static Excel.Workbook OutputWorkbook;
        static Excel.Worksheet OutPutWorksheet;
        static Excel.Worksheet InputWorksheet;
        static Excel.Workbook InputWorkbook;
        static Excel.Workbook ValidationWorkbook;
        static Excel.Worksheet ValidationWorksheet;
        static Excel.Workbook destWorkbook;
        static Excel.Worksheet destWorksheet;
        static Excel.Application excelApp;

        const string ExcelClassName = "XLMAIN";

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        public static void Start(ProgressBar progressbar, System.Windows.Forms.Label label)
        {
            int Iteration = 0;
            bool sysCome = false;
            bool assemblyCome = false;
            bool elementsCome = false;
            bool hasValue = false;
            bool duplicateValue1 = false;
            bool duplicateValue2 = false;
            bool groupLevel2Come = false;
            bool inputWorksheetOpened = false;
            bool inputWorkbookOpenend = false;
            bool outputWorksheetOpened = false;
            bool outputWorkbookOpened = false;
            bool verificationWorksheetOpened = false;
            bool verificationWorkBookOpened = false;
            bool tryHelper = false;
            string[] Code;
            string[] Path;
            string[] Name;
            string[] SequenceNumber;
            string[] FunctionType;
            string[] ComponentType;
            string[] ComponentClass;
            string[] Maker;
            string[] Model;
            string[] SerialNumber;
            string[] Criticality;
            string[] MaximoEquipment;
            string[] MaximoEquipmentDescription;
            string[] wrongRows;
            string[] maximoColor;
            string[] makerColor;
            string[] modelColor;
            string[] serialColor;
            string[] componentColor;
            int wrongRowsNumber = 0;
            uint excelProcessId;
            int maximoerrorrownumber = 0;
            int numRowsOfV;

            List<int> splitList = new List<int>();
            Dictionary<string, List<int>> jobCodeSheet = new Dictionary<string, List<int>>();
            Dictionary<string, List<int>> jobCmpclsSheet = new Dictionary<string, List<int>>();
            Dictionary<string, List<int>> outputCodeSheet = new Dictionary<string, List<int>>();
            Dictionary<string, List<int>> outputCmpclsSheet = new Dictionary<string, List<int>>();

            int increment = 0;
            string[] maximoerror;
            object[,] allCellValues;
            int numRowsOfI;
            int numColsOfI;
            int numColsOfV;
            bool helper = false;
            Excel.Range InputusedRange;
            Excel.Range ValidationUsedRange;
            object[,] allCellValuesOfValidationSheet;
            string InputFilePath = Form1.inputPath;
            string OutPutFilePath = Form1.selectedFolderPath;
            string ValidationFilePath = Form1.validationPath;
            string MaximoJobFilePath = Form1.maximoJobpath;
            string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string resourceName = "FinalProject.Template.xlsx";
            try
            {
               
                SaveEmbeddedResourceToFile(resourceName, OutPutFilePath);
                               
            }
            catch (Exception ex)
            {
                MessageBox.Show("an error occured while copying the files" + ex.Message);
            }
            label.Text = "Output File Created";
            label.Text = "Opening Excel Enviornment....";
            label.Text = "Please Wait For a while....";
            progressbar.Value = 1;
            int count = 0;
            uint PID = 0;
            // Create a new Excel application
            Dictionary<string, object> initialCOMAddInStates = new Dictionary<string, object>();

            excelApp = new Excel.Application();
            // Store the initial state of each COM add-in        
            try
            {
                if (TryGetExcelProcessId(out excelProcessId))
                {
                    PID = excelProcessId;
                }
                else
                {
                    MessageBox.Show("Excel is not running or not visible in the taskbar.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while accessing processid: {ex.Message}");
            }
            label.Text = "Excel App Running";
            progressbar.Value = 2;

            try
            {
                Workbook jobWorkbook = excelApp.Workbooks.Open(Form1.jobCodePath);
                Worksheet jobWorksheet = jobWorkbook.Worksheets[1];
                Range jobUsedRange = jobWorksheet.UsedRange;
                jobUsedRange.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;

                // Get the values of all cells in the used range
                object[,] allCellValuesOfJob = (object[,])jobUsedRange.Value;

                // Get the number of rows and columns in the used range
                int numRowsOfJ = allCellValuesOfJob.GetLength(0);
                int numColsOfJ = allCellValuesOfJob.GetLength(1);
                // Opening Files
                OpenFiles();
                // Extracting Range of coloumns and Rows, and values in 2D array
                ExtractRange();

                List<string> componentsNumber = new List<string>(numRowsOfV + 2);
                for (int i = 1; i <= numRowsOfV; i++)
                {
                    object cellValue1 = allCellValuesOfValidationSheet[i, 10];
                    string cellValueDescription = cellValue1?.ToString();
                    componentsNumber.Add(cellValueDescription);
                }

                progressbar.Value = 10;
                int rowsIncrease = 0;
                bool duplicateCode = false;
                bool duplicateName = false;
                string temp = "";
                label.Text = $"Checking the heirerchy problem in rows";
                for (int row = 2; row <= numRowsOfI; row++)
                {
                    duplicateCode = false;
                    duplicateName = false;
                    for (int col = 1; col <= numColsOfI; col++)
                    {
                        // Read the value from the current cell
                        object cellValue = allCellValues[row, col];
                        string cellValueDescription = cellValue?.ToString();
                        if (!String.IsNullOrWhiteSpace(cellValueDescription))
                        {
                            if (col == 7 || col == 9 || col == 12 || col == 15)
                            {
                                if (duplicateCode == true)
                                {
                                    if (rowsIncrease == 0)
                                    {
                                        temp = row.ToString();
                                        rowsIncrease++;
                                    }
                                    else
                                    {
                                        if (temp != row.ToString())
                                        {
                                            temp = row.ToString();
                                            rowsIncrease++;
                                        }
                                    }
                                }
                                duplicateCode = true;
                            }
                            if (col == 8 || col == 11 || col == 14 || col == 17)
                            {
                                if (duplicateName == true)
                                {
                                    if (rowsIncrease == 0)
                                    {
                                        temp = row.ToString();
                                        rowsIncrease++;
                                    }
                                    else
                                    {
                                        if (temp != row.ToString())
                                        {
                                            temp = row.ToString();
                                            rowsIncrease++;
                                        }
                                    }
                                }
                                duplicateName = true;

                            }
                        }
                    }
                }
                label.Text = $"There are heirerchy problem in {rowsIncrease} rows";
                // giving all string size
                Code = new string[numRowsOfI + 2 + rowsIncrease];
                Path = new string[numRowsOfI + 2 + rowsIncrease];
                Name = new string[numRowsOfI + 2 + rowsIncrease];
                SequenceNumber = new string[numRowsOfI + 2 + rowsIncrease];
                FunctionType = new string[numRowsOfI + 2 + rowsIncrease];
                ComponentType = new string[numRowsOfI + 2 + rowsIncrease];
                ComponentClass = new string[numRowsOfI + 2 + rowsIncrease];
                Maker = new string[numRowsOfI + 2 + rowsIncrease];
                Model = new string[numRowsOfI + 2 + rowsIncrease];
                SerialNumber = new string[numRowsOfI + 2 + rowsIncrease];
                MaximoEquipment = new string[numRowsOfI + 2 + rowsIncrease];
                MaximoEquipmentDescription = new string[numRowsOfI + 2 + rowsIncrease];
                Criticality = new string[numRowsOfI + 2 + rowsIncrease];
                wrongRows = new string[numRowsOfI + 2 + rowsIncrease];
                maximoColor = new string[numRowsOfI + 2 + rowsIncrease];
                makerColor = new string[numRowsOfI + 2 + rowsIncrease];
                modelColor = new string[numRowsOfI + 2 + rowsIncrease];
                serialColor = new string[numRowsOfI + 2 + rowsIncrease];
                componentColor = new string[numRowsOfV + 2];
                maximoerror = new string[numRowsOfI + 2 + rowsIncrease];
                label.Text = "Reading the data from input file";
                // Iterate through each row and column to read cell values
                for (int row = 2; row <= numRowsOfI; row++)
                {
                    count++;
                    hasValue = false;
                    duplicateValue1 = false;
                    duplicateValue2 = false;
                    for (int col = 1; col <= numColsOfI; col++)
                    {
                        // Read the value from the current cell
                        object cellValue = allCellValues[row, col];
                        string cellValueDescription = cellValue?.ToString();
                        if (!String.IsNullOrWhiteSpace(cellValueDescription))
                        {
                            if (col == 7 || col == 9 || col == 12 || col == 15)
                            {
                                if (duplicateValue1 == true)
                                {
                                    if (wrongRowsNumber == 0)
                                    {
                                        wrongRows[wrongRowsNumber] = row.ToString();
                                        wrongRowsNumber++;
                                        Iteration++;
                                    }
                                    else
                                    {
                                        if (wrongRows[wrongRowsNumber - 1] != row.ToString())
                                        {
                                            wrongRows[wrongRowsNumber] = row.ToString();
                                            wrongRowsNumber++;
                                            Iteration++;
                                        }
                                    }
                                }
                                duplicateValue1 = true;
                                Code[Iteration] = cellValueDescription;
                                hasValue = true;
                                switch (col)
                                {
                                    case 7:
                                        groupLevel2Come = false; break;
                                    case 9:
                                        sysCome = false; break;
                                    case 12:
                                        assemblyCome = false; break;
                                    case 15:
                                        elementsCome = false; break;
                                }
                            }
                            if (col == 8 || col == 11 || col == 14 || col == 17)
                            {
                                if (!groupLevel2Come && col != 8)
                                {
                                    duplicateValue2 = true;
                                }
                                if (!sysCome && col != 11)
                                {
                                    duplicateValue2 = true;
                                }
                                if (!assemblyCome && col != 14)
                                {
                                    duplicateValue2 = true;
                                }
                                if (!elementsCome && col != 17)
                                {
                                    duplicateValue2 = true;
                                }
                                if (duplicateValue2 == true)
                                {
                                    if (wrongRowsNumber == 0)
                                    {
                                        wrongRows[wrongRowsNumber] = row.ToString();
                                        wrongRowsNumber++;
                                        Iteration++;
                                    }
                                    else
                                    {
                                        if (wrongRows[wrongRowsNumber - 1] != row.ToString())
                                        {
                                            wrongRows[wrongRowsNumber] = row.ToString();
                                            wrongRowsNumber++;
                                            Iteration++;
                                        }
                                    }
                                }
                                duplicateValue2 = true;
                                Name[Iteration] = cellValueDescription;

                                if (cellValueDescription.Contains("E-Motor"))
                                {
                                    ComponentClass[Iteration] = "E-Motor";
                                }
                                else if (cellValueDescription.Contains("Cooler") || cellValueDescription.Contains("Heater"))
                                {
                                    ComponentClass[Iteration] = "Heat Exchanger";
                                }
                                else if (cellValueDescription.Contains("Pump Unit"))
                                {
                                    ComponentClass[Iteration] = "Pump Unit";
                                }
                                else if (cellValueDescription.Contains("Pump"))
                                {
                                    ComponentClass[Iteration] = "Pump";
                                }
                                hasValue = true;
                                switch (col)
                                {
                                    case 8:
                                        groupLevel2Come = false; break;
                                    case 11:
                                        sysCome = false; break;
                                    case 14:
                                        assemblyCome = false; break;
                                    case 17:
                                        elementsCome = false; break;
                                }
                            }
                            if (col == 10 || col == 13 || col == 16)
                            {
                                SequenceNumber[Iteration] = cellValueDescription;
                                hasValue = true;
                                switch (col)
                                {
                                    case 10:
                                        sysCome = false; break;
                                    case 13:
                                        assemblyCome = false; break;
                                    case 16:
                                        elementsCome = false; break;
                                }
                            }
                            if (col == 18)
                            {
                                if (cellValueDescription.Contains(','))
                                {
                                    maximoerror[maximoerrorrownumber] = row.ToString();
                                    maximoerrorrownumber++;
                                }
                                MaximoEquipment[Iteration] = cellValueDescription;
                                hasValue = true;
                            }
                            if (col == 19)
                            {
                                MaximoEquipmentDescription[Iteration] = cellValueDescription;
                                hasValue = true;
                            }
                            if (col == 20)
                            {
                                Maker[Iteration] = cellValueDescription;
                                hasValue = true;
                            }
                            if (col == 21)
                            {
                                Model[Iteration] = cellValueDescription;
                                hasValue = true;
                            }
                            if (col == 22)
                            {
                                SerialNumber[Iteration] = cellValueDescription;
                                hasValue = true;
                            }
                        }
                        if (!groupLevel2Come)
                        {
                            FunctionType[Iteration] = "Group Level 2";
                            Reset();
                        }
                        if (!sysCome)
                        {
                            FunctionType[Iteration] = "System";
                            Reset();
                        }
                        if (!assemblyCome)
                        {
                            FunctionType[Iteration] = "Assembly";
                            Reset();
                        }
                        if (!elementsCome)
                        {
                            FunctionType[Iteration] = "Element";
                            Reset();
                        }
                    }
                    if (hasValue)
                    {
                        Iteration++;
                    }
                    int percentage = 10 + (row * (10 / numRowsOfI));
                    progressbar.Value = percentage;
                }
                InputWorkbook.Close(false);
                Console.WriteLine("All {0} rows successfully read. ", (count + 1));
                Console.WriteLine("Total {0} rows after filter", (Iteration + 1));
                // for path 
                string tempSys = "";
                string tempAsm = "";
                progressbar.Value = 20;
                label.Text = "Comparing the values with Validation sheet";
                for (int i = 0; i < Iteration; i++)
                {
                    if (FunctionType[i] == "System")
                    {
                        tempSys = Name[i];
                    }
                    if (FunctionType[i] == "Assembly")
                    {
                        tempAsm = Name[i];
                        Path[i] = tempSys;
                    }
                    if (FunctionType[i] == "Element")
                    {
                        Path[i] = tempSys + "/" + tempAsm;
                    }
                    string valueToFind = MaximoEquipment[i];

                    if (!String.IsNullOrWhiteSpace(valueToFind))
                    {
                        if (valueToFind.Contains(","))
                        {
                            valueToFind = valueToFind.Remove(valueToFind.IndexOf(","));
                            maximoColor[i] = "Yellow";
                        }
                        int index = componentsNumber.IndexOf(valueToFind);
                        if (index != -1)
                        {
                            componentColor[index + 1] = "Green";
                            object statusObject = allCellValuesOfValidationSheet[index + 1, 16];
                            string statusString = statusObject?.ToString();
                            Criticality[i] = statusString;

                            object makerObject = allCellValuesOfValidationSheet[index + 1, 11];
                            string makerString = makerObject?.ToString();
                            string maker1 = null, maker2 = null;

                            if (!String.IsNullOrWhiteSpace(makerString))
                            {
                                string[] parts = makerString.Split(new string[] { "||" }, StringSplitOptions.None);
                                if (parts.Length >= 2)
                                {
                                    maker1 = String.IsNullOrWhiteSpace(parts[0]) ? null : parts[0];
                                    maker2 = String.IsNullOrWhiteSpace(parts[1]) ? null : parts[1];
                                }
                                else if (parts.Length == 1)
                                {
                                    maker1 = String.IsNullOrWhiteSpace(parts[0]) ? null : parts[0];
                                }
                            }
                            if (String.IsNullOrWhiteSpace(maker1) && String.IsNullOrWhiteSpace(maker2) && !String.IsNullOrWhiteSpace(Maker[i]))
                            {
                                //color will be Blue
                                makerColor[i] = "Blue";
                            }
                            else if ((!String.IsNullOrWhiteSpace(maker1) || !String.IsNullOrWhiteSpace(maker2)) && String.IsNullOrWhiteSpace(Maker[i]))
                            {
                                //color will be orange and value updated
                                Maker[i] = maker1 ?? maker2;
                                makerColor[i] = "Orange";
                            }
                            else if ((!String.IsNullOrWhiteSpace(maker1) || !String.IsNullOrWhiteSpace(maker2)) && !String.IsNullOrWhiteSpace(Maker[i]))
                            {
                                if (maker1 == Maker[i] || maker2 == Maker[i])
                                {
                                    //value will be compared color will be green
                                    makerColor[i] = "Green";

                                }
                                else
                                {
                                    // color will be red
                                    makerColor[i] = "Red";
                                }
                            }

                            object modelObject = allCellValuesOfValidationSheet[index + 1, 20];
                            string modelString = modelObject?.ToString();
                            if (String.IsNullOrWhiteSpace(modelString) && !String.IsNullOrWhiteSpace(Model[i]))
                            {
                                //color will be blue
                                modelColor[i] = "Blue";
                            }
                            else if (!String.IsNullOrWhiteSpace(modelString) && String.IsNullOrWhiteSpace(Model[i]))
                            {
                                //color will be orange and value updated

                                Model[i] = modelString;
                                modelColor[i] = "Orange";
                            }
                            else if (!String.IsNullOrWhiteSpace(modelString) && !String.IsNullOrWhiteSpace(Model[i]))
                            {
                                if (modelString == Model[i])
                                {
                                    //color will be green
                                    modelColor[i] = "Green";
                                    //   Console.WriteLine($"validation string is {modelString} and row is {index + 1} and output model is {Model[i]} and index is {i} ");
                                }
                                else
                                {
                                    // color will be red
                                    modelColor[i] = "Red";
                                }
                            }

                            object serialNumberObject = allCellValuesOfValidationSheet[index + 1, 18];
                            string serialNumberString = serialNumberObject?.ToString();
                            if (String.IsNullOrWhiteSpace(serialNumberString) && !String.IsNullOrWhiteSpace(SerialNumber[i]))
                            {
                                //color will be red
                                serialColor[i] = "Blue";
                            }
                            else if (!String.IsNullOrWhiteSpace(serialNumberString) && String.IsNullOrWhiteSpace(SerialNumber[i]))
                            {
                                //color will be orange and value updated
                                SerialNumber[i] = serialNumberString;
                                serialColor[i] = "Orange";
                            }
                            else if (!String.IsNullOrWhiteSpace(serialNumberString) && !String.IsNullOrWhiteSpace(SerialNumber[i]))
                            {
                                if (serialNumberString == SerialNumber[i])
                                {
                                    //color will be green
                                    serialColor[i] = "Green";
                                }
                                else
                                {
                                    // color will be red
                                    serialColor[i] = "Red";
                                }
                            }
                            // giving function status
                            Excel.Range cell = OutPutWorksheet.Cells[i + 2, 11];
                            if (makerColor[i] == "Green" && modelColor[i] == "Green")
                            {
                                cell.Value = "Details Provided are Correct";
                            }
                        }
                    }
                    int percentage = 20 + (i * (20 / Iteration));
                    progressbar.Value = percentage;
                }

                label.Text = "All values have been compared";
                label.Text = "Data written is in process....";
                progressbar.Value = 40;
                OutPutWorksheet.Cells.NumberFormat = "@";
                Excel.Range columnRange = OutPutWorksheet.Range["I:I"];
                columnRange.NumberFormat = "General";
                // Get a Range object representing the starting cell
                Excel.Range startRange1 = OutPutWorksheet.Range["A2"];
                Excel.Range startRange2 = OutPutWorksheet.Range["B2"];
                Excel.Range startRange3 = OutPutWorksheet.Range["C2"];
                Excel.Range startRange4 = OutPutWorksheet.Range["D2"];
                Excel.Range startRange5 = OutPutWorksheet.Range["E2"];
                Excel.Range startRange6 = OutPutWorksheet.Range["F2"];
                Excel.Range startRange9 = OutPutWorksheet.Range["I2"];
                Excel.Range startRange10 = OutPutWorksheet.Range["J2"];
                Excel.Range startRange12 = OutPutWorksheet.Range["L2"];
                Excel.Range startRange13 = OutPutWorksheet.Range["M2"];
                Excel.Range startRange14 = OutPutWorksheet.Range["N2"];
                Excel.Range startRange15 = OutPutWorksheet.Range["O2"];
                Excel.Range startRange16 = OutPutWorksheet.Range["P2"];
                progressbar.Value = 42;
                // Convert the string array to a transposed 2D array
                object[,] codeArray = new object[Code.Length, 1];
                object[,] codePath = new object[Code.Length, 1];
                object[,] codeName = new object[Code.Length, 1];
                object[,] codeSequence = new object[Code.Length, 1];
                object[,] codeCriticality = new object[Code.Length, 1];
                object[,] codeFunctionType = new object[Code.Length, 1];
                object[,] codeComponentClass = new object[Code.Length, 1];
                object[,] codeMaker = new object[Code.Length, 1];
                object[,] codeModel = new object[Code.Length, 1];
                object[,] codeSerial = new object[Code.Length, 1];
                object[,] codeME = new object[Code.Length, 1];
                object[,] codeMED = new object[Code.Length, 1];
                progressbar.Value = 44;
                for (int i = 0; i < Code.Length; i++)
                {
                    codeArray[i, 0] = Code[i] == "NULL" ? null : Code[i];
                    codePath[i, 0] = Path[i] == "NULL" ? null : Path[i];
                    codeName[i, 0] = Name[i] == "NULL" ? null : Name[i]; ;
                    codeSequence[i, 0] = SequenceNumber[i] == "NULL" ? null : SequenceNumber[i];
                    codeCriticality[i, 0] = Criticality[i] == "NULL" ? null : Criticality[i];
                    codeFunctionType[i, 0] = FunctionType[i] == "NULL" ? null : FunctionType[i];
                    codeComponentClass[i, 0] = ComponentClass[i] == "NULL" ? null : ComponentClass[i];
                    codeMaker[i, 0] = Maker[i] == "NULL" ? null : Maker[i];
                    codeModel[i, 0] = Model[i] == "NULL" ? null : Model[i];
                    codeSerial[i, 0] = SerialNumber[i] == "NULL" ? null : SerialNumber[i];
                    codeME[i, 0] = MaximoEquipment[i] == "NULL" ? null : MaximoEquipment[i];
                    codeMED[i, 0] = MaximoEquipmentDescription[i] == "NULL" ? null : MaximoEquipmentDescription[i];
                    int percentage = 44 + ((i * 5) / Code.Length);
                    progressbar.Value = percentage;
                }

                // Write the transposed data to the worksheet
                startRange1.Resize[Code.Length, 1].Value = codeArray;
                startRange2.Resize[Code.Length, 1].Value = codePath;
                startRange3.Resize[Code.Length, 1].Value = codeName;
                startRange4.Resize[Code.Length, 1].Value = codeSequence;
                startRange5.Resize[Code.Length, 1].Value = codeFunctionType;
                startRange6.Resize[Code.Length, 1].Value = codeCriticality;
                startRange10.Resize[Code.Length, 1].Value = codeComponentClass;
                startRange12.Resize[Code.Length, 1].Value = codeMaker;
                startRange13.Resize[Code.Length, 1].Value = codeModel;
                startRange14.Resize[Code.Length, 1].Value = codeSerial;
                startRange15.Resize[Code.Length, 1].Value = codeME;
                startRange16.Resize[Code.Length, 1].Value = codeMED;

                /// concatination will be using forwardslash
                string formula = "=CONCATENATE(RC[1], IF(ISBLANK(RC[3]), \"\", \"/\"), RC[3], IF(ISBLANK(RC[4]), \"\", \"/\"), RC[4])";
                startRange9.Resize[Code.Length, 1].FormulaR1C1 = formula;
                label.Text = "All data written";
                label.Text = "Filling the Color";
                progressbar.Value = 50;
                for (int i = 0; i < modelColor.Length; i++)
                {
                    if (!string.IsNullOrEmpty(makerColor[i]))
                    {
                        Excel.Range cell = OutPutWorksheet.Cells[i + 2, 12];
                        GetColorValue(makerColor[i], cell);
                    }

                    if (!string.IsNullOrEmpty(modelColor[i]))
                    {
                        Excel.Range cell = OutPutWorksheet.Cells[i + 2, 13];
                        GetColorValue(modelColor[i], cell);
                    }

                    if (!string.IsNullOrEmpty(serialColor[i]))
                    {
                        Excel.Range cell = OutPutWorksheet.Cells[i + 2, 14];
                        GetColorValue(serialColor[i], cell);
                    }
                    if (!string.IsNullOrEmpty(maximoColor[i]))
                    {
                        Excel.Range cell = OutPutWorksheet.Cells[i + 2, 15];
                        GetColorValue(maximoColor[i], cell);
                    }
                    if (i < componentColor.Length)
                    {
                        if (!string.IsNullOrEmpty(componentColor[i]))
                        {
                            Excel.Range cell = ValidationWorksheet.Cells[i, 10];
                            GetColorValue(componentColor[i], cell);
                        }
                    }
                    int percentage = 50 + ((i * 30) / modelColor.Length);
                    progressbar.Value = percentage;
                }
                label.Text = "All colors filled";
                progressbar.Value = 80;
                string validationOutputPath = OutPutFilePath.Remove(OutPutFilePath.LastIndexOf("\\")) + $"\\New_Validation_Output_{currentDateTime}.xlsx";
                ValidationWorkbook.SaveAs(validationOutputPath);
                ValidationWorkbook.Close(false);

                label.Text = "New Validation file with color coding has been created";



                for (int i = 2; i <= numRowsOfJ; i++)
                {
                    string value = allCellValuesOfJob[i, 1]?.ToString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        if (!jobCodeSheet.ContainsKey(value))
                        {
                            jobCodeSheet.Add(value, new List<int> { i });
                        }
                        else
                        {
                            jobCodeSheet[value].Add(i);
                        }
                    }
                    else
                    {
                        string value2 = allCellValuesOfJob[i, 10]?.ToString();
                        if (!string.IsNullOrWhiteSpace(value2))
                        {
                            if (!jobCmpclsSheet.ContainsKey(value2))
                            {
                                jobCmpclsSheet.Add(value2, new List<int> { i });

                            }
                            else
                            {
                                jobCmpclsSheet[value2].Add(i);
                            }
                        }
                    }
                }
                label.Text = "job code sheet read successfully please wait...";
                Range outPutWorksheetUsedRange = OutPutWorksheet.UsedRange;
                object[,] allcellvaluesOutput = (object[,])outPutWorksheetUsedRange.Value;
                for (int i = 2; i <= outPutWorksheetUsedRange.Rows.Count; i++)
                {
                    string value = allcellvaluesOutput[i, 1]?.ToString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        if (!outputCodeSheet.ContainsKey(value))
                        {
                            outputCodeSheet.Add(value, new List<int> { i });
                        }
                        else
                        {
                            outputCodeSheet[value].Add(i);
                        }
                    }
                    string value2 = allcellvaluesOutput[i, 10]?.ToString();
                    if (!string.IsNullOrWhiteSpace(value2))
                    {
                        if (!outputCmpclsSheet.ContainsKey(value2))
                        {
                            outputCmpclsSheet.Add(value2, new List<int> { i });
                        }
                        else
                        {
                            outputCmpclsSheet[value2].Add(i);
                        }
                    }
                }
                label.Text = "output sheet code stored please wait...";
                foreach (var item in jobCodeSheet)
                {
                    if (outputCodeSheet.ContainsKey(item.Key))
                    {
                        // int startRow = outputCodeSheet[item.Key][0];  //if we want to insert jobs only on first occurance 
                        int countIndex = jobCodeSheet[item.Key].Count;
                        foreach (int startRow in outputCodeSheet[item.Key])
                        {
                            CopyData(startRow + increment, countIndex);
                            int newStartRow = startRow + increment - (countIndex - 1);
                            int firstRow = jobCodeSheet[item.Key][0];
                            bool inSeries = true;
                            for (int i = 1; i < countIndex; i++)
                            {
                                int row = jobCodeSheet[item.Key][i];
                                if (firstRow + i != row)
                                {
                                    inSeries = false;
                                }
                            }
                            label.Text = $"{inSeries}  {startRow + increment}";
                            if (inSeries)
                            {
                                Excel.Range sourceRange = jobWorksheet.Range[$"O{firstRow}:AF{jobCodeSheet[item.Key][countIndex - 1]}"];
                                Excel.Range destRange = OutPutWorksheet.Range[$"T{newStartRow}"];
                                sourceRange.Copy(destRange);
                                for (int i = 0; i < countIndex; i++)
                                {
                                    int row = jobCodeSheet[item.Key][i];
                                    string CounterType = allCellValuesOfJob[row, 19]?.ToString();
                                    if (!String.IsNullOrEmpty(CounterType))
                                    {
                                        int.TryParse(allCellValuesOfJob[row, 18]?.ToString(), out int result);
                                        switch (CounterType)
                                        {
                                            case "Days":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 30)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Weeks":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(7 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(7 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 4)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Months":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(30 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(30 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 1)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Years":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(365 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(365 * result * 0.1, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";
                                                break;
                                            case "Hours":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 720)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                for (int i = 0; i < countIndex; i++)
                                {
                                    int row = jobCodeSheet[item.Key][i];
                                    OutPutWorksheet.Cells[newStartRow + i, 20].Value = allCellValuesOfJob[row, 15];  //Job code
                                    OutPutWorksheet.Cells[newStartRow + i, 21].Value = allCellValuesOfJob[row, 16];  // Job Name
                                    OutPutWorksheet.Cells[newStartRow + i, 22].Value = allCellValuesOfJob[row, 17];  // Job Description
                                    OutPutWorksheet.Cells[newStartRow + i, 23].Value = allCellValuesOfJob[row, 18];  // Interval
                                    OutPutWorksheet.Cells[newStartRow + i, 24].Value = allCellValuesOfJob[row, 19];  // counter type
                                    OutPutWorksheet.Cells[newStartRow + i, 25].Value = allCellValuesOfJob[row, 20];  // job category
                                    OutPutWorksheet.Cells[newStartRow + i, 26].Value = allCellValuesOfJob[row, 21];  // job type
                                                                                                                     //    OutPutWorksheet.Cells[newStartRow + i, 27].Value = allCellValuesOfJob[row, 22];  // reminder
                                                                                                                     //  OutPutWorksheet.Cells[newStartRow + i, 28].Value = allCellValuesOfJob[row, 23];  // window
                                    OutPutWorksheet.Cells[newStartRow + i, 29].Value = allCellValuesOfJob[row, 24];   // unit
                                    OutPutWorksheet.Cells[newStartRow + i, 30].Value = allCellValuesOfJob[row, 25]; //res depart
                                    OutPutWorksheet.Cells[newStartRow + i, 31].Value = allCellValuesOfJob[row, 26]; // round
                                    OutPutWorksheet.Cells[newStartRow + i, 32].Value = allCellValuesOfJob[row, 27]; // round title
                                                                                                                    //  OutPutWorksheet.Cells[newStartRow + i, 33].Value = allCellValuesOfJob[row, 28]; // scheduling type
                                    OutPutWorksheet.Cells[newStartRow + i, 37].Value = allCellValuesOfJob[row, 32]; //  job origin
                                    string CounterType = allCellValuesOfJob[row, 19]?.ToString();
                                    if (!String.IsNullOrEmpty(CounterType))
                                    {
                                        int.TryParse(allCellValuesOfJob[row, 18]?.ToString(), out int result);
                                        switch (CounterType)
                                        {
                                            case "Days":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 30)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Weeks":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(7 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(7 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 4)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Months":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(30 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(30 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 1)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Years":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(365 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(365 * result * 0.1, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";
                                                break;
                                            case "Hours":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 720)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                label.Text = "Function type job code has been written please wait...";
                foreach (var item in jobCmpclsSheet)
                {
                    if (outputCmpclsSheet.ContainsKey(item.Key))
                    {
                        int countIndex = jobCmpclsSheet[item.Key].Count;
                        foreach (int startRow in outputCmpclsSheet[item.Key])
                        {
                            CopyData(startRow + increment, countIndex);
                            int newStartRow = startRow + increment - (countIndex - 1);
                            int firstRow = jobCmpclsSheet[item.Key][0];
                            bool inSeries = true;
                            for (int i = 1; i < countIndex; i++)
                            {
                                int row = jobCmpclsSheet[item.Key][i];
                                if (firstRow + i != row)
                                {
                                    inSeries = false;
                                }
                            }
                            label.Text = $"{inSeries}  {startRow + increment}";
                            if (inSeries)
                            {
                                Excel.Range sourceRange = jobWorksheet.Range[$"O{firstRow}:AF{jobCmpclsSheet[item.Key][countIndex - 1]}"];
                                Excel.Range destRange = OutPutWorksheet.Range[$"T{newStartRow}"];
                                sourceRange.Copy(destRange);
                                for (int i = 0; i < countIndex; i++)
                                {
                                    int row = jobCmpclsSheet[item.Key][i];
                                    string CounterType = allCellValuesOfJob[row, 19]?.ToString();
                                    if (!String.IsNullOrEmpty(CounterType))
                                    {
                                        int.TryParse(allCellValuesOfJob[row, 18]?.ToString(), out int result);
                                        switch (CounterType)
                                        {
                                            case "Days":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 30)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Weeks":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(7 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(7 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 4)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Months":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(30 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(30 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 1)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Years":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(365 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(365 * result * 0.1, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";
                                                break;
                                            case "Hours":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 720)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                for (int i = 0; i < countIndex; i++)
                                {
                                    int row = jobCmpclsSheet[item.Key][i];
                                    OutPutWorksheet.Cells[newStartRow + i, 20].Value = allCellValuesOfJob[row, 15];  //Job code
                                    OutPutWorksheet.Cells[newStartRow + i, 21].Value = allCellValuesOfJob[row, 16];  // Job Name
                                    OutPutWorksheet.Cells[newStartRow + i, 22].Value = allCellValuesOfJob[row, 17];  // Job Description
                                    OutPutWorksheet.Cells[newStartRow + i, 23].Value = allCellValuesOfJob[row, 18];  // Interval
                                    OutPutWorksheet.Cells[newStartRow + i, 24].Value = allCellValuesOfJob[row, 19];  // counter type
                                    OutPutWorksheet.Cells[newStartRow + i, 25].Value = allCellValuesOfJob[row, 20];  // job category
                                    OutPutWorksheet.Cells[newStartRow + i, 26].Value = allCellValuesOfJob[row, 21];  // job type
                                    OutPutWorksheet.Cells[newStartRow + i, 27].Value = allCellValuesOfJob[row, 22];  // reminder
                                    OutPutWorksheet.Cells[newStartRow + i, 28].Value = allCellValuesOfJob[row, 23];  // window
                                    OutPutWorksheet.Cells[newStartRow + i, 29].Value = allCellValuesOfJob[row, 24];   // unit
                                    OutPutWorksheet.Cells[newStartRow + i, 30].Value = allCellValuesOfJob[row, 25]; //res depart
                                    OutPutWorksheet.Cells[newStartRow + i, 31].Value = allCellValuesOfJob[row, 26]; // round
                                    OutPutWorksheet.Cells[newStartRow + i, 32].Value = allCellValuesOfJob[row, 27]; // round title
                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = allCellValuesOfJob[row, 28]; // scheduling type
                                    OutPutWorksheet.Cells[newStartRow + i, 37].Value = allCellValuesOfJob[row, 32]; //  job origin
                                    string CounterType = allCellValuesOfJob[row, 19]?.ToString();
                                    if (!String.IsNullOrEmpty(CounterType))
                                    {
                                        int.TryParse(allCellValuesOfJob[row, 18]?.ToString(), out int result);
                                        switch (CounterType)
                                        {
                                            case "Days":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 30)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Weeks":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(7 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(7 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 4)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Months":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(30 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(30 * result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 1)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                            case "Years":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(365 * result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(365 * result * 0.1, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";
                                                break;
                                            case "Hours":
                                                OutPutWorksheet.Cells[newStartRow + i, 27].Value = Math.Round(result * 0.07, MidpointRounding.AwayFromZero);
                                                OutPutWorksheet.Cells[newStartRow + i, 28].Value = Math.Round(result * 0.1, MidpointRounding.AwayFromZero);
                                                if (result <= 720)
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Fixed";
                                                }
                                                else
                                                {
                                                    OutPutWorksheet.Cells[newStartRow + i, 33].Value = "Scheduled";

                                                }
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                label.Text = "componet class job has been written please wait...";
                OutputWorkbook.Save();
                //Maximo sheet 
                Excel.Range newUsedRange = OutPutWorksheet.UsedRange;
                object[,] tempeee = newUsedRange.Value;
                int tempeerow = tempeee.GetLength(0);
                Marshal.ReleaseComObject(jobWorkbook);
                Marshal.ReleaseComObject(jobWorksheet);
                for (int i = 2; i <= tempeerow; i++)
                {
                    if (tempeee[i, 5]?.ToString() == "Group Level 2")
                    {
                        splitList.Add(i);
                    }
                }


                label.Text = $"Making the {splitList.Count} worksheets based upon Group Level 2";
                for (int i = 0; i < splitList.Count; i++)
                {
                    int startRow = splitList[i];
                    int endRow;
                    if (i < splitList.Count - 1)
                    {
                        endRow = splitList[i + 1];
                    }
                    else
                    {
                        endRow = tempeerow;
                    }
                    string destPath = OutPutFilePath.Remove(OutPutFilePath.LastIndexOf("\\"));
                    destPath += $"\\{i + 1}. {tempeee[startRow, 1]?.ToString()}_Output_" + currentDateTime + ".xlsx";
                    try
                    {
                        SaveEmbeddedResourceToFile(resourceName, destPath);
                        destWorkbook = excelApp.Workbooks.Open(destPath);
                        destWorksheet = (Excel.Worksheet)destWorkbook.Worksheets[1];
                    }
                    catch (Exception ex)
                    {
                        destWorkbook = excelApp.Workbooks.Add();
                        destWorksheet = (Excel.Worksheet)destWorkbook.Worksheets[1];
                        destWorkbook.SaveAs(destPath);
                    }


                    Excel.Range sourceRangeo = OutPutWorksheet.Range[$"A{startRow}:AN{endRow - 1}"];
                    Excel.Range destRange = destWorksheet.Range["A2"];
                    // Copy the data from sourceRange to destRange
                    sourceRangeo.Copy(destRange);
                    destWorkbook.Save();
                    destWorkbook.Close(true);
                    label.Text = $"File {i + 1} has been created";
                    int percentage = 80 + (i * 20 / splitList.Count);
                    progressbar.Value = percentage;
                }
                OutputWorkbook.Save();
                OutputWorkbook.Close(true);
                textfile();
                label.Text = $"All completed and Total {splitList.Count + 3} files have been created ";
                progressbar.Value = 100;
                label.Text = "Completed ";
                MessageBox.Show("Please check the log file ", "Successfully created", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (!outputWorkbookOpened)
                {
                    MessageBox.Show($"There is some issue while opening the Output Workbook", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (!outputWorksheetOpened)
                {
                    MessageBox.Show($"There is some issue while opening the Output Worksheet", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (!inputWorkbookOpenend)
                {
                    MessageBox.Show($"There is some issue while opening the Input Workbook", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (!inputWorksheetOpened)
                {
                    MessageBox.Show($"There is some issue while opening the Input Worksheet", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (!verificationWorkBookOpened)
                {
                    MessageBox.Show($"There is some issue while opening the verification Workbook", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (!verificationWorksheetOpened)
                {
                    MessageBox.Show($"There is some issue while opening the Verification Worksheet", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    Thread.Sleep(500);
                    try
                    {
                        Process process = Process.GetProcessById((int)PID);
                    }
                    catch
                    {
                        tryHelper = true;
                        MessageBox.Show($"Excel App has been automatically closed", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    if (!tryHelper)
                    {
                        MessageBox.Show($"{ex.Message}", "An error occurred", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
            }
            finally
            {
                ResourceRelease(PID, tryHelper);
            }
            void MaximoJobSheet(string outputFilePath,string maximoJobFilePath)
            {
                Workbook maximoJobworkbook = excelApp.Workbooks.Open(Form1.maximoJobpath);
                Worksheet maximoJobWorksheet = maximoJobworkbook.Worksheets[1];
                Range maximoUsedRange = maximoJobWorksheet.UsedRange;
                maximoUsedRange.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;

                
                object[,] allCellValuesOfMaxmio = (object[,])maximoUsedRange.Value;

                int numRowsOfM = allCellValuesOfMaxmio.GetLength(0);
                int numColsOfM = allCellValuesOfMaxmio.GetLength(1);
                // Opening Files
                OpenFiles();
                

                Dictionary<string,List<int>> assetNumber = new Dictionary<string, List<int>>(numRowsOfV + 2);
                for (int i = 1; i <= numRowsOfM; i++)
                {
                    object cellValue1 = allCellValuesOfValidationSheet[i, 10];
                    string cellValueDescription = cellValue1?.ToString();
                   // componentsNumber.Add(cellValueDescription);
                }


            }
            void CopyData(int startRow, int countRow)
            {
                if (countRow > 1)
                {
                    Excel.Range copiedRow = OutPutWorksheet.Range[$"A{startRow}"].EntireRow;
                  

                    // Insert three new rows below the first row
                    Excel.Range insertRange = copiedRow.Resize[countRow - 1]; // Resize the range to include three new rows
                    insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    // Get the range of the newly inserted rows
                    Excel.Range newRows = OutPutWorksheet.Range[$"A{startRow + 1}:A{startRow + countRow - 1}"]; // Adjust the range as needed

                    // Paste the copied data into the new rows
                    copiedRow.Copy(newRows);
                   // newRows.PasteSpecial(Excel.XlPasteType.xlPasteAll);

                    // Clear the clipboard
                    excelApp.CutCopyMode = 0;
                    increment += countRow - 1;
                }


            }
            void OpenFiles()
            {
                try
                {
                    if (helper)
                    {
                        OutputWorkbook = excelApp.Workbooks.Add();
                        label.Text = "Output workbook add";
                        outputWorkbookOpened = true;
                        OutPutWorksheet = (Excel.Worksheet)OutputWorkbook.Worksheets[1];
                        label.Text = "output worksheet add";
                        outputWorksheetOpened = true;
                        OutputWorkbook.SaveAs(OutPutFilePath);
                        label.Text = "Output File Created";
                        progressbar.Value = 2;

                    }
                    else
                    {

                        OutputWorkbook = excelApp.Workbooks.Open(OutPutFilePath);
                        label.Text = "Output workbook opened";

                        outputWorkbookOpened = true;
                        OutPutWorksheet = (Excel.Worksheet)OutputWorkbook.Worksheets[1];

                        outputWorksheetOpened = true;
                        label.Text = "Output File Opened";
                        progressbar.Value = 2;

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Output file not opened" + ex.Message);
                }
                try
                {

                    InputWorkbook = excelApp.Workbooks.Open(InputFilePath);
                    inputWorkbookOpenend = true;
                    label.Text = "Input workbook Opened";

                    InputWorksheet = (Excel.Worksheet)InputWorkbook.Worksheets[1];
                    inputWorksheetOpened = true;
                    label.Text = "Input File Opened";
                    progressbar.Value = 4;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Input file not opened\n" + ex.Message);
                }
                try
                {

                    ValidationWorkbook = excelApp.Workbooks.Open(ValidationFilePath);
                    verificationWorkBookOpened = true;
                    label.Text = "Validation Workbook Opened";

                    ValidationWorksheet = (Excel.Worksheet)ValidationWorkbook.Worksheets[1];
                    verificationWorksheetOpened = true;
                    label.Text = "Validation File Opened";
                    progressbar.Value = 6;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Verification file not opened\n" + ex.Message);
                }

            }
            void ExtractRange()
            {
                // Get the used range of the Input worksheet
                InputusedRange = InputWorksheet.UsedRange;

                // Get the values of all cells in the used range
                allCellValues = (object[,])InputusedRange.Value;

                // Get the number of rows and columns in the used range
                numRowsOfI = allCellValues.GetLength(0);
                numColsOfI = allCellValues.GetLength(1);

                // Get the used range of the Validation worksheet
                ValidationUsedRange = ValidationWorksheet.UsedRange;

                // Get the values of all cells in the ValidationUsedRange
                allCellValuesOfValidationSheet = (object[,])ValidationUsedRange.Value;

                // Get the number of rows and columns in the ValidationUsedRange
                numRowsOfV = allCellValuesOfValidationSheet.GetLength(0);
                numColsOfV = allCellValuesOfValidationSheet.GetLength(1);

            }
            void Reset()
            {
                groupLevel2Come = true;
                sysCome = true;
                assemblyCome = true;
                elementsCome = true;
            }
            void textfile()
            {
                
                string filePath = OutPutFilePath.Remove(OutPutFilePath.LastIndexOf("\\")) + "\\logs_" + currentDateTime + ".txt";

                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    if (wrongRowsNumber > 0)
                    {
                        writer.WriteLine("\nThere is some heirerchy problem in {0} rows", wrongRowsNumber);
                        for (int wrongRowsNum = 0; wrongRowsNum < wrongRowsNumber; wrongRowsNum++)
                        {
                            writer.WriteLine($"[{wrongRows[wrongRowsNum]}]");
                        }
                    }
                    else
                    {
                        writer.WriteLine("There is no heirerchy error");
                    }
                    if (maximoerror.Length > 0)
                    {
                        writer.WriteLine("There is some issues in Maximo Equipment number coloumn O because coma (,) is there");
                        writer.WriteLine("Total error rows are " + maximoerrorrownumber);
                        writer.WriteLine("Rows are the following in input data");
                        for (int i = 0; i < maximoerrorrownumber; i++)
                        {
                            writer.WriteLine(maximoerror[i]);
                        }
                    }
                    
                    writer.Close();
                }
            }

        }
        static void PrintValuesByKey(Dictionary<string, List<int>> dictionary, string key)
        {
            if (dictionary.ContainsKey(key))
            {
                Console.WriteLine("Values for key '" + key + "': " + string.Join(", ", dictionary[key]));
            }
            else
            {
                Console.WriteLine("Key '" + key + "' not found.");
            }
        }
        public static void ResourceRelease(uint PID, bool tryHelper)
        {

            Thread.Sleep(500);
            try
            {
                Process process = Process.GetProcessById((int)PID);
                if (process != null)
                {
                    process.Kill();
                    excelApp = null;
                    // MessageBox.Show($"Process with PID {PID} terminated successfully.");
                }
            }
            catch
            {
                if (!tryHelper)
                {
                    MessageBox.Show($"Excel Automatically Terminated");
                }
            }
            ObjectCheck(excelApp);
            ObjectCheck(OutputWorkbook);
            ObjectCheck(OutPutWorksheet);
            ObjectCheck(InputWorkbook);
            ObjectCheck(InputWorksheet);
            ObjectCheck(ValidationWorkbook);
            ObjectCheck(ValidationWorksheet);
            ObjectCheck(destWorkbook);
            ObjectCheck(destWorksheet);
        }
        private static void ObjectCheck(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch
            {
                MessageBox.Show("An error occurred while releasing COM object: ");
            }
        }
        static bool TryGetExcelProcessId(out uint processId)
        {
            processId = 0;

            // Use the Excel class name (XLMAIN) to find the main window
            IntPtr excelWindowHandle = FindWindow(ExcelClassName, null);

            if (excelWindowHandle != IntPtr.Zero)
            {
                // Get the process ID associated with the Excel window
                GetWindowThreadProcessId(excelWindowHandle, out processId);
                return true;
            }

            return false;
        }
        static Process GetExcelProcessById(uint processId)
        {
            try
            {
                // Attempt to get the Excel process by ID
                return Process.GetProcessById((int)processId);
            }
            catch (ArgumentException)
            {
                // Process with the specified ID not found
                return null;
            }
            catch (Exception ex)
            {
                // Handle other exceptions as needed
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }
        static void GetColorValue(string colorName, Excel.Range cell)
        {
            // Add more color cases as needed
            switch (colorName.ToLower())
            {
                case "red":
                    cell.Font.Color = XlRgbColor.rgbWhite;
                    cell.Interior.Color = XlRgbColor.rgbRed;
                    break;
                case "green":
                    cell.Font.Color = XlRgbColor.rgbWhite;
                    cell.Interior.Color = XlRgbColor.rgbGreen;
                    break;
                case "orange":
                    cell.Font.Color = XlRgbColor.rgbWhite;
                    cell.Interior.Color = XlRgbColor.rgbOrange;
                    break;
                case "yellow":
                    cell.Font.Color = XlRgbColor.rgbBlack;
                    cell.Interior.Color = XlRgbColor.rgbYellow;
                    break;
                case "blue":
                    cell.Font.Color = XlRgbColor.rgbWhite;
                    cell.Interior.Color = XlRgbColor.rgbBlue;
                    break;
                default:
                    break; // Default to no color
            }
        }
        private static Stream GetEmbeddedResourceStream(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            return assembly.GetManifestResourceStream(resourceName);
        }

        public static void SaveEmbeddedResourceToFile(string resourceName, string filePath)
        {
            using (Stream resourceStream = GetEmbeddedResourceStream(resourceName))
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                {
                    resourceStream.CopyTo(fileStream);
                }
            }
        }
        public static void SaveOtherFile(string resourceName, string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
            {
                File.Copy(resourceName, filePath);
            }
        }
    }
}
