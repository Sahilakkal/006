
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using App = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Font = System.Drawing.Font;
using SysApp = System.Windows.Forms.Application;

namespace ExcelHierarchyConversion_InterOp
{
    public partial class ExcelHierarchyCon : Form
    {
        private Task conversionTask;
        private App excelApp;
        private Workbook verificationWorkbook;
        private Workbook inputWorkbook;
        private Workbook outputWorkbook;
        private Worksheet inputWorksheet;
        private Worksheet verificationWorksheet;
        private Worksheet outputWorksheet;
        bool isExcelRunning = true;

        public ExcelHierarchyCon()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
        }

        private static bool IsWindowVisible(IntPtr hWnd)
        {
            return IsWindowVisible(hWnd.ToInt32());
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(int hWnd);


        private void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            folderBrowse.ShowDialog();
            outputPathTextBox.Text = folderBrowse.SelectedPath;
        }

        private void uploadButton_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            fileDialog.ShowDialog();
            inputPathTextBox.Text = fileDialog.FileName;
        }

        /// <summary>
        ///  This function Reads the data from input worksheet and stores them into a list of list of string
        /// </summary>
        /// <param name="inputWorksheet"> This is the inputsheet from where data is to be readed</param>
        /// <param name="inputWorkbook"> This is the input workbook where data is to be readed</param>
        /// <returns> returns the list of list of string which represents the rows </returns>
        private List<List<string>> ReadData(Worksheet inputWorksheet, Workbook inputWorkbook)
        {
            Excel.Range usedRange = inputWorksheet.UsedRange;
            object[,] data = usedRange.Value;

            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            //---------------------Reading the data and storing it in a list of list------------------------//
            List<List<string>> transformedData = new List<List<string>>();

            int chunkSize = 1000;

            for (int rowIdx = 2; rowIdx <= rowCount; rowIdx += chunkSize)
            {
                int rowsToRead = Math.Min(chunkSize, rowCount - rowIdx + 1);

                for (int i = rowIdx; i < rowIdx + rowsToRead; i++)
                {
                    List<string> rowData = new List<string>();

                    for (int j = 7; j < 23; j++)
                    {
                        string cellData = Convert.ToString(data[i, j]);

                        if (cellData != "" && !string.IsNullOrWhiteSpace(cellData))
                        {
                            rowData.Add(cellData);

                        }

                        else
                        {
                            rowData.Add("");
                        }

                    }

                    transformedData.Add(rowData);
                }
            }
            return transformedData;



        }

        /// <summary>
        ///  This function writes the updated data to an Excel worksheet, utilizing chunked writing for optimized performance. It also updates cell colors in specific columns based on conditions. The final workbook is saved to the specified outputFilePath.
        /// </summary>
        /// <param name="updatedData"> Reresents list of list of strings which contains data in from of row</param>
        /// <param name="outputFilePath"> Path where this file is to be saved</param>
        /// <param name="excelApp"> instance of App</param>
        /// <param name="templateFilePath"> Represents the path where template file is stored </param>
        private void WriteData(List<List<string>> updatedData, string outputFilePath, App excelApp, string templateFilePath)
        {
            outputWorkbook = excelApp.Workbooks.Open(templateFilePath);
            outputWorksheet = outputWorkbook.Sheets[1];

            try
            {

                int rows = updatedData.Count;
                int cols = 19;
                int startRow = 2;
                int startCol = 1;
                int chunkSize = 1000;

                for (int rowIdx = 0; rowIdx < rows; rowIdx += chunkSize)
                {
                    int rowsToWrite = Math.Min(chunkSize, rows - rowIdx);
                    object[,] chunk = new object[rowsToWrite, cols];

                    for (int i = 0; i < rowsToWrite; i++)
                    {
                        List<string> rowData = updatedData[rowIdx + i];
                        for (int j = 0; j < rowData.Count; j++)
                        {
                            //---------rowData[16] represents the color of MAKER cell ---------------//
                            //---------rowData[17] represents the color of MODEL cell ---------------//
                            //---------rowData[18] represents the color of SERIAL cell ---------------//
                            //---------rowData[19] represents the color of MAXIMO EQ cell ---------------//

                            if ((j == 16 || j == 17 || j == 18 || j == 19))
                            {
                                if (j == 16 && rowData[16] != null)
                                {
                                    Range makerCell = outputWorksheet.Cells[startRow + rowIdx + i, startCol + 11];
                                    if (rowData[j] == "Red")
                                    {
                                        makerCell.Interior.Color = System.Drawing.Color.Red;
                                    }
                                    if (rowData[j] == "Green")
                                    {
                                        makerCell.Interior.Color = System.Drawing.Color.Green;
                                    }
                                    if (rowData[j] == "Orange")
                                    {
                                        makerCell.Interior.Color = System.Drawing.Color.Orange;
                                    }

                                    if (rowData[j] == "Blue")
                                    {
                                        makerCell.Interior.Color = System.Drawing.Color.Blue;

                                    }

                                }  // Setting the intended color for MAKER Column
                                if (j == 17 && rowData[17] != null)
                                {
                                    Range modelCell = outputWorksheet.Cells[startRow + rowIdx + i, startCol + 12];
                                    if (rowData[j] == "Red")
                                    {
                                        modelCell.Interior.Color = System.Drawing.Color.Red;
                                    }
                                    if (rowData[j] == "Green")
                                    {
                                        modelCell.Interior.Color = System.Drawing.Color.Green;
                                    }
                                    if (rowData[j] == "Orange")
                                    {
                                        modelCell.Interior.Color = System.Drawing.Color.Orange;
                                    }


                                    if (rowData[j] == "Blue")
                                    {
                                        modelCell.Interior.Color = System.Drawing.Color.Blue;

                                    }
                                } // Setting the intended color for MODEL Column
                                if (j == 18 && rowData[18] != null)
                                {

                                    Range SerialCell = outputWorksheet.Cells[startRow + rowIdx + i, startCol + 13];
                                    if (rowData[j] == "Red")
                                    {
                                        SerialCell.Interior.Color = System.Drawing.Color.Red;
                                    }
                                    if (rowData[j] == "Green")
                                    {
                                        SerialCell.Interior.Color = System.Drawing.Color.Green;
                                    }
                                    if (rowData[j] == "Orange")
                                    {
                                        SerialCell.Interior.Color = System.Drawing.Color.Orange;
                                    }

                                    if (rowData[j] == "Blue")
                                    {
                                        SerialCell.Interior.Color = System.Drawing.Color.Blue;

                                    }
                                } // Setting the intended color for SERIAL Column

                                if (j == 19 && rowData[18] != null)
                                {

                                    Range maximoCell = outputWorksheet.Cells[startRow + rowIdx + i, startCol + 14];
                                    if (rowData[j] == "Yellow")
                                    {
                                        maximoCell.Interior.Color = System.Drawing.Color.Yellow;
                                    }

                                } // Setting the intended color for MAXIMO Column
                            } // Adding colors into the desired columns


                            else
                            {

                                chunk[i, j] = rowData[j];

                            } // Adding data into the excel sheet 
                        }
                    }

                    Excel.Range textColumnsRange = outputWorksheet.Range[outputWorksheet.Cells[startRow, startCol], outputWorksheet.Cells[startRow + rows - 1, startCol + cols - 1]];
                    textColumnsRange.Columns[1].NumberFormat = "@"; // Column 0 (A)
                    textColumnsRange.Columns[4].NumberFormat = "@"; // Column 3 (D)
                    textColumnsRange.Columns[14].NumberFormat = "@"; // Column 13 (N)

                    // Write the chunk to Excel
                    Excel.Range writeRange = outputWorksheet.Range[outputWorksheet.Cells[startRow + rowIdx, startCol], outputWorksheet.Cells[startRow + rowIdx + rowsToWrite - 1, startCol + cols - 1]];
                    writeRange.Value = chunk;
                }
            }
            finally
            {
                outputWorkbook.SaveAs(outputFilePath);
            }
        }

        /// <summary>
        /// This function reads data from an Excel worksheet designated for verification. It focuses on specific columns, filters out empty cells, and utilizes a chunked reading approach to efficiently handle large datasets. The processed verification data is returned as a list of lists.
        /// </summary>
        /// <param name="verificationWorksheet">Excel worksheet containing data for verification.</param>
        /// <returns>List of lists (verificationData) containing the processed verification data.</returns>
        private List<List<string>> ReadDataForVerification(Worksheet verificationWorksheet)
        {
            Excel.Range usedRange = verificationWorksheet.UsedRange;
            object[,] data = usedRange.Value;
            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            //   ---------------------Reading the data and storing it in a list of list------------------------//
            List<List<string>> verificationData = new List<List<string>>();
            int chunkSize = 500;
            for (int rowIdx = 2; rowIdx <= rowCount; rowIdx += chunkSize)
            {
                int rowsToRead = Math.Min(chunkSize, rowCount - rowIdx + 1);

                for (int i = rowIdx; i < rowIdx + rowsToRead; i++)
                {
                    List<string> rowData = new List<string>();
                    for (int j = 10; j < 22; j++)
                    {

                        if (j == 10 || j == 11 || j == 16 || j == 18 || j == 20 || j == 21)
                        {
                            string cellData = Convert.ToString(data[i, j]);

                            if ((cellData != "" && !string.IsNullOrWhiteSpace(cellData)) && j != 21)
                            {
                                rowData.Add(cellData);

                            }


                            else
                            {
                                rowData.Add("");
                            }

                        }

                        else
                        {
                            continue;
                        }


                    }


                    verificationData.Add(rowData);

                }
            }
            return verificationData;


        }

        /// <summary>
        /// This function verifies and compares data between an updated data list (updList) and a verification data list (verList). It checks for matches between Maximo equipment numbers and component numbers, updates color codes and values accordingly, and logs errors for identified issues.
        /// </summary>
        /// <param name="updList">Reference to the list of lists representing updated data.</param>
        /// <param name="verList">Reference to the list of lists representing verification data.</param>
        private void VerifyData(ref List<List<string>> updList, ref List<List<string>> verList)
        {
            int maximoIndex = 14;
            int componentIndex = 0;

            int updIdx = -1;
            int verIdx = -1;

            // Get distinct component numbers from verList
            var distinctComponentNumbers = verList.Select(row => row[componentIndex]).Distinct();
            foreach (var componentNo in distinctComponentNumbers)
            {
                for (int i = 0; i < updList.Count; i++)
                {
                    if (updList[i].Count > maximoIndex && updList[i][maximoIndex].Contains(componentNo))
                    {
                        string maximoValue = updList[i][maximoIndex];
                        string trimmedMaximoValue = maximoValue;

                        // Check if maximoValue contains a comma
                        if (maximoValue.Contains(","))
                        {
                            // Trim the string after the comma for comparison
                            trimmedMaximoValue = maximoValue.Split(',')[0].Trim();
                            updList[i][19] = "Yellow";
                        }

                        if (trimmedMaximoValue == componentNo)
                        {
                            for (int j = 0; j < verList.Count; j++)
                            {
                                if (verList[j].Count > componentIndex && verList[j][componentIndex] == componentNo)
                                {
                                    updIdx = i; // updList index
                                    verIdx = j; // verList index



                                    if (!string.IsNullOrEmpty(componentNo) && !string.IsNullOrWhiteSpace(componentNo))
                                    {

                                        verList[verIdx][5] = "Green";
                                        updList[updIdx][5] = verList[verIdx][2];         // adding CRITICALITY from STATUS

                                        string makerNameInVer = verList[verIdx][1];    // MAKER Names
                                        string makerNameInUpd = updList[updIdx][11];
                                        string[] substrings = makerNameInVer.Split(new string[] { " || " }, StringSplitOptions.None);
                                        string makerFullNameInVer;
                                        string makerShortNameInVer;
                                        if (substrings.Length == 2)
                                        {
                                            makerFullNameInVer = substrings[0];
                                            makerShortNameInVer = substrings[1];
                                        }

                                        else
                                        {
                                            makerFullNameInVer = makerNameInVer;
                                            makerShortNameInVer = makerNameInVer;
                                        }

                                        string serialNoInVer = verList[verIdx][3];     // SERIAL
                                        string serialNoInUpd = updList[updIdx][13];

                                        string modelInVer = verList[verIdx][4];        // MODEL
                                        string modelInUpd = updList[updIdx][12];

                                        string colorInVer = verList[verIdx][5];       //represents color for COMPONENT no in verification sheet
                                        string makercolorInUpd = updList[updIdx][16]; // represents color for MAKER in Output
                                        string modelcolorInUpd = updList[updIdx][17]; // represents color for MODEL in Output
                                        string serialcolorInUpd = updList[updIdx][18];// // represents color for SERIAL in Output


                                        if ((makerShortNameInVer == makerNameInUpd || makerFullNameInVer == makerNameInUpd) || (serialNoInVer == serialNoInUpd) || modelInVer == modelInUpd)
                                        {

                                            if (((makerShortNameInVer == makerNameInUpd) || (makerFullNameInVer == makerNameInUpd)) && (!string.IsNullOrEmpty(makerShortNameInVer) || !string.IsNullOrEmpty(makerFullNameInVer)))
                                            {
                                                updList[updIdx][16] = "Green";
                                            }

                                            if (modelInVer == modelInUpd && !string.IsNullOrEmpty(modelInVer))
                                            {
                                                updList[updIdx][17] = "Green";

                                            }

                                            if (serialNoInUpd == serialNoInVer && !string.IsNullOrEmpty(serialNoInVer))
                                            {
                                                updList[updIdx][18] = "Green";
                                            }




                                        }  // when MAKER MODEL SERIAL  any one is same in lists

                                        if (((makerFullNameInVer != makerNameInUpd) || makerShortNameInVer != makerNameInUpd || (serialNoInVer != serialNoInUpd) || modelInVer != modelInUpd))
                                        {

                                            if (string.IsNullOrEmpty(makerNameInUpd) && (!string.IsNullOrEmpty(makerFullNameInVer) || !string.IsNullOrEmpty(makerShortNameInVer)))
                                            {

                                                updList[updIdx][11] = makerShortNameInVer;     //copying values of MAKER from ver to upd
                                                updList[updIdx][16] = "Orange";

                                            } // when MAKER is not present in output but present in verification sheet

                                            if (string.IsNullOrEmpty(modelInUpd) && !string.IsNullOrEmpty(modelInVer))
                                            {
                                                updList[updIdx][12] = modelInVer;
                                                updList[updIdx][17] = "Orange";

                                            }// when MODEL is not present in output but present in verification sheet


                                            if (string.IsNullOrEmpty(serialNoInUpd) && !string.IsNullOrEmpty(serialNoInVer))
                                            {
                                                updList[updIdx][13] = serialNoInVer;
                                                updList[updIdx][18] = "Orange";

                                            } // when SERIAL is not present in output but present in verification sheet


                                            if (!string.IsNullOrEmpty(makerNameInUpd) && (!string.IsNullOrEmpty(makerNameInVer) || !string.IsNullOrEmpty(makerFullNameInVer)) && ((makerNameInUpd != makerFullNameInVer) || (makerNameInUpd != makerShortNameInVer)))
                                            {

                                                if (makerNameInVer.Contains(makerNameInUpd) && makerNameInVer.Contains("||"))
                                                {
                                                    updList[updIdx][16] = "Green";
                                                }
                                                else
                                                {
                                                    updList[updIdx][16] = "Red";

                                                }


                                            }  // when both MAKERS are non empty and not equal

                                            if (!string.IsNullOrEmpty(modelInUpd) && !string.IsNullOrEmpty(modelInVer) && modelInUpd != modelInVer)
                                            {
                                                updList[updIdx][17] = "Red";

                                            }  // when both MODEL are non empty and not equal

                                            if (!string.IsNullOrEmpty(serialNoInUpd) && !string.IsNullOrEmpty(serialNoInVer) && serialNoInUpd != serialNoInVer)
                                            {
                                                updList[updIdx][18] = "Red";

                                            }  // when both SERIAL number are non empty and not equal

                                            if (!string.IsNullOrEmpty(makerNameInUpd) && (string.IsNullOrEmpty(makerFullNameInVer) || string.IsNullOrEmpty(makerShortNameInVer)))
                                            {
                                                updList[updIdx][16] = "Blue";

                                            }   // when MAKER in verification sheet is empty and it is not empty in output

                                            if (!string.IsNullOrEmpty(modelInUpd) && string.IsNullOrEmpty(modelInVer))
                                            {
                                                updList[updIdx][17] = "Blue";

                                            }   // when MODEL in verification sheet is empty and it is not empty in output

                                            if (!string.IsNullOrEmpty(serialNoInUpd) && string.IsNullOrEmpty(serialNoInVer))
                                            {
                                                updList[updIdx][18] = "Blue";

                                            }    // when SERIIAL in verification sheet is empty and it is not empty in output

                                        }  //Handling MAKER MODEL SERIAL NO color scheme

                                        if (updList[updIdx][11].Contains("||"))
                                        {
                                            string[] parts = updList[updIdx][11].Split(new string[] { "||" }, StringSplitOptions.None);

                                            if (parts.Length == 2)
                                            {
                                                updList[updIdx][11] = parts[1];
                                            }


                                        }    //  Handling MAKER full name and MAKER short Name
                                    }  //comparing data between verififcation sheet and output sheet

                                    if (updList[updIdx][16] == "Green" && updList[updIdx][17] == "Green" && updList[updIdx][18] == "Green")
                                    {
                                        // set Green color for this row
                                        updList[updIdx][10] = "Details Provided are Correct";


                                    }



                                    break;
                                }  // handling the Case where maximo Index Matches with component number
                            }
                            break;
                        }
                    }
                }
            }

        }

        /// <summary>
        /// This function processes input data, extracts relevant information, and performs conditional operations to create an updated list. It handles the formatting of specific columns and logs errors, including duplicate Maximo equipment numbers, if the corresponding checkbox is checked.
        /// </summary>
        /// <param name="transformedData">List of lists containing input data to be processed.</param>
        /// <param name="logFilePath">String specifying the path to the log file.</param>
        /// <returns></returns>
        private List<List<string>> ProcessData(List<List<string>> transformedData, string logFilePath)
        {
            //---------Performing Operations on data---------//
            List<List<string>> updatedData = new List<List<string>>();
            String pathAssembly = "", pathElement = "";
            int countRows = 2;
            foreach (List<string> item in transformedData)
            {
                List<string> updatedItem = new List<string>(new string[20]);
                string maximoEqDescription = "", maximoEq = "", maker = "", serialNum = "", modelType = "";
                bool isBlank = true;
                int count = 0;

                if ((item[0] != "" && !string.IsNullOrWhiteSpace(item[0])) || (item[2] != "" && !string.IsNullOrWhiteSpace(item[2])) || (item[5] != "" && !string.IsNullOrWhiteSpace(item[5])) || (item[8] != "" && !string.IsNullOrWhiteSpace(item[8])))
                {
                    isBlank = false;

                    if (item[0] != "" && !string.IsNullOrWhiteSpace(item[0]))
                    {
                        count++;
                        updatedItem[0] = item[0];

                    }
                    if (item[2] != "" && !string.IsNullOrWhiteSpace(item[2]))
                    {
                        updatedItem[0] = item[2];
                        count++;

                    }

                    if (item[5] != "" && !string.IsNullOrWhiteSpace(item[5]))
                    {
                        updatedItem[0] = item[5];
                        count++;

                    }

                    if (item[8] != "" && !string.IsNullOrWhiteSpace(item[8]))
                    {
                        count++;
                        updatedItem[0] = item[8];

                    }

                }    // handling the column code

                if (item[1] != "" || item[4] != "" || item[7] != "" || item[10] != "")
                {
                    if (item[1] != "")
                    {
                        updatedItem[2] = item[1];


                        updatedItem[4] = "Group Level 2";
                    }

                    if (item[4] != "")
                    {
                        updatedItem[2] = item[4];


                        updatedItem[4] = "System";

                        pathAssembly = updatedItem[2];
                    }
                    if (item[7] != "")
                    {
                        updatedItem[2] = item[7];


                        updatedItem[4] = "Assembly";
                        updatedItem[1] = pathAssembly;
                        pathElement = pathAssembly + "/" + updatedItem[2];
                    }

                    if (item[10] != "")
                    {
                        updatedItem[2] = item[10];


                        updatedItem[4] = "Element";
                        updatedItem[1] = pathElement;
                    }


                } // handling the Column Name

                if (item[3] != "" || item[6] != "" || item[9] != "")
                {
                    if (item[6] != "")
                    {
                        updatedItem[3] = item[6];
                    }

                    if (item[9] != "")
                    {
                        updatedItem[3] = item[9];
                    }

                    if (item[3] != "")
                    {
                        updatedItem[3] = item[3];
                    }
                }  // handling the column Sequence No

                if (item[11] != "")
                {
                    maximoEq = item[11];
                }   // handling the maximo Equipment

                if (item[12] != "")
                {
                    maximoEqDescription = item[12];
                }  // handling the maximo Equipment Description

                if (item[13] != "" && item[13] != "NULL")
                {
                    maker = item[13];
                }  // handling the maker

                if (item[14] != "" && item[14] != "NULL")
                {
                    modelType = item[14];
                }  // handling the modelType
                if (item[15] != "" && item[15] != "NULL")
                {
                    serialNum = item[15];
                }  //handling the serial Num  

                if (!string.IsNullOrEmpty(updatedItem[2]))
                {

                    if (updatedItem[2].Contains("Pump") && updatedItem[2].Contains("Unit") && updatedItem[2].Contains("Pump Unit"))
                    {
                        updatedItem[9] = "Pump Unit";
                    }

                    else if (updatedItem[2].Contains("Pump") && !updatedItem[2].Contains("Unit") && !updatedItem[2].Contains("E-Motor"))
                    {
                        updatedItem[9] = "Pump";
                    }

                    else if (updatedItem[2].Contains("E-Motor"))
                    {
                        updatedItem[9] = "E-Motor";
                    }

                    else if (updatedItem[2].Contains("Cooler") || updatedItem[2].Contains("Heater"))
                    {
                        updatedItem[9] = "Heat Exchanger";

                    }


                }  // --------- Handling the Column Component Class------------------//

                if (maximoEq.Contains(","))
                {
                    string errorMessage = string.Format($"Duplicate Maximo equimpent Number Found at {transformedData.IndexOf(item) + 2} row");

                    if (checkBox_LogErrors.Checked)
                    {

                        LogDuplicateMaximoError(errorMessage, logFilePath);
                    }
                }   // handling the condition where maximo Equipment number is more than 1

                updatedItem[11] = maker;      // MAKER
                updatedItem[12] = modelType;  // MODEL
                updatedItem[13] = serialNum;  // SERIAL 
                updatedItem[14] = maximoEq;   // MAXIMO EQUIPMENT
                updatedItem[15] = maximoEqDescription; // MAXIMO EQUIPMENT DESCRIPTION


                updatedItem[16] = "";  // maker color
                updatedItem[17] = "";  // model color
                updatedItem[18] = "";  // serial color
                updatedItem[19] = "";  // maximo Eq color


                if (count > 1)
                {
                    if (checkBox_LogErrors.Checked)
                    {

                        LogError($"Data Mismatched at [{(transformedData.IndexOf(item) + 2)}] Row\n", logFilePath);
                    }

                    continue;
                }


                if (!isBlank)
                {
                    updatedItem[8] = $"=IF(AND(ISBLANK(J{countRows}), ISBLANK(L{countRows}), ISBLANK(M{countRows})), \"\", IF(AND(ISBLANK(J{countRows}), ISBLANK(L{countRows})), M{countRows}, IF(ISBLANK(J{countRows}), IF(ISBLANK(L{countRows}), M{countRows}, CONCATENATE(L{countRows}, \"/\", M{countRows})), CONCATENATE(J{countRows}, IF(ISBLANK(L{countRows}), \"\", CONCATENATE(\"/\", L{countRows})), IF(ISBLANK(M{countRows}), \"\", CONCATENATE(\"/\", M{countRows}))))))";


                    updatedData.Add(updatedItem);
                    countRows++;

                }  // checking whether row is not blank

            }
            return updatedData;

        }

        /// <summary>
        ///This function dynamically sets the visibility of multiple UI elements in the form based on the boolean parameter. It is designed to enable or disable the specified elements to control user interaction with the form. 
        /// </summary>
        /// <param name="visiblity">Boolean flag indicating whether the elements should be visible or not.</param>
        public void SetVisiblityOfElements(bool visiblity)
        {
            CheckBox_splitFiles.Enabled = visiblity;
            checkBox_LogErrors.Enabled = visiblity;
            uploadButton.Enabled = visiblity;
            inputPathTextBox.Enabled = visiblity;
            outputButton.Enabled = visiblity;
            outputPathTextBox.Enabled = visiblity;
            convertButton.Enabled = visiblity;
            uploadVerificationButton.Enabled = visiblity;
            verificationPathTextBox.Enabled = visiblity;
            this.ControlBox = visiblity;
        }

        /// <summary>
        /// This function serves as a wrapper for the LogError function, specifically designed for logging duplicate Maximo errors. It delegates the logging process to the LogError function, allowing consistent error logging across different error types.
        /// </summary>
        /// <param name="errorMessage">String representing the error message to be logged.</param>
        /// <param name="logFilePath">String specifying the path to the log file.</param>
        public void LogDuplicateMaximoError(string errorMessage, string logFilePath)
        {
            LogError(errorMessage, logFilePath);
        }

        /// <summary>
        /// This function logs an error message to a file at the specified logFilePath. It creates a new log file if it's the first error or appends to the existing log file for subsequent errors. Any exceptions during the logging process are caught and reported to the console.
        /// </summary>
        /// <param name="errorMessage"> String representing the error message to be logged.</param>
        /// <param name="logFilePath">String specifying the path to the log file.</param>
        public void LogError(string errorMessage, string logFilePath)
        {

            try
            {
                // If it's the first error, create a new log file
                if (!File.Exists(logFilePath))
                {
                    using (StreamWriter writer = new StreamWriter(logFilePath))
                    {
                        writer.WriteLine($"{DateTime.Now} - {errorMessage}");
                    }
                }
                else
                {
                    // If it's a subsequent error, append to the existing log file
                    using (StreamWriter writer = new StreamWriter(logFilePath, true))
                    {
                        writer.WriteLine($"{DateTime.Now} - {errorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error logging to file: {ex.Message}");
            }
        }

        /// <summary>
        /// Writes data to an Excel worksheet, applying color formatting to target cells, and saves the modified workbook.
        /// </summary>
        /// <param name="verificationData">List of lists containing data to be written to an Excel worksheet</param>
        /// <param name="verificationWorksheet">Target Excel worksheet.</param>
        /// <param name="verificationWorkbook">Workbook containing the target worksheet</param>
        /// <param name="verificationFileSavePath">File path for saving the modified workbook</param>
        private void WriteDataInVerificationList(List<List<string>> verificationData, Worksheet verificationWorksheet, Workbook verificationWorkbook, string verificationFileSavePath)
        {
            try
            {
                int startRow = 2;
                int startCol = 1;

                foreach (List<string> row in verificationData)
                {
                    // Access the 5th element (index 4) in each inner list
                    string colorValue = row.Count > 5 ? row[5] : "";

                    if (colorValue == "Green" && row[0] != "")
                    {
                        // Apply color to the cell in column J (10th column, 0-based index)
                        Excel.Range cell = (Excel.Range)verificationWorksheet.Cells[startRow, startCol + 9];
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                    }

                    // Move to the next row
                    startRow++;
                }
            }



            finally
            {
                verificationWorkbook.SaveAs(verificationFileSavePath);

            }

        }

        /// <summary>
        /// This function splits a list of lists, representing data, into subgroups based on keyword "Group level 2".
        /// </summary>
        /// <param name="outputData">List of lists containing data to be split based on keyword "Group level 2"</param>
        /// <returns>returns a list of lists of lists, where each inner list represents a group of rows that share the same splitKeyword</returns>
        /// 
        public List<List<List<string>>> SplitData(List<List<string>> outputData)
        {
            string splitKeyword = "Group Level 2";
            List<List<List<string>>> result = new List<List<List<string>>>();
            List<List<string>> currentSplit = new List<List<string>>();

            foreach (var row in outputData)
            {
                if (row.Count > 0 && row[4] == splitKeyword)
                {
                    if (currentSplit.Any())
                    {
                        result.Add(new List<List<string>>(currentSplit));
                        currentSplit.Clear();
                    }
                }

                currentSplit.Add(new List<string>(row));
            }

            // Add the last split if there are any remaining rows
            if (currentSplit.Any())
            {
                result.Add(new List<List<string>>(currentSplit));
            }

            return result;
        }



        private  void convertButton_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(inputPathTextBox.Text) || string.IsNullOrEmpty(outputPathTextBox.Text))
            {
                if (string.IsNullOrEmpty(inputPathTextBox.Text))
                {
                    MessageBox.Show("Please Upload the Excel file", "Try uploading file again", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                else
                {
                    MessageBox.Show("Please Upload the Excel file", "Try uploading file again", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }

            }

            else
            {
                label_operationStatus.Font = new Font(label_operationStatus.Font, FontStyle.Bold);

                progressBar1.Value = 0;
                progressBar1.Visible = true;
                this.ControlBox = false;

                SetVisiblityOfElements(false);

                DateTime startTime = DateTime.Now;
                string inputFilePath = inputPathTextBox.Text;
                DateTime dateTime = DateTime.Now;
                string inputDirectoryPath = outputPathTextBox.Text;
                string verificationFilePath = verificationPathTextBox.Text;

                string verificationFileName = Path.GetFileNameWithoutExtension(verificationFilePath);
                string currentDate = dateTime.ToString("yyyyMMdd_HHmmss");
                string inputFileName = Path.GetFileNameWithoutExtension(inputFilePath);

                string logFilePath = Path.Combine(inputDirectoryPath, $"ErrorLogs_{currentDate}.txt");
                string templateFilePath = Path.Combine(SysApp.StartupPath, "Output template.xlsx");
                string templateFilePath_JobSheet = Path.Combine(SysApp.StartupPath, "Jobs Sheet.xlsx");
                string outputFilePath = Path.Combine(inputDirectoryPath, $"{inputFileName}_Output_{currentDate}.xlsx");
                string verificationFileSavePath = Path.Combine(inputDirectoryPath, $"{verificationFileName}_OUTPUT_{currentDate}.xlsx");

                try
                {

                    label_operationStatus.Text = "Opening Excel Application";
                    excelApp = new App();
                    excelApp.ScreenUpdating = false;

                    label_operationStatus.Text = "Opening Input File";
                    inputWorkbook = excelApp.Workbooks.Open(inputFilePath);   //opening the input file in excel -----Reading sheet
                    inputWorksheet = inputWorkbook.Sheets[1];              //opening the input sheet

                    label_operationStatus.Text = "Opening Verification File";
                    verificationWorkbook = excelApp.Workbooks.Open(verificationFilePath);
                    verificationWorksheet = verificationWorkbook.Sheets[1];

                    List<List<string>> storedData;                     //------------for storing Data Present in Input sheet------------------//
                    List<List<string>> verificationData;               //------------for storing Data Present in Verification sheet------------------//
                    List<List<string>> updatedData;                    //------------Transforming the stored Data present In input sheet according to given format------------------//
                    List<List<List<string>>> splittedData;             //------------for storing Data for creating multiple Workbooks of output sheet------------------//



                    storedData = ReadData(inputWorksheet, inputWorkbook);   //Reading the Data
                    verificationData = ReadDataForVerification(verificationWorksheet);   //Reading the verification Data
                    updatedData = ProcessData(storedData, logFilePath);  //Processing the Data
                    VerifyData(ref updatedData, ref verificationData);   // verifying the data


                    //----------------------Handling the Jobsheets Part [006] ------------------\\

                    Workbook jobSheetWorkbook = excelApp.Workbooks.Open(templateFilePath_JobSheet);
                    Worksheet jobSheetWorksheet = jobSheetWorkbook.Sheets[1];

                    List<JobSheetData> jobSheetData;

                    JobSheetData obj_Jobsheet = new JobSheetData();
                    jobSheetData = obj_Jobsheet.ReadDataFromJobSheet(jobSheetWorksheet);

                    OutputSheetData obj_OutputSheetData= new OutputSheetData();
                    
                    obj_OutputSheetData.MapDataToOutputSheet(updatedData,jobSheetData);
                   

                        label_operationStatus.Text = "Writing In verification sheet";
                      //  WriteDataInVerificationList(verificationData, verificationWorksheet, verificationWorkbook, verificationFileSavePath);   // Coloring the component no column in verification sheet
                        progressBar1.Value = 30;
                        label_operationStatus.Text = "Writing In output sheet";
                      //  WriteData(updatedData, outputFilePath, excelApp, templateFilePath);                                                //Writing the Data

                        DateTime endTime = DateTime.Now;
                        TimeSpan duration = endTime - startTime;
                        string formattedTime = $"{(int)duration.TotalMinutes} minutes {duration.Seconds} seconds";
                        MessageBox.Show($"File Converted Successfully. Time taken: {formattedTime}", "Thank You For using Excel hierarchy Converter", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);




                        if (!CheckBox_splitFiles.Checked)
                        {
                            progressBar1.Value = 100;
                        }    // if split files check box is checked then no need to further increase the vakue of progressBar
                        progressBar1.Value = 70;

                        if (CheckBox_splitFiles.Checked)
                        {
                            label_operationStatus.Text = "Splitting Files";
                            MessageBox.Show("Splitting Files.........", "It may take while", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            splittedData = SplitData(updatedData);
                            foreach (var workBookk in splittedData)
                            {

                                string firstValue = workBookk[0][0];

                                string folderPath = Path.Combine(inputDirectoryPath, "SplittedFiles");
                                Directory.CreateDirectory(folderPath);

                                // Modify the file name using the first value
                                string outputFile = $"{firstValue}_Output_{currentDate}.xlsx";

                                // Full path including the folder
                                string fullOutputPath = Path.Combine(folderPath, outputFile);

                                label_operationStatus.Text = $" Creating {firstValue}.xlsx";
                                WriteData(workBookk, fullOutputPath, excelApp, templateFilePath);
                            }

                            MessageBox.Show($"All Files Splitted Successfully", "Thank You For using Excel hierarchy Converter", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                        }
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    isExcelRunning = false;

                    if (comEx.HResult == -2147023174) // 0x800706BA (RPC server unavailable) error code
                    {
                        MessageBox.Show("Excel is not running or not registered as an active object.");
                    }
                    else
                    {
                        MessageBox.Show("An unexpected COMException occurred.");
                        // Set isExcelRunning to false, as there is an issue with Excel
                    }
                }

                finally
                {

                    ReleaseResources();

                }
                progressBar1.Value = 100;
            }
        }
        private void ReleaseResources()
        {
            List<int> excelPID = new List<int>();

            // Get all processes
            Process[] prs = Process.GetProcesses();

            foreach (Process p in prs)
            {
                if (p.ProcessName == "EXCEL.EXE")
                {
                    // Check if the Excel process has a main window and is visible
                    if (IsWindowVisible(p.MainWindowHandle))
                    {
                        Console.WriteLine($"Excel process with PID {p.Id} is visible.");
                    }
                    else
                    {
                        excelPID.Add(p.Id);
                    }
                }
            }

            prs = Process.GetProcesses();

            foreach (Process p in prs)
            {
                if (p.ProcessName == "EXCEL" && !excelPID.Contains(p.Id))
                {
                    // Check if the Excel process has a main window and is visible
                    if (IsWindowVisible(p.MainWindowHandle))
                    {
                        Console.WriteLine($"Excel process with PID {p.Id} is visible.");
                    }
                    else
                    {
                        try
                        {
                            p.Kill();

                        }

                        catch
                        {
                            MessageBox.Show("Excel File not running in Background");
                            MessageBox.Show("Hello");
                            System.Windows.Forms.Application.Restart();
                        }
                        Console.WriteLine($"Excel process with PID {p.Id} killed.");
                    }
                }
            }
        }
        private void exitButton_Click(object sender, EventArgs e)
        {
            DialogResult result;

            if (progressBar1.Visible)
            {
                result = MessageBox.Show("Current Process Will be Terminated, Are you Sure you want to Exit?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            else
            {
                result = MessageBox.Show("Are you Sure you want to Exit?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }

            if (result == DialogResult.Yes)
            {
                ReleaseResources();
                Environment.Exit(0); // Exit the application
            }
            // If the user selects "No", do nothing and let the application continue
        }

        private void uploadVerificationButton_Click(object sender, EventArgs e)
        {
            fileDialog.ShowDialog();
            verificationPathTextBox.Text = fileDialog.FileName;
        }

        private void inputPathTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}