using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using DataTable = System.Data.DataTable;

namespace ExcelHierarchyConversion_InterOp
{
    internal class OutputSheetData
    {
        public string CodeInOutput { get; set; } // A [1]
        public string Path { get; set; } // B [2]
        public string Name { get; set; } // C [3]
        public string SequenceNo { get; set; } // D [4]
        public string FunctionType { get; set; } // E [5]
        public string Criticality { get; set; } //  F [6]
        public string Location { get; set; } //  G [7]
        public string ComponentTypeCode { get; set; } // H [8]
        public string ComponentType { get; set; } //  I [9]

        public string MakerColor { get; set; }
        public string ModelColor { get; set; }
        public string SerialColor { get; set; }
        public string MaximoEqColor { get; set; }



        public string ComponentClass { get; set; } //  J [10]
        public string FunctionStatus { get; set; } //  K [11]
        public string Maker { get; set; } //  L [12]
        public string Model { get; set; } //  M [13]
        public string SerialNo { get; set; } //  N [14]
        public string MaximoEq { get; set; } //  O [15]
        public string MaximoEqDescription { get; set; } //  P [16]
        public string ColorYellow { get; set; }
        public string ColorGreen { get; set; }
        public int RowsToBeAdd { get; set; }
        public JobSheetData dataFromJobSheet;
        public MaximoSheetData dataFromMaximoSheet;

        public OutputSheetData()
        {
            CodeInOutput = string.Empty;
            Path = string.Empty;
            Name = string.Empty;
            SequenceNo = string.Empty;
            FunctionType = string.Empty;
            Criticality = string.Empty;
            Location = string.Empty;
            ComponentTypeCode = string.Empty;
            ComponentType = string.Empty;
            FunctionStatus = string.Empty;
            Maker = string.Empty;
            Model = string.Empty;
            SerialNo = string.Empty;
            MaximoEq = string.Empty;
            RowsToBeAdd = 0;
            MaximoEqDescription = string.Empty;
            dataFromJobSheet = new JobSheetData();
            dataFromMaximoSheet = new MaximoSheetData();
        }
        public List<OutputSheetData> MapDataToOutputSheet(List<List<string>> outputData, List<JobSheetData> jobSheetData, List<MaximoSheetData> maximoSheetData)
        {
            //lalit will work on this function i have made a parameter option you will work on this and map the dataFromMaximoSheet propertty of outputsheet class with objects present in maximoSheetdata which is passed as paramter in function . I have already map the JobsheetData and set the property datafromJobsheet of this class
            List<OutputSheetData> outputSheetData = new List<OutputSheetData>();

            foreach (List<string> row in outputData)
            {
                OutputSheetData singleRow = new OutputSheetData();

                singleRow.CodeInOutput = row[0];
                singleRow.Path = row[1];
                singleRow.Name = row[2];
                singleRow.SequenceNo = row[3];
                singleRow.FunctionType = row[4];
                singleRow.Criticality = row[5];
                singleRow.Location = row[6];
                singleRow.ComponentTypeCode = row[7];
                singleRow.ComponentType = row[8];
                singleRow.ComponentClass = row[9];
                singleRow.FunctionStatus = row[10];
                singleRow.Maker = row[11];
                singleRow.Model = row[12];
                singleRow.SerialNo = row[13];
                singleRow.MaximoEq = row[14];
                singleRow.MaximoEqDescription = row[15];

                singleRow.MakerColor = row[16];
                singleRow.ModelColor = row[17];
                singleRow.SerialColor = row[18];
                singleRow.MaximoEqColor = row[19];

                for (int i = 0; i < jobSheetData.Count; i++)
                {
                    JobSheetData singleRowJobData = jobSheetData[i];
                    if (singleRowJobData.CodeInJob == singleRow.CodeInOutput)
                    {

                        singleRow.dataFromJobSheet = singleRowJobData;
                        break;
                    }

                }

                for (int i = 0; i < maximoSheetData.Count; i++)
                {
                    MaximoSheetData singleRowMaximoData = maximoSheetData[i];
                    if (singleRowMaximoData.AssetNumber == singleRow.MaximoEq)
                    {
                        singleRow.dataFromMaximoSheet = singleRowMaximoData;

                        //  MessageBox.Show("data added" + singleRowMaximoData.AssetNumber);
                        break;
                    }
                }

                outputSheetData.Add(singleRow);
            }

            return outputSheetData;

        }

        public async void WriteDataInOutputAsync(List<OutputSheetData> outputSheetData, Worksheet worksheet,Workbook workbook,string path)
        {
            int numRows = outputSheetData.Count;

            System.Data.DataTable dataTable = new System.Data.DataTable();
            DataRow dataRow;
            string data = "Code\tPath\tName\tSequence No\tFunction Type\tCriticality\tLocation\tComponent Type Code\tComponent Type\tComponent Class\tFunction Status\tMaker\tModel\tSerial No.\tMaximo Equipment\tMaximo Equipment Description\tMaximo PM Details\tMaximo Job Plan Number\tMaximo Job Plan Task Number And Details\tJob Code\tJob Name\tJob Descriptions\tInterval\tCounter Type\tJob Category\tJob Type\tReminder\tWindow\tReminder / Window Unit\tResponsible Department\tRound\tScheduling Type\tLast Done Date\tLast Done Value\tLast Done Life\tJob Origin\tCriticalitiiy\tJob only linked to Function\tApproved By Boskalis\tx\ty\tz\ta\tb\tc";

            // Split the data into columns based on the tab character
            string[] dataTableColumns = data.Split('\t');
            MessageBox.Show(dataTableColumns.Length.ToString());
            for (int col = 0; col < dataTableColumns.Length; col++)
            {
                
                dataTable.Columns.Add(dataTableColumns[col]?.ToString() ?? $"Column{col}");
            }

            for (int i = 0; i < numRows; i++)

            {
                int rowsAdded = 0;
                OutputSheetData rowData = outputSheetData[i];
                int countMaximo = 0;
                int countJob = 0;
                dataRow = dataTable.Rows.Add();
                AddStaticColumns(i, rowData, ref dataRow);

                while (countJob != rowData.dataFromJobSheet.JobCode.Count)
                {
                    if (countJob == 0)
                    {
                        dataRow[19] = rowData.dataFromJobSheet.JobCode[countJob];
                        dataRow[20] = rowData.dataFromJobSheet.JobName[countJob];

                        dataRow[22] = rowData.dataFromJobSheet.Interval[countJob];
                        dataRow[23] = rowData.dataFromJobSheet.CounterType[countJob];
                        dataRow[24] = rowData.dataFromJobSheet.JobCategory[countJob];
                        dataRow[25] = rowData.dataFromJobSheet.JobType[countJob];
                        if (dataRow[22].ToString() != "")
                        {
                            dataRow[26] = rowData.dataFromJobSheet.Reminder[countJob];
                            dataRow[27] = rowData.dataFromJobSheet.Window[countJob];
                            dataRow[31] = rowData.dataFromJobSheet.SchedulingType[countJob];

                        }
                        dataRow[29] = rowData.dataFromJobSheet.ResponsibleDepartment[countJob];
                        dataRow[28] = rowData.dataFromJobSheet.ReminderWindowUnit[countJob];

                        dataRow[39] = rowData.MakerColor;
                        dataRow[40] = rowData.ModelColor;
                        dataRow[41] = rowData.SerialColor;
                        dataRow[42] = rowData.MaximoEqColor;
                        dataRow[43] = "True";
                    }
                    else
                    {
                        dataRow = dataTable.Rows.Add();
                        rowsAdded++;
                        AddStaticColumns(i, rowData, ref dataRow);
                        dataRow[19] = rowData.dataFromJobSheet.JobCode[countJob];
                        dataRow[20] = rowData.dataFromJobSheet.JobName[countJob];
                        dataRow[22] = rowData.dataFromJobSheet.Interval[countJob];
                        dataRow[23] = rowData.dataFromJobSheet.CounterType[countJob];
                        dataRow[24] = rowData.dataFromJobSheet.JobCategory[countJob];
                        dataRow[25] = rowData.dataFromJobSheet.JobType[countJob];
                        if (dataRow[22].ToString() != "")
                        {
                            dataRow[26] = rowData.dataFromJobSheet.Reminder[countJob];
                            dataRow[27] = rowData.dataFromJobSheet.Window[countJob];
                            dataRow[31] = rowData.dataFromJobSheet.SchedulingType[countJob];

                        }
                        dataRow[28] = rowData.dataFromJobSheet.ReminderWindowUnit[countJob];
                        dataRow[29] = rowData.dataFromJobSheet.ResponsibleDepartment[countJob];

                        dataRow[39] = rowData.MakerColor;
                        dataRow[40] = rowData.ModelColor;
                        dataRow[41] = rowData.SerialColor;
                        dataRow[42] = rowData.MaximoEqColor;
                        dataRow[43] = "True";
                       
                    }

                    countJob++;
                }

                while (countMaximo != rowData.dataFromMaximoSheet.MaximoJobPlanNumber.Count)
                {
                    dataRow = dataTable.Rows.Add();
                    rowsAdded++;
                    AddStaticColumns(i, rowData, ref dataRow);
                    dataRow[16] = rowData.dataFromMaximoSheet.MaximoPMDetails[countMaximo];
                    dataRow[17] = rowData.dataFromMaximoSheet.MaximoJobPlanNumber[countMaximo];
                    dataRow[18] = rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo];
                  //  dataRow[21] = rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo];
                    dataRow[20] = rowData.dataFromMaximoSheet.MaximoPMDetails[countMaximo];  // Job Name
                    dataRow[21] = MakeJobdescription(rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo]); // job Descriptions
                    dataRow[22] = rowData.dataFromMaximoSheet.Interval[countMaximo];
                    dataRow[23] = rowData.dataFromMaximoSheet.CounterType[countMaximo];
                    dataRow[26] = rowData.dataFromMaximoSheet.Reminder[countMaximo];
                    dataRow[27] = rowData.dataFromMaximoSheet.Window[countMaximo];
                    dataRow[29] = rowData.dataFromMaximoSheet.ResponsibleDepartment[countMaximo];
                    dataRow[31] = rowData.dataFromMaximoSheet.SchedulingType[countMaximo];
                    dataRow[32] = rowData.dataFromMaximoSheet.LastDoneDate[countMaximo];
                    dataRow[33] = rowData.dataFromMaximoSheet.LastDoneValue[countMaximo];
                    dataRow[35] = "Fleet Maintenance System";

                    dataRow[39] = rowData.MakerColor;
                    dataRow[40] = rowData.ModelColor;
                    dataRow[41] = rowData.SerialColor;
                    dataRow[42] = rowData.MaximoEqColor;
                    
                    dataRow[44] = "True";
                    countMaximo++;

                }

            }


            int rows = dataTable.Rows.Count;
            int cols = dataTable.Columns.Count;

            object[,] array2D = new Object[rows, cols];


            if (true)
            {

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        array2D[i, j] = dataTable.Rows[i][j];
                    }
                }
            }


            Range outputRange = worksheet.Range[$"A2:AM{rows}"];
            outputRange.Value = array2D;
            workbook.SaveAs(path);


            await Task.Run(() => WriteWorkBookColor(worksheet, dataTable))
                .ContinueWith(task =>
                {
                    // Handle completion or errors if needed
                    if (task.IsFaulted)
                    {
                        MessageBox.Show("Error: " + task.Exception?.ToString());
                    }
                    else
                    {
                        MessageBox.Show("output Coloring Done");
                        
                        worksheet.Columns.AutoFit();
                        workbook.Save();
                        ExcelHierarchyCon.ReleaseResources();
                    }
                    //ReleaseWriteWorkBook();
                });

        }
      /*  dataRow[16] = rowData.dataFromMaximoSheet.MaximoPMDetails[countMaximo];
                    dataRow[17] = rowData.dataFromMaximoSheet.MaximoJobPlanNumber[countMaximo];
                    dataRow[18] = rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo];
                    dataRow[32] = rowData.dataFromMaximoSheet.LastDoneDate[countMaximo];
                    dataRow[33] = rowData.dataFromMaximoSheet.LastDoneValue[countMaximo];
                    dataRow[21] = rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo];
                    dataRow[21] = MakeJobdescription(rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo]); // job Descriptions
        dataRow[20] = rowData.dataFromMaximoSheet.MaximoPMDetails[countMaximo];  // Job Name
                    dataRow[22] = rowData.dataFromMaximoSheet.Interval[countMaximo];
                    dataRow[31] = rowData.dataFromMaximoSheet.SchedulingType[countMaximo];
                    dataRow[26] = rowData.dataFromMaximoSheet.Reminder[countMaximo];
                    dataRow[27] = rowData.dataFromMaximoSheet.Window[countMaximo];
                    dataRow[29] = rowData.dataFromMaximoSheet.ResponsibleDepartment[countMaximo];
                    dataRow[23] = rowData.dataFromMaximoSheet.CounterType[countMaximo];
                    dataRow[35] = "Fleet Maintenance System";*/
        
        public static async Task WriteWorkBookColor(Worksheet WriteWorksheet, DataTable dataTable)
        {
            MessageBox.Show("Coloring is processing In background Please wait!");
            await Task.Run(() =>
            {
                for (int j = 0; j < dataTable.Rows.Count; j++)
                {
                    if (dataTable.Rows[j][43].ToString() == "True")
                    {
                        WriteWorksheet.Cells[j + 2, 20].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                        WriteWorksheet.Cells[j + 2, 21].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                        WriteWorksheet.Cells[j + 2, 23].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                        WriteWorksheet.Cells[j + 2, 24].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                        WriteWorksheet.Cells[j + 2, 25].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                        WriteWorksheet.Cells[j + 2, 26].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                        WriteWorksheet.Cells[j + 2, 29].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                        WriteWorksheet.Cells[j + 2, 30].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbForestGreen;
                    }
                    if (dataTable.Rows[j][44].ToString() == "True")
                    {
                        WriteWorksheet.Cells[j + 2, 17].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 18].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 19].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 33].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 34].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 21].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 22].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 23].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 24].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 27].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 28].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 30].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 30].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 32].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        WriteWorksheet.Cells[j + 2, 36].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                    }
                    if (dataTable.Rows[j][39].ToString() == "Green")
                    {
                        WriteWorksheet.Cells[j + 2, 12].Interior.Color = XlRgbColor.rgbGreen;
                    }
                    else if (dataTable.Rows[j][39].ToString() == "Orange")
                    {
                        WriteWorksheet.Cells[j + 2, 12].Interior.Color = XlRgbColor.rgbOrange;
                    }
                    else if (dataTable.Rows[j][39].ToString() == "Blue")
                    {
                        WriteWorksheet.Cells[j + 2, 12].Interior.Color = XlRgbColor.rgbBlue;
                    }
                    else if (dataTable.Rows[j][39].ToString() == "Red")
                    {
                        WriteWorksheet.Cells[j + 2, 12].Interior.Color = XlRgbColor.rgbRed;
                    }
                    if (dataTable.Rows[j][40].ToString() == "Green")
                    {
                        WriteWorksheet.Cells[j + 2, 13].Interior.Color = XlRgbColor.rgbGreen;
                    }
                    else if (dataTable.Rows[j][40].ToString() == "Orange")
                    {
                        WriteWorksheet.Cells[j + 2, 13].Interior.Color = XlRgbColor.rgbOrange;
                    }
                    else if (dataTable.Rows[j][40].ToString() == "Blue")
                    {
                        WriteWorksheet.Cells[j + 2, 13].Interior.Color = XlRgbColor.rgbBlue;
                    }
                    else if (dataTable.Rows[j][40].ToString() == "Red")
                    {
                        WriteWorksheet.Cells[j + 2, 13].Interior.Color = XlRgbColor.rgbRed;
                    }
                    if (dataTable.Rows[j][41].ToString() == "Green")
                    {
                        WriteWorksheet.Cells[j + 2, 14].Interior.Color = XlRgbColor.rgbGreen;
                    }
                    else if (dataTable.Rows[j][41].ToString() == "Orange")
                    {
                        WriteWorksheet.Cells[j + 2, 14].Interior.Color = XlRgbColor.rgbOrange;
                    }
                    else if (dataTable.Rows[j][41].ToString() == "Blue")
                    {
                        WriteWorksheet.Cells[j + 2, 14].Interior.Color = XlRgbColor.rgbBlue;
                    }
                    else if (dataTable.Rows[j][41].ToString() == "Red")
                    {
                        WriteWorksheet.Cells[j + 2, 14].Interior.Color = XlRgbColor.rgbRed;
                    }
                    if (dataTable.Rows[j][42].ToString() == "Yellow")
                    {
                        WriteWorksheet.Cells[j + 2, 15].Interior.Color = XlRgbColor.rgbYellow;
                    }
                }

            });
        }

        public string MakeJobdescription(string jobdescription)
        {
            string jobDesc = "Procedure: \n";
            int i = jobdescription.IndexOf('-');
            string[] arr = jobdescription.Split('\n');
            foreach (string s in arr)
            {
                if (s.Length > i)
                {
                    string temp = s.Substring(i).Trim();
                    jobDesc += "-" + temp + "\n";
                }
            }

            return jobDesc;

        }
        public void AddStaticColumns(int i1, OutputSheetData rowData, ref DataRow dataRow)
        {

            dataRow[0] = rowData.CodeInOutput;
            dataRow[1] = rowData.Path;
            dataRow[2] = rowData.Name;
            dataRow[3] = rowData.SequenceNo;
            dataRow[4] = rowData.FunctionType;
            dataRow[5] = rowData.Criticality;
            dataRow[6] = rowData.Location;
            dataRow[7] = rowData.ComponentTypeCode;
            dataRow[8] = rowData.ComponentType;
            dataRow[9] = rowData.ComponentClass;
            dataRow[10] = rowData.FunctionStatus;
            dataRow[11] = rowData.Maker;
            dataRow[12] = rowData.Model;
            dataRow[13] = rowData.SerialNo;
            dataRow[14] = rowData.MaximoEq;
            dataRow[15] = rowData.MaximoEqDescription;


        }

        public List<List<OutputSheetData>> SplitData(List<OutputSheetData> outputData)
        {
            string splitKeyword = "Group Level 2";
            string codeNo = "";
            List<List<OutputSheetData>> result = new List<List<OutputSheetData>>();
            List<OutputSheetData> currentSplit = new List<OutputSheetData>();

            int count = 0;

            foreach (OutputSheetData row in outputData)
            {
                if (count == 0)
                {
                    codeNo = row.CodeInOutput;
                }

                if (row.FunctionType == splitKeyword && row.CodeInOutput != codeNo)
                {
                    if (currentSplit.Any())
                    {
                        result.Add(new List<OutputSheetData>(currentSplit));
                        currentSplit.Clear();
                    }
                }

                currentSplit.Add(row);
                count++;
            }

            // Add the last split if there are any remaining rows
            if (currentSplit.Any())
            {
                result.Add(new List<OutputSheetData>(currentSplit));
            }

            return result;
        }

    }
}