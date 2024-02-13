﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Label = System.Windows.Forms.Label;

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

            List<OutputSheetData> outputSheetData = new List<OutputSheetData>();
            int countRows = 2;
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
                singleRow.ComponentType = $"=IF(AND(ISBLANK(J{countRows}), ISBLANK(L{countRows}), ISBLANK(M{countRows})), \"\", IF(AND(ISBLANK(J{countRows}), ISBLANK(L{countRows})), M{countRows}, IF(ISBLANK(J{countRows}), IF(ISBLANK(L{countRows}), M{countRows}, CONCATENATE(L{countRows}, \"/\", M{countRows})), CONCATENATE(J{countRows}, IF(ISBLANK(L{countRows}), \"\", CONCATENATE(\"/\", L{countRows})), IF(ISBLANK(M{countRows}), \"\", CONCATENATE(\"/\", M{countRows}))))))"; ;
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
                countRows++;

                for (int i = 0; i < jobSheetData.Count; i++)
                {
                    JobSheetData singleRowJobData = jobSheetData[i];
                    if (singleRowJobData.CodeInJob == singleRow.CodeInOutput)
                    {

                        singleRow.dataFromJobSheet = singleRowJobData;
                        break;
                    }

                }  // addinf =

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

        public void WriteDataInOutputAsync(List<OutputSheetData> outputSheetData, Worksheet worksheet, Workbook workbook, string path, [Optional] Label label)
        {
            int numRows = outputSheetData.Count;

            System.Data.DataTable dataTable = new System.Data.DataTable();
            DataRow dataRow;
            string data = "Code\tPath\tName\tSequence No\tFunction Type\tCriticality\tLocation\tComponent Type Code\tComponent Type\tComponent Class\tFunction Status\tMaker\tModel\tSerial No.\tMaximo Equipment\tMaximo Equipment Description\tMaximo PM Details\tMaximo Job Plan Number\tMaximo Job Plan Task Number And Details\tJob Code\tJob Name\tJob Descriptions\tInterval\tCounter Type\tJob Category\tJob Type\tReminder\tWindow\tReminder / Window Unit\tResponsible Department\tRound\tScheduling Type\tLast Done Date\tLast Done Value\tLast Done Life\tJob Origin\tCriticalitiiy\tJob only linked to Function\tApproved By Boskalis\tx\ty\tz\ta\tb\tc";

            // Split the data into columns based on the tab character
            string[] dataTableColumns = data.Split('\t');
            if (label != null)
            {
                label.Text = "Writing Headers";
            }
            for (int col = 0; col < dataTableColumns.Length; col++)
            {
                dataTable.Columns.Add(dataTableColumns[col]?.ToString() ?? $"Column{col}");
            }

            if (label != null)
            {
                label.Text = "Merging Data From Job Sheet and Maximo Sheet To Output Sheet";
            }
            int rowsAdded = 1;
            for (int i = 0; i < numRows; i++)
            {

                OutputSheetData rowData = outputSheetData[i];

                int countMaximo = 0;
                int countJob = 0;
                dataRow = dataTable.Rows.Add();
                rowsAdded++;

                AddStaticColumns(i, rowData, ref dataRow, rowsAdded);

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
                        dataRow[26] = rowData.dataFromJobSheet.Reminder[countJob];
                        dataRow[27] = rowData.dataFromJobSheet.Window[countJob];
                        dataRow[32] = rowData.dataFromJobSheet.SchedulingType[countJob];
                        dataRow[29] = rowData.dataFromJobSheet.ResponsibleDepartment[countJob];
                        dataRow[28] = rowData.dataFromJobSheet.ReminderWindowUnit[countJob];

                        dataRow[43] = "True";
                    }
                    else
                    {
                        dataRow = dataTable.Rows.Add();
                        rowsAdded++;
                        AddStaticColumns(i, rowData, ref dataRow, rowsAdded);
                        dataRow[19] = rowData.dataFromJobSheet.JobCode[countJob];
                        dataRow[20] = rowData.dataFromJobSheet.JobName[countJob];
                        dataRow[22] = rowData.dataFromJobSheet.Interval[countJob];
                        dataRow[23] = rowData.dataFromJobSheet.CounterType[countJob];
                        dataRow[24] = rowData.dataFromJobSheet.JobCategory[countJob];
                        dataRow[25] = rowData.dataFromJobSheet.JobType[countJob];
                        dataRow[26] = rowData.dataFromJobSheet.Reminder[countJob];
                        dataRow[27] = rowData.dataFromJobSheet.Window[countJob];
                        dataRow[32] = rowData.dataFromJobSheet.SchedulingType[countJob];
                        dataRow[28] = rowData.dataFromJobSheet.ReminderWindowUnit[countJob];
                        dataRow[29] = rowData.dataFromJobSheet.ResponsibleDepartment[countJob];

                        dataRow[43] = "True";

                    }

                    countJob++;
                }

                while (countMaximo != rowData.dataFromMaximoSheet.MaximoJobPlanNumber.Count)
                {
                    dataRow = dataTable.Rows.Add();
                    rowsAdded++;


                    AddStaticColumns(i, rowData, ref dataRow, rowsAdded);
                    dataRow[16] = rowData.dataFromMaximoSheet.MaximoPMDetails[countMaximo];
                    dataRow[17] = rowData.dataFromMaximoSheet.MaximoJobPlanNumber[countMaximo];
                    dataRow[18] = rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo];
                    dataRow[20] = rowData.dataFromMaximoSheet.MaximoPMDetails[countMaximo];  // Job Name
                    dataRow[21] = MakeJobdescription(rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo]); // job Descriptions
                    dataRow[22] = rowData.dataFromMaximoSheet.Interval[countMaximo];
                    dataRow[23] = rowData.dataFromMaximoSheet.CounterType[countMaximo];
                    dataRow[26] = rowData.dataFromMaximoSheet.Reminder[countMaximo];
                    dataRow[27] = rowData.dataFromMaximoSheet.Window[countMaximo];
                    dataRow[28] = rowData.dataFromMaximoSheet.ReminderWindowUnit[countMaximo];
                    dataRow[32] = rowData.dataFromMaximoSheet.SchedulingType[countMaximo];
                    dataRow[29] = rowData.dataFromMaximoSheet.ResponsibleDepartment[countMaximo];
                    dataRow[33] = rowData.dataFromMaximoSheet.LastDoneDate[countMaximo];
                    dataRow[34] = rowData.dataFromMaximoSheet.LastDoneValue[countMaximo];
                    dataRow[36] = "Fleet Maintenance System";



                    dataRow[44] = "True";
                    countMaximo++;

                }

            }


            int rows = dataTable.Rows.Count;
            int cols = dataTable.Columns.Count;

            object[,] array2D = new Object[rows, cols - 6];


            if (true)
            {

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols - 6; j++)
                    {
                        array2D[i, j] = dataTable.Rows[i][j];
                    }
                }
            }


            Range outputRange = worksheet.Range[$"A2:AM{rows + 1}"];
            outputRange.Value = array2D;
            workbook.SaveAs(path);

            if (label != null)
            {
                label.Text = "Coloring Cells";
            }

            if (label != null)
            {
                WriteWorkBookColor(worksheet, dataTable, label);
            }

            else
            {
                WriteWorkBookColor(worksheet, dataTable);
            }

            worksheet.Columns.AutoFit();

            if (label != null)
            {
                label.Text = "Saving Output File";
            }
            workbook.Save();



        }


        public void WriteWorkBookColor(Worksheet WriteWorksheet, DataTable dataTable, [Optional] Label label)
        {


            for (int j = 0; j < dataTable.Rows.Count; j++)
            {

                if (label != null)
                {
                    label.Text = $"Coloring {j + 2} row / {WriteWorksheet.UsedRange.Rows.Count}";
                }

                if (dataTable.Rows[j][43].ToString() == "True")
                {
                    string s1 = $"T{j + 2}:U{j + 2}" + "," + $"W{j + 2}:Z{j + 2}" + "," + $"AC{j + 2}:AD{j + 2}";

                    Range r1 = WriteWorksheet.Range[s1];



                    r1.Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;

                    if (dataTable.Rows[j][32].ToString() != "")
                    {
                        WriteWorksheet.Cells[j + 2, 33].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrange;
                    }
                    WriteWorksheet.Cells[j + 2, 27].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrange;
                    WriteWorksheet.Cells[j + 2, 28].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrange;

                }
                if (dataTable.Rows[j][44].ToString() == "True")
                {

                    string s2 = $"Q{j + 2}:S{j + 2}" + "," + $"U{j + 2}:X{j + 2}" + "," + $"AC{j + 2}:AD{j + 2}";

                    Range r2 = WriteWorksheet.Range[s2];
                    r2.Interior.Color = XlRgbColor.rgbGreen;
                    WriteWorksheet.Cells[j + 2, 27].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrange;
                    WriteWorksheet.Cells[j + 2, 28].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrange;

                    if (dataTable.Rows[j][32].ToString() != "")
                    {
                        WriteWorksheet.Cells[j + 2, 33].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrange;
                    }
                    WriteWorksheet.Cells[j + 2, 37].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;

                    if (dataTable.Rows[j][33].ToString() != "")
                    {

                        WriteWorksheet.Cells[j + 2, 34].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                    }

                    if (dataTable.Rows[j][34].ToString() != "")
                    {

                        WriteWorksheet.Cells[j + 2, 35].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                    }

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


        }



        public string MakeJobdescription(string jobdescription)
        {
            string jobDesc = "Procedure: \n";
            int i = jobdescription.IndexOf('-');

            if (i < 0)
            {
                return "";
            }


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
        public void AddStaticColumns(int i1, OutputSheetData rowData, ref DataRow dataRow, int countRows)
        {

            dataRow[0] = rowData.CodeInOutput;
            dataRow[1] = rowData.Path;
            dataRow[2] = rowData.Name;
            dataRow[3] = rowData.SequenceNo;
            dataRow[4] = rowData.FunctionType;
            dataRow[5] = rowData.Criticality;
            dataRow[6] = rowData.Location;
            dataRow[7] = rowData.ComponentTypeCode;
            dataRow[8] = $"=IF(AND(ISBLANK(J{countRows}), ISBLANK(L{countRows}), ISBLANK(M{countRows})), \"\", IF(AND(ISBLANK(J{countRows}), ISBLANK(L{countRows})), M{countRows}, IF(ISBLANK(J{countRows}), IF(ISBLANK(L{countRows}), M{countRows}, CONCATENATE(L{countRows}, \"/\", M{countRows})), CONCATENATE(J{countRows}, IF(ISBLANK(L{countRows}), \"\", CONCATENATE(\"/\", L{countRows})), IF(ISBLANK(M{countRows}), \"\", CONCATENATE(\"/\", M{countRows}))))))";
            dataRow[9] = rowData.ComponentClass;
            dataRow[10] = rowData.FunctionStatus;
            dataRow[11] = rowData.Maker;
            dataRow[12] = rowData.Model;
            dataRow[13] = rowData.SerialNo;
            dataRow[14] = rowData.MaximoEq;
            dataRow[15] = rowData.MaximoEqDescription;

            dataRow[39] = rowData.MakerColor;
            dataRow[40] = rowData.ModelColor;
            dataRow[41] = rowData.SerialColor;
            dataRow[42] = rowData.MaximoEqColor;


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