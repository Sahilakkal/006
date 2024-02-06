using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
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




        public string ComponentClass { get; set; } //  J [10]
        public string FunctionStatus { get; set; } //  K [11]
        public string Maker { get; set; } //  L [12]
        public string Model { get; set; } //  M [13]
        public string SerialNo { get; set; } //  N [14]
        public string MaximoEq { get; set; } //  O [15]
        public string MaximoEqDescription { get; set; } //  P [16]
        public string ColorYellow { get; set; }
        public string ColorGreen { get; set; }
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

                        MessageBox.Show("data added" + singleRowMaximoData.AssetNumber);
                        break;
                    }
                }

                outputSheetData.Add(singleRow);
            }

            return outputSheetData;

        }

        public void WriteDataInOutput(List<OutputSheetData> outputSheetData,Worksheet worksheet)
        {


            int startRow = 10;
            int numRows = outputSheetData.Count;
            int numColumns = 18;

            /*  Range dataRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow + numRows - 1, numColumns]];

              dataRange.NumberFormat = "@";*/

            System.Data.DataTable dataTable = new System.Data.DataTable();
            DataRow dataRow;
            string data = "Code\tPath\tName\tSequence No\tFunction Type\tCriticality\tLocation\tComponent Type Code\tComponent Type\tComponent Class\tFunction Status\tMaker\tModel\tSerial No.\tMaximo Equipment\tMaximo Equipment Description\tMaximo PM Details\tMaximo Job Plan Number\tMaximo Job Plan Task Number And Details\tJob Code\tJob Name\tJob Descriptions\tInterval\tCounter Type\tJob Category\tJob Type\tReminder\tWindow\tReminder / Window Unit\tResponsible Department\tRound\tScheduling Type\tLast Done Date\tLast Done Value\tLast Done Life\tJob Origin\tCriticalitiiy\tJob only linked to Function\tApproved By Boskalis";

            // Split the data into columns based on the tab character
            string[] dataTableColumns = data.Split('\t');

            for (int col = 0; col < dataTableColumns.Length; col++)
            {
                dataTable.Columns.Add(dataTableColumns[col]?.ToString() ?? $"Column{col}");
            }

            for (int i = 0; i < numRows; i++)
            {
                OutputSheetData rowData = outputSheetData[i];
                int countMaximo = 0;
                int countJob = 0;


                while (countMaximo != rowData.dataFromMaximoSheet.MaximoJobPlanNumber.Count || countJob != rowData.dataFromJobSheet.JobCode.Count)
                {
                    dataRow = dataTable.Rows.Add();
                    AddStaticColumns(i, rowData, ref dataRow);

                    if (countMaximo != rowData.dataFromMaximoSheet.MaximoJobPlanNumber.Count)
                    {
                        dataRow[16] = rowData.dataFromMaximoSheet.MaximoPMDetails[countMaximo];
                        dataRow[17] = rowData.dataFromMaximoSheet.MaximoJobPlanNumber[countMaximo];
                        dataRow[18] = rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countMaximo];
                        //dataRow[26] = rowData.dataFromMaximoSheet.Reminder[countJob];
                        //dataRow[27] = rowData.dataFromMaximoSheet.Window[countJob];
                        // dataRow[29] = rowData.dataFromMaximoSheet.ResponsibleDepartment[countJob];
                        dataRow[32] = rowData.dataFromMaximoSheet.LastDoneDate[countMaximo];
                        dataRow[33] = rowData.dataFromMaximoSheet.LastDoneValue[countMaximo];
                        dataRow[21] = rowData.dataFromMaximoSheet.MaximoJobPlanTaskNumberAndDetails[countJob];
                        dataRow[22] = rowData.dataFromMaximoSheet.Interval[countJob];
                        countMaximo++;

                    }

                    if (countJob != rowData.dataFromJobSheet.JobCode.Count)
                    {
                        dataRow[19] = rowData.dataFromJobSheet.JobCode[countJob];
                        dataRow[20] = rowData.dataFromJobSheet.JobName[countJob];

                        dataRow[23] = rowData.dataFromJobSheet.CounterType[countJob];
                        dataRow[24] = rowData.dataFromJobSheet.JobCategory[countJob];
                        dataRow[25] = rowData.dataFromJobSheet.JobType[countJob];
                        dataRow[28] = rowData.dataFromJobSheet.ReminderWindowUnit[countJob];
                        countJob++;
                        //dataRow[30] = rowData.dataFromJobSheet.Round[countJob];
                    }
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

    }
}

