﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHierarchyConversion_InterOp
{
    internal class MaximoSheetData
    {

        public string AssetNumber { get; set; }
        public List<string> Interval { get; set; } //  // Maps with Frequency
        public List<string> Counter { get; set; } //  // Maps with Frequency Duration
        public List<string> LastDoneDate { get; set; } // Maps with Last Done Date
        public List<string> LastDoneValue { get; set; } // Maps with At reading
        public List<string> MaximoJobPlanNumber { get; set; }    // Maps With JP Number [insert rows] 
        public List<string> MaximoJobPlanTaskNumberAndDetails { get; set; }    // merge all the Job task number and job task description containing same JP Number
        public List<string> MaximoPMDetails { get; set; }
        public List<string> ResponsibleDepartment { get; set; }
        public List<string> Reminder { get; set; }
        public List<string> Window { get; set; }
        public List<string> SchedulingType { get; set; }

        public MaximoSheetData()
        {
            AssetNumber = "";
            Interval = new List<string>();
            Counter = new List<string>();
            LastDoneDate = new List<string>();
            LastDoneValue = new List<string>();
            MaximoJobPlanNumber = new List<string>();
            MaximoJobPlanTaskNumberAndDetails = new List<string>();
            MaximoPMDetails = new List<string>();
            ResponsibleDepartment = new List<string>();
            Reminder = new List<string>();
            Window = new List<string>();
            SchedulingType = new List<string>();

        }

        public List<MaximoSheetData> ReadDataFromMaximoSheet(Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
            object[,] data = usedRange.Value;

            int rowCount = data.GetLength(0);

            //---------------------Reading the data and storing it in a list of MaximoSheetData------------------------//
            List<MaximoSheetData> rows = new List<MaximoSheetData>();

            int chunkSize = 1000;

            for (int rowIdx = 2; rowIdx <= rowCount; rowIdx += chunkSize)
            {
                int rowsToRead = Math.Min(chunkSize, rowCount - rowIdx + 1);

                for (int i = rowIdx; i < rowIdx + rowsToRead; i++)
                {
                    int temp = i;

                    string jobTaskNumber = "";
                    string jobTaskDesc = "";
                    string mergedData = "";
                    int jpNumberCount = 0;

                    MaximoSheetData singleRow = new MaximoSheetData();

                    singleRow.AssetNumber = Convert.ToString(data[i, 5]);
                    singleRow.MaximoJobPlanNumber.Add(Convert.ToString(data[i, 7]));  //Jp number

                    singleRow.Interval.Add(Convert.ToString(data[i, 16])); //Frequency
                    singleRow.Counter.Add(Convert.ToString(data[i, 17]));// Frequency Duration 
                    singleRow.LastDoneDate.Add(Convert.ToString(data[i, 18]));// LastDoneDte 
                    singleRow.LastDoneValue.Add(Convert.ToString(data[i, 19]));// At reading 
                    singleRow.MaximoPMDetails.Add(Convert.ToString(data[i, 2])); // PM description

                    jobTaskNumber = Convert.ToString(data[i, 8]);  //10-20 like 
                    jobTaskDesc = Convert.ToString(data[i, 9]);   // Job task Desc
                    mergedData = "\n" + jobTaskNumber + " -" + jobTaskDesc;

                    if (temp < rowCount - 1 && singleRow.AssetNumber != Convert.ToString(data[temp + 1, 5]))
                    {
                        singleRow.MaximoJobPlanTaskNumberAndDetails.Add(mergedData);
                    }

                    else
                    {
                        while (temp < rowCount - 1 && singleRow.AssetNumber == Convert.ToString(data[temp + 1, 5]))
                        {
                            if (singleRow.MaximoJobPlanNumber[jpNumberCount] != Convert.ToString(data[temp + 1, 7]))
                            {


                                singleRow.MaximoJobPlanNumber.Add(Convert.ToString(data[temp + 1, 7]));  // thena add Jp number
                                singleRow.Interval.Add(Convert.ToString(data[temp + 1, 16])); //Frequency
                                singleRow.Counter.Add(Convert.ToString(data[temp + 1, 17]));// Frequency Duration 
                                singleRow.LastDoneDate.Add(Convert.ToString(data[temp + 1, 18]));// LastDoneDte 
                                singleRow.LastDoneValue.Add(Convert.ToString(data[temp + 1, 19]));// At reading 
                                singleRow.MaximoPMDetails.Add(Convert.ToString(data[temp + 1, 2])); // PM description
                                singleRow.MaximoJobPlanTaskNumberAndDetails.Add(mergedData);
                                mergedData = "";
                                jpNumberCount++;

                            } //if next JpNumber not matched



                            jobTaskNumber = Convert.ToString(data[temp + 1, 8]);  //10-20 like 
                            jobTaskDesc = Convert.ToString(data[temp + 1, 9]);   // Job task Desc
                            mergedData = mergedData + "\n" + jobTaskNumber + " -" + jobTaskDesc;

                            temp++;
                            i++;
                        } // this loop is for Asset Number when next is same

                        if (mergedData != "")
                        {
                            singleRow.MaximoJobPlanTaskNumberAndDetails.Add(mergedData);

                        }


                        //----------------------Handling ReminderWindow and schedulign type ------------------------\\

                        /*for (int i = 0; i < singleRow.Interval.Count, i++)
                        {


                            string unit = rows[i].CounterType;

                            if (string.Equals(unit, "Weeks", StringComparison.OrdinalIgnoreCase))
                            {
                                rows[i].Reminder = Math.Round(0.07 * interval * 7);
                                rows[i].Window = Math.Round(0.1 * interval * 7);
                            }

                            else if (string.Equals(unit, "Months", StringComparison.OrdinalIgnoreCase))
                            {
                                rows[i].Reminder = Math.Round(0.07 * interval * 30);
                                rows[i].Window = Math.Round(0.1 * interval * 30);
                            }

                            else if (string.Equals(unit, "Years", StringComparison.OrdinalIgnoreCase))
                            {
                                rows[i].Reminder = Math.Round(0.07 * interval * 365);
                                rows[i].Window = Math.Round(0.1 * interval * 365);
                            }

                            else if (string.Equals(unit, "Days", StringComparison.OrdinalIgnoreCase))
                            {
                                rows[i].Reminder = Math.Round(0.07 * interval);
                                rows[i].Window = Math.Round(0.1 * interval);
                            }

                            else if (string.Equals(unit, "Hours", StringComparison.OrdinalIgnoreCase))
                            {
                                rows[i].Reminder = Math.Round(0.07 * interval);
                                rows[i].Window = Math.Round(0.1 * interval);
                            }
                        }*/
                    }
                    rows.Add(singleRow);
                }
            }
            return rows;

        }
    }
}