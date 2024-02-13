using Microsoft.Office.Interop.Excel;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHierarchyConversion_InterOp
{
    internal class JobSheetData
    {
        public string CodeInJob { get; set; }   // A [1]
        public List<string> ComponentClass { get; set; } // J [10]
        public List<string> JobOrigin { get; set; } // [AF] 32
        public List<string> JobCode { get; set; } // O [15]
        public List<string> JobName { get; set; }  // P [16]
        public List<string> Interval { get; set; } // R [18]
        public List<string> CounterType { get; set; } // S [19]
        public List<string> JobCategory { get; set; } // T [20]
        public List<string> JobType { get; set; } // U [21]    
        public List<string> ReminderWindowUnit { get; set; } // X [24]
        public List<string> ResponsibleDepartment { get; set; } // Y [25]
        public List<string> Round { get; set; } // Z[26]
        public List<string> Reminder { get; set; }
        public List<string> Window { get; set; }
        public List<string> SchedulingType { get; set; }
        public int rowNumber { get; set; } = 0;



        public JobSheetData()
        {
            CodeInJob = string.Empty;
            ComponentClass = new List<string>();
            JobCode = new List<string>();
            JobName = new List<string>();
            Interval = new List<string>();
            CounterType = new List<string>();
            JobCategory = new List<string>();
            JobType = new List<string>();
            ResponsibleDepartment = new List<string>();
            Round = new List<string>();
            ReminderWindowUnit = new List<string>();
            Reminder = new List<string>();
            Window = new List<string>();
            SchedulingType = new List<string>();
            JobOrigin = new List<string>();
        }


        public List<JobSheetData> ReadDataFromJobSheet(Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
            object[,] data = usedRange.Value;

            int rowCount = data.GetLength(0);

            //---------------------Reading the data and storing it in a list of RowData------------------------//
            List<JobSheetData> rows = new List<JobSheetData>();

            int chunkSize = 1000;

            for (int rowIdx = 2; rowIdx <= rowCount; rowIdx += chunkSize)
            {
                int rowsToRead = Math.Min(chunkSize, rowCount - rowIdx + 1);

                for (int i = rowIdx; i < rowIdx + rowsToRead; i++)
                {
                    int temp = i;
                    JobSheetData singleRow = new JobSheetData();  // Holds data for single Row
                    singleRow.rowNumber = i;
                    singleRow.CodeInJob = Convert.ToString(data[i, 1]);
                    singleRow.ComponentClass.Add(Convert.ToString(data[i, 10]));
                    singleRow.JobCode.Add(Convert.ToString(data[i, 15]));
                    singleRow.JobName.Add(Convert.ToString(data[i, 16]));
                    singleRow.Interval.Add(Convert.ToString(data[i, 18]));
                    singleRow.CounterType.Add(Convert.ToString(data[i, 19]));
                    singleRow.JobCategory.Add(Convert.ToString(data[i, 20]));
                    singleRow.JobType.Add(Convert.ToString(data[i, 21]));
                    singleRow.ReminderWindowUnit.Add(Convert.ToString(data[i, 24]));
                    singleRow.ResponsibleDepartment.Add(Convert.ToString(data[i, 25]));
                    singleRow.Round.Add(Convert.ToString(data[i, 26]));
                    singleRow.JobOrigin.Add((data[i, 32]??"").ToString());

                    if (singleRow.CodeInJob != "")
                    {

                        while (temp < rowCount - 1 && Convert.ToString(data[temp + 1, 1]) == singleRow.CodeInJob)
                        {
                            singleRow.ComponentClass.Add(Convert.ToString(data[temp + 1, 10]));
                            singleRow.JobCode.Add(Convert.ToString(data[temp + 1, 15]));
                            singleRow.JobName.Add(Convert.ToString(data[temp + 1, 16]));
                            singleRow.Interval.Add(Convert.ToString(data[temp + 1, 18]));
                            singleRow.CounterType.Add(Convert.ToString(data[temp + 1, 19]));
                            singleRow.JobCategory.Add(Convert.ToString(data[temp + 1, 20]));
                            singleRow.JobType.Add(Convert.ToString(data[temp + 1, 21]));
                            singleRow.ReminderWindowUnit.Add(Convert.ToString(data[temp + 1, 24]));
                            singleRow.ResponsibleDepartment.Add(Convert.ToString(data[temp + 1, 25]));
                            singleRow.Round.Add(Convert.ToString(data[temp + 1, 26]));
                            singleRow.JobOrigin.Add((data[i, 32] ?? "").ToString());

                            i++;
                            temp++;
                        }

                    }

                    else
                    {
                        singleRow.rowNumber = i;
                        singleRow.CodeInJob = Convert.ToString(data[i, 1]);
                        singleRow.ComponentClass.Add(Convert.ToString(data[i, 10]));
                        singleRow.JobCode.Add(Convert.ToString(data[i, 15]));
                        singleRow.JobName.Add(Convert.ToString(data[i, 16]));
                        singleRow.Interval.Add(Convert.ToString(data[i, 18]));
                        singleRow.CounterType.Add(Convert.ToString(data[i, 19]));
                        singleRow.JobCategory.Add(Convert.ToString(data[i, 20]));
                        singleRow.JobType.Add(Convert.ToString(data[i, 21]));
                        singleRow.ReminderWindowUnit.Add(Convert.ToString(data[i, 24]));
                        singleRow.ResponsibleDepartment.Add(Convert.ToString(data[i, 25]));
                        singleRow.Round.Add(Convert.ToString(data[i, 26]));
                        singleRow.Round.Add((data[i, 32] ?? "").ToString());

                    }
                    for (int j = 0; j < singleRow.Interval.Count; j++)
                    {
                        string unit = singleRow.CounterType[j];
                        if (int.TryParse(singleRow.Interval[j], out int interval))
                        {

                            if (string.Equals(unit, "Weeks", StringComparison.OrdinalIgnoreCase))
                            {
                                singleRow.Reminder.Add((Math.Round(0.07 * interval * 7)).ToString());
                                singleRow.Window.Add((Math.Round(0.1 * interval * 7)).ToString());
                                if (interval <= 4)
                                {
                                    singleRow.SchedulingType.Add("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType.Add("Scheduled");
                                }
                            }

                            else if (string.Equals(unit, "Months", StringComparison.OrdinalIgnoreCase))
                            {
                                singleRow.Reminder.Add((Math.Round(0.07 * interval * 30)).ToString());
                                singleRow.Window.Add((Math.Round(0.1 * interval * 30)).ToString());
                                if (interval <= 1)
                                {
                                    singleRow.SchedulingType.Add("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType.Add("Scheduled");
                                }
                            }

                            else if (string.Equals(unit, "Years", StringComparison.OrdinalIgnoreCase))
                            {
                                singleRow.Reminder.Add((Math.Round(0.07 * interval * 365)).ToString());
                                singleRow.Window.Add((Math.Round(0.1 * interval * 365)).ToString());

                                singleRow.SchedulingType.Add("Scheduled");

                            }

                            else if (string.Equals(unit, "Days", StringComparison.OrdinalIgnoreCase))
                            {
                                singleRow.Reminder.Add((Math.Round(0.07 * interval)).ToString());
                                singleRow.Window.Add((Math.Round(0.1 * interval)).ToString());
                                if (interval <= 30)
                                {
                                    singleRow.SchedulingType.Add("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType.Add("Scheduled");
                                }
                            }

                            else if (string.Equals(unit, "HR", StringComparison.OrdinalIgnoreCase))
                            {
                                singleRow.Reminder.Add((Math.Round(0.07 * interval)).ToString());
                                singleRow.Window.Add((Math.Round(0.1 * interval)).ToString());
                                if (interval <= 720)
                                {
                                    singleRow.SchedulingType.Add("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType.Add("Scheduled");
                                }

                            }
                        }

                        else
                        {
                            singleRow.Reminder.Add("");
                            singleRow.SchedulingType.Add("");
                            singleRow.Window.Add("");
                        }


                    }
                    rows.Add(singleRow);
                }
            }
            return rows;

        }
    }
}

