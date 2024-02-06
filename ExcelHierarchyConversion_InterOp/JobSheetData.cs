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
        public List<string> JobCode { get; set; } // O [15]
        public List<string> JobName { get; set; }  // P [16]
        public List<string> Interval { get; set; } // R [18]
        public List<string> CounterType { get; set; } // S [19]
        public List<string> JobCategory { get; set; } // T [20]
        public List<string> JobType { get; set; } // U [21]    
        public List<string> ReminderWindowUnit { get; set; } // X [24]
        public List<string> ResponsibleDepartment { get; set; } // Y [25]
        public List<string> Round { get; set; } // Z[26]


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

                            i++;
                            temp++;
                        }

                    }

                    rows.Add(singleRow);

                }
            }
            return rows;

        }




    }
}
