using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHierarchyConversion_InterOp
{
    internal class MaximoSheetData
    {

        public string AssetNumber { get; set; }
        public string JpNumber { get; set; }
        public string JobTaskDescription { get; set; }
        public string PMDescription { get; set; }
        public string Frequency { get; set; }
        public List<string> Interval { get; set; } //  // Maps with Frequency
        public List<string> Counter { get; set; } //  // Maps with Counter Column
        public List<string> LastDoneDate { get; set; } // Maps with Last Done Date
        public List<string> LastDoneValue { get; set; } // Maps with Last Done Date
        public List<string> MaximoJobPlanNumber { get; set; }    // Maps With JP Number [insert rows] 
        public List<string> MaximoJobPlanTaskNumberAndDetails { get; set; }    // merge all the Job task number and job task description containing same JP Number
        public List<string> MaximoPMDetails { get; set; }
        public List<string> ResponsibleDepartment { get; set; }
        public List<string> Reminder { get; set; }
        public List<string> Window { get; set; }
        public List<string> SchedulingType { get; set; }

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
                    MaximoSheetData singleRow=new MaximoSheetData();

                    singleRow.AssetNumber = Convert.ToString(data[i,])

                }
            }

        }
    }
}