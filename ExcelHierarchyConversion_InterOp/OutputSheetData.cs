using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

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
        }
        public List<OutputSheetData> MapDataToOutputSheet(List<List<string>> outputData, [Optional] List<JobSheetData> jobSheetData)
        {
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

                outputSheetData.Add(singleRow);
            }

            return outputSheetData;

        }

    }
}
