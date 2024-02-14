using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHierarchyConversion_InterOp
{
    internal class MaximoSheetData
    {

        public string AssetNumber { get; set; }
        public List<string> Interval { get; set; } //  // Maps with Frequency
        public List<string> CounterType { get; set; } //  // Maps with Frequency Duration
        public List<string> LastDoneDate { get; set; } // Maps with Last Done Date
        public List<string> LastDoneValue { get; set; } // Maps w
                                                        //
                                                        // ith At reading
        public List<string> MaximoJobPlanNumber { get; set; }    // Maps With JP Number [insert rows] 
        public List<string> MaximoJobPlanTaskNumberAndDetails { get; set; }    // merge all the Job task number and job task description containing same JP Number
        public List<string> MaximoPMDetails { get; set; }
        public List<string> ResponsibleDepartment { get; set; }
        public List<string> Reminder { get; set; }
        public List<string> Window { get; set; }
        public List<string> SchedulingType { get; set; }
        public int rowNumber { get; set; } = 0;
        public List<string> ReminderWindowUnit { get; set; }




        public MaximoSheetData()
        {
            AssetNumber = "";
            Interval = new List<string>();
            CounterType = new List<string>();
            LastDoneDate = new List<string>();
            LastDoneValue = new List<string>();
            MaximoJobPlanNumber = new List<string>();
            MaximoJobPlanTaskNumberAndDetails = new List<string>();
            MaximoPMDetails = new List<string>();
            ResponsibleDepartment = new List<string>();
            Reminder = new List<string>();
            Window = new List<string>();
            SchedulingType = new List<string>();
            ReminderWindowUnit = new List<string>();

        }

        public List<MaximoSheetData> ReadDataFromMaximoSheet(Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
            object[,] data = usedRange.Value;

            int rowCount = data.GetLength(0);

            //
            //
            //
            //--------------------Reading the data and storing it in a list of MaximoSheetData------------------------//
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
                    string pmDescription = "";
                    bool isEqNoPresent = false;

                    MaximoSheetData singleRow = new MaximoSheetData();
                    singleRow.rowNumber = i;

                    singleRow.AssetNumber = Convert.ToString(data[i, 5]);
                    singleRow.MaximoJobPlanNumber.Add(Convert.ToString(data[i, 7]));  //Jp number

                    singleRow.Interval.Add(Convert.ToString(data[i, 16])); //Frequency
                    singleRow.CounterType.Add(Convert.ToString(data[i, 17]));// Frequency Duration 
                    singleRow.LastDoneDate.Add(Convert.ToString(data[i, 18]));// LastDoneDte 
                    singleRow.LastDoneValue.Add(Convert.ToString(data[i, 19]));// At reading 
                    singleRow.MaximoPMDetails.Add(Convert.ToString(data[i, 2]));
                    pmDescription = Convert.ToString(data[i, 2]);// PM description

                    if ((data[i, 26] ?? "").ToString() != "")
                    {
                        isEqNoPresent = true;
                    }//EqNumber



                    jobTaskNumber = Convert.ToString(data[i, 8]);  //10-20 like 
                    jobTaskDesc = Convert.ToString(data[i, 9]);   // Job task Desc
                    mergedData = "\n" + jobTaskNumber + "-" + jobTaskDesc;

                    if (temp < rowCount && singleRow.AssetNumber != Convert.ToString(data[temp + 1, 5]))
                    {
                        singleRow.MaximoJobPlanTaskNumberAndDetails.Add(mergedData);
                    }

                    else
                    {
                        while (temp < rowCount && singleRow.AssetNumber == Convert.ToString(data[temp + 1, 5]))
                        {

                            if ((data[temp + 1, 26] ?? "").ToString() != "")
                            {
                                isEqNoPresent = true;
                            }//EqNumber
                            if (singleRow.MaximoJobPlanNumber[jpNumberCount] != Convert.ToString(data[temp + 1, 7]))
                            {


                                singleRow.MaximoJobPlanNumber.Add(Convert.ToString(data[temp + 1, 7]));  // thena add Jp number
                                singleRow.Interval.Add(Convert.ToString(data[i, 16])); //Frequency
                                singleRow.CounterType.Add(Convert.ToString(data[i, 17]));// Frequency Duration 
                                singleRow.LastDoneDate.Add(Convert.ToString(data[i, 18]));// LastDoneDte 
                                singleRow.LastDoneValue.Add(Convert.ToString(data[i, 19]));// At reading 
                                singleRow.MaximoPMDetails.Add(Convert.ToString(data[temp + 1, 2])); // PM description
                                singleRow.MaximoJobPlanTaskNumberAndDetails.Add(mergedData);
                                mergedData = "";
                                jpNumberCount++;

                            } //if next JpNumber not matched



                            jobTaskNumber = Convert.ToString(data[temp + 1, 8]);  //10-20 like 
                            jobTaskDesc = Convert.ToString(data[temp + 1, 9]);   // Job task Desc
                            mergedData = mergedData + "\n" + jobTaskNumber + "-" + jobTaskDesc;

                            temp++;
                            i++;
                        } // this loop is for Asset Number when next is same

                        if (mergedData != "")
                        {
                            singleRow.MaximoJobPlanTaskNumberAndDetails.Add(mergedData);

                        }
                    }




                    //----------------------Handling Reminder Window and Scheduling type ------------------------\\

                    for (int j = 0; j < singleRow.Interval.Count; j++)
                    {
                        string unit = singleRow.CounterType[j];


                        int interval;
                        string intervalStr = singleRow.Interval[j];


                        if (intervalStr.Contains(','))
                        {
                            intervalStr = intervalStr.Replace(",", "");
                            interval = Convert.ToInt32(intervalStr);
                        }
                        else
                        {
                            interval = Convert.ToInt32(intervalStr);
                        }

                        if (string.Equals(unit, "Weeks", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit.Add("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval[j]) && !string.IsNullOrWhiteSpace(singleRow.Interval[j]))
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

                            else
                            {
                                singleRow.Reminder.Add("");
                                singleRow.Window.Add("");
                                singleRow.SchedulingType.Add("");

                            }
                        }

                        else if (string.Equals(unit, "Months", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit.Add("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval[j]) && !string.IsNullOrWhiteSpace(singleRow.Interval[j]))
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

                            else
                            {
                                singleRow.Reminder.Add("");
                                singleRow.Window.Add("");
                                singleRow.SchedulingType.Add("");

                            }
                        }

                        else if (string.Equals(unit, "Years", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit.Add("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval[j]) && !string.IsNullOrWhiteSpace(singleRow.Interval[j]))
                            {
                                singleRow.Reminder.Add((Math.Round(0.07 * interval * 365)).ToString());
                                singleRow.Window.Add((Math.Round(0.1 * interval * 365)).ToString());

                                singleRow.SchedulingType.Add("Scheduled");
                            }

                            else
                            {
                                singleRow.Reminder.Add("");
                                singleRow.Window.Add("");
                                singleRow.SchedulingType.Add("");

                            }
                        }

                        else if (string.Equals(unit, "DAYS", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit.Add("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval[j]) && !string.IsNullOrWhiteSpace(singleRow.Interval[j]))
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

                            else
                            {
                                singleRow.Reminder.Add("");
                                singleRow.Window.Add("");
                                singleRow.SchedulingType.Add("");

                            }
                        }

                        else if (string.Equals(unit, "HR", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit.Add("Hours");
                            if (!string.IsNullOrEmpty(singleRow.Interval[j]) && !string.IsNullOrWhiteSpace(singleRow.Interval[j]))
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

                            else
                            {
                                singleRow.Reminder.Add("");
                                singleRow.Window.Add("");
                                singleRow.SchedulingType.Add("");

                            }

                        }

                        else
                        {
                            singleRow.ReminderWindowUnit.Add("");
                            singleRow.Reminder.Add("");
                            singleRow.Window.Add("");
                            singleRow.SchedulingType.Add("");

                        }
                        if (pmDescription.Contains("ENGR:"))
                        {
                            singleRow.ResponsibleDepartment.Add("Engine");
                        }
                        else if (pmDescription.Contains("DECK:"))
                        {

                            singleRow.ResponsibleDepartment.Add("Deck");

                        }
                        else if (pmDescription.Contains("ELEC:"))
                        {
                            singleRow.ResponsibleDepartment.Add("Electrical");

                        }
                        else
                        {
                            singleRow.ResponsibleDepartment.Add("");
                        }
                    }


                    if (!isEqNoPresent)
                    {

                        rows.Add(singleRow);
                    }
                }
            }
            return rows;

        }


    }

    internal class EquipmentNo
    {
        public string EqNumber { get; set; }
        public string Interval { get; set; } //  // Maps with Frequency
        public string CounterType { get; set; } //  // Maps with Frequency Duration
        public string LastDoneDate { get; set; } // Maps with Last Done Date
        public string LastDoneValue { get; set; } // Maps w
                                                  //
                                                  // ith At reading
        public string MaximoJobPlanNumber { get; set; }    // Maps With JP Number [insert rows] 
        public string MaximoJobPlanTaskNumberAndDetails { get; set; }    // merge all the Job task number and job task description containing same JP Number
        public string MaximoPMDetails { get; set; }
        public string ResponsibleDepartment { get; set; }
        public string Reminder { get; set; }
        public string Window { get; set; }
        public string SchedulingType { get; set; }
        public int rowNumber { get; set; } = 0;
        public string ReminderWindowUnit { get; set; }

        public EquipmentNo()
        {
            EqNumber = string.Empty;
            Interval = String.Empty;
            CounterType = String.Empty;
            LastDoneDate = String.Empty;
            LastDoneValue = String.Empty;
            MaximoJobPlanNumber = String.Empty;
            MaximoJobPlanTaskNumberAndDetails = String.Empty;
            MaximoPMDetails = String.Empty;
            ResponsibleDepartment = String.Empty;
            Reminder = String.Empty;
            Window = String.Empty;
            SchedulingType = String.Empty;
            ReminderWindowUnit = String.Empty;

        }

        public Dictionary<string, List<EquipmentNo>> ReadDataFromMaximoSheetEqNo(Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
            object[,] data = usedRange.Value;

            int rowCount = data.GetLength(0);

            //
            //
            //
            //--------------------Reading the data and storing it in a list of MaximoSheetData------------------------//
            Dictionary<string, List<EquipmentNo>> dict_EqNo = new Dictionary<string, List<EquipmentNo>>();
            List<EquipmentNo> rows = new List<EquipmentNo>();

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
                    string pmDescription = "";
                    bool isEqNoPresent = false;

                    EquipmentNo singleRow = new EquipmentNo();
                    singleRow.rowNumber = i;

                    singleRow.EqNumber = Convert.ToString(data[i, 26]);
                    singleRow.MaximoJobPlanNumber = (Convert.ToString(data[i, 7]));  //Jp number

                    singleRow.Interval = (Convert.ToString(data[i, 16])); //Frequency
                    singleRow.CounterType = (Convert.ToString(data[i, 17]));// Frequency Duration 
                    singleRow.LastDoneDate = (Convert.ToString(data[i, 18]));// LastDoneDte 
                    singleRow.LastDoneValue = (Convert.ToString(data[i, 19]));// At reading 
                    singleRow.MaximoPMDetails = (Convert.ToString(data[i, 2]));
                    pmDescription = Convert.ToString(data[i, 2]);// PM description

                    if ((data[i, 26] ?? "").ToString() != "")
                    {
                        isEqNoPresent = true;
                    }//EqNumber



                    jobTaskNumber = Convert.ToString(data[i, 8]);  //10-20 like 
                    jobTaskDesc = Convert.ToString(data[i, 9]);   // Job task Desc
                    mergedData = "\n" + jobTaskNumber + "-" + jobTaskDesc;


                    singleRow.MaximoJobPlanTaskNumberAndDetails = mergedData;





                    //----------------------Handling Reminder Window and Scheduling type ------------------------\\


                    {
                        string unit = singleRow.CounterType;


                        int interval;
                        string intervalStr = singleRow.Interval;


                        if (intervalStr.Contains(','))
                        {
                            intervalStr = intervalStr.Replace(",", "");
                            interval = Convert.ToInt32(intervalStr);
                        }
                        else
                        {
                            interval = Convert.ToInt32(intervalStr);
                        }

                        if (string.Equals(unit, "Weeks", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit = ("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval) && !string.IsNullOrWhiteSpace(singleRow.Interval))
                            {

                                singleRow.Reminder = ((Math.Round(0.07 * interval * 7)).ToString());
                                singleRow.Window = ((Math.Round(0.1 * interval * 7)).ToString());
                                if (interval <= 4)
                                {
                                    singleRow.SchedulingType = ("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType = ("Scheduled");
                                }
                            }

                            else
                            {
                                singleRow.Reminder = ("");
                                singleRow.Window = ("");
                                singleRow.SchedulingType = ("");

                            }
                        }

                        else if (string.Equals(unit, "Months", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit = ("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval) && !string.IsNullOrWhiteSpace(singleRow.Interval))
                            {
                                singleRow.Reminder = ((Math.Round(0.07 * interval * 30)).ToString());
                                singleRow.Window = ((Math.Round(0.1 * interval * 30)).ToString());
                                if (interval <= 1)
                                {
                                    singleRow.SchedulingType = ("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType = ("Scheduled");
                                }
                            }

                            else
                            {
                                singleRow.Reminder = ("");
                                singleRow.Window = ("");
                                singleRow.SchedulingType = ("");

                            }
                        }

                        else if (string.Equals(unit, "Years", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit = ("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval) && !string.IsNullOrWhiteSpace(singleRow.Interval))
                            {
                                singleRow.Reminder = ((Math.Round(0.07 * interval * 365)).ToString());
                                singleRow.Window = ((Math.Round(0.1 * interval * 365)).ToString());

                                singleRow.SchedulingType = ("Scheduled");
                            }

                            else
                            {
                                singleRow.Reminder = ("");
                                singleRow.Window = ("");
                                singleRow.SchedulingType = ("");

                            }
                        }

                        else if (string.Equals(unit, "DAYS", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit = ("Days");
                            if (!string.IsNullOrEmpty(singleRow.Interval) && !string.IsNullOrWhiteSpace(singleRow.Interval))
                            {
                                singleRow.Reminder = ((Math.Round(0.07 * interval)).ToString());
                                singleRow.Window = ((Math.Round(0.1 * interval)).ToString());
                                if (interval <= 30)
                                {
                                    singleRow.SchedulingType = ("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType = ("Scheduled");
                                }
                            }

                            else
                            {
                                singleRow.Reminder = ("");
                                singleRow.Window = ("");
                                singleRow.SchedulingType = ("");

                            }
                        }

                        else if (string.Equals(unit, "HR", StringComparison.OrdinalIgnoreCase))
                        {
                            singleRow.ReminderWindowUnit = ("Hours");
                            if (!string.IsNullOrEmpty(singleRow.Interval) && !string.IsNullOrWhiteSpace(singleRow.Interval))
                            {
                                singleRow.Reminder = ((Math.Round(0.07 * interval)).ToString());
                                singleRow.Window = ((Math.Round(0.1 * interval)).ToString());
                                if (interval <= 720)
                                {
                                    singleRow.SchedulingType = ("Fixed");
                                }
                                else
                                {
                                    singleRow.SchedulingType = ("Scheduled");
                                }
                            }

                            else
                            {
                                singleRow.Reminder = ("");
                                singleRow.Window = ("");
                                singleRow.SchedulingType = ("");

                            }

                        }

                        else
                        {
                            singleRow.ReminderWindowUnit = ("");
                            singleRow.Reminder = ("");
                            singleRow.Window = ("");
                            singleRow.SchedulingType = ("");

                        }
                        if (pmDescription.Contains("ENGR:"))
                        {
                            singleRow.ResponsibleDepartment = ("Engine");
                        }
                        else if (pmDescription.Contains("DECK:"))
                        {

                            singleRow.ResponsibleDepartment = ("Deck");

                        }
                        else if (pmDescription.Contains("ELEC:"))
                        {
                            singleRow.ResponsibleDepartment = ("Electrical");

                        }
                        else
                        {
                            singleRow.ResponsibleDepartment = ("");
                        }
                    }


                    if (isEqNoPresent)
                    {
                        List<EquipmentNo> list_eqNo;

                        if (dict_EqNo.ContainsKey(singleRow.EqNumber))
                        {

                            List<EquipmentNo> obj = dict_EqNo[singleRow.EqNumber];
                            bool foundMatchingJobPlan = false;

                            for (int i1 = obj.Count - 1; i1 >= 0; i1--)
                            {
                                if (obj[i1].MaximoJobPlanNumber == singleRow.MaximoJobPlanNumber)
                                {
                                    obj[i1].MaximoJobPlanTaskNumberAndDetails += "\n" + singleRow.MaximoJobPlanTaskNumberAndDetails;
                                    foundMatchingJobPlan = true; // Set flag indicating a match was found
                                    break; // Exit the loop since we found a match
                                }
                            }

                            if (!foundMatchingJobPlan)
                            {
                                obj.Add(singleRow); // Add singleRow only if no matching job plan was found
                            }

                        }

                        else
                        {
                            list_eqNo = new List<EquipmentNo>();
                            string eqNoKey = singleRow.EqNumber;
                            list_eqNo.Add(singleRow);
                            dict_EqNo.Add(eqNoKey, list_eqNo);
                        }

                    }
                }

            }
            return dict_EqNo;

        }

    }
}