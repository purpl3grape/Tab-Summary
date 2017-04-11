using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Application2 = System.Windows.Forms;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TabSummaryGenerator
{
    class Program
    {
        private static string _subject = "";
        public static string subject { get { return _subject; } set { _subject = value; } }

        public static string FilePath = "";
        public static string programMode = "1";
        public static string searchStringTotalsLabel = "";
        public static int TotalLabelRow = -999;
        public static int StartingCountsRow = -999;
        public static int EndCountsRow = -999;
        public static string[] WordsToOmit = null;
        public static string[] LabelsForTotal = null;
        public static string FileName = "5";
        public static string currentWeek = "1";
        public static string mReportDetailEmail = "";
        public static string mReportDetailLog = "";
        public static string mCodingDetailEmail = "";
        public static string mCodingDetailLog = "";
        public static string mSampleFileDetailEmail = "";
        public static string DialogLogDetails = "";
        public static Range foundRange = null;
        public static bool isCodingError = false;

        public static int[] missingAgentRow_Initial_List = new int[10000];
        public static int missingAgentRow_Initial_Index = 0;
        public static string[] InitialSurveyId_List = new string[10000];
        public static int InitialSurveyId_Index = 0;

        public static int[] missingAgentRow_Transferred_List = new int[10000];
        public static int missingAgentRow_Transferred_Index = 0;
        public static string[] TransferredSurveyId_List = new string[10000];
        public static int TransferredSurveyId_Index = 0;

        public static string SaveFilePathName = "";

        public static string qid = "";

        public static string[] questionList = new string[10000];
        public static int questionIndex = 0;
        public static string[] headerList = new string[10000];
        public static int headerIndex = 0;

        public static string[] filePathToAttach = new string[10];
        public static int fileID = 0;

        private static int _lastRow = 1;
        public static int lastRow { get { return _lastRow; } set { _lastRow = value; } }
        public static int lastCol = -999;

        private static string _SampleFilePath = "";
        public static string SampleFilePath { get { return _SampleFilePath; } set { _SampleFilePath = value; } }

        public static Hashtable QuestionIndexHashTable = new Hashtable();
        public static Hashtable QuestionTotalHashTable = new Hashtable();
        public static Hashtable QuestionCheckSumHashTable = new Hashtable();
        public static int integerValue = -999;
        public static string stringValue = "";


        delegate void MessageDelegate1(MailMessage mailMessage);
        public static void MessageDisplay(MailMessage mailMsg, bool showBody)
        {
            MessageDelegate1 md1 = new MessageDelegate1(ShowMessageSubject);
            md1 += ShowMessageBody;
            if (!showBody)
                md1 -= ShowMessageBody;
            md1(mailMsg);
        }

        /*  METHOD: SHOW MESSAGE SUBJECT CALLBACK   */
        public static void ShowMessageSubject(MailMessage msg)
        {
            System.Diagnostics.Debug.WriteLine("message subject: " + msg.Subject + "\n");
        }

        /*  METHOD: SHOW MESSAGE BODY CALLBACK   */
        public static void ShowMessageBody(MailMessage msg)
        {
            System.Diagnostics.Debug.WriteLine("message body: " + msg.Body + "\n");
        }

        /* METHOD: RETURNS LAST ROW NUMBER THAT HAS VALUES */
        public static int getLastRowNumberWithValues(Worksheet sheet, string ColumnRange)
        {
            Microsoft.Office.Interop.Excel.Range endRange = sheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Microsoft.Office.Interop.Excel.Range range = sheet.get_Range("A1", endRange);

            lastRow = endRange.Row;
            string cellVal = "";
            for (int i = 1; i < lastRow; i++)
            {
                try
                {
                    cellVal = sheet.Range[ColumnRange + i.ToString() + ":" + ColumnRange + i.ToString()].Value.ToString();
                }
                catch
                {
                    cellVal = "";
                }
                if (cellVal == "")
                {
                    return (i - 1);
                }
            }
            return lastRow;
        }

        /* METHOD: RETURNS LAST ROW NUMBER OF GIVEN EXCEL FILE */
        public static void getLastRowNumber(Worksheet sheet)
        {
            Microsoft.Office.Interop.Excel.Range endRange = sheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Microsoft.Office.Interop.Excel.Range range = sheet.get_Range("A1", endRange);

            lastRow = endRange.Row;
            //Store Range values
            int dictionaryID = 1;
            int startRow = 2;
            int currentRow = startRow;

            //STORE QIDS IN DICTIONARY
            while (currentRow <= lastRow)
            {
                currentRow++;
                dictionaryID++;
            }
        }

        public static void GetCellValue(Worksheet sh, string rng)
        {
            try { integerValue = (int)(sh.Range[rng].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(sh.Range[rng].Value2); integerValue = -999; }
        }

        public static int GetRow_FromColumn_ForString(Worksheet sh, string column)
        {
            foundRange = sh.Range[column + ":" + column].Find(searchStringTotalsLabel, Missing.Value, XlFindLookIn.xlValues, XlLookAt.xlWhole, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
            int foundRow = foundRange.Row;
            return foundRow;
        }

        public static int GetNewRowAfterTotals(Worksheet sh, string column)
        {
            foundRange = sh.Range[column + ":" + column].Find(searchStringTotalsLabel, Missing.Value, XlFindLookIn.xlValues, XlLookAt.xlWhole, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
            foundRange = sh.Range[column + (foundRange.Row) + ":" + column + (1000)].Find("*", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
            string foundRangeValue = foundRange.Value;
            int foundRow = foundRange.Row;

            for (int i = 0; i < WordsToOmit.Length; i++)
            {
                if (!foundRangeValue.Equals(WordsToOmit[i]))
                {
                    continue;
                }
                else
                {
                    foundRange = sh.Range[column + (foundRow - 1) + ":" + column + (1000)].Find("*", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
                    foundRange = sh.Range[column + (foundRow - 1) + ":" + column + (1000)].FindNext(foundRange);
                    foundRangeValue = foundRange.Value;
                    foundRow = foundRange.Row;
                    i = 0;  //Restart Loop
                }
            }
            return foundRow;
        }

        public static int GetNewRowAfterStubs(Worksheet sh, string column, string lastFoundString)
        {
            foundRange = sh.Range[column + foundRange.Row + ":" + column + (1000)].Find(lastFoundString, Missing.Value, XlFindLookIn.xlValues, XlLookAt.xlWhole, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
            foundRange = sh.Range[column + (foundRange.Row) + ":" + column + (1000)].Find("*", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
            string foundRangeValue = foundRange.Value;
            int foundRow = foundRange.Row;

            for (int i = 0; i < WordsToOmit.Length; i++)
            {
                if (!foundRangeValue.Equals(WordsToOmit[i]))
                {
                    continue;
                }
                else
                {
                    foundRange = sh.Range[column + (foundRow - 1) + ":" + column + (1000)].Find("*", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
                    foundRange = sh.Range[column + (foundRow - 1) + ":" + column + (1000)].FindNext(foundRange);
                    foundRangeValue = foundRange.Value;
                    foundRow = foundRange.Row;
                    i = 0;  //Restart Loop
                }
            }
            return foundRow;
        }


        public static void SummarizeTabFile(Application xlApp, string mySampleFilePath)
        {
            DateTimeOffset dateToday;
            string dateTodayString = "";
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("yyyyMMdd");

            SaveFilePathName = FilePath + @"\_Summary_" + FileName + " - " + dateTodayString + ".xlsx";
            System.IO.File.Delete(SaveFilePathName);

            Console.WriteLine("Analyzing Table Summary for:\t" + mySampleFilePath);
            int workSheetIndex = 0;

            //OPEN UP (NECESSARY FILES) AND OBTAIN LAST ROW NUMBER USED RESPECTIVELY
            SampleFilePath = mySampleFilePath;
            Microsoft.Office.Interop.Excel.Workbook DataFile_WB = xlApp.Workbooks.Open(SampleFilePath);
            Worksheet DataFile_SHEET = (Worksheet)DataFile_WB.Sheets[1];
            //getLastRowNumber(DataFile_SHEET);

            foreach (Worksheet currentSheet in DataFile_WB.Worksheets)
            {
                currentSheet.Activate();
                getLastRowNumber(currentSheet);
                Console.WriteLine("Summarizing sheet:\t" + currentSheet.Name);

                //Data list to omit from VLOOKUPS
                for (int i = 0; i < WordsToOmit.Length; i++)
                {
                    currentSheet.Range["AAB" + (i+1).ToString()].Value = WordsToOmit[i];
                }

                //currentSheet.Names.Add("OmmissionList", currentSheet.Range["AAB1:AAB" + WordsToOmit.Length]);


                for (int i = 1; i <= lastRow; i++)
                {
                    if(currentSheet.Range["A" + i].Value == searchStringTotalsLabel)
                    {
                        TotalLabelRow = i;
                    }
                    currentSheet.Range["AAC" + i].Formula = "=IFERROR(IF(VLOOKUP(A" + i + ",AAB:AAB,1,FALSE)<>\"\",\"\",B" + i + "),B" + i + ")";
                }

                currentSheet.Range["AAC1:AAC" + lastRow].Copy();
                currentSheet.Range["AAC1:AAC" + lastRow].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
                //1) (CALCULATING THE CHECKSUM) AND ADD IT TO HASHTABLE OF KEY: SHEETNAME
                //2) STORE TOTALS IN HASHTABLE
                //3) KEY IS THE NAME OF THE SHEET


                // try { TotalLabelRow = GetRow_FromColumn_ForString(currentSheet, "A"); }
                // catch
                // {
                //     Console.WriteLine("\r\n\r\n\tAPPOLOGIES\r\n\tThe Tab Checker feels really really sad because\r\n\tit was unable to help you find the 'TOTALS' label:\t" + searchStringTotalsLabel + ".\r\n\tin the Tables File.\r\n\r\nTO CONTINUE PRESS ENTER\r\nTO QUIT ENTER [0]\r\n\r\n\r\n");
                //     programMode = Console.ReadLine();
                //     if (programMode.Equals("0"))
                //     {
                //         System.Environment.Exit(1);
                //     }
                //     else
                //     {
                //         xlApp.Quit();
                //         RunAll();
                //     }
                // }
                // try { StartingCountsRow = GetNewRowAfterTotals(currentSheet, "A"); }
                // catch
                // {
                //     Console.WriteLine("\r\n\r\n\tAPPOLOGIES\r\n\tThe Tab Checker feels really really sad because\r\n\tit was unable to help you calculate the 'CHECKSUM' for label:\t" + searchStringTotalsLabel + ".\r\n\tin the Tables File.\r\n\r\nTO CONTINUE PRESS ENTER\r\nTO QUIT ENTER [0]\r\n\r\n\r\n");
                //     programMode = Console.ReadLine();
                //     if (programMode.Equals("0"))
                //     {
                //         System.Environment.Exit(1);
                //     }
                //     else
                //     {
                //         xlApp.Quit();
                //         RunAll();
                //     }
                // }



                //currentSheet.Range["ZZ1"].Formula = "=ROUND(SUM(AAC" + StartingCountsRow + ":AAC1000),0)";
                currentSheet.Range["ZZ1"].Formula = "=ROUND(SUM(AAC1:AAC" + lastRow + "),0)";
                currentSheet.Range["ZZ1"].Copy();
                currentSheet.Range["ZZ1"].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
                //string sss = Console.ReadLine();

                currentSheet.Range["ZZ2"].Formula = "=ROUND(B" + TotalLabelRow + ",0)";
                currentSheet.Range["ZZ2"].Copy();
                currentSheet.Range["ZZ2"].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

                currentSheet.Range["A1"].Select();

                try { QuestionTotalHashTable.Add(currentSheet.Name, currentSheet.Range["ZZ2"].Value); } catch { Console.WriteLine("Failed To Hash Totals at Sheet: " + currentSheet.Name); }
                try { QuestionCheckSumHashTable.Add(currentSheet.Name, currentSheet.Range["ZZ1"].Value); } catch { Console.WriteLine("Failed To Hash CheckSum at Sheet: " + currentSheet.Name); }
                try { QuestionIndexHashTable.Add(currentSheet.Name, workSheetIndex); } catch { Console.WriteLine("Failed To Hash Worksheet Index at Sheet: " + workSheetIndex); }
                //currentSheet.Range["ZZ:ZZ"].Delete();
                currentSheet.Range["ZZ:AAC"].Delete();

                //Hyperlinks
                currentSheet.Hyperlinks.Add(currentSheet.Range["B1"], string.Empty, "'SUMMARY'!A1", "Back to summary.", "Return to summary");

                workSheetIndex++;
            }

            Console.WriteLine("\r\nPreparing Summary Page");


            //ADD SUMMARY PAGE AND DISPLAY DETAILS
            Worksheet TableSummarySheet = xlApp.Worksheets.Add();
            TableSummarySheet.Name = "SUMMARY";
            TableSummarySheet.Activate();
            TableSummarySheet.Range["A1"].Value = "INDEX";
            TableSummarySheet.Range["B1"].Value = "QUESTION";
            TableSummarySheet.Range["C1"].Value = "TOTAL";
            TableSummarySheet.Range["D1"].Value = "CHECK SUM";
            TableSummarySheet.Range["E1"].Value = "REMARKS";
            TableSummarySheet.Range["F1"].Value = "LINK TO SHEET";

            System.Drawing.Color highlightColor = System.Drawing.Color.FromArgb(35, 45, 55);
            System.Drawing.Color textColor = System.Drawing.Color.FromArgb(100, 255, 255);
            int SummaryIndex = 2;
            ICollection QuestionKeyCollection = QuestionCheckSumHashTable.Keys;
            Worksheet currentworksheet = null;
            foreach (string key in QuestionKeyCollection)
            {
                TableSummarySheet.Range["A" + SummaryIndex].Value = QuestionIndexHashTable[key];
                TableSummarySheet.Range["B" + SummaryIndex].Value = key;
                TableSummarySheet.Range["C" + SummaryIndex].Value = QuestionTotalHashTable[key];
                TableSummarySheet.Range["D" + SummaryIndex].Value = QuestionCheckSumHashTable[key];
                TableSummarySheet.Range["E" + SummaryIndex].Formula = "=IF(C" + SummaryIndex + "=D" + SummaryIndex + ",\"GOOD\",IF(C" + SummaryIndex + "<D" + SummaryIndex + ",\"MULTI-CHOICE\",\"ERROR\"))";
                TableSummarySheet.Range["E" + SummaryIndex].Copy();
                TableSummarySheet.Range["E" + SummaryIndex].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

                textColor = System.Drawing.Color.FromArgb(0, 0, 0);
                if (TableSummarySheet.Range["E" + SummaryIndex].Value == "GOOD") { highlightColor = System.Drawing.Color.FromArgb(150, 250, 50); }
                else if (TableSummarySheet.Range["E" + SummaryIndex].Value == "MULTI-CHOICE") { highlightColor = System.Drawing.Color.FromArgb(100, 200, 100); }
                else { highlightColor = System.Drawing.Color.FromArgb(250, 50, 150); }
                TableSummarySheet.Range["E" + SummaryIndex].Interior.Color = highlightColor;
                TableSummarySheet.Range["E" + SummaryIndex].Font.Color = textColor;

                //Hyperlinks
                TableSummarySheet.Hyperlinks.Add(TableSummarySheet.Range["F" + SummaryIndex], string.Empty, "'" + key + "'!A1", "Click here to view sheet: " + key, "Sheet: " + key);
                SummaryIndex++;
            }

            Console.WriteLine("\r\nFinalizing");

            TableSummarySheet.Range["A1:F1"].Font.Bold = true;
            TableSummarySheet.Range["A1:F1"].Font.Size = "10";
            TableSummarySheet.Range["A1:F1"].ColumnWidth = 20f;
            TableSummarySheet.Range["A1:F1"].Font.Name = "Verdana";
            TableSummarySheet.Range["A1:F1"].RowHeight = 30f;
            TableSummarySheet.Range["A1:F1"].WrapText = true;
            TableSummarySheet.Range["A1:F1"].AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            TableSummarySheet.Range["A1:F1"].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            TableSummarySheet.Range["A1:F1"].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            highlightColor = System.Drawing.Color.FromArgb(35, 45, 55);
            textColor = System.Drawing.Color.FromArgb(100, 255, 255);
            TableSummarySheet.Range["A1:F1"].Interior.Color = highlightColor;
            TableSummarySheet.Range["A1:F1"].Font.Color = textColor;
            TableSummarySheet.Range["B2"].Activate();
            TableSummarySheet.Range["B2"].Select();
            TableSummarySheet.Range["B2"].Application.ActiveWindow.FreezePanes = true;
            TableSummarySheet.Move(DataFile_WB.Sheets[1]);
            getLastRowNumber(TableSummarySheet);
            highlightColor = System.Drawing.Color.FromArgb(50, 200, 200);
            textColor = System.Drawing.Color.FromArgb(250, 50, 150);
            TableSummarySheet.Range["A2:D" + lastRow].Interior.Color = highlightColor;
            TableSummarySheet.Range["A2:D" + lastRow].Font.Color = textColor;
            TableSummarySheet.Range["F2:F" + lastRow].Interior.Color = highlightColor;
            TableSummarySheet.Range["F2:F" + lastRow].Font.Color = textColor;
            TableSummarySheet.Range["A1:F" + lastRow].Sort(TableSummarySheet.Range["A1:F" + lastRow], XlSortOrder.xlAscending, Type.Missing, Type.Missing, XlSortOrder.xlAscending, Type.Missing, XlSortOrder.xlAscending, XlYesNoGuess.xlYes, Type.Missing, Type.Missing, XlSortOrientation.xlSortColumns, XlSortMethod.xlPinYin, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal);


            Console.WriteLine("\r\nSaving");
            DataFile_WB.SaveAs(SaveFilePathName, XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);

            //Close all excel workbooks
            DataFile_WB.Close(false);

        }

        /* MAIN */
        [STAThreadAttribute]
        static void Main(string[] args)
        {
            RunAll();
        }

        /* RUN ALL */
        //[STAThreadAttribute]
        public static void RunAll()
        {
            //INITIALIZE THE (EXCEL APP) BY DECLARING IT ONCE AT THE START OF MAIN
            Application xlApp = new Application();
            xlApp.DisplayAlerts = false;

            DateTimeOffset dateToday;
            string dateTodayString = "";
            //TIMESTAMP PER RESPONDENT
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("yyyyMMdd");

            DateTimeOffset dateWeekend;
            string dateWeekendString = "";
            //TIMESTAMP PER RESPONDENT
            dateWeekend = DateTimeOffset.Now.AddDays(-2);
            dateWeekendString = dateWeekend.ToString("yyyyMMdd");

            Stream myStream = null;
            Application2.OpenFileDialog theDialog = new Application2.OpenFileDialog();

            //GET LIST OF WORDS (LABELS) TO OMIT WHEN DOING THE CHECKSUMS
            try { WordsToOmit = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\OmissionList.log"); }
            catch
            {
                System.IO.StreamWriter ommisionLogFile = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\OmissionList.log");
                ommisionLogFile.Close();
                WordsToOmit = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\OmissionList.log");
            }

            //GET LIST OF TOTAL TYPE (LABELS) TO OMIT WHEN DOING THE CHECKSUMS
            try { LabelsForTotal = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\TotalsList.log"); }
            catch
            {
                System.IO.StreamWriter totalsLabelLogFile = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\TotalsList.log");
                totalsLabelLogFile.Close();
                LabelsForTotal = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\TotalsList.log");
            }



            string[] logSampleFileContent;
            try
            {
                logSampleFileContent = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\DialogLog.log");
            }
            catch
            {
                System.IO.StreamWriter newLogFile = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\DialogLog.log");
                newLogFile.Close();
                logSampleFileContent = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\DialogLog.log");
            }

            foreach (string line in logSampleFileContent)
            {
                DialogLogDetails += line;
            }
            if (DialogLogDetails.Equals(""))
            {
                theDialog.InitialDirectory = @"G:\";
            }
            else
            {
                theDialog.InitialDirectory = DialogLogDetails;
            }

            System.IO.StreamWriter logFile = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\DialogLog.log");
            logFile.WriteLine((DialogLogDetails));
            logFile.Close();

            theDialog.Title = "Open Excel File";
            theDialog.Filter = "Excel files|*.xlsx";
            if (theDialog.ShowDialog() == Application2.DialogResult.OK)
            {
                try
                {
                    if ((myStream = theDialog.OpenFile()) != null)
                    {
                        FilePath = theDialog.FileName;
                        FileName = theDialog.SafeFileName;
                        FilePath = FilePath.Replace(FileName, "");
                    }
                }
                catch { Console.WriteLine("Error: Could not read file from disk. Original error: "); }
            }
            Console.WriteLine("File Path Selected:\t" + FilePath);
            Console.WriteLine("File Name Selected:\t" + FileName);

            logFile = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\TabSummaryGenerator\TabSummaryGenerator\DialogLog.log");
            logFile.WriteLine(FilePath);
            logFile.Close();


            Console.Write("\r\n\r\nTAB FILE SUMMARY\r\nLOAD LABELS TO CHECK ENTER [1] (Make sure that Totals Label is on Top).\r\n\r\nTO END PROGRAM ENTER [0].\r\n\r\nEnter option:");
            searchStringTotalsLabel = Console.ReadLine();
            if (searchStringTotalsLabel.Equals("0"))
            {
                System.Environment.Exit(1);
            }
            else
            {
                Console.Write("Validating Counts For:\t" + searchStringTotalsLabel);
                try { SummarizeTabFile(xlApp, FilePath + @"\" + FileName); } catch { Console.WriteLine("Table Summary Program Failed"); System.Diagnostics.Debug.WriteLine("Table Summary Program Failed"); }
            }

            System.Environment.Exit(1);

        }


    }
}