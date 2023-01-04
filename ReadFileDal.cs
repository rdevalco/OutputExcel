using System;
using System.Text;
using System.IO;
using System.Collections;
using System.Threading;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace OutputExcel
{
    /// <summary>
    /// Summary description for ReadFileDal.
    /// </summary>
    public class ReadFileDal
    {

        #region Member Variables

        private Excel.Application excelApp;
        private Excel.Workbook excelWorkbook;
        private Excel.Sheets excelSheets;
        private Excel.Worksheet excelWorksheet;

        private static object mmissing;

        private static object mvisible;

        private bool mapp_visible;

        private object mfilename;

        private object mupdate_links;
        private object mread_only;
        private object mformat;
        private object mpassword;
        private object mwrite_res_password;
        private object mignore_read_only_recommend;
        private object morigin;
        private object mdelimiter;
        private object meditable;
        private object mnotify;
        private object mconverter;
        private object madd_to_mru;
        private object mlocal;
        private object mcorrupt_load;

        private object msave_changes;
        private object mroute_workbook;

        private ArrayList m_al;

        #endregion

        #region Constructors

        public ReadFileDal()
        {
            Initialize();
            this.startExcel();
        }

        public ReadFileDal(bool visible)
        {
            Initialize();
            this.mapp_visible = visible;
            this.startExcel();
        }

        ~ReadFileDal()
        {
            
            this.stopExcel();
        }

        #endregion

        #region Private Methods

        private void Initialize()
        {

            excelApp = null;
            excelWorkbook = null;
            excelSheets = null;
            excelWorksheet = null;

            mmissing = System.Reflection.Missing.Value;

            mvisible = true;

            mapp_visible = false;

            mupdate_links = 0;
            mread_only = true;
            mformat = 1;
            mpassword = mmissing;
            mwrite_res_password = mmissing;
            mignore_read_only_recommend = true;
            morigin = mmissing;
            mdelimiter = mmissing;
            meditable = false;
            mnotify = false;
            mconverter = mmissing;
            madd_to_mru = false;
            mlocal = false;
            mcorrupt_load = false;

            msave_changes = false;
            mroute_workbook = false;

        }

        private void startExcel()
        {
            if( this.excelApp == null )
            {
                this.excelApp = new Excel.ApplicationClass();
            }

            // Make Excel Visible
            this.excelApp.Visible = this.mapp_visible;
        }

        private string[] ConvertToStringArray(System.Array values)
        {
            string[] newArray = new string[values.Length];
            int i = 0;
            int j = 0;
            int index = 0;
            for ( i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++ )
            {
                for ( j = values.GetLowerBound(1); j <= values.GetUpperBound(1); j++ )
                {
                    if(values.GetValue(i,j)==null)
                    {
                        newArray[index]="";
                    }
                    else
                    {
                        newArray[index]=(string)values.GetValue(i,j).ToString();
                    }
                    index++;
                }
            }
            return newArray;
        }
        
        #endregion

        public void stopExcel()
        {
            if( this.excelApp != null )
            {
                try
                {
                    this.excelApp.Quit();
                    Process[] pProcess;
                    pProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL.exe");
                    Debug.WriteLine("Process Count : " + pProcess.Length.ToString());
                    foreach (System.Diagnostics.Process p in pProcess)
                    {
                        p.Kill();
                    }
                }
                catch
                {
                    // do nothing
                }
            }
        }

        public string OpenFile(string fileName, string password)
        {
            mfilename = fileName;

            if( password.Length > 0 )
            {
                mpassword = password;
            }

            try
            {
                // Open a workbook in Excel
                this.excelWorkbook = this.excelApp.Workbooks.Open(fileName,
                                                                  mupdate_links,
                                                                  mread_only,
                                                                  mformat,
                                                                  mpassword,
                                                                  mwrite_res_password,
                                                                  mignore_read_only_recommend,
                                                                  morigin,
                                                                  mdelimiter,
                                                                  meditable,
                                                                  mnotify,
                                                                  mconverter,
                                                                  madd_to_mru,
                                                                  mlocal,
                                                                  mcorrupt_load);
            }
            catch(Exception e)
            {
                this.CloseFile();
                return e.Message;
            }
            return "OK";
        }

        public void CloseFile()
        {
            excelWorkbook.Close( msave_changes, mfilename, mroute_workbook );
        }

        public void GetExcelSheets()
        {
            if( this.excelWorkbook != null )
            {
                excelSheets = excelWorkbook.Worksheets;
            }
        }

        public bool FindExcelWorksheet(string worksheetName)
        {
            bool ATP_SHEET_FOUND = false;

            if( this.excelSheets != null )
            {
                for( int i=1; i<=this.excelSheets.Count; i++ )
                {
                    this.excelWorksheet = (Excel.Worksheet)excelSheets.get_Item((object)i);
                    if( this.excelWorksheet.Name.Equals(worksheetName) )
                    {
                        this.excelWorksheet.Activate();
                        ATP_SHEET_FOUND = true;
                        return ATP_SHEET_FOUND;
                    }
                }
            }
            return ATP_SHEET_FOUND;
        }

        public string[] GetRange(string range)
        {
            Excel.Range workingRangeCells = excelWorksheet.get_Range(range,Type.Missing);
            System.Array array = (System.Array)workingRangeCells.Cells.Value2;
            string[] arrayS = this.ConvertToStringArray(array);

            return arrayS;
        }

        public object GetRangeObj(string range)
        {
            Excel.Range workingRangeCells = excelWorksheet.get_Range(range,Type.Missing);
            object obj = workingRangeCells.Cells.Value2;
            return obj;
        }

        private void SplitUpText(string strText, int nMaxCount)
        {

            int nLength = 0;
            int nWhereAt = 0;

            char nChar;
            char[] characters = strText.ToCharArray();

            string strCurrentLine = "";
            string strTrimmedText = "";
            string strCurrentWord = "";

            int nWordLength = 0;

            m_al.Clear();

            nLength = characters.Length;

            while (nWhereAt < nLength)
            {

                nChar = characters[nWhereAt];

                if ((nChar == 32) || (nChar == '\r'))
                {

                    strTrimmedText = strCurrentWord.Replace("\r\n",".").Trim();
                    nWordLength = strTrimmedText.Length;

                    // add to string here

                    if ((strCurrentLine.Length + nWordLength + 1) > nMaxCount)
                    {

                        m_al.Add(strCurrentLine);
                        strCurrentLine = strTrimmedText + " ";

                    } // if ((strCurrentLine.Length + nWordLength + 1) > nMaxCount)
                    else
                    {

                        strCurrentLine = strCurrentLine + strTrimmedText + " ";

                    } // else (strCurrentLine.Length + nWordLength + 1) < nMaxCount

                    strCurrentWord = "";

                } // if ((nChar == 32) || (nChar == '\r'))
                else
                {
                    

                    //if ((nChar != 10) && (nChar != 13))
                    //{

                        strCurrentWord = strCurrentWord + nChar;

                    //} // if ((nChar != 10) && (nChar != 13))

                } // else nChar != 32 && nChar == '\r'

                nWhereAt++;

            } // while (nWhereAt < nLength)
            if (string.IsNullOrEmpty(strCurrentWord) == false)
            {
                strTrimmedText = strCurrentWord.Replace("\r\n", ".").Trim();
                nWordLength = strTrimmedText.Length;

                // add to string here

                if ((strCurrentLine.Length + nWordLength + 1) > nMaxCount)
                {

                    m_al.Add(strCurrentLine);
                    strCurrentLine = strTrimmedText + " ";

                } // if ((strCurrentLine.Length + nWordLength + 1) > nMaxCount)
                else
                {

                    strCurrentLine = strCurrentLine + strTrimmedText + " ";

                } // else (strCurrentLine.Length + nWordLength + 1) < nMaxCount
            }

            if (strCurrentLine.Length > 0)
            {

                strTrimmedText = strCurrentLine.Trim();
                m_al.Add(strTrimmedText);

            } // if (strCurrentLine.Length > 0)

        } // private void SplitUpText(string strText,int nMaxCount)

        public string GetSpaces(int spaceCount)
        {

            string message = "";

            if (spaceCount > 0)
            {
                message = message.PadRight(spaceCount, ' ');
            }

            return message;

        }

        public void ReadExcel(string outputPath, string fileName, string startCell, string stopCell)
        {
            
            string formattedOutputFile = "";
            DateTime dateTime = DateTime.Now;
            StreamWriter streamWriter = null;
            string messageText = "";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook;
            Excel._Worksheet xlWorksheet;
            Excel.Range range;
            System.Array myvalues;
            string[] strArray;
            
            int count = 0;

            char[] specialCharacters = null;
            ArrayList alInformation = null;
            string lastStatus = "";
            int spaceCount = 0;
            try
            {

                specialCharacters = new char[3];

                specialCharacters[0] = (char)0x0C;
                specialCharacters[1] = (char)0x0D;
                specialCharacters[2] = (char)0x0A;

                formattedOutputFile = outputPath;

                if (formattedOutputFile.EndsWith("\\") == false)
                {
                    formattedOutputFile += "\\";
                }

                formattedOutputFile += string.Format("{0}_{1}.txt","bugs",dateTime.ToString("MM_dd_yyyy"));
                streamWriter = new StreamWriter(formattedOutputFile, false);

                streamWriter.Write(specialCharacters[0]);
                streamWriter.Write(specialCharacters[1]);
                streamWriter.Write(specialCharacters[2]);

                m_al = new ArrayList();

                xlWorkbook = xlApp.Workbooks.Open(fileName);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];


                alInformation = new ArrayList();
                lastStatus = "";

                range = xlWorksheet.get_Range(startCell, stopCell);
                myvalues = (System.Array)range.Cells.Value;
                strArray = ConvertToStringArray(myvalues);

                count = 1;
                foreach (string s in strArray)
                {
                    switch (count)
                    {
                        case 1:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 2:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 3:

                            messageText = "Id";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);

                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 4:
                            messageText = "Issue Created On";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 5:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 6:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 7:
                            messageText = "Priority";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 8:
                            messageText = "Issue Type";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 9:
                            messageText = "Category";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 10:
                            messageText = "Functionality";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 11:
                            messageText = "Issue Title";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 12:
                            messageText = "Detailed Description of Problem";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 13:
                            messageText = "User Level";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 14:
                            messageText = "Steps To Reproduce";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;

                        case 15:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 16:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 17:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 18:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 19:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 20:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 21:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 22:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 23:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 24:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 25:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 26:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 27:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 28:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;

                        case 29:
                            messageText = "Issue Status";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            lastStatus = s;
                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }
                        break;
                        case 30:
                            messageText = "Testing Notes/Comments";
                            //streamWriter.WriteLine(messageText);
                            alInformation.Add(messageText);
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);

                            messageText = GetSpaces(4);
                            foreach (string splitUpString in m_al)
                            {
                                //streamWriter.WriteLine("    " + splitUpString);
                                alInformation.Add(messageText + splitUpString);
                            }

                            if (lastStatus.Contains("Deployed to Test") == true)
                            {
                                spaceCount = 16;
                            }
                            else
                            {
                                if (lastStatus.Contains("Ready for QA") == true)
                                {
                                    spaceCount = 20;
                                }
                                else
                                {

                                    if (lastStatus.Contains("Deployed to QA") == true)
                                    {
                                        spaceCount = 24;
                                    }
                                    else
                                    {

                                        if (lastStatus.Contains("Ready for Build") == true)
                                        {
                                            spaceCount = 28;
                                        }
                                        else
                                        {
                                            if (lastStatus.Contains("Return for Review") == true)
                                                spaceCount = 4;
                                            else
                                                spaceCount = 8;
                                        }
                                    }
                                }
                            }

                            messageText = GetSpaces(spaceCount);
                            foreach (string s2 in alInformation)
                            {
                                
                                streamWriter.WriteLine(messageText + s2);
                            }

                            streamWriter.WriteLine();
                            streamWriter.Write(specialCharacters[0]);
                            streamWriter.Write(specialCharacters[1]);
                            streamWriter.Write(specialCharacters[2]);
                            alInformation = new ArrayList();

                        break;
                        case 31:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                        break;
                        case 32:
                            count = 0;

                        break;
                    }

                    count++;

                }

                if (lastStatus.Contains("Deployed to Test") == true)
                {
                    spaceCount = 16;
                }
                else
                {
                    if (lastStatus.Contains("Ready for QA") == true)
                    {
                        spaceCount = 20;
                    }
                    else
                    {

                        if (lastStatus.Contains("Deployed to QA") == true)
                        {
                            spaceCount = 24;
                        }
                        else
                        {

                            if (lastStatus.Contains("Ready for Build") == true)
                            {
                                spaceCount = 28;
                            }
                            else
                            {
                                if (lastStatus.Contains("Return for Review") == true)
                                    spaceCount = 4;
                                else
                                    spaceCount = 8;
                            }
                        }
                    }
                }

                messageText = GetSpaces(spaceCount);
                foreach (string s2 in alInformation)
                {

                    streamWriter.WriteLine(messageText + s2);
                }

                streamWriter.WriteLine();
                streamWriter.Write(specialCharacters[0]);
                streamWriter.Write(specialCharacters[1]);
                streamWriter.Write(specialCharacters[2]);
                alInformation = new ArrayList();


            }
            catch
            {
                throw;
            }
            finally
            {
                xlApp.Workbooks.Close();
                xlApp.Quit();

                if (streamWriter != null)
                {
                    streamWriter.Flush();
                    streamWriter.Close();
                    streamWriter = null;
                }

            }

        }

        public void ReadExcelForCenterStudentUpdate(string outputPath, string fileName, string startCell, string stopCell)
        {

            string formattedOutputFile = "";
            DateTime dateTime = DateTime.Now;
            StreamWriter streamWriter = null;
            string messageText = "";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook;
            Excel._Worksheet xlWorksheet;
            Excel.Range range;
            System.Array myvalues;
            string[] strArray;

            int count = 0;

            string oesid = "";
            string EnrollmentCenterID = "";
            string CalculatedSessionID = "";
            string EnrollmentID = "";
            string Id = "";
            string SessionID = "";
            string CenterID = "";
            string SessionType = "";

            try
            {

                formattedOutputFile = outputPath;

                if (formattedOutputFile.EndsWith("\\") == false)
                {
                    formattedOutputFile += "\\";
                }

                formattedOutputFile += "Run.sql";
                streamWriter = new StreamWriter(formattedOutputFile, false);

                m_al = new ArrayList();

                xlWorkbook = xlApp.Workbooks.Open(fileName);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];


                range = xlWorksheet.get_Range(startCell, stopCell);
                myvalues = (System.Array)range.Cells.Value;
                strArray = ConvertToStringArray(myvalues);

                count = 1;
                foreach (string s in strArray)
                {
                    switch (count)
                    {
                        case 1:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            oesid = m_al[0].ToString();
                        break;
                        case 2:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            EnrollmentCenterID = m_al[0].ToString();
                        break;
                        case 3:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            CalculatedSessionID = m_al[0].ToString();
                        break;
                        case 4:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            EnrollmentID = m_al[0].ToString();
                        break;
                        case 5:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            Id = m_al[0].ToString();
                        break;
                        case 6:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            SessionID = m_al[0].ToString();
                        break;
                        case 7:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            CenterID = m_al[0].ToString();
                        break;
                        case 8:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            SessionType = m_al[0].ToString();

                            messageText = "declare @SessionID int;";
                            streamWriter.WriteLine(messageText);
                            messageText = "declare @SessionType int;";
                            streamWriter.WriteLine(messageText);
                            messageText = "declare @Id uniqueidentifier;";
                            streamWriter.WriteLine(messageText);
                            messageText = "declare @Who nvarchar(50);";
                            streamWriter.WriteLine(messageText);
                            messageText = "declare @SavedWhen datetime;";
                            streamWriter.WriteLine(messageText);

                            streamWriter.WriteLine("");

                            messageText = "set @Who = SUSER_SNAME();";
                            streamWriter.WriteLine(messageText);
                            messageText = "set @SavedWhen = getdate();";
                            streamWriter.WriteLine(messageText);
                            streamWriter.WriteLine("");
 
                            messageText = string.Format("set @Id = '{0}';",Id);
                            streamWriter.WriteLine(messageText);
                            messageText = string.Format("set @SessionID = {0};",CalculatedSessionID);
                            streamWriter.WriteLine(messageText);
                            messageText = string.Format("set @SessionType = {0};", SessionType);
                            streamWriter.WriteLine(messageText);
                            streamWriter.WriteLine("");
 
                            messageText = "exec [dbo].[UpdateCenterStudentSessionID] @Id,@SessionID,@Who,@SavedWhen;";
                            streamWriter.WriteLine(messageText);
                            streamWriter.WriteLine("");
 
                            messageText = "go";
                            streamWriter.WriteLine(messageText);
                            streamWriter.WriteLine("");

                            count = 0;

                        break;
                    }

                    count++;

                }

            }
            catch
            {
                throw;
            }
            finally
            {
                xlApp.Workbooks.Close();
                xlApp.Quit();

                if (streamWriter != null)
                {
                    streamWriter.Flush();
                    streamWriter.Close();
                    streamWriter = null;
                }

            }

        }

        public void ReadExcelForSessionStudentsUpdate(string outputPath, string fileName, string startCell, string stopCell)
        {

            string formattedOutputFile = "";
            DateTime dateTime = DateTime.Now;
            StreamWriter streamWriter = null;
            string messageText = "";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook;
            Excel._Worksheet xlWorksheet;
            Excel.Range range;
            System.Array myvalues;
            string[] strArray;

            int count = 0;
            string LocalId = "";
            string EnrollmentID = "";
            string EnrollmentCenterID = "";
            string CalculatedSessionID = "";
            string SessionStudentID = "";
            string OESID = "";
            string SessionID = "";
            string CenterID = "";
            string SessionTypeDescription = "";
            string CenterName = "";
            string EntryDate = "";
            string ExitDate = "";
            string SessionStartDate = "";
            string SessionEndDate = "";
            string IsActive = "";
            ArrayList updatedInformation = null;

            try
            {

                updatedInformation = new ArrayList();

                formattedOutputFile = outputPath;

                if (formattedOutputFile.EndsWith("\\") == false)
                {
                    formattedOutputFile += "\\";
                }

                formattedOutputFile += "Run.sql";
                streamWriter = new StreamWriter(formattedOutputFile, false);

                m_al = new ArrayList();

                xlWorkbook = xlApp.Workbooks.Open(fileName);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];


                range = xlWorksheet.get_Range(startCell, stopCell);
                myvalues = (System.Array)range.Cells.Value;
                strArray = ConvertToStringArray(myvalues);

                count = 1;
                foreach (string s in strArray)
                {
                    switch (count)
                    {
                        case 1:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            LocalId = m_al[0].ToString();
                            break;
                        case 2:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            EnrollmentID = m_al[0].ToString();
                            break;
                        case 3:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            EnrollmentCenterID = m_al[0].ToString();
                        break;
                        case 4:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            CalculatedSessionID = m_al[0].ToString();
                            break;
                        case 5:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            SessionStudentID = m_al[0].ToString();
                            break;
                        case 6:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            OESID = m_al[0].ToString();
                            break;
                        case 7:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            SessionID = m_al[0].ToString();
                            break;
                        case 8:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            CenterID = m_al[0].ToString();
                            break;
                        case 9:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            SessionTypeDescription = m_al[0].ToString();
                        break;
                        case 10:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            CenterName = m_al[0].ToString();
                        break;
                        case 11:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            EntryDate = m_al[0].ToString();
                        break;
                        case 12:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            ExitDate = m_al[0].ToString();
                        break;
                        case 13:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            SessionStartDate = m_al[0].ToString();
                        break;
                        case 14:
                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            SessionEndDate = m_al[0].ToString();
                        break;
                        case 15:

                            SplitUpText(s.Replace("\n\n", "\r\n"), 80);
                            IsActive = m_al[0].ToString();

                            if (updatedInformation.Contains(SessionStudentID) == false)
                            {

                                messageText = "declare @SessionID int;";
                                streamWriter.WriteLine(messageText);
                                messageText = "declare @SessionStudentID int;";
                                streamWriter.WriteLine(messageText);
                                messageText = "declare @Who nvarchar(50);";
                                streamWriter.WriteLine(messageText);
                                messageText = "declare @SavedWhen datetime;";
                                streamWriter.WriteLine(messageText);

                                streamWriter.WriteLine("");

                                messageText = "set @Who = SUSER_SNAME();";
                                streamWriter.WriteLine(messageText);
                                messageText = "set @SavedWhen = getdate();";
                                streamWriter.WriteLine(messageText);
                                streamWriter.WriteLine("");

                                messageText = string.Format("set @SessionStudentID = {0};", SessionStudentID);
                                streamWriter.WriteLine(messageText);
                                messageText = string.Format("set @SessionID = {0};", CalculatedSessionID);
                                streamWriter.WriteLine(messageText);
                                streamWriter.WriteLine("");

                                messageText = "begin tran _tempTran;";
                                streamWriter.WriteLine(messageText);
                                streamWriter.WriteLine("");
                                messageText = "update";
                                streamWriter.WriteLine(messageText);
                                messageText = "    [dbo].[SessionStudents]";
                                streamWriter.WriteLine(messageText);
                                messageText = "set";
                                streamWriter.WriteLine(messageText);
                                messageText = "    SessionID = @SessionID,";
                                streamWriter.WriteLine(messageText);
                                messageText = "    UpdatedByUser = @Who,";
                                streamWriter.WriteLine(messageText);
                                messageText = "    UpdatedByTime = @SavedWhen";
                                streamWriter.WriteLine(messageText);
                                messageText = "where";
                                streamWriter.WriteLine(messageText);
                                messageText = "    SessionStudentID = @SessionStudentID;";
                                streamWriter.WriteLine(messageText);
                                streamWriter.WriteLine("");
                                messageText = "commit tran _tempTran;";
                                streamWriter.WriteLine(messageText);
                                streamWriter.WriteLine("");

                                messageText = "go";
                                streamWriter.WriteLine(messageText);
                                streamWriter.WriteLine("");
                                
                                updatedInformation.Add(SessionStudentID);

                            }

                            count = 0;

                        break;
                    }

                    count++;

                }

            }
            catch
            {
                throw;
            }
            finally
            {
                xlApp.Workbooks.Close();
                xlApp.Quit();

                if (streamWriter != null)
                {
                    streamWriter.Flush();
                    streamWriter.Close();
                    streamWriter = null;
                }

            }

        }


    }

}
