using System;
using System.IO;
using System.Collections;
using System.Threading;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using CodeGeneration;
using CodeGeneration.BusinessLogic;
using CodeGeneration.Entity;
using System.Data;

namespace OutputExcel
{
	/// <summary>
	/// Summary description for WriteFileDal.
	/// </summary>
	public class WriteFileDal
	{

        private Excel.Application mExcelApplication;
        private Excel.Workbook mExcelWorkBook;
        private Excel.Worksheet mExcelWorksheet;
        private Excel.Sheets mExcelSheets;

        public WriteFileDal()
		{
            Initialize();
		}
       
        protected void Initialize()
        {

            mExcelApplication = null;
            mExcelWorkBook = null;
            mExcelWorksheet = null;
            mExcelSheets = null;
        }

        private void SetupApplication()
        {
            if (mExcelApplication == null)
            {
                mExcelApplication = new Excel.ApplicationClass();
            }
        }

        public bool FindExcelWorksheet(string worksheetName)
        {
            bool ATP_SHEET_FOUND = false;
            mExcelSheets = mExcelWorkBook.Worksheets;

            if( mExcelSheets != null )
            {
                for( int i=1; i <= mExcelSheets.Count; i++ )
                {
                    mExcelWorksheet = (Excel.Worksheet)mExcelSheets.get_Item((object)i);
                    if( this.mExcelWorksheet.Name.Equals(worksheetName) )
                    {
                        this.mExcelWorksheet.Activate();
                        ATP_SHEET_FOUND = true;
                        break;
                    }
                }
            }
            return ATP_SHEET_FOUND;
        }

/*
        public void SetupHeaderRow(string sSheetName)
        {
            
            Excel.Range rng = null;
            int nWhereAt = 1;
            if (FindExcelWorksheet(sSheetName) == true)
            {

                rng = mExcelWorksheet.get_Range("A"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCP_PERF_TS";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("B"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "HEAD";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("C"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCP_STA_ID";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("D"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "OPEN";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("E"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCP_HEAD_OPEN_IND";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("F"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCP_SENT";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("G"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCP_SENT_TS";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("H"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCP_LOAD_YYYYQ";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("I"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCP_SET_NO";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("J"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "RCT_TITLE_NAME";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

                rng = mExcelWorksheet.get_Range("K"+nWhereAt.ToString(),Type.Missing);
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rng.Value2 = "OPS_LP_PERF.PERFID";
                rng.ColumnWidth = (rng.Value2.ToString().Length * 1.667);

            }

        }

 */

        private TableColumnEntity ConvertTableColumnEntity(DataRow dataRow)
        {

            TableColumnEntity tableColumnEntity = new TableColumnEntity();

            if ((dataRow != null) && (tableColumnEntity != null))
            {

                if (dataRow.IsNull("TABLE_QUALIFIER") == false)
                {

                    tableColumnEntity.TableQualifier = Convert.ToString(dataRow["TABLE_QUALIFIER"]);

                }

                if (dataRow.IsNull("TABLE_OWNER") == false)
                {

                    tableColumnEntity.TableOwner = Convert.ToString(dataRow["TABLE_OWNER"]);

                }

                if (dataRow.IsNull("TABLE_NAME") == false)
                {

                    tableColumnEntity.TableName = Convert.ToString(dataRow["TABLE_NAME"]);

                }

                if (dataRow.IsNull("COLUMN_NAME") == false)
                {

                    tableColumnEntity.ColumnName = Convert.ToString(dataRow["COLUMN_NAME"]);

                }

                if (dataRow.IsNull("DATA_TYPE") == false)
                {

                    tableColumnEntity.DataType = Convert.ToInt16(dataRow["DATA_TYPE"]);

                }

                if (dataRow.IsNull("TYPE_NAME") == false)
                {

                    tableColumnEntity.TypeName = Convert.ToString(dataRow["TYPE_NAME"]);

                }

                if (dataRow.IsNull("PRECISION") == false)
                {

                    tableColumnEntity.ColumnPrecision = Convert.ToInt32(dataRow["PRECISION"]);

                }

                if (dataRow.IsNull("LENGTH") == false)
                {

                    tableColumnEntity.ColumnLength = Convert.ToInt32(dataRow["LENGTH"]);

                }

                if (dataRow.IsNull("SCALE") == false)
                {

                    tableColumnEntity.ColumnScale = Convert.ToInt16(dataRow["SCALE"]);

                }

                if (dataRow.IsNull("RADIX") == false)
                {

                    tableColumnEntity.ColumnRadix = Convert.ToInt16(dataRow["RADIX"]);

                }

                if (dataRow.IsNull("NULLABLE") == false)
                {

                    tableColumnEntity.Nullable = Convert.ToInt16(dataRow["NULLABLE"]);

                }

                if (dataRow.IsNull("REMARKS") == false)
                {

                    tableColumnEntity.Remarks = Convert.ToString(dataRow["REMARKS"]);

                }

                if (dataRow.IsNull("COLUMN_DEF") == false)
                {

                    tableColumnEntity.ColumnDefault = Convert.ToString(dataRow["COLUMN_DEF"]);

                }

                if (dataRow.IsNull("SQL_DATA_TYPE") == false)
                {

                    tableColumnEntity.SqlDataType = Convert.ToInt16(dataRow["SQL_DATA_TYPE"]);

                }

                if (dataRow.IsNull("SQL_DATETIME_SUB") == false)
                {

                    tableColumnEntity.SqlDateTimeSub = Convert.ToInt16(dataRow["SQL_DATETIME_SUB"]);

                }

                if (dataRow.IsNull("CHAR_OCTET_LENGTH") == false)
                {

                    tableColumnEntity.CharOctetLength = Convert.ToInt32(dataRow["CHAR_OCTET_LENGTH"]);

                }

                if (dataRow.IsNull("ORDINAL_POSITION") == false)
                {

                    tableColumnEntity.OrdinalPosition = Convert.ToInt32(dataRow["ORDINAL_POSITION"]);

                }

                if (dataRow.IsNull("IS_NULLABLE") == false)
                {

                    tableColumnEntity.IsNullable = Convert.ToString(dataRow["IS_NULLABLE"]);

                }

                if (dataRow.IsNull("SS_DATA_TYPE") == false)
                {

                    tableColumnEntity.SsDataType = Convert.ToByte(dataRow["SS_DATA_TYPE"]);

                }

            } // if ((dataRow != null) && (tableColumnEntity != null))

            return tableColumnEntity;

        } // public TableColumnEntity ConvertTableColumnEntity(DataRow dataRow)

        public void CreateExcel(string sPathFile)
        {
            Exception next = null;
            string messageText = "";
            string connectionString = "server=(local); integrated security=sspi; connection reset=false;connection lifetime=15;min pool size=1;max pool size=1000;database=ArtistEngine";

            MetaDataLogic metaDataLogic = null;
            DataSet dataSetTableColumns = null;
            DataSet dataSetTableNames = null;
            DataRow dataRow = null;
            string tableName = "";
            int whereAt = 0;
            TableColumnEntity tableColumnEntity = null;
            Excel.Range rng = null;
            int tableWhereAt = 0;
            DataRow dataRow2 = null;
            int excelColumnIndex = 0;

            string sqlOutput = "";

            try
            {
                //connectionString = "Data Source=ArtistEngineSer\\ProductionServer; Initial Catalog=ArtistEngine; User ID=sa; Password=Platinumpen7";
                metaDataLogic = new MetaDataLogic(connectionString);

                if (metaDataLogic != null)
                {

                    SetupApplication();

                    mExcelWorkBook = mExcelApplication.Workbooks.Add(Type.Missing);

                    mExcelSheets = mExcelWorkBook.Worksheets;

                    dataSetTableNames = metaDataLogic.GetTableNames("ArtistEngine");

                    if (dataSetTableNames != null)
                    {

                        if (dataSetTableNames.Tables["TableNames"] != null)
                        {

                            if (dataSetTableNames.Tables["TableNames"].Rows.Count > 0)
                            {

                                tableWhereAt = dataSetTableNames.Tables["TableNames"].Rows.Count-1;
                                do
                                {
                                    dataRow2 = dataSetTableNames.Tables["TableNames"].Rows[tableWhereAt];
                                    if (dataRow2 != null)
                                    {
                                        
                                        if (dataRow2.IsNull("TABLE_NAME") == false)
                                        {

                                            tableName = dataRow2["TABLE_NAME"].ToString().Trim();
                                            if (FindExcelWorksheet(tableName) == false)
                                            {
                                                mExcelWorksheet = (Excel.Worksheet)mExcelSheets.Add(Type.Missing,
                                                                                                    Type.Missing,
                                                                                                    Type.Missing,
                                                                                                    Type.Missing);

                                                mExcelWorksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                                                mExcelWorksheet.Name = tableName;

                                                rng = mExcelWorksheet.get_Range("A1", Type.Missing);
                                                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                                rng.Value2 = tableName;
                                                rng.ColumnWidth = (rng.Value2.ToString().Length * 2.25);

                                                sqlOutput = string.Format("select isnull(count(*),0) as [rowCount] from [dbo].[{0}];", tableName);
                                                System.Diagnostics.Debug.WriteLine(sqlOutput);
                                                System.Diagnostics.Debug.WriteLine("");

                                                dataSetTableColumns = metaDataLogic.GetTableColumns(tableName);

                                                if (dataSetTableColumns != null)
                                                {

                                                    if (dataSetTableColumns.Tables["TableColumns"] != null)
                                                    {

                                                        if (dataSetTableColumns.Tables["TableColumns"].Rows.Count > 0)
                                                        {
                                                            excelColumnIndex = 2;
                                                            whereAt = 0;

                                                            while (whereAt < dataSetTableColumns.Tables["TableColumns"].Rows.Count)
                                                            {

                                                                dataRow = dataSetTableColumns.Tables["TableColumns"].Rows[whereAt];

                                                                if (dataRow != null)
                                                                {

                                                                    tableColumnEntity = ConvertTableColumnEntity(dataRow);

                                                                    if (tableColumnEntity != null)
                                                                    {

                                                                        if ((tableColumnEntity.TypeName == "tinyint") ||
                                                                            (tableColumnEntity.TypeName == "smallint") ||
                                                                            (tableColumnEntity.TypeName == "int") ||
                                                                            (tableColumnEntity.TypeName == "decimal") ||
                                                                            (tableColumnEntity.TypeName == "bigint") ||
                                                                            (tableColumnEntity.TypeName == "tinyint identity") ||
                                                                            (tableColumnEntity.TypeName == "smallint identity") ||
                                                                            (tableColumnEntity.TypeName == "int identity") ||
                                                                            (tableColumnEntity.TypeName == "bigint identity"))
                                                                        {

                                                                            if ((tableColumnEntity.TypeName == "tinyint") || (tableColumnEntity.TypeName == "tinyint identity"))
                                                                            {
                                                                                sqlOutput = string.Format("select distinct isnull({0},255) as [{1}] from [dbo].[{2}] order by isnull({3},255);", tableColumnEntity.ColumnName, tableColumnEntity.ColumnName, tableName, tableColumnEntity.ColumnName);
                                                                            }
                                                                            else
                                                                            {
                                                                                sqlOutput = string.Format("select distinct isnull({0},-1) as [{1}] from [dbo].[{2}] order by isnull({3},-1);", tableColumnEntity.ColumnName, tableColumnEntity.ColumnName, tableName, tableColumnEntity.ColumnName);
                                                                            }


                                                                        }
                                                                        else
                                                                        {

                                                                            if (tableColumnEntity.TypeName == "xml")
                                                                            {
                                                                                sqlOutput = string.Format("select isnull({0},'OOPS') as [{1}] from [dbo].[{2}];", tableColumnEntity.ColumnName, tableColumnEntity.ColumnName, tableName);
                                                                            }
                                                                            else
                                                                            {

                                                                                if (tableColumnEntity.TypeName == "bit")
                                                                                {
                                                                                    sqlOutput = string.Format("select distinct isnull({0},-1) as [{1}] from [dbo].[{2}] order by isnull({3},-1);", tableColumnEntity.ColumnName, tableColumnEntity.ColumnName, tableName, tableColumnEntity.ColumnName);
                                                                                }
                                                                                else
                                                                                {

                                                                                    if (tableColumnEntity.TypeName == "uniqueidentifier")
                                                                                    {
                                                                                        sqlOutput = string.Format("select distinct isnull({0},'{1}') as [{2}] from [dbo].[{3}] order by isnull({4},'{5}');", tableColumnEntity.ColumnName, Guid.Empty.ToString(), tableColumnEntity.ColumnName, tableName, tableColumnEntity.ColumnName, Guid.Empty.ToString());
                                                                                    }
                                                                                    else
                                                                                    {

                                                                                        if ((tableColumnEntity.TypeName == "datetime") || (tableColumnEntity.TypeName == "date"))
                                                                                        {
                                                                                            sqlOutput = string.Format("select distinct isnull({0},'1900-01-01 00:00:00.000') as [{1}] from [dbo].[{2}] order by isnull({3},'1900-01-01 00:00:00.000');", tableColumnEntity.ColumnName, tableColumnEntity.ColumnName, tableName, tableColumnEntity.ColumnName);
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            sqlOutput = string.Format("select distinct isnull({0},'OOPS') as [{1}] from [dbo].[{2}] order by isnull({3},'OOPS');", tableColumnEntity.ColumnName, tableColumnEntity.ColumnName, tableName, tableColumnEntity.ColumnName);
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }


                                                                        System.Diagnostics.Debug.WriteLine(sqlOutput);

                                                                        rng = mExcelWorksheet.get_Range("A" + excelColumnIndex.ToString(), Type.Missing);
                                                                        rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                                                        rng.Value2 = tableColumnEntity.ColumnName;
                                                                        rng.ColumnWidth = (rng.Value2.ToString().Length * 2.25);

                                                                        rng = mExcelWorksheet.get_Range("B" + excelColumnIndex.ToString(), Type.Missing);
                                                                        rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                                                                        if ((tableColumnEntity.TypeName == "binary") || (tableColumnEntity.TypeName == "char") || (tableColumnEntity.TypeName == "nchar") || (tableColumnEntity.TypeName == "nvarchar") || (tableColumnEntity.TypeName == "varbinary") || (tableColumnEntity.TypeName == "varchar"))
                                                                        {
                                                                            messageText = string.Format("{0} ({1})", tableColumnEntity.TypeName, tableColumnEntity.ColumnLength);

                                                                        }
                                                                        else
                                                                        {
                                                                            messageText = string.Format("{0}", tableColumnEntity.TypeName);

                                                                        }

                                                                        rng.Value2 = messageText;

                                                                        rng.ColumnWidth = (rng.Value2.ToString().Length * 2.25);

                                                                        excelColumnIndex++;


                                                                    }

                                                                }

                                                                whereAt++;

                                                            }

                                                            System.Diagnostics.Debug.WriteLine("");

                                                        }
                                                    }
                                                }
                                            }
                                            
                                        }
                                    }
                                    tableWhereAt--;
                                } while (tableWhereAt >= 0);

                                if (FindExcelWorksheet("Sheet1") == true)
                                {
                                    mExcelWorksheet.Delete();
                                }

                                if (FindExcelWorksheet("Sheet2") == true)
                                {
                                    mExcelWorksheet.Delete();
                                }

                                if (FindExcelWorksheet("Sheet3") == true)
                                {
                                    mExcelWorksheet.Delete();
                                    
                                }

                                tableWhereAt = 0;
                                while (tableWhereAt < dataSetTableNames.Tables["TableNames"].Rows.Count)
                                {
                                    dataRow2 = dataSetTableNames.Tables["TableNames"].Rows[tableWhereAt];
                                    if (dataRow2 != null)
                                    {
                                        
                                        if (dataRow2.IsNull("TABLE_NAME") == false)
                                        {

                                            tableName = dataRow2["TABLE_NAME"].ToString().Trim();
                                            System.Diagnostics.Debug.WriteLine(tableName);
                                        }
                                    }
                                    tableWhereAt++;
                                }

                            }
                        }
                    }

                }

                mExcelWorkBook.SaveAs(sPathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlShared, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing);

                mExcelApplication.Workbooks.Close();

            }
            catch (Exception ex)
            {
                next = ex;

                while (next != null)
                {

                    Debug.WriteLine("Message: " + next.Message + "\r\nSource: " + next.Source + "\r\nStackTrace: " + next.StackTrace + "\r\n");

                    next = next.InnerException;

                } // while (next != null)
            }
            finally
            {
                

                mExcelApplication.Quit();
            }
        }

	}
}
