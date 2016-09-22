using System;
using System.Collections.Generic;
using System.Text;
using CLIUtility;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Office.Interop.Excel;

namespace ExcelReportCreator01.CmdLineCommands01
{
    class DatabaseQueryCommand : Command
    {
        public DatabaseQueryCommand() : base()
        {
            Name = "Database Query";
            CommandArguments.Add("-db");
            CommandArguments.Add("--database");
            CommandArguments.Add("-database");
            HelpText = "Creates an excel file from the given query and database connection\n";
            Arguments.Add(new RequiredArgument("Query"));
            Arguments.Add(new RequiredArgument("DB Conn String"));
            Arguments.Add(new RequiredArgument("FileName"));
            Arguments.Add(new OptionalArgument("FileLocation", System.IO.Directory.GetCurrentDirectory()));
            
        }
        /// <summary>
        /// Executes the create command
        /// </summary>
        public override int DoCommand()
        {
            int retval = 0;
            DataSet ds = null;
            try
            {

                String DBQuery = Arguments[0].Value;
                String DBConnStr = Arguments[1].Value;
                String FileName = Arguments[2].Value;
                String FileDir = Arguments[3].Value;

                Console.Out.WriteLine("Querying the Database.");
                Console.Out.WriteLine(" The Query is: " + DBQuery);
                Console.Out.WriteLine(" The ConnStr is: " + DBConnStr);
                ds = QueryDatabase(DBQuery, DBConnStr);
                
                String ExcelFile = FileDir + "\\" + ConvertFileName(FileName,DateTime.Now);
                Console.Out.WriteLine("Creating the Excel File.");
                Console.Out.WriteLine(" The File is named: " + ConvertFileName(FileName, DateTime.Now));
                Console.Out.WriteLine(" The Directory is: " + FileDir);
                WriteDataSetToExcelFile(ds, ExcelFile);
                Console.Out.WriteLine("Finished");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                retval = 1;
            }
            finally
            {
                if (ds != null)
                {
                    ds.Dispose();
                    ds = null;
                }               
            }

            return retval;
        }
        private String ConvertFileName(String FileName, DateTime date)
        {
            String retval = FileName;

            Int32 firstIndex = 0;
            Int32 secondIndex = 0;
            while (firstIndex >= 0)
            {
                firstIndex = FileName.IndexOf('[');
                if (firstIndex >= 0)
                {
                    secondIndex = FileName.IndexOf(']', firstIndex);
                    if (secondIndex >= 0)
                    {
                        String strbit = FileName.Substring(firstIndex, (secondIndex - firstIndex)+1);
                        String replacestrbit = (String)strbit.Clone();
                        replacestrbit = replacestrbit.Replace("[", "");
                        replacestrbit = replacestrbit.Replace("]", "");
                        replacestrbit = date.ToString(replacestrbit);
                        FileName = FileName.Replace(strbit, replacestrbit);
                    }
                }
            }
            return FileName;
        }

        

        private DataSet QueryDatabase(String Query, String DBConnString)
        {
            DataSet retval = new DataSet();
            if (Query == null)
            {
                throw new ArgumentNullException("Query");
            }
            if (DBConnString == null)
            {
                throw new ArgumentNullException("DBConnString");
            }
            SqlConnection sqlConn = null;
            SqlCommand sqlCmd = null;
            SqlDataAdapter sda = null;
            try
            {
                sqlConn = new SqlConnection(DBConnString);
                sqlConn.Open();
                sqlCmd = sqlConn.CreateCommand();
                sqlCmd.CommandTimeout = 3600;
                sqlCmd.CommandText = Query;
                sda = new SqlDataAdapter(sqlCmd);
                sda.Fill(retval);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (sda != null)
                {
                    sda.Dispose();
                    sda = null;
                }
                if (sqlCmd != null)
                {
                    sqlCmd.Dispose();
                    sqlCmd = null;
                }
                if (sqlConn != null)
                {
                    sqlConn.Close();
                    sqlConn.Dispose();
                    sqlConn = null;
                }
            }
            return retval;
        }
        private void WriteDataSetToExcelFile(DataSet ds, String ExcelFile)
        {
            DataTableReader dtr = null;
            System.Data.DataTable dt = null;
            DataTableReader dsr = null;
            Application excelApplication = null;
            Workbook excelWorkbook = null;
            Worksheet excelWorksheet = null;
            try
            {

                dtr = ds.CreateDataReader();
                dt = dtr.GetSchemaTable();
                dsr = dt.CreateDataReader();
                Int32 ColumnSize = dtr.VisibleFieldCount;
                String[] ColumnNames = new String[ColumnSize];
                Int32 index = 0;
                while (dsr.Read())
                {
                    ColumnNames[index] = dsr.GetString(0);
                    index++;
                }
                excelApplication = new Application();
                excelApplication.DisplayAlerts = false;
                //excelApplication.Visible = true;
                excelWorkbook = excelApplication.Workbooks.Add(Type.Missing);
                excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                excelApplication.Calculation = XlCalculation.xlCalculationManual;

                Int32 ColIdx = 1;
                Int32 RowIdx = 1;
                
                foreach (String ColumnName in ColumnNames)
                {
                    excelWorksheet.Cells[RowIdx, ColIdx] = ColumnName;
                    ColIdx++;
                }
                ColIdx = 1;
                RowIdx = 2;
                Int32 Maxrows = ds.Tables[0].Rows.Count;
                if (dtr.Read())
                {
                    for (ColIdx = 1; ColIdx <= ColumnSize; ColIdx++)
                    {
                        if (dtr.GetFieldType(ColIdx - 1) == typeof(String))
                        {
                            ((Range)excelWorksheet.Cells[RowIdx, ColIdx]).EntireColumn.NumberFormat = "@";
                        }
                        else if (dtr.GetFieldType(ColIdx - 1) == typeof(Decimal))
                        {
                            ((Range)excelWorksheet.Cells[RowIdx, ColIdx]).EntireColumn.NumberFormat = "#,##0.00_);(#,##0.00)";
                        }
                        else if (dtr.GetFieldType(ColIdx - 1) == typeof(DateTime))
                        {
                            ((Range)excelWorksheet.Cells[RowIdx, ColIdx]).EntireColumn.NumberFormat = "m/d/yyyy";
                        }
                        else
                        {
                            ((Range)excelWorksheet.Cells[RowIdx, ColIdx]).EntireColumn.NumberFormat = "General";
                        }
                        excelWorksheet.Cells[RowIdx, ColIdx] = dtr.GetValue(ColIdx - 1);
                    }
                    RowIdx++;
                }
                while (dtr.Read())
                {
                    for (ColIdx = 1; ColIdx <= ColumnSize; ColIdx++)
                    {
                        excelWorksheet.Cells[RowIdx, ColIdx] = dtr.GetValue(ColIdx-1);
                    }
                    RowIdx++;
                }
                excelApplication.Calculation = XlCalculation.xlCalculationAutomatic;
                excelWorkbook.SaveAs(ExcelFile,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,XlSaveAsAccessMode.xlShared,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (excelWorkbook != null)
                {
                    excelWorkbook.Close(Type.Missing,Type.Missing,Type.Missing);
                    excelWorkbook = null;
                }
                if (excelApplication != null)
                {
                    excelApplication.DisplayAlerts = true;
                    excelApplication.Quit();
                    excelApplication = null;
                }
                if (dsr != null)
                {
                    dsr.Close();
                    dsr.Dispose();
                    dsr = null;
                }
                if (dt != null)
                {
                    dt.Dispose();
                    dt = null;
                }
                if (dtr != null)
                {
                    dtr.Close();
                    dtr.Dispose();
                    dtr = null;
                }
            }
        }
        
    }
}
