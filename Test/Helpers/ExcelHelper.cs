using Aspose.Cells;
using Excel;
using LumenWorks.Framework.IO.Csv;
using OpenQA.Selenium;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Test.Config;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;
using xl = Microsoft.Office.Interop.Excel;
namespace Test.Helpers
{
    public class ExcelHelper
    {
        public static string clipboardText;
        public static string downloadFilepath;
        public static string logFilePath;
        public static string folder = null;
        public static string filepath;
        private static List<Datacollection> _dataCol = new List<Datacollection>();
        public static string latestfile = "";
        public string xlFilePath;
        xl.Application xlApp = null;
        xl.Workbooks workbooks = null;
        xl.Workbook workbook = null;
        Hashtable sheets;
        public ExcelHelper(string xlFilePath)
        {
            this.xlFilePath = xlFilePath;
        }
        public void OpenExcel()
        {
            xlApp = new xl.Application();
            workbooks = xlApp.Workbooks;
            workbook = workbooks.Open(xlFilePath);
            sheets = new Hashtable();
            int count = 1;
            // Storing worksheet names in Hashtable.
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }
        }
        public void CloseExcel()
        {
            workbook.Close(false, xlFilePath, null); // Close the connection to workbook
            Marshal.FinalReleaseComObject(workbook); // Release unmanaged object references.
            workbook = null;
            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;
            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }
        public string getExcelSheetName(string xlFilePath,int index)
        {
            string strWorksheetName = "";
            try
            {
                xlApp = new xl.Application();
                workbooks = xlApp.Workbooks;
                workbook = workbooks.Open(xlFilePath);
                xl.Sheets sheets = workbook.Worksheets;
                xl.Worksheet worksheet = (xl.Worksheet)sheets.get_Item(index);
                strWorksheetName = worksheet.Name;
                //ADDED NEWLY
                workbooks.Close();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return strWorksheetName;
        }
        public int GetRowCount(string sheetName)
        {
            OpenExcel();
            int rowCount = 0;
            int sheetValue = 0;
            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets) // Iterate over Hashtable
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                // Getting particular worksheet using index/key from workbook
                xl.Worksheet worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange; // Range of cells which is having content.
                rowCount = range.Rows.Count;
            }
            CloseExcel();
            return rowCount;
        }
        public int GetColumnCount(string sheetName)
        {
            OpenExcel();
            int columnCount = 0;
            int sheetValue = 0;
            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;
                columnCount = range.Columns.Count;
            }
            CloseExcel();
            return columnCount;
        }
        public string GetCellData(string sheetName, int colNumber, int rowNumber)
        {
            OpenExcel();
            string value = string.Empty;
            int sheetValue = 0;
            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;
                value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Value2);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            CloseExcel();
            return value;
        }
        public string GetCellData(string sheetName, string colName, int rowNumber)
        {
            OpenExcel();
            string value = string.Empty;
            int sheetValue = 0;
            int colNumber = 0;
            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;
                for (int i = 1; i <= range.Columns.Count; i++)
                {
                    string colNameValue = Convert.ToString((range.Cells[1, i] as xl.Range).Value2);
                    if (colNameValue.ToLower() == colName.ToLower())
                    {
                        colNumber = i;
                        break;
                    }
                }
                value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Value2);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            CloseExcel();
            return value;
        }
        public string SetCellData(string sheetName, string colName, int rowNumber, string value)
        {
            OpenExcel();
            int sheetValue = 0;
            int colNumber = 0;
            try
            {
                if (sheets.ContainsValue(sheetName))
                {
                    foreach (DictionaryEntry sheet in sheets)
                    {
                        if (sheet.Value.Equals(sheetName))
                        {
                            sheetValue = (int)sheet.Key;
                        }
                    }
                    xl.Worksheet worksheet = null;
                    worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                    xl.Range range = worksheet.UsedRange;
                    for (int i = 1; i <= range.Columns.Count; i++)
                    {
                        string colNameValue = Convert.ToString((range.Cells[1, i] as xl.Range).Value2);
                        if (colNameValue.ToLower() == colName.ToLower())
                        {
                            colNumber = i;
                            break;
                        }
                    }
                    range.Cells[rowNumber, colNumber] = value;
                    workbook.Save();
                    Marshal.FinalReleaseComObject(worksheet);
                    worksheet = null;
                    CloseExcel();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return value;

        }
        public static void PopulateInCollection(string fileName)
        {
            DataTable table = ExcelToDataTable(fileName);
            //Iterate through the rows and columns of the Table
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    Datacollection dtTable = new Datacollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    //Add all the details for each row
                    _dataCol.Add(dtTable);
                }
            }
        }
        ///// <summary>
        ///// Reading all the datas from Excelsheet
        ///// </summary>
        ///// <param name="fileName"></param>
        ///// <returns></returns>
        //public static DataTable ExcelToDataTable(string fileName)
        //{
        //    using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
        //    {
        //        using (var reader = ExcelReaderFactory.CreateReader(stream))
        //        {
        //            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
        //            {
        //                ConfigureDataTable = (data) => new ExcelDataTableConfiguration()
        //                {
        //                    UseHeaderRow = true
        //                }
        //            });
        //            //Get all the Tables
        //            DataTableCollection table = result.Tables;
        //            //Store it in DataTable
        //            DataTable resultTable = table["Sheet1"];
        //            //return
        //            return resultTable;
        //        }
        //    }
        //}
        /// <summary>
        /// Reading all the datas from Excelsheet
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string fileName)
        {
            //Open file and return as stream
            FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            //CreateOpenXmlReader via  ExcelReaderFactory
            IExcelDataReader excelReaderxslx = ExcelReaderFactory.CreateOpenXmlReader(stream);//.xlsx
            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //  IExcelDataReader excelReaderxlsx = ExcelReaderFactory.CreateBinaryReader(stream);//.xls
            //for .xlsx
            //Set First Row as Column Name
            excelReaderxslx.IsFirstRowAsColumnNames = true;
            //Return as DataSet
            DataSet result = excelReaderxslx.AsDataSet();
            //Get all the Tables
            DataTableCollection table = result.Tables;
            //Store it in DataTable
            DataTable resultTable = table["Sheet1"];
            //return
            return resultTable;
        }
        public static string ReadData(int rowNumber, string columnName)
        {
            try
            {
                //Retriving Data using LINQ to reduce much of iterations
                string data = (from colData in _dataCol
                               where colData.colName == columnName && colData.rowNumber == rowNumber
                               select colData.colValue).SingleOrDefault();
                //var datas = dataCol.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;
                return data.ToString();
            }
            catch (Exception e)
            {
                return null;
            }
        }
        public class Datacollection
        {
            public int rowNumber { get; set; }
            public string colName { get; set; }
            public string colValue { get; set; }
        }
        public static bool CheckExcelContains(String filePath, string TextToFind, bool bSortAndCompare = false)
        {
            try
            {
                var wb1 = new Workbook(filePath);
                int sheetcount = wb1.Worksheets.Count;
                StringBuilder fileContent = new StringBuilder();
                Worksheet worksheet1 = wb1.Worksheets[0];
                Cells SourceCells = worksheet1.Cells;
                for (int i = 0; i < SourceCells.MaxDataRow + 1; i++)
                {
                    for (int j = 0; j < SourceCells.MaxDataColumn + 1; j++)
                    {
                        string sourceValue = SourceCells[i, j].StringValue.Trim();
                        if (sourceValue.Contains(TextToFind))
                        {
                            bSortAndCompare = true;
                            fileContent.AppendLine("The text" + TextToFind + " is  in the Cell [" + i + "," + j + "]");
                            fileContent.AppendLine("Excel value:" + sourceValue);
                            break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                bSortAndCompare = false;
                Console.WriteLine(e);
                return bSortAndCompare;
            }
            return bSortAndCompare;
        }
        //Shwetabh Srivastava----Add text to excel (parameters-filepath,filename,rownumber,columnnumber and data to write)
        public static void WriteToExcel(IWebDriver driver, String filepath, String fileName, int RowNumber,
            int ColNumber, String dataToWrite)
        {
            //************************Write using Interop****************************************************
            //**************************************************************************************************
            //int rw = 0;
            //int cl = 0;
            //StackFrame stackFrame = new StackFrame();
            //MethodBase methodBase = stackFrame.GetMethod();
            //WriteLog(driver, filepath, "Executing Method :- " + methodBase.Name);
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filepath + "\\" + latestfile, 0, false, 5, "", "", false,
            //    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //Excel.Range xlRange = xlWorkSheet.UsedRange;
            //int rowNumber = xlRange.Rows.Count + 1;
            //rw = xlRange.Rows.Count;
            //cl = xlRange.Columns.Count;
            //xlWorkSheet.Cells[RowNumber, ColNumber] = dataToWrite;
            //WriteLog(driver, filepath, dataToWrite + "-> Data added successfully at Column Number-> " + cl + " and Row Number ->" + rw);
            //// Disable file override confirmaton message  
            //xlApp.DisplayAlerts = false;
            //xlWorkBook.SaveAs(filepath + "\\" + latestfile, Excel.XlFileFormat.xlOpenXMLWorkbook,
            //    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
            //    Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
            //    Missing.Value, Missing.Value);
            //xlWorkBook.Close();
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
            //Console.BackgroundColor = ConsoleColor.DarkBlue;
            //Console.WriteLine("\nRecords Added successfully...");
            //Console.BackgroundColor = ConsoleColor.Black;
            //************************Write using Aspose.Cell****************************************************
            //WORKING NOW
            //**************************************************************************************************
            Aspose.Cells.License license = new Aspose.Cells.License();
            //   license.SetLicense("C:\\Automation\\Aspose.Total.lic");
            license.SetLicense("D:\\CWorkspace\\ShellWorkspace\\ShellTest\\Aspose.Total.lic");
            Workbook wb = new Workbook(filepath + "\\" + fileName);
            Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            worksheet.Cells[RowNumber, ColNumber].PutValue(dataToWrite);
            LogHelper.WriteLog( filepath, " Data--> (" + dataToWrite + ") added successfully at Column Number-> " + ColNumber + " and Row Number ->" + RowNumber);
            wb.Save(filepath + "\\" + latestfile);
        }
        //Shwetabh Srivastava----Read contexts of excel file and print in log file
        public static void readExcel(IWebDriver driver, String filepath, String fileName)
        {
            //****************************************************************************************************************
            //***************************************************************Working code thr Interop****************
            //****************************************************************************************************
            //Excel.Application xlApp;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //Excel.Range range;
            //string str;
            //int rCnt;
            //int cCnt;
            //int rw = 0;
            //int cl = 0;
            //StackFrame stackFrame = new StackFrame();
            //MethodBase methodBase = stackFrame.GetMethod();
            //WriteLog(driver, filepath, methodBase.Name);
            //xlApp = new Excel.Application();
            //xlWorkBook = xlApp.Workbooks.Open(filepath + "\\" + fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //range = xlWorkSheet.UsedRange;
            //rw = range.Rows.Count;
            //cl = range.Columns.Count;
            //WriteLog(driver, filepath, "Reading Excel contents of file " + filepath + "\\" + fileName + "   are :-");
            //for (rCnt = 1; rCnt <= rw; rCnt++)
            //{
            //    for (cCnt = 1; cCnt <= cl; cCnt++)
            //    {
            //        str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
            //        WriteLog(driver, filepath, str + "\n");
            //    }
            //}
            ////xlWorkBook.SaveAs( latestfile, Excel.XlFileFormat.xlOpenXMLWorkbook,
            ////    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
            ////    Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
            ////    Missing.Value, Missing.Value);
            ////     xlWorkBook.Close();
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
            //WriteLog(driver, filepath, "Data Read Successfully");
            //****************************************************************************************************************
            //***************************************************************Working code with Aspose.Cell****************
            //****************************************************************************************************************
            //StackFrame stackFrame = new StackFrame();
            //MethodBase methodBase = stackFrame.GetMethod();
            //WriteLog(driver, filepath, methodBase.Name);
            ////Open your template file.
            //Workbook wb = new Workbook(filepath + "\\" + fileName);
            ////Get the first worksheet.
            //Worksheet worksheet = wb.Worksheets[0];
            ////Get the cells collection.
            //Cells cells = worksheet.Cells;
            ////Define the list.
            //List<string> myList = new List<string>();
            ////Get the AA column index. (Since "Status" is always @ AA column.
            //int col = CellsHelper.ColumnNameToIndex("A");
            ////Get the last row index in AA column.
            //int last_row = worksheet.Cells.GetLastDataRow(col);
            ////Loop through the "Status" column while start collecting values from row 9
            ////to save each value to List
            //for (int i = 8; i <= last_row; i++)
            //{
            //    myList.Add(cells[i, col].Value.ToString());
            //}
            //WriteLog(driver, filepath, "Reading Excel contents of file " + filepath + "\\" + fileName + "   are :-");
            //foreach (string element  in myList)
            //{
            //    WriteLog(driver, filepath, element + "\n");
            //}
            //WriteLog(driver, filepath, "Data Read Successfully");
            Workbook wb = new Workbook(filepath + "\\" + fileName);
            Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            LogHelper.WriteLog( filepath, "Reading Excel contents of file " + filepath + "\\" + fileName + "   are :-");
            for (int i = 0; i < cells.MaxDataRow + 1; i++)
            {
                for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                {
                    string s = cells[i, j].StringValue.Trim();
                    LogHelper.WriteLog( filepath, s + "\n");
                }
            }
            LogHelper.WriteLog( filepath, "Data Read Successfully");
            //**********************************not working for ***********************************
            //Not working for   Read and Write Excel Documents Using OLEDB
            //**********************************************************************************************
            // DataSet ds = new DataSet();
            // string connectionString = GetConnectionString(filepath, fileName);
            // using (OleDbConnection conn = new OleDbConnection(connectionString))
            // {
            //     conn.Open();
            //     OleDbCommand cmd = new OleDbCommand();
            //     cmd.Connection = conn;
            //     // Get all Sheets in Excel File
            //     DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //     // Loop through all Sheets to get data
            //     foreach (DataRow dr in dtSheet.Rows)
            //     {
            //         string sheetName = dr["TABLE_NAME"].ToString();
            //         if (!sheetName.EndsWith("$"))
            //             continue;
            //         // Get all rows from the Sheet
            //         cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
            //         DataTable dt = new DataTable();
            //         dt.TableName = sheetName;
            //         OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            //         da.Fill(dt);
            //         ds.Tables.Add(dt);
            //     }
            //     WriteLog(driver, filepath, "ds");
            //     cmd = null;
            //     conn.Close();
            // }
            //// return ds;
        }
        //Shwetabh Srivastava----Read specified contexts of excel file and print in log file
        public static string readExcelrowcolumnwise(IWebDriver driver, String filepath, String fileName, int sRowNumber,
            int eRowNumber, int sColNumber, int eColNumber)
        {
            //***************************************************************Working code****************
            //Excel.Application xlApp;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //Excel.Range range;
            //int rw = 0;
            //int cl = 0;
            //StackFrame stackFrame = new StackFrame();
            //MethodBase methodBase = stackFrame.GetMethod();
            //WriteLog(driver, filepath, "Executing Method :- " + methodBase.Name);
            //xlApp = new Excel.Application();
            //xlWorkBook = xlApp.Workbooks.Open(filepath + "\\" + fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //range = xlWorkSheet.UsedRange;
            //rw = range.Rows.Count;
            //cl = range.Columns.Count;
            //WriteLog(driver, filepath, "Reading Excel contents of file " + filepath + "\\" + fileName);
            //string test = xlWorkSheet.Cells[RowNumber, ColNumber].Value.ToString();
            //WriteLog(driver, filepath, "Excel data for Row Number ->" + RowNumber + " and Coumn Number " + ColNumber + " is :" + test);
            ////xlWorkBook.SaveAs( latestfile, Excel.XlFileFormat.xlOpenXMLWorkbook,
            ////    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
            ////    Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
            ////    Missing.Value, Missing.Value);
            ////     xlWorkBook.Close();
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
            //WriteLog(driver, filepath, "Data Read Successfully \n \n");
            //*********************************************
            //     using aspose.cells
            //**********************************************
            LogHelper.WriteLog( filepath, "**************************************");
            LogHelper.WriteLog( filepath, "Inside method readExcelrowcolumnwise\n \n");
            LogHelper.WriteLog( filepath, "********************************************");
            string s = null;
            Workbook wb = new Workbook(filepath + "\\" + fileName);
            Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            LogHelper.WriteLog( filepath, "Reading Excel contents of file " + filepath + "\\" + fileName + "   are :-");
            for (int i = sRowNumber; i <= eRowNumber; i++)
            {
                for (int j = sColNumber; j <= eColNumber; j++)
                {
                    //       s = cells[i, j].StringValue.Trim();
                    s = Convert.ToString(cells[i, j].StringValue.Trim());
                    LogHelper.WriteLog( filepath, "Excel data for Row Number ->" + i + " and Column Number " + j + " is :" + s);
                }
            }
            LogHelper.WriteLog( filepath, "Data Read Successfully \n \n");
            return s;
        }
        //Shwetabh Srivastava----Read specified contexts of excel file and print in log file
        public static string ReaddataonmultiplesheetExcelrowcolumnwise(IWebDriver driver, String filepath, String fileName, int sRowNumber, int eRowNumber, int sColNumber, int eColNumber, string SheetName)
        {
            //*********************************************
            //     using aspose.cells
            //********************************
            LogHelper.WriteLog( filepath, "**************************************");
            LogHelper.WriteLog( filepath, "Inside method ReaddataonmultiplesheetExcelrowcolumnwise \n \n");
            LogHelper.WriteLog( filepath, "********************************************");
            string s = null;
            Workbook wb = new Workbook(filepath + "\\" + fileName);
            Worksheet worksheet = wb.Worksheets[SheetName];
            //    Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            LogHelper.WriteLog( filepath, "Reading Excel contents of file " + filepath + "\\" + fileName + "   are :-");
            for (int i = sRowNumber; i <= eRowNumber; i++)
            {
                for (int j = sColNumber; j <= eColNumber; j++)
                {
                    //       s = cells[i, j].StringValue.Trim();
                    s = Convert.ToString(cells[i, j].StringValue.Trim());
                    LogHelper.WriteLog( filepath, "Actual Excel data for Row Number ->" + i + " and Column Number " + j + " in Sheet  is :" + s);
                }
            }
            LogHelper.WriteLog( filepath, "Data Read Successfully \n \n");
            return s;
        }
        //Shwetabh Srivastava----Read specified contexts of excel file and print in log file
        public static string ReaddataonmultiplesheetExcelrowcolumnwise1(IWebDriver driver, String filepath, String fileName, int sRowNumber, int eRowNumber, int sColNumber, int eColNumber, string SheetName)
        {
            //*********************************************
            //     using aspose.cells
            //********************************
            LogHelper.WriteLog( filepath, "**************************************");
            LogHelper.WriteLog( filepath, "Inside method ReaddataonmultiplesheetExcelrowcolumnwise \n \n");
            LogHelper.WriteLog( filepath, "********************************************");
            string s = null;
            Workbook wb = new Workbook(filepath + "\\" + fileName);
            Worksheet worksheet = wb.Worksheets[SheetName];
            //    Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            LogHelper.WriteLog( filepath, "Reading Excel contents of file " + filepath + "\\" + fileName + "   are :-");
            for (int i = sRowNumber; i <= eRowNumber; i++)
            {
                for (int j = sColNumber; j <= eColNumber; j++)
                {
                    //       s = cells[i, j].StringValue.Trim();
                    s = Convert.ToString(cells[i, j].StringValue.Trim());
                    LogHelper.WriteLog( filepath, "Actual Excel data for Row Number ->" + i + " and Column Number " + j + " in Sheet  is :" + s);
                }
            }
            LogHelper.WriteLog( filepath, "Data Read Successfully \n \n");
            return s;
        }
        //Shwetabh Srivastava----Read specified contexts of excel file and print in log file
        public static bool verifydataonmultiplesheetExcelrowcolumnwise(IWebDriver driver, String filepath, String fileName, int sRowNumber, int eRowNumber, int sColNumber, int eColNumber, string SheetName, string texttoverify)
        {
            //*********************************************
            //     using aspose.cells
            //********************************
            bool result = false;
            LogHelper.WriteLog( filepath, "**************************************");
            LogHelper.WriteLog( filepath, "Inside method verifydataonmultiplesheetExcelrowcolumnwise \n \n");
            LogHelper.WriteLog( filepath, "********************************************");
            string s = null;
            Workbook wb = new Workbook(filepath + "\\" + fileName);
            Worksheet worksheet = wb.Worksheets[SheetName];
            //    Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            LogHelper.WriteLog( filepath, "Reading Excel contents of file " + filepath + "\\" + fileName + "   are :-");
            for (int i = sRowNumber; i <= eRowNumber; i++)
            {
                for (int j = sColNumber; j <= eColNumber; j++)
                {
                    //       s = cells[i, j].StringValue.Trim();
                    s = Convert.ToString(cells[i, j].StringValue.Trim());
                    LogHelper.WriteLog( filepath, "Excel data for Row Number ->" + i + " and Column Number " + j + " is :" + s);
                    if (s.Equals(texttoverify))
                    {
                        LogHelper.WriteLog( filepath, "Actual Excel Data for Row Number ->" + i + " and Column Number -->" + j + " is :[" + s + "] &  Expected data is -->[" + texttoverify + "]. Hence ,verification is Passed ");
                        return true;
                    }
                    else
                    {
                        LogHelper.WriteLog( filepath, "Actual Excel Data for Row Number ->" + i + " and Column Number -->" + j + " is :[" + s + "] &  Expected data -->[" + texttoverify + "]. Hence ,verification is failed ");
                        throw new Exception("Actual Excel Data for Row Number ->" + i + " and Column Number -->" + j + " is :[" + s + "] &  Expected data -->[" + texttoverify + "]. Hence ,verification is failed ");
                        return false;
                    }
                }
            }
            LogHelper.WriteLog( filepath, "Data Read Successfully \n \n");
            return result;
        }
        public static void Readfromcsv(IWebDriver driver, string filepath, string filename, string header)
        {
            LogHelper.WriteLog( filepath, "**********************");
            LogHelper.WriteLog( filepath, "Inside method Readfromcsv");
            LogHelper.WriteLog( filepath, "***********************");
            TextReader textReader = null;
            List<string> columns = new List<string>();
            try
            {
                using (CsvReader csv = new CsvReader(new StreamReader(@filepath + "\\" + filename + ".csv"), true))
                {
                    int fieldCount = csv.FieldCount;
                    string[] headers;
                    int i;
                    int idx;
                    for (i = 0; i <= fieldCount - 1; i++)
                    {
                        //  Display the column names:
                        headers = csv.GetFieldHeaders();
                        Console.WriteLine("Column names present in csv are ->" + Convert.ToString(i) + ": " + headers[i]);
                        LogHelper.WriteLog( filepath, "Column names present in csv at index ->" + Convert.ToString(i) + " is ->: " + headers[i]);
                        //  The following line demonstrates  to get the column index given a column name:
                        idx = csv.GetFieldIndex(headers[i]);
                        Console.WriteLine(headers + " is at column index " + Convert.ToString(idx));
                        LogHelper.WriteLog( filepath, headers[i] + " is present at column index " + Convert.ToString(idx));
                        if (headers[i].Contains(header))
                        {
                            while (csv.ReadNextRecord())
                            {
                                columns.Add(csv[i]);
                                foreach (var item in columns)
                                {
                                    LogHelper.WriteLog( filepath, "  Data present in CSV in Column Header ->" + header + "  are--> " + item);
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                if (textReader != null)
                {
                    textReader.Close();
                }
            }
        }
        public static string Readcsvcolumnnamerownumberwise(IWebDriver driver, string csvfilepath, string csvfilename, string header, int rownumber)
        {
            string logpath = ConfigReader.logFilePath;
            LogHelper.WriteLog( logpath, "**********************");
            LogHelper.WriteLog( logpath, "Inside method Readfromcsvcolumnnamerownumberwise");
            LogHelper.WriteLog( logpath, "***********************");
            TextReader textReader = null;
            List<string> columns = new List<string>();
            string FileLines = System.IO.File.ReadAllText(@csvfilepath + "\\" + csvfilename + ".csv");
            string cellvaluenew = null;
            string cellvalue = null;
            // Split into lines.
            FileLines = FileLines.Replace('\n', '\r');
            string[] lines = FileLines.Split(new char[] { '\r' }, StringSplitOptions.RemoveEmptyEntries);
            // See how many rows and columns there are.
            // 6 rows
            int num_rows = lines.Length;
            try
            {
                using (CsvReader csv = new CsvReader(new StreamReader(@csvfilepath + "\\" + csvfilename + ".csv"), true))
                {
                    int fieldCount = csv.FieldCount;
                    string[] headers;
                    int r = 0, c = 0;
                    int idx;
                    //3 columns
                    int num_cols = lines[0].Split(',').Length;
                    // Allocate the data array.
                    string[,] al = new string[num_rows, num_cols];
                    List<string> al1 = new List<string>(num_rows * num_cols);
                    // Load the array.
                    //1..2..3.4..5
                    for (r = 1; r <= num_rows - 1; r++)
                    {
                        //split records line wise
                        string[] line_r = lines[r].Split(',');
                        //  1	shwetabah	India
                        //  2   charu china
                        //   3   ROHIT RUSSIA
                        //   4   Mohit usa
                        //   5   shipra canada
                        if (r.Equals(rownumber))
                        {
                            for (c = 0; c <= fieldCount - 1; c++)
                            {
                                //  Display the column names:
                                headers = csv.GetFieldHeaders();
                                //  if (headers[c].Contains(header))
                                if (headers[c].Equals(header))
                                {
                                    al[r, c] = line_r[c];
                                    al1.Add(al[r, c]);
                                    cellvalue = al[r, c];
                                    LogHelper.WriteLog( logpath, "Data present in CSV at  Row [" + r + "] Column [ " + header + " ] has value -->[ " + al[r, c] + "]");
                                    LogHelper.WriteLog( logpath, "Data present in CSV at  Row [" + r + "] Column [ " + header + " ] has value -->[ " + cellvalue + "]");
                                    headers = csv.GetFieldHeaders();
                                    LogHelper.WriteLog( logpath, "Column names present in csv at index ->" + Convert.ToString(c) + " is ->: " + headers[c]);
                                    //  The following line demonstrates  to get the column index given a column name:
                                    idx = csv.GetFieldIndex(headers[c]);
                                    LogHelper.WriteLog( logpath, headers[c] + "->is present at column index -->" + Convert.ToString(idx));
                                    //while (csv.ReadNextRecord())
                                    //{
                                    //    columns.Add(csv[c]);//1,2,3,4,5
                                    //    foreach (var item in columns)
                                    //    {
                                    //        LogHelper.WriteLog(driver, filepath, "  Data present in CSV in Column Header ->" + header + "  are--> " + item);
                                    //        al[r, c] = line_r[c];
                                    //        al1.Add(al[r, c]);
                                    //   //    LogHelper.WriteLog(driver, csvfilepath, "  Data present in CSV at  Row [" + r + "] Column [ " + c + " ] has value -->[ " + al[r, c] + "]");
                                    //    }
                                    //}
                                }//end of if
                            }//end of for
                        }//end of if
                    }//end of for
                }//end of using
            }//end of try
            finally
            {
                if (textReader != null)
                {
                    textReader.Close();
                }
            }
            return cellvalue;
        }
        public static string Readfromcsvrowcolumnwise(IWebDriver driver, string csvfilepath, string csvfilename, int srownum, int erownum, int scolnum, int ecolnum)
        {
            LogHelper.WriteLog( csvfilepath, "**********************");
            LogHelper.WriteLog( csvfilepath, "Inside method Readfromcsvrowrowcolumnwise \n \n");
            LogHelper.WriteLog( csvfilepath, "***********************");
            string FileLines = System.IO.File.ReadAllText(@csvfilepath + "\\" + csvfilename + ".csv");
            // Split into lines.
            FileLines = FileLines.Replace('\n', '\r');
            string[] lines = FileLines.Split(new char[] { '\r' },
                StringSplitOptions.RemoveEmptyEntries);
            // See how many rows and columns there are.
            int num_rows = lines.Length;
            int num_cols = lines[0].Split(',').Length;
            // Allocate the data array.
            string[,] al = new string[num_rows, num_cols];
            List<string> al1 = new List<string>(num_rows * num_cols);
            string cellvalue = null;
            // Load the array.
            for (int r = srownum; r <= erownum; r++)
            {
                string[] line_r = lines[r].Split(',');
                for (int c = scolnum; c <= ecolnum; c++)
                {
                    al[r, c] = line_r[c];
                    al1.Add(al[r, c]);
                    cellvalue = al[r,c];
                    //    LogHelper.WriteLog(driver, csvfilepath, "  Data present in CSV at  Row [" + r + "] Column [ " + c + " ] has value -->[ " + al1[r, c] + "]");
                    LogHelper.WriteLog( csvfilepath, "  Data present in CSV at  Row [" + r + "] Column [ " + c + " ] has value -->[ " + al[r, c] + "]");
                }
            }
            return cellvalue;
        }
        public static bool csvcompare(IWebDriver driver, string sourcecsvfilepath, string sourcecsvfilename, string destinationcsvfilepath, string destinationcsvfilename)
        {
            LogHelper.WriteLog( sourcecsvfilepath, "**********************");
            LogHelper.WriteLog( sourcecsvfilepath, "Inside method csvcompare ");
            LogHelper.WriteLog( sourcecsvfilepath, "***********************");
            string[] fileContentsOne = File.ReadAllLines(@sourcecsvfilepath + "\\" + sourcecsvfilename + ".csv");
            string[] fileContentsTwo = File.ReadAllLines(@destinationcsvfilepath + "\\" + destinationcsvfilename + ".csv");
            System.IO.StreamReader file1 = new System.IO.StreamReader(@sourcecsvfilepath + "\\" + sourcecsvfilename + ".csv");
            System.IO.StreamReader file2 = new System.IO.StreamReader(@destinationcsvfilepath + "\\" + destinationcsvfilename + ".csv");
            //Compare length of files
            bool result = false;
            if (!fileContentsOne.Length.Equals(fileContentsTwo.Length))
                return false;
            for (int i = 0; i < fileContentsOne.Length; i++)
            {
                string[] columnsOne = fileContentsOne[i].Split(new char[] { ';' });
                string[] columnsTwo = fileContentsTwo[i].Split(new char[] { ';' });
                //If length of files are equal , Compare number of columns on each row
                if (!columnsOne.Length.Equals(columnsTwo.Length))
                    return false;
                //If Columns length are equal, Compare column values
                for (int j = 0; j < columnsOne.Length; j++)
                {
                    LogHelper.WriteLog( destinationcsvfilepath, "  Data present in CSV1 at row [ " + i + " ] has value -->[ " + columnsOne[j] + " ]");
                    LogHelper.WriteLog( destinationcsvfilepath, "  Data present in CSV2 at row [ " + i + " ] has value -->[ " + columnsTwo[j] + " ]");
                    if (columnsOne[j].Equals(columnsTwo[j])) //if line 1 equals line 2
                    {
                        LogHelper.WriteLog( destinationcsvfilepath, " No differences found"); //no differences found
                        continue;
                    }
                    else if (!columnsOne[j].Equals(columnsTwo[j]))
                    {
                        LogHelper.WriteLog( destinationcsvfilepath, " Differences found at Line Number [ " + i + " ] "); //no differences found
                        throw new Exception(" Differences found at Line Number [ " + i + " ]");
                        return false;
                    }
                }
            }
            return result;
        }
        public static bool csvComparerowcolumnnwise(IWebDriver driver, string sourcecsvfilepath, string sourcecsvfilename, int sourcesrownum, int sourceerownum, int sourcescolnum, int sourceecolnum, string destinationcsvfilepath, string destinationcsvfilename, int destinationsrownum, int destinationerownum, int destinationscolnum, int destinationecolnum)
        {
            LogHelper.WriteLog( sourcecsvfilepath, "**********************");
            LogHelper.WriteLog( sourcecsvfilepath, "Inside method csvComparerowcolumnnwise");
            LogHelper.WriteLog( sourcecsvfilepath, "***********************");
            //CSV 1 
            //Reading CSV
            LogHelper.WriteLog( sourcecsvfilepath, "**********************");
            LogHelper.WriteLog( sourcecsvfilepath, "Reading CSV1 Data");
            LogHelper.WriteLog( sourcecsvfilepath, "***********************");
            //     string FileLinesone = System.IO.File.ReadAllText(@sourcecsvfilepath + "\\" + sourcecsvfilename + ".csv");
            // Split into lines.
            //        FileLinesone = FileLinesone.Replace('\n', '\r');
            //   string[] linesone = FileLinesone.Split(new char[] { '\r' }, StringSplitOptions.RemoveEmptyEntries);
            string[] linesone = File.ReadAllLines(@sourcecsvfilepath + "\\" + sourcecsvfilename + ".csv");
            // See how many rows and columns there are.
            int num_rowsone = linesone.Length;
            int max_num_colsone = 0;
            int num_colsone = 0;
            for (int i = 0; i < num_rowsone; i++)
            {
                LogHelper.WriteLog( @sourcecsvfilepath, " Reading Row ->" + i + "in CSV 1 ");
                num_colsone = linesone[i].Split(',').Length;
                LogHelper.WriteLog( @sourcecsvfilepath, " Number of columns in CSV 1 at Row  ->" + i + " are " + num_colsone);
                if (max_num_colsone < num_colsone)
                {
                    max_num_colsone = num_colsone;
                }
            }
            // Allocate the data array.
            string[,] al = new string[num_rowsone, max_num_colsone];
            List<string> al1 = new List<string>(num_rowsone * max_num_colsone);
            // Load the array.
            for (int r1 = sourcesrownum; r1 <= sourceerownum; r1++)
            {
                string[] line_r = linesone[r1].Split(',');
                for (int c1 = sourcescolnum; c1 <= sourceecolnum; c1++)
                {
                    al[r1, c1] = line_r[c1];
                    al1.Add(al[r1, c1]);
                    //    LogHelper.WriteLog(driver, csvfilepath, "  Data present in CSV at  Row [" + r + "] Column [ " + c + " ] has value -->[ " + al1[r, c] + "]");
                    LogHelper.WriteLog( @sourcecsvfilepath, "  Data present in CSV1 at  Row [ " + r1 + " ] Column [ " + c1 + " ] has value -->[ " + al[r1, c1] + " ]");
                }
            }
            //CSV 2
            //Reading CSV
            bool result = false;
            LogHelper.WriteLog( sourcecsvfilepath, "**********************");
            LogHelper.WriteLog( sourcecsvfilepath, "Reading CSV2 Data");
            LogHelper.WriteLog( sourcecsvfilepath, "***********************");
            //string FileLinestwo = System.IO.File.ReadAllText(@destinationcsvfilepath + "\\" + destinationcsvfilename + ".csv");
            //// Split into lines.
            //FileLinestwo = FileLinestwo.Replace('\n', '\r');
            //string[] linestwo = FileLinestwo.Split(new char[] { '\r' },StringSplitOptions.RemoveEmptyEntries);
            string[] linestwo = File.ReadAllLines(@destinationcsvfilepath + "\\" + destinationcsvfilename + ".csv");
            // See how many rows and columns there are.
            int num_rowstwo = linestwo.Length;
            int max_num_colstwo = 0;
            int num_colstwo = 0;
            for (int j = 0; j < num_rowstwo; j++)
            {
                LogHelper.WriteLog( @destinationcsvfilepath, " Reading Row ->" + j + "in CSV2 ");
                num_colstwo = linestwo[j].Split(',').Length;
                LogHelper.WriteLog( @destinationcsvfilepath, " Number of columns in CSV 2 at Row  ->" + j + " are " + num_colstwo);
                if (max_num_colstwo < num_colstwo)
                {
                    max_num_colstwo = num_colstwo;
                }
            }
            // Allocate the data array.
            string[,] all = new string[num_rowstwo, max_num_colstwo];
            List<string> all1 = new List<string>(num_rowstwo * max_num_colstwo);
            // Load the array.
            for (int r2 = destinationsrownum; r2 <= destinationerownum; r2++)
            {
                string[] line_r1 = linestwo[r2].Split(',');
                for (int c2 = destinationscolnum; c2 <= destinationecolnum; c2++)
                {
                    all[r2, c2] = line_r1[c2];
                    all1.Add(all[r2, c2]);
                    //    LogHelper.WriteLog(driver, csvfilepath, "  Data present in CSV at  Row [" + r + "] Column [ " + c + " ] has value -->[ " + al1[r, c] + "]");
                    LogHelper.WriteLog( destinationcsvfilepath, "  Data present in CSV2 at  Row [ " + r2 + " ] Column [ " + c2 + " ] has value -->[ " + all[r2, c2] + " ]");
                }
            }
            LogHelper.WriteLog( sourcecsvfilepath, "**********************");
            LogHelper.WriteLog( sourcecsvfilepath, "Comparison of CSV1 & CSV2  Starts \n \n");
            LogHelper.WriteLog( sourcecsvfilepath, "***********************");
            List<string> al3 = new List<string>();
            string result1 = null;
            string result2 = null;
            string result3 = null;
            //foreach (string temp in al1)
            //{
            //    result1 = temp;
            foreach (string item in al1)
            {
                result2 = item;
                al3.Add(all1.Contains(item) ? "Pass" : "Fail");
                foreach (string cc in al3)
                {
                    result3 = cc;
                }
            }
            //}
            LogHelper.WriteLog( destinationcsvfilepath, "Data Comparison between CSV1 [" + sourcecsvfilename + " ] and CSV2  [ " + destinationcsvfilename + " ] is-- >  [ " + result3 + " ]");
            if (al3.Contains("Fail"))
            {
                LogHelper.WriteLog( destinationcsvfilepath, "Data Comparison between CSV1 [" + sourcecsvfilename + " ] and CSV2  [ " + destinationcsvfilename + " ]  is failed ");
                throw new Exception("Data Comparison between CSV1 [" + sourcecsvfilename + " ] and CSV2  [ " + destinationcsvfilename + " ]  is failed ");
                return false;
            }
            else
            {
                LogHelper.WriteLog( destinationcsvfilepath, "Data Comparison between CSV1 [" + sourcecsvfilename + " ] and CSV2  [ " + destinationcsvfilename + " ]  is Passed ");
                return true;
            }
            return result;
        }
        public static bool csvCompareExcel(IWebDriver driver, string sourcecsvfilepath, string sourcecsvfilename, int sourcesrownum, int sourceerownum, int sourcescolnum, int sourceecolnum, string destinationfilepath, string destinationfilename, int destinationsrownum, int destinationerownum, int destinationscolnum, int destinationecolnum, string SheetName)
        {
            LogHelper.WriteLog( sourcecsvfilepath, "**********************");
            LogHelper.WriteLog( sourcecsvfilepath, "Inside method csvCompareexcel \n \n");
            LogHelper.WriteLog( sourcecsvfilepath, "***********************");
            LogHelper.WriteLog( sourcecsvfilepath, "**********************");
            LogHelper.WriteLog( sourcecsvfilepath, "Reading CSV Data \n\n ");
            LogHelper.WriteLog( sourcecsvfilepath, "***********************");
            //Reading CSV
            //CSV 1 
            //     string FileLinesone = System.IO.File.ReadAllText(@sourcecsvfilepath + "\\" + sourcecsvfilename + ".csv");
            // Split into lines.
            //        FileLinesone = FileLinesone.Replace('\n', '\r');
            //   string[] linesone = FileLinesone.Split(new char[] { '\r' }, StringSplitOptions.RemoveEmptyEntries);
            string[] linesone = File.ReadAllLines(@sourcecsvfilepath + "\\" + sourcecsvfilename + ".csv");
            // See how many rows and columns there are.
            int num_rowsone = linesone.Length;
            int max_num_colsone = 0;
            int num_colsone = 0;
            for (int i = 0; i < num_rowsone; i++)
            {
                // LogHelper.WriteLog(driver, @sourcecsvfilepath, " Reading Row ->" + i + "in CSV  ");
                num_colsone = linesone[i].Split(',').Length;
                //   LogHelper.WriteLog(driver, @sourcecsvfilepath, " Number of columns in CSV  at Row  ->" + i + " are " + num_colsone);
                if (max_num_colsone < num_colsone)
                {
                    max_num_colsone = num_colsone;
                }
            }
            // Allocate the data array.
            string[,] al = new string[num_rowsone, max_num_colsone];
            List<string> al1 = new List<string>(num_rowsone * max_num_colsone);
            // Load the array.
            for (int r1 = sourcesrownum; r1 <= sourceerownum; r1++)
            {
                string[] line_r = linesone[r1].Split(',');
                for (int c1 = sourcescolnum; c1 <= sourceecolnum; c1++)
                {
                    al[r1, c1] = line_r[c1];
                    al1.Add(al[r1, c1]);
                    //    LogHelper.WriteLog(driver, csvfilepath, "  Data present in CSV at  Row [" + r + "] Column [ " + c + " ] has value -->[ " + al1[r, c] + "]");
                    LogHelper.WriteLog( @sourcecsvfilepath, "  Data present in CSV ->[" + sourcecsvfilename + " ] at  Row [ " + r1 + " ] Column [ " + c1 + " ] has value -->[ " + al[r1, c1] + " ]");
                }
            }
            //Reading EXCEL
            //*********************************************
            //     using aspose.cells
            //************************************************
            LogHelper.WriteLog( destinationfilepath, "**************************************");
            LogHelper.WriteLog( destinationfilepath, "Reading Excel Data \n \n");
            LogHelper.WriteLog( destinationfilepath, "********************************************");
            string s = null;
            Workbook wb = new Workbook(destinationfilepath + "\\" + destinationfilename);
            //   bool result = false;
            Worksheet worksheet = wb.Worksheets[SheetName];
            //    Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            LogHelper.WriteLog( destinationfilepath, "Reading Excel contents of file " + destinationfilepath + "\\" + destinationfilename + "   are :-");
            LogHelper.WriteLog( destinationfilepath, "**************************************");
            List<string> al2 = new List<string>();
            for (int i = destinationsrownum; i <= destinationerownum; i++)
            {
                for (int j = destinationscolnum; j <= destinationecolnum; j++)
                {
                    //       s = cells[i, j].StringValue.Trim();
                    s = Convert.ToString(cells[i, j].StringValue.Trim());
                    LogHelper.WriteLog( destinationfilepath, "Data Present in Excel [ " + destinationfilename + " ] at  Row Number ->[ " + i + " ] and Column Number[ " + j + " ] in Sheet [ " + SheetName + " ]  is :-->[ " + s + " ]");
                    al2.Add(s);
                }
            }
            LogHelper.WriteLog( destinationfilepath, "Data Read Successfully \n \n");
          //  --**********************************
            //Comparison of CSV & Excel  Starts
         //   --*********************************
            LogHelper.WriteLog( destinationfilepath, "**************************************");
            LogHelper.WriteLog( destinationfilepath, "Comparison of CSV & Excel  Starts \n \n");
            LogHelper.WriteLog( destinationfilepath, "*******************************************");
            List<string> al3 = new List<string>();
            string result1 = null;
            string result2 = null;
            string result3 = null;
            bool result = false;
            for (int i = 0; i < al1.Count; i++)
            {
                al3.Add(al1[i].Equals(al2[i]) ? "Pass" : "Fail");
                foreach (string cc in al3)
                {
                    result3 = cc;
                }
                LogHelper.WriteLog( destinationfilepath, "Data Comparison between CSV [" + sourcecsvfilename + " ] and Excel  [ " + destinationfilename + " ] for Data [ " + al1[i] + " ] is-- >  [ " + result3 + " ]");
            }
            if (al3.Contains("Fail"))
            {
                LogHelper.WriteLog( destinationfilepath, "********************************************\n\n");
                LogHelper.WriteLog( destinationfilepath, "Data Comparison between CSV [" + sourcecsvfilename + " ] and Excel  [ " + destinationfilename + " ]  is failed \n\n ");
                LogHelper.WriteLog( destinationfilepath, "********************************************\n\n");
                throw new Exception("Data Comparison between CSV [" + sourcecsvfilename + " ] and Excel  [ " + destinationfilename + " ]  is failed ");
                return false;
            }
            else
            {
                LogHelper.WriteLog( destinationfilepath, "********************************************\n\n");
                LogHelper.WriteLog( destinationfilepath, "Data Comparison between CSV [" + sourcecsvfilename + " ] and Excel  [ " + destinationfilename + " ]  is Passed \n\n ");
                LogHelper.WriteLog( destinationfilepath, "********************************************\n\n");
                return true;
            }
            return result;
        }
    }
}
