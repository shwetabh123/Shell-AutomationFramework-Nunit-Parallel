using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LumenWorks.Framework.IO.Csv;
using System.Configuration;
using Aspose.Slides;
using System.Collections;
using Test.Config;
namespace Test.Helpers
{
    public class DBHelper
    {
        public static string connectionString;

        public static string LogFileName = @"C:\Automation\AnalyzeAutomationLog.txt";
        public static string getSurveyURLSByDistribution(int SurveyID, string userName, string password, string serverName)
        {
            try
            {
                connectionString = "data source =" + serverName + ";Initial Catalog=TCES;User ID=" + userName +
                ";Password=" + password;
                string playerURL = null;
                using (SqlConnection sqlConn = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = sqlConn;
                    if (command.Connection.State == ConnectionState.Closed)
                        if (command.Connection.State == ConnectionState.Closed)
                            sqlConn.Open();
                    command.CommandText = @"SELECT TOP 1 CONCAT('http://qa-surveys.cebglobal.com/Pulse/Player/Start/',SurveyWaveID,'/1/',Guid) FROM SurveyWaveParticipant WHERE SurveyID = @SurveyID AND [Status] = 1 AND IsSystemGenerated = 0 AND SurveyStatusFlag = 0";
                    //command.CommandText = "SELECT TOP 1 SurveyWaveID FROM SurveyWaveParticipant WHERE SurveyID = @SurveyID AND [Status] = 1 AND IsSystemGenerated = 0 AND SurveyStatusFlag = 0";
                    command.Parameters.AddWithValue("@SurveyID", SurveyID);
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        playerURL = reader.GetString(0);
                    }
                    reader.Close();
                    sqlConn.Close();
                }
                return playerURL;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }






        public static void executesqlandwritetocsv(IWebDriver driver, string filepath, string filename, string InputSPText)
        {
            GenericHelper.WriteLog(driver, filepath, "**************************************");
            GenericHelper.WriteLog(driver, filepath, "Inside executesqlandwritetocsv \n \n");
            GenericHelper.WriteLog(driver, filepath, "********************************************");
            DataSet ds = new DataSet();
            object cell = null;
            string s = null;
            try
            {
                //common login
                //connectionString = "Data Source =" + DBConnectionParameters.TCESReportingserverName + ";Initial Catalog="
                //+ DBConnectionParameters.TCESReportingDB + ";User ID=" + DBConnectionParameters.userName + ";Password=" +
                //DBConnectionParameters.password;


       //         connectionString = "Server=" + ConfigReader.TCESReportingserverName() + ";Database=" +
       //     ConfigReader.TCESReportingDB() + ";User ID = " + ConfigReader.userName() +
       //";Password=" + ConfigReader.password() + ";";

                connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_OLTP_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";



                string playerURL = null;
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string queryString = InputSPText;
                    SqlCommand command = new SqlCommand(queryString, connection);
                    command.Connection = connection;
                    command.CommandTimeout = 3600;
                    command.CommandType = CommandType.Text;
                    command.CommandText = InputSPText;
                    Console.WriteLine("Executing the Input SP query");
                    GenericHelper.WriteLog(driver, filepath, "Executing the Input SP query");
                    GenericHelper.WriteLog(driver, filepath, "Connecting database ::" + connectionString);
                    GenericHelper.WriteLog(driver, filepath, "Connected database successfully...");
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(ds);
                    var outCsvFile = @filepath + "\\" + filename + ".csv";
                    foreach (DataTable dt in ds.Tables)
                    {
                        int countRow = dt.Rows.Count;
                        int countCol = dt.Columns.Count;
                        StringBuilder sb = new StringBuilder();
                        foreach (DataColumn col in dt.Columns)
                        {
                            sb.Append(col.ColumnName + ',');
                        }
                        sb.Remove(sb.Length - 1, 1);
                        sb.AppendLine();
                        foreach (DataRow row in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                sb.Append(row[i].ToString() + ",");
                            }
                            sb.AppendLine();
                        }
                        File.WriteAllText(outCsvFile, sb.ToString());
                        GenericHelper.WriteLog(driver, filepath, "Contents of sql csv are :-" + sb.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }



        public static void executesqlandwritetocsvmultipletables(IWebDriver driver, string filepath, string filename, string DBType, string InputSPText)
        {

            string logpath = ConfigReader.logFilePath;
            GenericHelper.WriteLog(driver, filepath, "**************************************");
            GenericHelper.WriteLog(driver, filepath, "Inside executesqlandwritetocsvmultipletables \n \n");
            GenericHelper.WriteLog(driver, filepath, "********************************************");
            DataSet ds = new DataSet();
            object cell = null;
            string s = null;
            try
            {
                //common login
                if (DBType.Equals("OLTP"))
                {

                    connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_OLTP_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";

                }
                else if(DBType.Equals("BATCH"))

                {
                    connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_BATCH_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";


                }

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string queryString = InputSPText;
                    SqlCommand command = new SqlCommand(queryString, connection);
                    command.Connection = connection;
                    command.CommandTimeout = 3600;
                    command.CommandType = CommandType.Text;
                    command.CommandText = InputSPText;
                    Console.WriteLine("Executing the Input SP query");
                    GenericHelper.WriteLog(driver, filepath, "Executing the Input SP query");
                    GenericHelper.WriteLog(driver, filepath, "**********************");
                    GenericHelper.WriteLog(driver, filepath, "Connecting database::\n\n" + connectionString);
                    GenericHelper.WriteLog(driver, filepath, "**********************");
                    GenericHelper.WriteLog(driver, filepath, "Connected database successfully...\n\n");
                    GenericHelper.WriteLog(driver, filepath, "**********************");
                    GenericHelper.WriteLog(driver, filepath, "Executing the SP query  :--\n\n" + InputSPText);
                    GenericHelper.WriteLog(driver, filepath, "**********************");
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(ds);
                    var outCsvFile = @filepath + "\\" + filename + ".csv";
                    StringBuilder sb = new StringBuilder();
                    foreach (DataTable dt in ds.Tables)
                    {
                        int countRow = dt.Rows.Count;
                        int countCol = dt.Columns.Count;
                        foreach (DataColumn col in dt.Columns)
                        {
                            sb.Append(col.ColumnName + ',');
                        }
                        sb.Remove(sb.Length - 1, 1);
                        sb.AppendLine();
                        foreach (DataRow row in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                sb.Append(row[i].ToString() + ",");

                                //sb.Append("'");
                                //if(sb.ToString().Contains(","))
                                //{
                                //    sb.Replace("'","");


                                //}

                                //sb.Replace("'", "");


                            }
                            sb.AppendLine();
                        }
                    }
                    File.WriteAllText(@filepath + "\\" + filename + ".csv", String.Empty);//  Remove all previous text before writing
                  
                    File.AppendAllText(outCsvFile, sb.ToString()); //Append all Tables ( Table 1... to Table n )data one below other

                    LogHelper.WriteLog(logpath, "Contents of csv file are :-"+sb.ToString());


                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                LogHelper.WriteLog(logpath, "Error is :-" + e.Message);

            }
        }

        //public static void executesqlandwritetocsvmultipletables(IWebDriver driver, string filepath, string filename, DBConnectionStringDTO DBConnectionParameters, string InputSPText)
        //{
        //    GenericHelper.WriteLog(driver, filepath, "**************************************");
        //    GenericHelper.WriteLog(driver, filepath, "Inside executesqlandwritetocsvmultipletables \n \n");
        //    GenericHelper.WriteLog(driver, filepath, "********************************************");
        //    DataSet ds = new DataSet();
        //    object cell = null;
        //    string s = null;
        //    try
        //    {
        //        //common login
        //        connectionString = "Data Source =" + DBConnectionParameters.TCESReportingserverName + ";Initial Catalog="
        //        + DBConnectionParameters.TCESReportingDB + ";User ID=" + DBConnectionParameters.userName + ";Password=" + DBConnectionParameters.password;



        //        string playerURL = null;
        //        using (SqlConnection connection = new SqlConnection(connectionString))
        //        {
        //            string queryString = InputSPText;
        //            SqlCommand command = new SqlCommand(queryString, connection);
        //            command.Connection = connection;
        //            command.CommandTimeout = 3600;
        //            command.CommandType = CommandType.Text;
        //            command.CommandText = InputSPText;
        //            Console.WriteLine("Executing the Input SP query");
        //            //        GenericHelper.WriteLog(driver, filepath, "Executing the Input SP query");
        //            GenericHelper.WriteLog(driver, filepath, "**********************");
        //            GenericHelper.WriteLog(driver, filepath, "Connecting database::\n\n" + connectionString);
        //            GenericHelper.WriteLog(driver, filepath, "**********************");
        //            GenericHelper.WriteLog(driver, filepath, "Connected database successfully...\n\n");
        //            GenericHelper.WriteLog(driver, filepath, "**********************");
        //            GenericHelper.WriteLog(driver, filepath, "Executing the SP query  :--\n\n" + InputSPText);
        //            GenericHelper.WriteLog(driver, filepath, "**********************");
        //            SqlDataAdapter adapter = new SqlDataAdapter(command);
        //            adapter.Fill(ds);
        //            var outCsvFile = @filepath + "\\" + filename + ".csv";
        //            StringBuilder sb = new StringBuilder();
        //            foreach (DataTable dt in ds.Tables)
        //            {
        //                int countRow = dt.Rows.Count;
        //                int countCol = dt.Columns.Count;
        //                foreach (DataColumn col in dt.Columns)
        //                {
        //                    sb.Append(col.ColumnName + ',');
        //                }
        //                sb.Remove(sb.Length - 1, 1);
        //                sb.AppendLine();
        //                foreach (DataRow row in dt.Rows)
        //                {
        //                    for (int i = 0; i < dt.Columns.Count; i++)
        //                    {
        //                        sb.Append(row[i].ToString() + ",");
        //                    }
        //                    sb.AppendLine();
        //                }
        //            }
        //            File.WriteAllText(@filepath + "\\" + filename + ".csv", String.Empty);//  Remove all previous text before writing
        //            File.AppendAllText(outCsvFile, sb.ToString()); //Append all Tables ( Table 1... to Table n )data one below other
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //    }
        //}
        public static bool verifyCSVColumnTextWithPPT(IWebDriver driver, string csvfilepath, string csvfilename, int srownum, int erownum, int scolnum, int ecolnum, string pptfilepath, string pptfilename, string slidenumber)
        {
            GenericHelper.WriteLog(driver, csvfilepath, "**********************");
            GenericHelper.WriteLog(driver, csvfilepath, "Inside verifyCSVColumnTextWithPPT \n \n");
            GenericHelper.WriteLog(driver, csvfilepath, "***********************");
            // Get the file's text.
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
            // Load the array.
            for (int r = srownum; r <= erownum; r++)
            {
                string[] line_r = lines[r].Split(',');
                for (int c = scolnum; c <= ecolnum; c++)
                {
                    al[r, c] = line_r[c];
                    al1.Add(al[r, c]);
                    GenericHelper.WriteLog(driver, csvfilepath, "  Data present in CSV at  Row [" + r + "] Column [ " + c + " ] has value -->[ " + al[r, c] + "]");
                }
            }
            //***************************Print PPT Data  using Aspose ****************************************
            bool result = false;
            string s = null;
            Aspose.Slides.License license = new Aspose.Slides.License();
            // license.SetLicense("C:\\Automation\\Aspose.Slides.lic");
            license.SetLicense("D:\\CWorkspace\\ShellWorkspace\\Aspose.Slides.lic");
            List<string> al2 = new List<string>();
            Presentation pres = new Presentation(pptfilepath + "\\" + pptfilename);
            int slideCount = pres.Slides.Count;
            //Get an Array of ITextFrame objects from all slides in the PPTX
            //ITextFrame[] textFramesPPT = Aspose.Slides.Util.SlideUtil.GetAllTextFrames(pres, true);
            int slidenumber1 = Int32.Parse(slidenumber);
            ISlide slide = pres.Slides[slidenumber1];
            //Get an Array of TextFrameEx objects from the first slide
            ITextFrame[] textFramesPPT = Aspose.Slides.Util.SlideUtil.GetAllTextBoxes(slide);
            //Loop through the Array of TextFrames
            for (int i = 0; i < textFramesPPT.Length; i++)
            {
                //Loop through paragraphs in current ITextFrame
                foreach (IParagraph para in textFramesPPT[i].Paragraphs)
                {
                    //Loop through portions in the current IParagraph
                    foreach (IPortion port in para.Portions)
                    {
                        //Display text in the current portion
                        Console.WriteLine(port.Text);
                        s = port.Text;
                        al2.Add(s);
                    }
                }
                foreach (string item in al2)
                {
                    GenericHelper.WriteLog(driver, pptfilepath, "  Data in slide number  ---> " + slidenumber + "    is -->" + item);
                }
            }
            //*********************************************
            //Comparing csv data with ppt data
            //Storing the comparison output in ArrayList<String>
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
                al3.Add(al2.Contains(item) ? "Pass" : "Fail");
                foreach (string cc in al3)
                {
                    result3 = cc;
                }
            }
            //}
            GenericHelper.WriteLog(driver, pptfilepath, "Data Comparison between Csv data and PPT data is -->  [" + result3 + " ]");
            if (al3.Contains("Pass"))
            {
                GenericHelper.WriteLog(driver, pptfilepath, "Data Comparison between csv and ppt is passed ");
                return true;
            }
            else
            {
                GenericHelper.WriteLog(driver, pptfilepath, "Data Comparison between csv and ppt is failed ");
                throw new Exception("Data Comparison between csv and ppt is failed ");
                return false;
            }
            return result;
        }
        public static bool DeleteCompanyDirectoryParticipant(string emailid)
        {
            try
            {
                using (SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataModel"].ToString()))
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = sqlConn;
                    if (command.Connection.State == ConnectionState.Closed)
                        if (command.Connection.State == ConnectionState.Closed)
                            sqlConn.Open();
                    command.CommandText = @"update CompanyDirectoryParticipant set [status]=0 where emailaddress = '" + emailid + "'";
                    //command.CommandText = "SELECT TOP 1 SurveyWaveID FROM SurveyWaveParticipant WHERE SurveyID = @SurveyID AND [Status] = 1 AND IsSystemGenerated = 0 AND SurveyStatusFlag = 0";
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        string data = reader.GetString(0);
                    }
                    reader.Close();
                    sqlConn.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public static int GetCompanyDirectoryParticipant(string emailid)
        {
            int data = 0;
            try
            {
                using (SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataModel"].ToString()))
                {
                    SqlCommand command = new SqlCommand();
                    command.Connection = sqlConn;
                    if (command.Connection.State == ConnectionState.Closed)
                        if (command.Connection.State == ConnectionState.Closed)
                            sqlConn.Open();
                    command.CommandText = @"select count(*) from CompanyDirectoryParticipant where [status]= 1 and  emailaddress = '" + emailid + "'";
                    //command.CommandText = "SELECT TOP 1 SurveyWaveID FROM SurveyWaveParticipant WHERE SurveyID = @SurveyID AND [Status] = 1 AND IsSystemGenerated = 0 AND SurveyStatusFlag = 0";
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        data = reader.GetInt32(0);
                    }
                    reader.Close();
                    sqlConn.Close();
                    return data;
                }
            }
            catch (Exception ex)
            {
                return data;
            }
        }
        public static DataSet ExecuteSPwithInputParameters(string InputSPText)
        {
            try
            {
                DataSet ds = new DataSet();
                //  connectionString = "data source =" + DBConnectionParameters.TCESReportingserverName + ";Initial Catalog="
                //    + DBConnectionParameters.TCESReportingDB + ";User ID=" + DBConnectionParameters.userName +
                //";Password=" + DBConnectionParameters.password;
                // connectionString = "Server=" + DBConnectionParameters.TCESReportingserverName + ";Database=" +
                //     DBConnectionParameters.TCESReportingDB + ";User ID = " + DBConnectionParameters.userName +
                //";Password=" + DBConnectionParameters.password + ";";


       //         connectionString = "Server=" + ConfigReader.TCESReportingserverName() + ";Database=" +
       //     ConfigReader.TCESReportingDB() + ";User ID = " + ConfigReader.userName() +
       //";Password=" + ConfigReader.password() + ";";

                connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_OLTP_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";


                using (SqlConnection sqlConn = new SqlConnection(connectionString))
                {
                    if (sqlConn.State == ConnectionState.Closed)
                    {
                        sqlConn.Open();
                        Console.WriteLine("Connection is Opened");
                    }
                    SqlCommand command = new SqlCommand(InputSPText, sqlConn);
                    command.Connection = sqlConn;
                    command.CommandTimeout = 3600;
                    command.CommandType = CommandType.Text;
                    command.CommandText = InputSPText;
                    Console.WriteLine("Executing the Input SP query2");
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    command.CommandTimeout = 3600;
                    adapter.Fill(ds);
                    Console.WriteLine("SP execution is completed");
                    sqlConn.Close();
                }
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                using (StreamWriter w = File.AppendText(LogFileName))
                {
                    // ReportDataValidation.Log("ERROR : " + e.Message, w);
                    // ReportDataValidation.Log("Input Query : " + InputSPText, w);
                }
                return null;
            }
        }
        public static DataSet ExecuteSPwithInputParameters_TCES(string InputSPText, DBConnectionStringDTO DBConnectionParameters)
        {
            try
            {
                DataSet ds = new DataSet();
                //  connectionString = "data source =" + DBConnectionParameters.TCESReportingserverName + ";Initial Catalog="
                //    + DBConnectionParameters.TCESReportingDB + ";User ID=" + DBConnectionParameters.userName +
                //";Password=" + DBConnectionParameters.password;

                //connectionString = "Server=" + DBConnectionParameters.TCESserverName + ";Database=" + DBConnectionParameters.TCESDB + ";User ID = " + DBConnectionParameters.userName +
                //     ";Password=" + DBConnectionParameters.password + ";";

                connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_OLTP_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";


                using (SqlConnection sqlConn = new SqlConnection(connectionString))
                {
                    if (sqlConn.State == ConnectionState.Closed)
                    {
                        sqlConn.Open();
                        Console.WriteLine("Connection is Opened");
                    }
                    SqlCommand command = new SqlCommand(InputSPText, sqlConn);
                    command.Connection = sqlConn;
                    command.CommandTimeout = 3600;
                    command.CommandType = CommandType.Text;
                    command.CommandText = InputSPText;
                    Console.WriteLine("Executing the Input SP query");
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(ds);
                    Console.WriteLine("SP execution is completed");
                    sqlConn.Close();
                }
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static DataSet ExecuteTestSP(SqlConnection sqlConn, string SPQuery, int SurveyID, string DemoQuery, string TimePeriodQuery)
        {
            try
            {
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand(SPQuery, sqlConn);
                command.Connection = sqlConn;
                command.CommandTimeout = 3600;
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter _SurveyID = command.Parameters.Add("@SurveyID", SqlDbType.BigInt);
                _SurveyID.Direction = ParameterDirection.Input;
                _SurveyID.Value = SurveyID;
                SqlParameter _DemoQuery = command.Parameters.Add("@DemoQuery", SqlDbType.NVarChar);
                _DemoQuery.Direction = ParameterDirection.Input;
                _DemoQuery.Value = DemoQuery;
                SqlParameter _TimePeriodQuery = command.Parameters.Add("@TImePeriodQuery", SqlDbType.NVarChar);
                _TimePeriodQuery.Direction = ParameterDirection.Input;
                _TimePeriodQuery.Value = TimePeriodQuery;
                Console.WriteLine("Test SP - Executing the Command " + command.CommandText);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                Console.WriteLine("Test SP - Command execution completed");
                foreach (System.Data.DataTable table in ds.Tables)
                {
                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            Console.Write(table.Rows[j].ItemArray[k].ToString() + " ");
                        }
                        Console.WriteLine("");
                    }
                }
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static DataSet ExecuteTestSP_MultipleDistribution(SqlConnection sqlConn, string SPQuery, string SurveyIDs, string DemoQuery, string TimePeriodQuery)
        {
            try
            {
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand(SPQuery, sqlConn);
                command.Connection = sqlConn;
                command.CommandTimeout = 3600;
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter _SurveyID = command.Parameters.Add("@SurveyIDs", SqlDbType.NVarChar);
                _SurveyID.Direction = ParameterDirection.Input;
                _SurveyID.Value = SurveyIDs;
                SqlParameter _DemoQuery = command.Parameters.Add("@DemoQuery", SqlDbType.NVarChar);
                _DemoQuery.Direction = ParameterDirection.Input;
                _DemoQuery.Value = DemoQuery;
                SqlParameter _TimePeriodQuery = command.Parameters.Add("@TImePeriodQuery", SqlDbType.NVarChar);
                _TimePeriodQuery.Direction = ParameterDirection.Input;
                _TimePeriodQuery.Value = TimePeriodQuery;
                Console.WriteLine("Test SP - Executing the Command " + command.CommandText);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                Console.WriteLine("Test SP - Command execution completed");
                foreach (System.Data.DataTable table in ds.Tables)
                {
                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            Console.Write(table.Rows[j].ItemArray[k].ToString() + " ");
                        }
                        Console.WriteLine("");
                    }
                }
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static DataSet ExecuteTrendSP(SqlConnection sqlConn, string Query)
        {
            try
            {
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand(Query, sqlConn);
                command.Connection = sqlConn;
                command.CommandTimeout = 3600;
                command.CommandType = CommandType.Text;
                command.CommandText = Query;
                Console.WriteLine("Trend SP - Executing the Command");
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                Console.WriteLine("Trend SP - Command execution completed");
                //foreach (System.Data.DataTable table in ds.Tables)
                //{
                //    for (int j = 0; j < table.Rows.Count; j++)
                //    {
                //        for (int k = 0; k < table.Columns.Count; k++)
                //        {
                //            Console.Write(table.Rows[j].ItemArray[k].ToString() + " ");
                //        }
                //        Console.WriteLine("");
                //    }
                //}
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static DataSet ExecuteTestSP_NR(SqlConnection sqlConn, string SPQuery, int SurveyID, int SurveyFormItemID, string DemoQuery, string TimePeriodQuery)
        {
            try
            {
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand(SPQuery, sqlConn);
                command.Connection = sqlConn;
                command.CommandTimeout = 3600;
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter _SurveyID = command.Parameters.Add("@SurveyID", SqlDbType.BigInt);
                _SurveyID.Direction = ParameterDirection.Input;
                _SurveyID.Value = SurveyID;
                SqlParameter _SurveyFormItemID = command.Parameters.Add("@SurveyFormItemID", SqlDbType.BigInt);
                _SurveyFormItemID.Direction = ParameterDirection.Input;
                _SurveyFormItemID.Value = SurveyFormItemID;
                SqlParameter _DemoQuery = command.Parameters.Add("@DemoQuery", SqlDbType.NVarChar);
                _DemoQuery.Direction = ParameterDirection.Input;
                _DemoQuery.Value = DemoQuery;
                SqlParameter _TimePeriodQuery = command.Parameters.Add("@TImePeriodQuery", SqlDbType.NVarChar);
                _TimePeriodQuery.Direction = ParameterDirection.Input;
                _TimePeriodQuery.Value = TimePeriodQuery;
                Console.WriteLine("Test SP - Executing the Command " + command.CommandText);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                Console.WriteLine("Test SP - Command execution completed");
                //foreach (System.Data.DataTable table in ds.Tables)
                //{
                //    for (int j = 0; j < table.Rows.Count; j++)
                //    {
                //        for (int k = 0; k < table.Columns.Count; k++)
                //        {
                //            Console.Write(table.Rows[j].ItemArray[k].ToString() + " ");
                //        }
                //        Console.WriteLine("");
                //    }
                //}
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static DataSet ExecuteTestSP_NR_MultipleDistribution(SqlConnection sqlConn, string SPQuery, string SurveyIDs, int SurveyFormItemID, string DemoQuery, string TimePeriodQuery)
        {
            try
            {
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand(SPQuery, sqlConn);
                command.Connection = sqlConn;
                command.CommandTimeout = 3600;
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter _SurveyID = command.Parameters.Add("@SurveyIDs", SqlDbType.NVarChar);
                _SurveyID.Direction = ParameterDirection.Input;
                _SurveyID.Value = SurveyIDs;
                SqlParameter _SurveyFormItemID = command.Parameters.Add("@SurveyFormItemID", SqlDbType.BigInt);
                _SurveyFormItemID.Direction = ParameterDirection.Input;
                _SurveyFormItemID.Value = SurveyFormItemID;
                SqlParameter _DemoQuery = command.Parameters.Add("@DemoQuery", SqlDbType.NVarChar);
                _DemoQuery.Direction = ParameterDirection.Input;
                _DemoQuery.Value = DemoQuery;
                SqlParameter _TimePeriodQuery = command.Parameters.Add("@TImePeriodQuery", SqlDbType.NVarChar);
                _TimePeriodQuery.Direction = ParameterDirection.Input;
                _TimePeriodQuery.Value = TimePeriodQuery;
                Console.WriteLine("Test SP - Executing the Command " + command.CommandText);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                Console.WriteLine("Test SP - Command execution completed");
                foreach (System.Data.DataTable table in ds.Tables)
                {
                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            Console.Write(table.Rows[j].ItemArray[k].ToString() + " ");
                        }
                        Console.WriteLine("");
                    }
                }
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static DataSet ExecuteTestHeatMapSP(SqlConnection sqlConn, string SPName, string SPQuery, int SurveyID, string GroupBy)
        {
            try
            {
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand(SPName, sqlConn);
                command.Connection = sqlConn;
                command.CommandTimeout = 3600;
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter _SurveyID = command.Parameters.Add("@SurveyID", SqlDbType.BigInt);
                _SurveyID.Direction = ParameterDirection.Input;
                _SurveyID.Value = SurveyID;
                SqlParameter _DemoQuery = command.Parameters.Add("@SPQuery", SqlDbType.NVarChar);
                _DemoQuery.Direction = ParameterDirection.Input;
                _DemoQuery.Value = SPQuery;
                SqlParameter _TimePeriodQuery = command.Parameters.Add("@GroupBy", SqlDbType.NVarChar);
                _TimePeriodQuery.Direction = ParameterDirection.Input;
                _TimePeriodQuery.Value = GroupBy;
                Console.WriteLine("Test SP - Executing the Command " + command.CommandText);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                Console.WriteLine("Test SP - Command execution completed");
                //foreach (System.Data.DataTable table in ds.Tables)
                //{
                //    for (int j = 0; j < table.Rows.Count; j++)
                //    {
                //        for (int k = 0; k < table.Columns.Count; k++)
                //        {
                //            Console.Write(table.Rows[j].ItemArray[k].ToString() + " ");
                //        }
                //        Console.WriteLine("");
                //    }
                //}
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static DataSet ExecuteTestHeatMapSP_NR(SqlConnection sqlConn, string SPName, string SPQuery, int SurveyID, string GroupBy, string SurveyFormItemId)
        {
            try
            {
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand(SPName, sqlConn);
                command.Connection = sqlConn;
                command.CommandTimeout = 3600;
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter _SurveyID = command.Parameters.Add("@SurveyID", SqlDbType.BigInt);
                _SurveyID.Direction = ParameterDirection.Input;
                _SurveyID.Value = SurveyID;
                SqlParameter _DemoQuery = command.Parameters.Add("@SPQuery", SqlDbType.NVarChar);
                _DemoQuery.Direction = ParameterDirection.Input;
                _DemoQuery.Value = SPQuery;
                SqlParameter _TimePeriodQuery = command.Parameters.Add("@GroupBy", SqlDbType.NVarChar);
                _TimePeriodQuery.Direction = ParameterDirection.Input;
                _TimePeriodQuery.Value = GroupBy;
                SqlParameter _SurveyFormItemId = command.Parameters.Add("@SurveyFormItemId", SqlDbType.BigInt);
                _SurveyFormItemId.Direction = ParameterDirection.Input;
                _SurveyFormItemId.Value = Int32.Parse(SurveyFormItemId);
                Console.WriteLine("Test SP - Executing the Command " + command.CommandText);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);
                Console.WriteLine("Test SP - Command execution completed");
                //foreach (System.Data.DataTable table in ds.Tables)
                //{
                //    for (int j = 0; j < table.Rows.Count; j++)
                //    {
                //        for (int k = 0; k < table.Columns.Count; k++)
                //        {
                //            Console.Write(table.Rows[j].ItemArray[k].ToString() + " ");
                //        }
                //        Console.WriteLine("");
                //    }
                //}
                return ds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static bool verifyTextnotpresentinPowerPoint(IWebDriver driver, string filepath, string filename, string slidenumber, string texttocompare)
        {
            bool result = false;
            //**************************************Interop ************************************
            //  string presentation_text = "";
            //  Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            //Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            //  Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(@"C:\Automation\SpotlightReport3627237636684590143688007.ppt", Microsoft.Office.Core.MsoTriState.msoTrue,
            //       Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);
            //  PowerPoint.Slides pptSlides = presentation.Slides;
            //      if (pptSlides != null)
            //      {
            //          var slidesCount = pptSlides.Count;
            //          StringBuilder sb = new StringBuilder();
            //          if (slidesCount > 0)
            //          {
            //              for (int slideIndex = 1; slideIndex <= slidesCount; slideIndex++)
            //              {
            //                  var slide = pptSlides[slideIndex];
            //                  foreach (PowerPoint.Shape textShape in slide.Shapes)
            //                  {
            //                      if (textShape.HasTextFrame == MsoTriState.msoTrue &&
            //                             textShape.TextFrame.HasText == MsoTriState.msoTrue)
            //                      {
            //                          PowerPoint.TextRange pptTextRange = textShape.TextFrame.TextRange;
            //                          if (pptTextRange != null && pptTextRange.Length > 0)
            //                          {
            //                              //     sb = stringBuilder.Append(" " + pptTextRange.Text);
            //                              var textRange = textShape.TextFrame.TextRange;
            //                              var text = textRange.Text;
            //                              presentation_text = text + " ";
            //                          }
            //                      }
            //                  }
            //              }
            //          }
            //      }
            //***************************Aspose ****************************************
            string s = null;
            Aspose.Slides.License license = new Aspose.Slides.License();
            license.SetLicense("D:\\CWorkspace\\ShellWorkspace\\Aspose.Slides.lic");
            ArrayList al2 = new ArrayList();
            //Load the desired the presentation
            //            Presentation pres = new Presentation(@"C:\Automation\SpotlightReport3627237636684590143688007.ppt");
            //           Presentation pres = new Presentation(filepath+"\\SpotlightReport3627237636684590143688007.ppt");
            Presentation pres = new Presentation(filepath + "\\" + filename);
            //using (Presentation prestg = new Presentation(@"C:\Automation\SpotlightReport3627237636684590143688007.ppt"))
            //{
            int slideCount = pres.Slides.Count;
            //Get an Array of ITextFrame objects from all slides in the PPTX
            //ITextFrame[] textFramesPPT = Aspose.Slides.Util.SlideUtil.GetAllTextFrames(pres, true);
            int slidenumber1 = Int32.Parse(slidenumber);
            ISlide slide = pres.Slides[slidenumber1];
            //Get an Array of TextFrameEx objects from the first slide
            ITextFrame[] textFramesPPT = Aspose.Slides.Util.SlideUtil.GetAllTextBoxes(slide);
            //Loop through the Array of TextFrames
            for (int i = 0; i < textFramesPPT.Length; i++)
            {
                //Loop through paragraphs in current ITextFrame
                foreach (IParagraph para in textFramesPPT[i].Paragraphs)
                {
                    //Loop through portions in the current IParagraph
                    foreach (IPortion port in para.Portions)
                    {
                        //Display text in the current portion
                        Console.WriteLine(port.Text);
                        s = port.Text;
                        //        GenericHelper.WriteLog(driver, filepath, s);
                        al2.Add(s);
                    }
                }
                foreach (var item in al2)
                {
                    GenericHelper.WriteLog(driver, filepath, "  Data in slide number  ---> " + slidenumber + "    is -->" + item);
                }
            }
            if (!(al2.Contains(texttocompare)))
            {
                GenericHelper.WriteLog(driver, filepath, "Data -->" + texttocompare + " -->is not  present in ppt !! Test Case Passed ");
                return true;
            }
            else
            {
                GenericHelper.WriteLog(driver, filepath, "Data -->" + texttocompare + " -->is  present in ppt !! Test Case Failed ");
                throw new Exception("Data -->" + texttocompare + " -->is   present in ppt ");
                return false;
            }
            return result;
        }
        public static bool verifyTextpresentinPowerPoint(IWebDriver driver, string filepath, string filename, string slidenumber, string texttocompare)
        {
            bool result = false;
            //**************************************Interop ************************************
            //  string presentation_text = "";
            //  Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            //Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            //  Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(@"C:\Automation\SpotlightReport3627237636684590143688007.ppt", Microsoft.Office.Core.MsoTriState.msoTrue,
            //       Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);
            //  PowerPoint.Slides pptSlides = presentation.Slides;
            //      if (pptSlides != null)
            //      {
            //          var slidesCount = pptSlides.Count;
            //          StringBuilder sb = new StringBuilder();
            //          if (slidesCount > 0)
            //          {
            //              for (int slideIndex = 1; slideIndex <= slidesCount; slideIndex++)
            //              {
            //                  var slide = pptSlides[slideIndex];
            //                  foreach (PowerPoint.Shape textShape in slide.Shapes)
            //                  {
            //                      if (textShape.HasTextFrame == MsoTriState.msoTrue &&
            //                             textShape.TextFrame.HasText == MsoTriState.msoTrue)
            //                      {
            //                          PowerPoint.TextRange pptTextRange = textShape.TextFrame.TextRange;
            //                          if (pptTextRange != null && pptTextRange.Length > 0)
            //                          {
            //                              //     sb = stringBuilder.Append(" " + pptTextRange.Text);
            //                              var textRange = textShape.TextFrame.TextRange;
            //                              var text = textRange.Text;
            //                              presentation_text = text + " ";
            //                          }
            //                      }
            //                  }
            //              }
            //          }
            //      }
            //***************************Aspose ****************************************
            string s = null;
            Aspose.Slides.License license = new Aspose.Slides.License();
            //   license.SetLicense("C:\\Automation\\Aspose.Slides.lic");
            license.SetLicense("D:\\CWorkspace\\ShellWorkspace\\Aspose.Slides.lic");
            ArrayList al2 = new ArrayList();
            //Load the desired the presentation
            //            Presentation pres = new Presentation(@"C:\Automation\SpotlightReport3627237636684590143688007.ppt");
            //           Presentation pres = new Presentation(filepath+"\\SpotlightReport3627237636684590143688007.ppt");
            Presentation pres = new Presentation(filepath + "\\" + filename);
            //using (Presentation prestg = new Presentation(@"C:\Automation\SpotlightReport3627237636684590143688007.ppt"))
            //{
            int slideCount = pres.Slides.Count;
            //Get an Array of ITextFrame objects from all slides in the PPTX
            //ITextFrame[] textFramesPPT = Aspose.Slides.Util.SlideUtil.GetAllTextFrames(pres, true);
            int slidenumber1 = Int32.Parse(slidenumber);
            ISlide slide = pres.Slides[slidenumber1];
            //Get an Array of TextFrameEx objects from the first slide
            ITextFrame[] textFramesPPT = Aspose.Slides.Util.SlideUtil.GetAllTextBoxes(slide);
            //Loop through the Array of TextFrames
            for (int i = 0; i < textFramesPPT.Length; i++)
            {
                //Loop through paragraphs in current ITextFrame
                foreach (IParagraph para in textFramesPPT[i].Paragraphs)
                {
                    //Loop through portions in the current IParagraph
                    foreach (IPortion port in para.Portions)
                    {
                        //Display text in the current portion
                        Console.WriteLine(port.Text);
                        s = port.Text;
                        //        GenericHelper.WriteLog(driver, filepath, s);
                        al2.Add(s);
                    }
                }
                foreach (var item in al2)
                {
                    GenericHelper.WriteLog(driver, filepath, "  Data in slide number  ---> " + slidenumber + "    is -->" + item);
                }
            }
            if (al2.Contains(texttocompare))
            {
                GenericHelper.WriteLog(driver, filepath, "Data -->" + texttocompare + " -->is present in ppt !! Test Case Passed ");
                return true;
            }
            else
            {
                GenericHelper.WriteLog(driver, filepath, "Data -->" + texttocompare + " -->is  not  present in ppt !! Test Case Failed ");
                throw new Exception("Data -->" + texttocompare + " -->is  not  present in ppt ");
                return false;
            }
            return result;
        }
        public static List<int> GetUserAnalysisIds(int userId, int surveyFormId, DBConnectionStringDTO DBConnectionParameters)
        {
            //connectionString = "Server=" + DBConnectionParameters.TCESReportingserverName + ";Database=" + DBConnectionParameters.TCESReportingDB + ";User ID = " + DBConnectionParameters.userName +
            //        ";Password=" + DBConnectionParameters.password + ";";

            connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_OLTP_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";


            List<int> ids = new List<int>();
            string query = "SELECT UserAnalyzeFilterID FROM UserAnalyzeFilter WHERE UserID=" + userId + " AND SurveyFormID=" + surveyFormId;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        ids.Add(Convert.ToInt32(reader["UserAnalyzeFilterID"]));
                    }
                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }
            }
            return ids;
        }
        public static void DeleteUserAnalysisRecords(int userAnalyseFilterId, DBConnectionStringDTO DBConnectionParameters)
        
            
         {
        //    connectionString = "Server=" + DBConnectionParameters.TCESReportingserverName + ";Database=" + DBConnectionParameters.TCESReportingDB + ";User ID = " + DBConnectionParameters.userName +
        //           ";Password=" + DBConnectionParameters.password + ";";

            connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_OLTP_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";



            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmdCommand = new SqlCommand();
            cmdCommand.CommandText = @"DELETE FROM UserAnalyzeFilter WHERE UserAnalyzeFilterID =" + userAnalyseFilterId
                                     + "; DELETE FROM DemoFilter WHERE UserAnalyzeFilterID =" + userAnalyseFilterId
                                     + "; DELETE FROM SurveyFilter WHERE UserAnalyzeFilterID =" + userAnalyseFilterId
                                     + "; DELETE FROM DemoRangeFilter WHERE UserAnalyzeFilterID =" + userAnalyseFilterId +
                                     ";";
            cmdCommand.CommandType = CommandType.Text;
            cmdCommand.Connection = conn;
            try
            {
                conn.Open();
                cmdCommand.ExecuteNonQuery();
            }
            finally
            {
                cmdCommand.Dispose();
                conn.Dispose();
            }
        }
        public static List<string> GetFavName(DBConnectionStringDTO DBConnectionParameters)
        {
            //connectionString = "Server=" + DBConnectionParameters.TCESReportingserverName + ";Database=" + DBConnectionParameters.TCESReportingDB + ";User ID = " + DBConnectionParameters.userName +
            //        ";Password=" + DBConnectionParameters.password + ";";

            connectionString = "Server =" + ConfigReader.DBServerName + ";Database=" + ConfigReader.PH_OLTP_DB + ";IntegratedSecurity=SSPI ;PersistSecurityInfo=False";



            List<string> ids = new List<string>();
            string query = "select distinct favouritename  from useranalyzefilter where clientid = 37400 and favouritename is not null";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        ids.Add(Convert.ToString(reader["favouritename"]));
                    }
                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }
            }
            return ids;
        }
    }
}