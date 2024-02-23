using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using ClosedXML.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using FullVersion.Models;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.Globalization;

namespace FullVersion.Controllers
{
    public class CostController : Controller
    {
        //
        // GET: /Cost//////

        //string path = Server.MapPath("~/FileValidation/");
        //IEnumerable<string> files1 = Directory.GetFiles(path);

        //foreach (string filedetails in files1)
        //{
        //}

        public static SqlConnection con = new SqlConnection(@"Data Source=nebula\mssqlserver1;Initial Catalog=EAS_PROD;uid=nebula_sql;pwd=python@123");


        [HttpGet]
        public ActionResult Index()
        {
            var list = SL("#menu5");
            

            return View(list);
        }


        [HttpPost]
        public ActionResult Index(string submit, List<SL_List> list, string finyear,string Type,string LoadType, HttpPostedFileBase uploadFile,
            string isFocus)
        {
            string tab = "";
            switch (submit)
            {
                case "MI Quick Summary Generate":
                    //SummarySheetGeneration(finyear);
                    SummarySheetGeneration_SecondHalf(finyear);
                    tab = "#menu5";
                    list =SL(tab);
                    break;
                case "MI Generate":

                    Cost_Opt_Generation(list, finyear, isFocus);
                    list = SL("#menu2");
                    break;
                case "Upload":
                    Summary_Upload(out list, finyear, Type, LoadType, uploadFile, out tab);
                    break;
                case "Submit":
                    //Data_Generation();
                    list = SL("#menu3");
                    break;
                case "Download Template":
                    //list = SL("#menu5");
                    if (Type == "Attrition") { 
                    string filename = "Attrition_Data_template.xlsx";
                    var path = Server.MapPath("~/Upload Template/" + filename);
                    var stream = new FileStream(path, FileMode.Open);
                    return File(stream, System.Net.Mime.MediaTypeNames.Application.Octet, filename);
                    }
                    else if (Type == "Talent Mobility")
                    {
                        string filename = "Talent Mobility_template.xlsx";
                        var path = Server.MapPath("~/Upload Template/" + filename);
                        var stream = new FileStream(path, FileMode.Open);
                        return File(stream, System.Net.Mime.MediaTypeNames.Application.Octet, filename);
                    }
                    break;
                default:
                    break;
            }

            if (submit == "MI Quick Summary Generate")
            {
                string filename = "EAS_MarginImprovement_QuickSummary_Report_"+ DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx";
                var path = Server.MapPath("~/Summary/" + filename);
                var stream = new FileStream(path, FileMode.Open);
                return File(stream, System.Net.Mime.MediaTypeNames.Application.Octet, filename);
            }
            else { 
            return View(list);
            }

           
        }

        private void Summary_Upload(out List<SL_List> list, string finyear, string Type, string LoadType, HttpPostedFileBase uploadFile, out string tab)
        {
            if (uploadFile != null && uploadFile.ContentLength > 0)
                try
                {
                    string date = DateTime.Now.ToString("dd/mmm/yyyy");
                    string path = Path.Combine(Server.MapPath("~/SummaryFiles"),
                                           Path.GetFileName(uploadFile.FileName));
                    uploadFile.SaveAs(path);

                    bool Issheetvalid = true;
                    bool IsValidFileName = true;

                    string FileName = uploadFile.FileName;

                    if (!FileName.Contains(Type)) { IsValidFileName = false; }

                    if (IsValidFileName)
                    {
                        string conString = GetConnectionstring(path);
                        DataTable dt = new DataTable();

                        Issheetvalid = OLEDB(conString, dt, Type);

                        if (Issheetvalid)
                            if (dt.Rows.Count > 0)
                            {
                                if (Type == "Attrition") { Upload_Attrition(dt, Type, LoadType, finyear); }
                                else if (Type == "Talent Mobility") { Upload_TalentMobility(dt, Type, LoadType, finyear); }
                            }
                            else { ViewBag.Message = "No records found in excel file..Please check.."; }
                        else { ViewBag.Message = "Please check the sheet name..Sheet name should be DATA"; }
                    }
                    else { ViewBag.Message = "Please check the file name.."; }

                }
                catch (Exception ex)
                {
                    ViewBag.Message = "ERROR:" + ex.Message.ToString();
                }
            else
            {
                //ViewBag.Message = "You have not specified a file.";
            }

            tab = "#menu5";
            list = SL(tab);
        }

        private string GetConnectionstring(string filePath)
        {

            string extension = Path.GetExtension(filePath);
            string conString = string.Empty;


            switch (extension)
            {
                case ".xls": //Excel 97-03 
                    conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"";
                    break;
                case ".xlsx": //Excel 07 or higher                    
                    conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1;Connect Timeout=0;\"";
                    break;
            }
            return conString;


        }

        private bool OLEDB(string conString, DataTable dt, string Type)
        {
            bool IsValidSheet = true;
            using (OleDbConnection connExcel = new OleDbConnection(conString))
            {
                using (OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {
                        cmdExcel.Connection = connExcel;

                        //Get the name of First Sheet.
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);


                        string[] sheet_list = dtExcelSchema.Rows.Cast<DataRow>().Select(x => x.Field<string>("TABLE_NAME")).ToArray();

                        sheet_list = sheet_list.Select(s => s.Replace("'", "")).ToArray();

                        string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();

                        if (Type == "Talent Mobility")
                        {
                            sheetName = "Data$"; if (!sheet_list.Contains(sheetName)) { IsValidSheet = false; }
                        }
                        else if (Type == "Attrition")
                        {
                            sheetName = "Data$"; if (!sheet_list.Contains(sheetName)) { IsValidSheet = false; }
                        }
                        


                        connExcel.Close();

                        if (IsValidSheet)
                        {
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                        }
                    }
                }
            }

            return IsValidSheet;
        }

        public void Upload_Attrition(DataTable dataTable, string Type,string LoadType, string finyear)
        {
            DataColumn dc = new DataColumn("SnapShotDate");
            dc.DataType = typeof(DateTime);
            dc.DefaultValue = System.DateTime.Now;
            dataTable.Columns.Add(dc);

            using (var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString))
            {
                SqlTransaction transaction = null;
                connection.Open();
                SqlCommand cmd = new SqlCommand("delete from tbl_auto_Attrition_summary where SnapShotDate='"+ System.DateTime.Now + "'", connection);
                cmd.ExecuteNonQuery();
                try
                {
                    transaction = connection.BeginTransaction();
                    using (var sqlBulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.TableLock, transaction))
                    {
                        sqlBulkCopy.DestinationTableName = "tbl_auto_Attrition_summary";
                       
                        sqlBulkCopy.ColumnMappings.Add("SL", "SL");
                        sqlBulkCopy.ColumnMappings.Add("Dimension", "Dimension");
                        sqlBulkCopy.ColumnMappings.Add("Type", "Type");
                        sqlBulkCopy.ColumnMappings.Add("YTD", "YTD");
                        sqlBulkCopy.ColumnMappings.Add("Q1", "Q1");
                        sqlBulkCopy.ColumnMappings.Add("Q2", "Q2");
                        sqlBulkCopy.ColumnMappings.Add("Q3", "Q3");
                        sqlBulkCopy.ColumnMappings.Add("Q4", "Q4");
                        sqlBulkCopy.ColumnMappings.Add("Apr", "M1");
                        sqlBulkCopy.ColumnMappings.Add("May", "M2");
                        sqlBulkCopy.ColumnMappings.Add("Jun", "M3");
                        sqlBulkCopy.ColumnMappings.Add("Jul", "M4");
                        sqlBulkCopy.ColumnMappings.Add("Aug", "M5");
                        sqlBulkCopy.ColumnMappings.Add("Sep", "M6");
                        sqlBulkCopy.ColumnMappings.Add("Oct", "M7");
                        sqlBulkCopy.ColumnMappings.Add("Nov", "M8");
                        sqlBulkCopy.ColumnMappings.Add("Dec", "M9");
                        sqlBulkCopy.ColumnMappings.Add("Jan", "M10");
                        sqlBulkCopy.ColumnMappings.Add("Feb", "M11");
                        sqlBulkCopy.ColumnMappings.Add("Mar", "M12");
                        sqlBulkCopy.ColumnMappings.Add("H1", "H1");
                        sqlBulkCopy.ColumnMappings.Add("H2", "H2");
                        sqlBulkCopy.ColumnMappings.Add("SnapShotDate", "snapshotdate");

                        sqlBulkCopy.WriteToServer(dataTable);
                    }
                    transaction.Commit();


                    cmd = new SqlCommand("sp_auto_Batch_Attrition_summary", connection);
                    cmd.Parameters.AddWithValue("@LoadType", LoadType);
                    cmd.Parameters.AddWithValue("@finyear", finyear);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
                    connection.Close();

                    ViewBag.Message= "Data Uploaded Successfully...";


                }
                catch (Exception ex)
                {
                    ViewBag.Message = "ERROR:" + ex.Message.ToString();
                }

            }
        }

        public void Upload_TalentMobility(DataTable dataTable, string Type,string LoadType,string finyear)
        {
            DataColumn dc = new DataColumn("SnapShotDate");
            dc.DataType = typeof(DateTime);
            dc.DefaultValue = System.DateTime.Now;
            dataTable.Columns.Add(dc);

            using (var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString))
            {
                SqlTransaction transaction = null;
                connection.Open();
                SqlCommand cmd = new SqlCommand("delete from tbl_auto_Talent_Mobility where SnapShotDate='" + System.DateTime.Now + "'", connection);
                cmd.ExecuteNonQuery();
                try
                {
                    transaction = connection.BeginTransaction();
                    using (var sqlBulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.TableLock, transaction))
                    {
                        sqlBulkCopy.DestinationTableName = "tbl_auto_Talent_Mobility";

                        sqlBulkCopy.ColumnMappings.Add("Emp No", "Emp No");
                        sqlBulkCopy.ColumnMappings.Add("Emp Name", "Emp Name");
                        sqlBulkCopy.ColumnMappings.Add("Company", "Company");
                        sqlBulkCopy.ColumnMappings.Add("SL Unit", "SL Unit");
                        sqlBulkCopy.ColumnMappings.Add("SL SubUnit", "SL SubUnit");
                        sqlBulkCopy.ColumnMappings.Add("Unit", "Unit");
                        sqlBulkCopy.ColumnMappings.Add("Subunit", "Subunit");
                        sqlBulkCopy.ColumnMappings.Add("PU", "PU");
                        sqlBulkCopy.ColumnMappings.Add("DU", "DU");
                        sqlBulkCopy.ColumnMappings.Add("Account", "Account");
                        //sqlBulkCopy.ColumnMappings.Add("Gdly", "Gdly");
                        sqlBulkCopy.ColumnMappings.Add("Job Sub Level", "Job Sub Level");
                        sqlBulkCopy.ColumnMappings.Add("Persongroupcode", "Persongroupcode");
                        sqlBulkCopy.ColumnMappings.Add("Country", "Country");
                        sqlBulkCopy.ColumnMappings.Add("Core", "Core");
                        sqlBulkCopy.ColumnMappings.Add("Account Tenure Age", "Account Tenure Age");
                        sqlBulkCopy.ColumnMappings.Add("Gt18", "Gt18");
                        sqlBulkCopy.ColumnMappings.Add("SnapShotDate", "snapshotdate");

                        sqlBulkCopy.WriteToServer(dataTable);
                    }
                    transaction.Commit();


                    cmd = new SqlCommand("sp_auto_Batch_Talent_Mobility_v1", connection);
                    cmd.Parameters.AddWithValue("@LoadType", LoadType);
                    cmd.Parameters.AddWithValue("@finyear", finyear);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    ViewBag.Message = "Data Uploaded Successfully...";


                }
                catch (Exception ex)
                {
                    ViewBag.Message = "ERROR:" + ex.Message.ToString();
                }

            }
        }


        private List<SL_List> SL(string tab)
        {
            var list = new List<SL_List>
            {
                 new SL_List{Id = 1, SL = "EAS", Checked = true},
                 new SL_List{Id = 2, SL = "SAP", Checked = true},
                 new SL_List{Id = 3, SL = "ORC", Checked = true},
                 new SL_List{Id = 4, SL = "ECAS", Checked = true},
                 new SL_List{Id = 5, SL = "EAIS", Checked = true},
                  new SL_List{hidTAB = tab}
            };

            return list;
        }

        private static void Data_Generation()
        {
                    System.Threading.Tasks.Parallel.Invoke(
              () => DL.ExcuteSP("sp_auto_Summary_Complete_Load", true)
            , () => DL.ExcuteSP("sp_auto_RPP_Complete_Load", true)
            , () => DL.ExcuteSP("sp_auto_RoleMix_Complete_Load", true)
            , () => DL.ExcuteSP("sp_auto_Utilization_Complete_Load", true)
            , () => DL.ExcuteSP("sp_auto_DelayedBilling_Complete_Load", true)
            , () => DL.ExcuteSP("sp_auto_TMUnbilled_Complete_Load", true)

           );
        }

        private void SummarySheetGeneration(string FinYear)
        {

            string[] SL = new string[] { "EAS","SAP","ORC","EAIS","ECAS" };

            //var SLlist = list.Where(i => i._FY != null).ToList();
            //var emptylist = list.Where(i => i._FY == null).ToList();

            //Config

            var workbook = new XLWorkbook(Server.MapPath("~/template/" + "SummarySheets.xlsx"));

            var config = workbook.Worksheet("Config");

            //config.Cell("C4").Value = 1;

            var summary = workbook.Worksheet("Summary Snapshot");

            string monthname = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(System.DateTime.Today.Month - 1);
            string Year = FinYear.Substring(2, 2);
            string Header = "EAS Snapshot for " + monthname + "'" + Year;
            summary.Cell(1, 1).Value = Header;

            for (int j = 0; j < SL.Length; j++)
            {
                List<Summary> list = FetchSummaryData(SL[j], FinYear);

                var ws = workbook.Worksheet(SL[j] + " Summary Trend");
                int row = 4;
                for (int i = 0; i < list.Count; i++)
                {
                    ws.Cell("C" + row.ToString()).Value = list[i]._H1;
                    ws.Cell("D" + row.ToString()).Value = list[i]._trend;
                    ws.Cell("E" + row.ToString()).Value = list[i]._blank;
                    ws.Cell("F" + row.ToString()).Value = list[i]._FY;
                    ws.Cell("G" + row.ToString()).Value = list[i]._blank;
                    ws.Cell("H" + row.ToString()).Value = list[i]._M1;
                    ws.Cell("I" + row.ToString()).Value = list[i]._M2;
                    ws.Cell("J" + row.ToString()).Value = list[i]._M3;
                    ws.Cell("K" + row.ToString()).Value = list[i]._M4;
                    ws.Cell("L" + row.ToString()).Value = list[i]._M5;
                    ws.Cell("M" + row.ToString()).Value = list[i]._M6;
                    ws.Cell("N" + row.ToString()).Value = list[i]._M7;
                    ws.Cell("O" + row.ToString()).Value = list[i]._M8;
                    ws.Cell("P" + row.ToString()).Value = list[i]._M9;
                    ws.Cell("Q" + row.ToString()).Value = list[i]._M10;
                    ws.Cell("R" + row.ToString()).Value = list[i]._M11;
                    ws.Cell("S" + row.ToString()).Value = list[i]._M12;
                    ws.Cell("W" + row.ToString()).Value = list[i]._PrevH2 ;
                    row++;
                }
               
                Header = SL[j] + " Snapshot for " +monthname +"'"+ Year;
                ws.Cell(2, 1).Value = Header;

            }
            string filename = "EAS_MarginImprovement_QuickSummary_Report_" + DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx";
            workbook.SaveAs(Server.MapPath("~/Summary/" + filename));



        }

        private void SummarySheetGeneration_SecondHalf(string FinYear)
        {

            string[] SL = new string[] { "EAS", "SAP", "ORC", "EAIS", "ECAS" };

            //var SLlist = list.Where(i => i._FY != null).ToList();
            //var emptylist = list.Where(i => i._FY == null).ToList();

            //Config

            var workbook = new XLWorkbook(Server.MapPath("~/template/" + "Summary_Second_Half.xlsx"));

            var config = workbook.Worksheet("Config");

            //config.Cell("C4").Value = 1;

            var summary = workbook.Worksheet("Summary Snapshot");

            string monthname = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(System.DateTime.Today.Month - 1);
            string Year = FinYear.Substring(2, 2);
            string Header = "EAS Snapshot for " + monthname + "'" + Year;
            summary.Cell(1, 1).Value = Header;

            for (int j = 0; j < SL.Length; j++)
            {
                List<Summary> list = FetchSummaryData_SecondHalf(SL[j], FinYear);

                var ws = workbook.Worksheet(SL[j] + " Summary Trend");
                int row = 4;
                for (int i = 0; i < list.Count; i++)
                {
                    ws.Cell("C" + row.ToString()).Value = list[i]._H2;
                    ws.Cell("D" + row.ToString()).Value = list[i]._trend;
                    ws.Cell("E" + row.ToString()).Value = list[i]._blank;
                    ws.Cell("F" + row.ToString()).Value = list[i]._FY;
                    ws.Cell("G" + row.ToString()).Value = list[i]._blank;
                    ws.Cell("H" + row.ToString()).Value = list[i]._M1;
                    ws.Cell("I" + row.ToString()).Value = list[i]._M2;
                    ws.Cell("J" + row.ToString()).Value = list[i]._M3;
                    ws.Cell("K" + row.ToString()).Value = list[i]._M4;
                    ws.Cell("L" + row.ToString()).Value = list[i]._M5;
                    ws.Cell("M" + row.ToString()).Value = list[i]._M6;
                    ws.Cell("N" + row.ToString()).Value = list[i]._M7;
                    ws.Cell("O" + row.ToString()).Value = list[i]._M8;
                    ws.Cell("P" + row.ToString()).Value = list[i]._M9;
                    ws.Cell("Q" + row.ToString()).Value = list[i]._M10;
                    ws.Cell("R" + row.ToString()).Value = list[i]._M11;
                    ws.Cell("S" + row.ToString()).Value = list[i]._M12;
                    ws.Cell("W" + row.ToString()).Value = list[i]._H1;
                    row++;
                }

                Header = SL[j] + " Snapshot for " + monthname + "'" + Year;
                ws.Cell(2, 1).Value = Header;

            }
            string filename = "EAS_MarginImprovement_QuickSummary_Report_" + DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx";
            workbook.SaveAs(Server.MapPath("~/Summary/" + filename));



        }


        public static List<Summary> FetchSummaryData(string SL, string Finyear)
        {
            //con.Open();
            //sp_Auto_Summary_Trend_Fetch_Online_FinYear_Second_Half
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand("sp_Auto_Summary_Trend_Fetch_Online_FinYear_First_Half_1", con);
            cmd.CommandTimeout = int.MaxValue;
            cmd.Parameters.AddWithValue("@SL", SL);
            cmd.Parameters.AddWithValue("@Finyear", Finyear);
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            sdr.Fill(ds);

            var summaryList = ds.Tables[0].AsEnumerable().Select(dataRow =>
                new Summary
                {
                    _parameter = dataRow.Field<string>("SubParam"),
                    _H1 = dataRow.Field<double?>("H1"),
                    _trend = dataRow.Field<string>("Trend"),
                    _blank = dataRow.Field<string>("Blank"),
                    _FY = dataRow.Field<double?>("FY"),
                    _blank1 = dataRow.Field<string>("Blank1"),
                    _M1 = dataRow.Field<double?>("M1"),
                    _M2 = dataRow.Field<double?>("M2"),
                    _M3 = dataRow.Field<double?>("M3"),
                    _M4 = dataRow.Field<double?>("M4"),
                    _M5 = dataRow.Field<double?>("M5"),
                    _M6 = dataRow.Field<double?>("M6"),
                    _M7 = dataRow.Field<double?>("M7"),
                    _M8 = dataRow.Field<double?>("M8"),
                    _M9 = dataRow.Field<double?>("M9"),
                    _M10 = dataRow.Field<double?>("M10"),
                    _M11 = dataRow.Field<double?>("M11"),
                    _M12 = dataRow.Field<double?>("M12"),
                    _PrevH2 = dataRow.Field<double?>("PrevH2")
                }).ToList();

            return summaryList;
        }


        public static List<Summary> FetchSummaryData_SecondHalf(string SL, string Finyear)
        {
            //con.Open();
            
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand("sp_Auto_Summary_Trend_Fetch_Online_FinYear_Second_Half", con);
            cmd.CommandTimeout = int.MaxValue;
            cmd.Parameters.AddWithValue("@SL", SL);
            cmd.Parameters.AddWithValue("@Finyear", Finyear);
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            sdr.Fill(ds);

            var summaryList = ds.Tables[0].AsEnumerable().Select(dataRow =>
                new Summary
                {
                    _parameter = dataRow.Field<string>("SubParam"),
                    _H2 = dataRow.Field<double?>("H2"),
                    _trend = dataRow.Field<string>("Trend"),
                    _blank = dataRow.Field<string>("Blank"),
                    _FY = dataRow.Field<double?>("FY"),
                    _blank1 = dataRow.Field<string>("Blank1"),
                    _M1 = dataRow.Field<double?>("M1"),
                    _M2 = dataRow.Field<double?>("M2"),
                    _M3 = dataRow.Field<double?>("M3"),
                    _M4 = dataRow.Field<double?>("M4"),
                    _M5 = dataRow.Field<double?>("M5"),
                    _M6 = dataRow.Field<double?>("M6"),
                    _M7 = dataRow.Field<double?>("M7"),
                    _M8 = dataRow.Field<double?>("M8"),
                    _M9 = dataRow.Field<double?>("M9"),
                    _M10 = dataRow.Field<double?>("M10"),
                    _M11 = dataRow.Field<double?>("M11"),
                    _M12 = dataRow.Field<double?>("M12"),
                    _H1 = dataRow.Field<double?>("H1")
                }).ToList();

            return summaryList;
        }

        private void Cost_Opt_Generation(List<SL_List> list, string FinYear,string isFocus)
        {

            string[] SL = new string[] { "EAS", "SAP", "ORC", "EAIS", "ECAS" };

            Data_Pack_Full_Version_File_Generate(list[0].SL.ToString(), list[0].Checked,isFocus);
            Data_Pack_Full_Version_File_Generate(list[1].SL.ToString(), list[1].Checked,isFocus);
            Data_Pack_Full_Version_File_Generate(list[2].SL.ToString(), list[2].Checked,isFocus);
            Data_Pack_Full_Version_File_Generate(list[3].SL.ToString(), list[3].Checked,isFocus);
            Data_Pack_Full_Version_File_Generate(list[4].SL.ToString(), list[4].Checked,isFocus);

        }

        public void Data_Pack_Full_Version_File_Generate(string Serviceline, bool check,string isFocus)
        {

            if (check)
            {
                string folder = "template";
                var MyDir = new DirectoryInfo(Server.MapPath(folder));

                //string PathToRef = "";
                //PathToRef = Server.MapPath("~/obj/Debug/scrrun.dll");
                Excel.Application oExcel1;
                Excel.Workbook oBook1 = default(Excel.Workbook);
                VBIDE.VBComponent oModule1;

                //PathToRef = Server.MapPath("~/obj/Debug/scrrun.dll");

                string templatefolder = "~/template";
                var templatemyDir = new DirectoryInfo(Server.MapPath(templatefolder));


                //string templatePath = templatemyDir.FullName + "\\DataPack_template_Full.xlsb";
                //string templatePath = MyDir.FullName + "\\Cost_Opt_template_" + Serviceline + ".xlsb";

                string templatePath = Server.MapPath("~/template/" + "Cost_Opt_template_" + Serviceline + ".xlsb");

                String sCode1;
                Object oMissing1 = System.Reflection.Missing.Value;


                oExcel1 = new Excel.Application();

                oBook1 = oExcel1.Workbooks.
                    Open(templatePath, 0, false, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                oModule1 = oBook1.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

                sCode1 = "Sub DataPack_" + Serviceline + "Part1()\r\n" +
                     GetVariableDeclaration("p", MacroDataType.String, Serviceline) +
                     GetVariableDeclaration("isFocus", MacroDataType.String, isFocus) +
                     //System.IO.File.ReadAllText(templatemyDir.FullName + "\\Fetch_Macro.txt") +
                     System.IO.File.ReadAllText(Server.MapPath("~/template/" + "Fetch_Macro.txt")) +
                         "\nend Sub";
                oModule1.CodeModule.AddFromString(sCode1);
                oExcel1.GetType().InvokeMember("Run",
                                System.Reflection.BindingFlags.Default |
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, oExcel1, new string[] { "DataPack_" + Serviceline + "Part1" });


                folder = "~/DataPack_Full_Version";
                MyDir = new DirectoryInfo(Server.MapPath(folder));


                string path = "";

                if(isFocus=="yes")
                {
                    path = Server.MapPath("~/DataPack_Full_Version/" + Serviceline + "_Datapack_full_Incl_FA_" + DateTime.Now.ToString("ddMMMyyyy hhmmss IST") + ".xlsb");
                }
                else
                {
                    path = Server.MapPath("~/DataPack_Full_Version/" + Serviceline + "_Datapack_full__Excl_FA_" + DateTime.Now.ToString("ddMMMyyyy hhmmss IST") + ".xlsb");
                }
               

                oBook1.SaveAs(path);
                oBook1.Close(false, templatePath, null);
                oExcel1.Quit();

                System.GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcel1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oModule1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook1);
            }
        }


        private void Cost_Opt_Generation_OLD(List<SL_List> list , string FinYear)
        {
            //Data Generation
            string[] SL = new string[] { "EAS", "SAP", "ORC", "EAIS", "ECAS" };
            Data_Generation_Client_Category(SL);


            string folder = "Cost_Opt_Reports//" + DateTime.Now.ToString("dd_MMM_yyyy");

            if (Directory.Exists(Server.MapPath(folder)))
                Directory.Delete(Server.MapPath(folder), true);

            Directory.CreateDirectory(Server.MapPath(folder));

            DL.ExcuteText("truncate table tbl_Auto_RPP_Fetch_ALL");
            DL.ExcuteText("truncate table tbl_Auto_DelayedBilling_Fetch_ALL");
            DL.ExcuteText("truncate table tbl_Auto_RoleMix_GMP_Fetch_ALL");
            DL.ExcuteText("truncate table tbl_Auto_RoleMix_Fetch_ALL");
            DL.ExcuteText("truncate table tbl_Auto_Summary_Fetch_ALL");
            DL.ExcuteText("truncate table tbl_Auto_TMUnbilled_Fetch_ALL");
            DL.ExcuteText("truncate table tbl_Auto_Utilization_Fetch_All");
            DL.ExcuteText("truncate table tbl_Auto_GMPLowMargin_Fetch_ALL");


            string IsEAS_ALL = "NO";
            Cost_Opt_File_Generate(list[0].SL.ToString(), list[0].Checked, IsEAS_ALL, FinYear);
            Cost_Opt_File_Generate(list[1].SL.ToString(), list[1].Checked, IsEAS_ALL, FinYear);
            Cost_Opt_File_Generate(list[2].SL.ToString(), list[2].Checked, IsEAS_ALL, FinYear);
            Cost_Opt_File_Generate(list[3].SL.ToString(), list[3].Checked, IsEAS_ALL, FinYear);
            Cost_Opt_File_Generate(list[4].SL.ToString(), list[4].Checked, IsEAS_ALL, FinYear);

            IsEAS_ALL = "YES";
            Cost_Opt_File_Generate("EAS", true, IsEAS_ALL, FinYear);
        }

        public void Cost_Opt_File_Generate(string Serviceline,bool check,string IsEAS_ALL,string FinYear)
        {

            if (check)
            {
                string folder = "template";
                var MyDir = new DirectoryInfo(Server.MapPath(folder));

                //string PathToRef = "";
                //PathToRef = Server.MapPath("~/obj/Debug/scrrun.dll");
                Excel.Application oExcel1;
                Excel.Workbook oBook1 = default(Excel.Workbook);
                VBIDE.VBComponent oModule1;


                //PathToRef = Server.MapPath("~/obj/Debug/scrrun.dll");

             

                string templatefolder = "~/template";
                var templatemyDir = new DirectoryInfo(Server.MapPath(templatefolder));


                string templatePath = templatemyDir.FullName + "\\Cost_Opt_template_" + Serviceline + ".xlsb";
                //string templatePath = MyDir.FullName + "\\Margin_Improvement_template_" + Serviceline + ".xlsb";
                String sCode1;
                Object oMissing1 = System.Reflection.Missing.Value;


                oExcel1 = new Excel.Application();

                oBook1 = oExcel1.Workbooks.
                    Open(templatePath, 0, false, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                oModule1 = oBook1.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

                sCode1 = "Sub CostOptMacro_" + Serviceline + "Part1()\r\n" +
                     GetVariableDeclaration("c", MacroDataType.String, Serviceline) +
                     GetVariableDeclaration("e", MacroDataType.String, IsEAS_ALL) +
                     GetVariableDeclaration("y", MacroDataType.String, FinYear) +
                     System.IO.File.ReadAllText(templatemyDir.FullName + "\\Data_Binding_Formating.txt") +
                         "\nend Sub";
                oModule1.CodeModule.AddFromString(sCode1);
                oExcel1.GetType().InvokeMember("Run",
                                System.Reflection.BindingFlags.Default |
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, oExcel1, new string[] { "CostOptMacro_" + Serviceline + "Part1" });

                //folder = "Cost_Opt_Reports//" + DateTime.Now.ToString("dd_MMM_yyyy");
                //folder = "Cost_Opt_Reports";
                folder= "~/Cost_Opt_Reports";
                MyDir = new DirectoryInfo(Server.MapPath(folder));

                string path = MyDir.FullName + "\\" + Serviceline + "_Margin_Improvement_Report_" + DateTime.Now.ToString("dd_MMM_yyyy hh-mm-ss tt") + ".xlsb";

                if (IsEAS_ALL =="YES")
                    path = MyDir.FullName + "\\" + Serviceline + "_ALL_Margin_Improvement_Report_" + DateTime.Now.ToString("dd_MMM_yyyy hh-mm-ss tt") + ".xlsb";

                oBook1.SaveAs(path);
                oBook1.Close(false, templatePath, null);
                oExcel1.Quit();

                System.GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcel1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oModule1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook1);
            }
        }

        enum MacroDataType
        {
            String, Integer

        }

        private string GetVariableDeclaration(string variableName, MacroDataType type, object value)
        {


            string formatString = "Dim {0} as {1} \n {0} =\"{2}\" \n";
            string formatNumber = "Dim {0} as {1} \n {0} ={2} \n";
            string returnValue = "";
            switch (type)
            {
                case MacroDataType.String:
                    returnValue = string.Format(formatString, variableName, type.ToString(), value);
                    break;
                case MacroDataType.Integer:
                    returnValue = string.Format(formatNumber, variableName, type.ToString(), value);
                    break;
                default:
                    break;
            }
            return returnValue;
        }


        // ---------------------------------------- Client Category ----------------------------------------------------------------------------------------------

        private void Data_Generation_Client_Category(string[] SL)
        {
            foreach (string _ddlServiceLine in SL)
            {
                if (_ddlServiceLine != "EAS")
                {
                    GenerateReportCC_SL(_ddlServiceLine, "Overall");
                    GenerateReportCC_SL(_ddlServiceLine, "Digital");
                }
                else
                {
                    GenerateReportCC("Overall");
                    GenerateReportCC("Digital");
                }

                Init_ClientCategory(_ddlServiceLine, "Overall");
                Init_ClientCategory(_ddlServiceLine, "Digital");

            }
        }

        void GenerateReportCC(string sheetName)
        {

            string macroName = "";

            if (sheetName == "Digital")
                macroName = "CC_Digital.txt";
            else
                macroName = "CC.txt";

            string fname = @"D:\Som_KP Prod\Source\SOM-KP\ExcelOperations\EAS_dashboard_chart3.xlsb";
            Microsoft.Office.Interop.Excel.Application oExcel;
            Microsoft.Office.Interop.Excel.Workbook oBook = default(Microsoft.Office.Interop.Excel.Workbook);
            VBIDE.VBComponent oModule;
            //try
            {

                string folder = "~/ExcelOperation";
                var myDir = new DirectoryInfo(Server.MapPath(folder));

                string templatefolder = "~/template/ClientCategory";
                var templatemyDir = new DirectoryInfo(Server.MapPath(templatefolder));

                String sCode;
                Object oMissing = System.Reflection.Missing.Value;
                //instance of excel
                oExcel = new Microsoft.Office.Interop.Excel.Application();

                oBook = oExcel.Workbooks.
                    //Open(templatemyDir.FullName + "\\" + fname, 0, false, 5, "", "", true,
                    Open(fname, 0, false, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Sheets WRss = oBook.Sheets;

                oModule = oBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

                sCode = "sub Macro()\r\n" +
                    //GetVariableDeclaration("SL", MacroDataType.String, ServiceLine) +

                    System.IO.File.ReadAllText(templatemyDir.FullName + "\\" + macroName) +
                        "\nend sub";
                oModule.CodeModule.AddFromString(sCode);

                oExcel.GetType().InvokeMember("Run",
                                System.Reflection.BindingFlags.Default |
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, oExcel, new string[] { "Macro" });



                /////////////////////////////////////

                oBook.Save();
                oBook.Close(false, myDir.FullName + "\\" + fname + "", null);


                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
                oBook = null;

                oExcel.Quit();
                oExcel = null;



                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();



            }

        }

        void GenerateReportCC_SL(string ServiceLine, string sheetName)
        {
            string fname = "EAS_Client_Category.xlsm";

            string macroName = "";

            if (sheetName == "Digital")
            {
                fname = "EAS_Client_Category_Digital.xlsm";
                macroName = "SL_Slicer_Digital.txt";
            }
            else
            {
                fname = "EAS_Client_Category.xlsm";
                macroName = "SL_Slicer.txt";
            }

            Microsoft.Office.Interop.Excel.Application oExcel;
            Microsoft.Office.Interop.Excel.Workbook oBook = default(Microsoft.Office.Interop.Excel.Workbook);
            VBIDE.VBComponent oModule;
            //try
            {

                string folder = "~/ExcelOperation";
                var myDir = new DirectoryInfo(Server.MapPath(folder));

                string templatefolder = "~/template/ClientCategory";
                var templatemyDir = new DirectoryInfo(Server.MapPath(templatefolder));

                String sCode;
                Object oMissing = System.Reflection.Missing.Value;
                //instance of excel
                oExcel = new Microsoft.Office.Interop.Excel.Application();

                oBook = oExcel.Workbooks.
                    Open(templatemyDir.FullName + "\\" + fname, 0, false, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Sheets WRss = oBook.Sheets;

                oModule = oBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

                sCode = "sub Macro()\r\n" +
                    GetVariableDeclaration("SL", MacroDataType.String, ServiceLine) +

                    System.IO.File.ReadAllText(templatemyDir.FullName + "\\" + macroName) +
                        "\nend sub";
                oModule.CodeModule.AddFromString(sCode);

                oExcel.GetType().InvokeMember("Run",
                                System.Reflection.BindingFlags.Default |
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, oExcel, new string[] { "Macro" });



                /////////////////////////////////////

                oBook.Save();
                oBook.Close(false, myDir.FullName + "\\" + fname + "", null);


                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
                oBook = null;

                oExcel.Quit();
                oExcel = null;



                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();



            }

        }

        private void Init_ClientCategory(string _ddlServiceLine, string sheetName)
        {
            var wbcc = new XLWorkbook();
            if (sheetName == "Overall")
                wbcc = new XLWorkbook(Server.MapPath("~/template/ClientCategory/EAS_Client_Category.xlsm"));
            else
                wbcc = new XLWorkbook(Server.MapPath("~/template/ClientCategory/EAS_Client_Category_Digital.xlsm"));

            var ws = wbcc.Worksheet(1);
            DataTable dtClientCategory = FetchClientCategory(ws, _ddlServiceLine, sheetName);
            DL.insertCC(dtClientCategory, _ddlServiceLine, sheetName);

        }

        private static DataTable FetchClientCategory(IXLWorksheet ws, string _ddlServiceLine, string Sheet)
        {
            DataTable dtClientCategory = new DataTable();

            var rngheader = ws.Range("C28:N28");

            dtClientCategory.Columns.Add("SL");
            dtClientCategory.Columns.Add("SheetName");

            foreach (var cell in rngheader.Cells())
            {
                string header = cell.Value.ToString();

                if (header != "")
                {
                    dtClientCategory.Columns.Add(header);
                }
            }

            var rows = ws.RangeUsed().RowsUsed().Skip(3);
            int headercount = dtClientCategory.Columns.Count;

            foreach (var overall_data in rows)
            {
                string text = overall_data.Cell(3).Value.ToString();

                if (text.Contains("Total"))
                {
                    if (text == "Total Total Billed Months ")
                    {
                        text = "Total Billed Months";
                    }
                    else
                        text = text.Replace("Total ", "");

                    DataRow dr = dtClientCategory.NewRow();
                    int dtcol = 2;
                    for (int col = 2; col < dtClientCategory.Columns.Count; col++)
                    {
                        dr[0] = _ddlServiceLine;
                        dr[1] = Sheet;

                        if (col == 2)
                        {
                            string r = overall_data.Cell(3).Value.ToString();

                            r = r == "Total # of MCC's" ? "Overall" : "";
                            dr[dtcol] = r;
                        }
                        else if (col == 3)
                        {
                            dr[dtcol] = text;
                        }
                        else
                        {
                            dr[dtcol] = overall_data.Cell(col + 1).Value;
                        }
                        dtcol++;
                    }
                    dtClientCategory.Rows.Add(dr);

                }
            }



            foreach (var overall_data in rows)
            {
                string text = overall_data.Cell(3).Value.ToString();
                string text1 = overall_data.Cell(4).Value.ToString();

                if (!text.Contains("Total") && !text1.Contains("Values") && text1 != "")
                {
                    DataRow dr1 = dtClientCategory.NewRow();
                    int dtcol = 2;
                    for (int col = 2; col < dtClientCategory.Columns.Count; col++)
                    {
                        dr1[0] = _ddlServiceLine;
                        dr1[1] = Sheet;
                        dr1[dtcol] = overall_data.Cell(col + 1).Value;
                        dtcol++;
                    }
                    dtClientCategory.Rows.Add(dr1);
                }
            }
            return dtClientCategory;
        }
    }
}
