using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
namespace FullVersion.Controllers
{
 

    public class DL
    {
        public static SqlConnection con = new SqlConnection(@"Data Source=nebula\mssqlserver1;Initial Catalog=EAS_PROD;uid=sa;pwd=Infy123+");

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

        public static void insertCC(DataTable dt, string SL, string SheetName)
        {
            // using (SqlConnection connection = new SqlConnection(con))
            {
                con.Open();

                SqlCommand cmmd = new SqlCommand("delete from [tbl_Client_Category_PD] where  SL ='" + SL + "' and SheetName='" + SheetName + "'", con);
                cmmd.ExecuteNonQuery();

                con.Close();

                SqlBulkCopy bulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
                bulkCopy.DestinationTableName = "tbl_Client_Category_PD";
                bulkCopy.BulkCopyTimeout = 1700;

                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    if (dt.Columns[k].ColumnName == "PBT Bucket")
                        bulkCopy.ColumnMappings.Add("PBT Bucket", "PBT Bucket");
                    else if (dt.Columns[k].ColumnName == "Values")
                        bulkCopy.ColumnMappings.Add("Values", "Values");
                    else if (dt.Columns[k].ColumnName == "50-100Mn")
                        bulkCopy.ColumnMappings.Add("50-100Mn", "50-100Mn");
                    else if (dt.Columns[k].ColumnName == "25-50Mn")
                        bulkCopy.ColumnMappings.Add("25-50Mn", "25-50Mn");
                    else if (dt.Columns[k].ColumnName == "10-25Mn")
                        bulkCopy.ColumnMappings.Add("10-25Mn", "10-25Mn");
                    else if (dt.Columns[k].ColumnName == "5-10Mn")
                        bulkCopy.ColumnMappings.Add("5-10Mn", "5-10Mn");
                    else if (dt.Columns[k].ColumnName == "1-5Mn")
                        bulkCopy.ColumnMappings.Add("1-5Mn", "1-5Mn");
                    else if (dt.Columns[k].ColumnName == "0.05-1Mn")
                        bulkCopy.ColumnMappings.Add("0.05-1Mn", "0.05-1Mn");
                    else if (dt.Columns[k].ColumnName == "0-0.05Mn")
                        bulkCopy.ColumnMappings.Add("0-0.05Mn", "0-0.05Mn");
                    else if (dt.Columns[k].ColumnName == "Zero")
                        bulkCopy.ColumnMappings.Add("Zero", "Zero");
                    else if (dt.Columns[k].ColumnName == "Negative")
                        bulkCopy.ColumnMappings.Add("Negative", "Negative");
                    else if (dt.Columns[k].ColumnName == "Grand Total")
                        bulkCopy.ColumnMappings.Add("Grand Total", "Grand Total");
                    else if (dt.Columns[k].ColumnName == "SL")
                        bulkCopy.ColumnMappings.Add("SL", "SL");
                    else if (dt.Columns[k].ColumnName == "SheetName")
                        bulkCopy.ColumnMappings.Add("SheetName", "SheetName");
                }

                con.Open();
                bulkCopy.WriteToServer(dt);
                con.Close();
            }
        }

        public static void ExcuteSP(string spName, bool check)
        {
            SqlCommand cmd = new SqlCommand(spName, con);
            cmd.CommandTimeout = int.MaxValue;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.ExecuteNonQuery();
        }

        public static void ExcuteText(string Text)
        {
            con.Open();
            SqlCommand cmd = new SqlCommand(Text, con);
            cmd.CommandTimeout = int.MaxValue;
            cmd.CommandType = CommandType.Text;
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}