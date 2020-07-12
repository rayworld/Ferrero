using System;
using System.Data;
using System.Data.OleDb;

namespace Ferrero.Common
{
    public static class ExcelHelper
    {
        #region ChkColumnsName
        /// <summary>
        /// 调用CheckColumnExists检查Excel中的列是否存在
        /// </summary>
        /// <param name="ColumnsName"></param>
        /// <returns></returns>
        public static string ChkColumnsName(string[] ColumnsName, DataTable dt)
        {
            string sErrMsg = "";
            foreach (string sColName in ColumnsName)
            {
                if (CheckColumnExist(sColName, dt) == false)
                {
                    sErrMsg += "\"" + sColName + "\",";
                }
            }
            return sErrMsg;
        }
        #endregion

        #region CheckColumnExist
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        private static bool CheckColumnExist(string sColumnName, DataTable dt)
        {
            bool retVal = false;

            foreach (DataColumn dc in dt.Columns)
            {
                if (dc.ColumnName == sColumnName)
                {
                    retVal = true;
                    break;
                }
            }
            return retVal;
        }
        #endregion

        #region 得到工作簿名称列表
        /// <summary>
        /// 得到工作簿名称列表
        /// </summary>
        /// <param name="excelFile">Excel文件名</param>
        /// <returns></returns>
        public static String[] GetExcelSheetNames(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;
            string connString = "";

            try
            {
                // Connection String. Change the excel file to the file you
                // will search.
                if (excelFile.IndexOf(".xlsx") > 0)
                {
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + excelFile + ";Extended Properties=Excel 12.0;";
                }
                else
                {
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";Extended Properties=Excel 8.0;";
                }
                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }
        #endregion

        #region 将选定的 Excel 数据转换成 DataTable 数据集
        /// <summary>
        /// 将选定的 Excel 数据转换成 DataTable 数据集
        /// </summary>
        /// <param name="sExcelFileName">Excel文件名</param>
        /// <param name="sSheetName">工作簿名称</param>
        /// <returns></returns>
        public static DataTable LoadDataFromExcel(string sExcelFileName, string sSheetName)
        {
            string sExcelConnectionString = "";
            try
            {
                // IMEX=1 可把混合型作为文本型读取，避免null值
                if (sExcelFileName.IndexOf(".xlsx") > 0)
                {
                    sExcelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sExcelFileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";
                }
                else if (sExcelFileName.IndexOf(".xls") > 0)
                {
                    sExcelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sExcelFileName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
                }
                else
                {
                    //string sExcelConnectionString = "";
                }
                OleDbConnection OleConn = new OleDbConnection(sExcelConnectionString);
                OleConn.Open();
                String sql = "SELECT * FROM [" + sSheetName + "]"; // 可更改 Sheet 名称    
                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet ds = new DataSet();
                OleDaExcel.Fill(ds, sSheetName);
                OleConn.Close();
                return ds.Tables[0];
            }

            catch (Exception ex)
            {
                Console.WriteLine("数据绑定Excel失败!失败原因：" + ex.Message);
                return null;
            }
        }
        #endregion
    }
}