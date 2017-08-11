using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Raysoft.Common
{
    public class Excel2DataTable
    {
        #region OpenExcelFile2DataTable

        /// <summary>
        /// 将Excel文件转制成数据表
        /// 注意这里的工作簿名称最好不用中文，否则可能打不开
        /// </summary>
        /// <returns>数据表</returns>
        public DataTable ExcelFile2DataTable()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Microsoft Excel 文件|*.xlsx;*.xls|所有文件|*.*";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK && ofd.FileName != "" && ofd.FileName != null)
            {
                return LoadDataFromExcel(ofd.FileName, "Sheet1$");
            }
            else
            {
                return (DataTable)null;
            }
        }
        #endregion

        #region LoadDataFromExcel
        /// <summary>
        /// 将选定的 Excel 数据转换成 DatatSet 数据集
        /// </summary>
        /// <param name="sExcelFileName">Excel文件名</param>
        /// <param name="sSheetName">工作簿名称</param>
        /// <returns></returns>
        private DataTable LoadDataFromExcel(string sExcelFileName, string sSheetName)
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
                MessageBox.Show("数据绑定Excel失败!失败原因：" + ex.Message, "提示信息", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return null;
            }
        }
        #endregion

        #region GetExcelSheetNames
        /// <summary>
        /// 得到Excel文件中所有工作簿列表
        /// </summary>
        /// <param name="excelFile">Excel文件名称</param>
        /// <param name="ExcelVersion">Excel文件版本，主要区分office 2003和2007</param>
        /// <returns></returns>
        public String[] GetExcelSheetNames(string excelFile)
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
                MessageBox.Show(ex.Message);
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

        #region ChkColumnsName
        /// <summary>
        /// 调用CheckColumnExists检查列是否存在
        /// </summary>
        /// <param name="ColumnsName"></param>
        /// <returns></returns>
        public bool ChkColumnsName(string[] ColumnsName,DataTable dt)
        {
            string sErrMsg = "";
            foreach (string sColName in ColumnsName)
            {
                if (CheckColumnExist(sColName,dt) == false)
                {
                    sErrMsg += "\"" + sColName + "\",";
                }
            }

            if (sErrMsg != "")
            {
                MessageBox.Show("请检查Excel文件，不存在的字段有:" + sErrMsg.Substring(0, sErrMsg.Length - 1), "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            return sErrMsg == "" ? true : false;
        }
        #endregion

        #region CheckColumnExist
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        private bool CheckColumnExist(string sColumnName,DataTable dt)
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

        #region bIsSpecialCustom
        /// <summary>
        /// 判定是否为大客户
        /// </summary>
        /// <param name="sCustomCode"></param>
        /// <returns></returns>
        public bool bIsSpecialCustom(string sCustomCode)
        {
            bool retVal = false;
            //百佳
            if (sCustomCode.Substring(0, 6) == "B.1013") retVal = true;
            //卜蜂莲花
            if (sCustomCode.Substring(0, 4) == "A.14") retVal = true;
            //华联
            if (sCustomCode.Substring(0, 4) == "A.15") retVal = true;
            //华润万家
            if (sCustomCode.Substring(0, 6) == "B.1006") retVal = true;
            //家乐福
            if (sCustomCode.Substring(0, 4) == "A.10") retVal = true;
            //麦德龙
            if (sCustomCode.Substring(0, 4) == "A.13") retVal = true;
            //屈臣氏
            if (sCustomCode.Substring(0, 4) == "A.09") retVal = true;
            //苏果
            if (sCustomCode.Substring(0, 6) == "B.1003") retVal = true;
            //沃尔玛
            if (sCustomCode.Substring(0, 4) == "A.12") retVal = true;
            //新一佳
            if (sCustomCode.Substring(0, 4) == "A.08") retVal = true;

            return retVal;
        }
        #endregion
    }
}
