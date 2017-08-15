using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using DevComponents.DotNetBar;

using Ray.Framework.Converter;

namespace Ferrero
{
    public partial class Form3 : Office2007Form 
    {
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        public Form3()
        {
            InitializeComponent();
        }


        #region 私有过程
        //private DataTable ReadExcelFile(string strFileName, string strSheetName)
        //{
        //    DataSet ds = new DataSet();
        //    if (strFileName != "")
        //    {
        //        try
        //        {
        //            //office 2003 
        //            string conn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";
        //            //office 2007
        //            //string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";  
        //            string sql = "select 商品长代码, 商品名称,规格型号, 批次, 单位, 期末结存数量 from [" + strSheetName + "$] ";
        //            //string sql = "SELECT * FROM OpenDataSource('Microsoft.Jet.OLEDB.4.0','Data Source=" + strFileName + ";Extended Properties='Excel 8.0;HDR=Yes;';Persist Security Info=False')...Sheet1$";
        //            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
        //            da.Fill(ds, "table1");
        //            MessageBox.Show("excel 导入成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //            //MessageBox.Show("导入文件时出错,文件可能正被打开", "提示");  
        //        }
        //        return ds.Tables[0];
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}
        #endregion

        private void panelEx1_Resize(object sender, EventArgs e)
        {
            dataGridView1.Width = panelEx1.Width / 2;
            dataGridView2.Width = panelEx1.Width / 2;
            dataGridView1.Height = panelEx1.Height;
            dataGridView2.Height = panelEx1.Height;
            dataGridView1.Left = 0;
            dataGridView2.Left = dataGridView1.Width;
            dataGridView1.Top = 0;
            dataGridView2.Top = 0;



        }


        /// <summary>
        /// 打开Ferrero即时库存Excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOepn1_Click(object sender, EventArgs e)
        {
            Excel2DataTable common = new Excel2DataTable();
            dt1 = common.ExcelFile2DataTable();
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    dt1.Rows.RemoveAt(index: dt1.Rows.Count - 1);
                    dataGridView2.DataSource = dt1;
                }
            };
        }
        
        private void btnCompare_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0 && dataGridView2.RowCount > 0)
            {
                //涂红友谊记录
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }

                //编号比对
                for (int k = 0; k < dataGridView1.RowCount; k++)
                {
                    for (int j = 0; j < dataGridView2.RowCount ; j++)
                    {
                        //费列罗物料代码
                        string yyid = dataGridView2.Rows[j].Cells["商品长代码"].Value.ToString();
                        //友谊物料代码
                        string feid = dataGridView1.Rows[k].Cells["商品长代码"].Value.ToString();
                        //费列罗批次
                        string yyBetchNo = dataGridView2.Rows[j].Cells["批次"].Value.ToString();
                        //友谊批次
                        string feBetchNo = dataGridView1.Rows[k].Cells["批次"].Value.ToString();
                        //批次==批次 and 物料代码 = 物料代码
                        if (yyid == feid && yyBetchNo == feBetchNo)
                        {
                            dataGridView1.Rows[k].DefaultCellStyle.BackColor = Color.White   ;
                            break;
                        }
                    }
                }

                //涂红ferrero记录
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }

                //编号比对
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    for (int k = 0; k < dataGridView2.RowCount; k++)
                    {
                        //费列罗物料代码
                        string yyid = dataGridView2.Rows[k].Cells["商品长代码"].Value.ToString();
                        //友谊物料代码
                        string feid = dataGridView1.Rows[j].Cells["商品长代码"].Value.ToString();
                        //费列罗批次
                        string yyBetchNo = dataGridView2.Rows[k].Cells["批次"].Value.ToString();
                        //友谊批次
                        string feBetchNo = dataGridView1.Rows[j].Cells["批次"].Value.ToString();
                        //批次==批次 and 物料代码 = 物料代码
                        if (yyid == feid && yyBetchNo == feBetchNo)
                        {
                            dataGridView2.Rows[k].DefaultCellStyle.BackColor = Color.White   ;
                            break;
                        }
                    }
                }

                //数量比对
                for (int l = 0; l < dataGridView1.RowCount; l++)
                {
                    for (int m = 0; m < dataGridView2.RowCount; m++)
                    {
                        if (dataGridView1.Rows[l].DefaultCellStyle.BackColor != Color.Red)
                        {
                            //费列罗物料代码
                            string yyid = dataGridView2.Rows[m].Cells["商品长代码"].Value.ToString();
                            //友谊物料代码
                            string feid = dataGridView1.Rows[l].Cells["商品长代码"].Value.ToString();
                            //费列罗批次
                            string yyBetchNo = dataGridView2.Rows[m].Cells["批次"].Value.ToString();
                            //友谊批次
                            string feBetchNo = dataGridView1.Rows[l].Cells["批次"].Value.ToString();
                            //批次==批次 and 物料代码 = 物料代码
                            if (yyid == feid && yyBetchNo == feBetchNo)
                            {
                                //友谊结存数量
                                int iyouyi = int.Parse(dataGridView1.Rows[l].Cells["期末结存数量"].Value.ToString());
                                //费列罗数量
                                int iferrero = int.Parse(dataGridView2.Rows[m].Cells["期末结存数量"].Value.ToString());
                                //数量不相等
                                if (iyouyi != iferrero)
                                {
                                    dataGridView1.Rows[l].DefaultCellStyle.BackColor = Color.Yellow;
                                    dataGridView2.Rows[m].DefaultCellStyle.BackColor = Color.Yellow;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }


        /// <summary>
        /// 打开‘友谊‘即时库存Excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpen_Click(object sender, EventArgs e)
        {
            Excel2DataTable common = new Excel2DataTable();
            dt = common.ExcelFile2DataTable();
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    dt.Rows.RemoveAt(index: dt.Rows.Count - 1);
                    dataGridView1.DataSource = dt;
                }
            };
        }
    }
}