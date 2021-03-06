﻿using DevComponents.DotNetBar;
using Ferrero.BLL;
using Ferrero.Common;
using Ferrero.Model;
using Ray.Framework.Converter;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Ferrero
{
    public partial class Form1  : Office2007Form 
    {
        //导入数据缓存
        DataTable dt = new DataTable();
        //用户名
        public string UserName { get; set; }
        //子公司名称
        public string SubCompany{get;set;}
        
        //数据库连接名称
        public string sConnectionString { get; set; }
        
        public Form1(string userName222, string subCompany, string connectionString)
        {
            this.UserName = userName222;
            this.SubCompany = subCompany;
            this.sConnectionString = connectionString;
            InitializeComponent();
        }

        #region 公共事件

        /// <summary>
        /// 显示行号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var dgv = (DataGridView)sender;
            if (dgv.RowHeadersVisible)
            {
                Rectangle rect = new Rectangle(e.RowBounds.Left, e.RowBounds.Top,
                               dgv.RowHeadersWidth, e.RowBounds.Height);
                rect.Inflate(-2, -2);
                TextRenderer.DrawText(e.Graphics,
                (e.RowIndex + 1).ToString(),
                e.InheritedRowStyle.Font,
                rect, e.InheritedRowStyle.ForeColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter
                );
            }
        }

        /// <summary>
        /// 删除数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            frmDelateConfirm frm = new frmDelateConfirm(1, sConnectionString);
            frm.ShowDialog();
        }

        /// <summary>
        /// 窗口加载事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            btnDelete.Visible = (UserName.ToLower() == "administrator") ? true : false;

            //sConnectionString = SubCompany == "" ? "" : SubCompany;

        }

        /// <summary>
        /// 导入数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void biImporData_Click(object sender, EventArgs e)
        {

            //合并Open\check和Import代码
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Microsoft Excel 文件|*.xlsx;*.xls|所有文件|*.*"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(openFileDialog.FileName))
            {
                string[] SheetNames = ExcelHelper.GetExcelSheetNames(openFileDialog.FileName);
                if (SheetNames != null && SheetNames.Length > 1)//Excel有不只一个sheet，弹出选择窗口
                {
                    FrmSheetsSelecter frmSheetsSelecter = new FrmSheetsSelecter();
                    frmSheetsSelecter.SheetList = SheetNames;
                    if (frmSheetsSelecter.ShowDialog() != DialogResult.Retry)
                    {
                        if (frmSheetsSelecter.SelectedSheetName != "")
                        {
                            dt = ExcelHelper.LoadDataFromExcel(openFileDialog.FileName, frmSheetsSelecter.SelectedSheetName);
                        }
                        else
                        {
                            dt = (DataTable)null;
                            dataGridView1.DataSource = dt;
                            MessageBox.Show("请先选择一个工作簿",
                                "系统错误",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Stop);
                            return;
                        }
                    }
                }
                else if(SheetNames != null && SheetNames.Length == 1)//只有一个sheet
                {
                    string SelectedSheetName = SheetNames[0];
                    dt = ExcelHelper.LoadDataFromExcel(openFileDialog.FileName, SelectedSheetName);
                }
                else
                {
                    dt = (DataTable)null;
                    dataGridView1.DataSource = dt;
                    MessageBox.Show("打开Excel文件出错！",
                        "系统错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                dt = (DataTable)null;
                dataGridView1.DataSource = dt;
                MessageBox.Show("请先打开Excel文件或者文件出错！",
                    "系统错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (dt != null && dt.Rows.Count > 0)
            {
                ///去掉统计行
                dt.Rows.RemoveAt(index: dt.Rows.Count - 1);
                //绑定、显示数据
                dataGridView1.DataSource = dt;
                //1、检查Excel文件中列完整性
                string ErrMsg = ExcelHelper.ChkColumnsName(new string[] { "单据编号", "日期", "商品长代码", "批号", "单位", "实收数量", "单价", "收料仓库", "基本单位实收数量", "验收", "保管", "制单" }, dt);
                if (ErrMsg == "")
                {
                    //2、检查数据同步情况
                    int ErrCount = checkData(dt);
                    if (ErrCount == 0)
                    {
                        //3、如果检查数据成功就开始导入数据
                        if (SubCompany.ToLower() == "wuhan")
                        {
                            Excel2DB(dt);
                        }
                        else
                        {
                            Excel2DB1(dt);
                        }
                    }
                    else
                    {
                        MessageBoxEx.Show(String.Format("总共测试数据 {0} 行，其中测试失败记录有 {1} 项！", dt.Rows.Count, ErrCount),
                            "系统提示",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Stop);
                        return;
                    }
                }
                else
                {
                    //dt = (DataTable)null;
                    //dataGridView1.DataSource = dt;
                    MessageBox.Show("请检查Excel文件，不存在的字段有:" + ErrMsg.Substring(0, ErrMsg.Length - 1), 
                        "系统提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop);
                    return;
                }
            }
            else
            {
                dt = (DataTable)null;
                dataGridView1.DataSource = dt;
                MessageBox.Show(" 打开Excel文件失败或无数据",
                    "系统提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Stop);
                return;
            }
        }

        #endregion

        #region 私有过程

        #region 检查数据

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        private int checkData(DataTable dt)
        {
            int ierrData = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string productName = dt.Rows[i]["商品名称"].ToString();
                string productCode = dt.Rows[i]["商品长代码"].ToString();
                ICStockBillService bStockBill = new  ICStockBillService();
                ICStockBillEntryService bStockBillEntry = new ICStockBillEntryService();

                //单位编号
                string strUnit = dt.Rows[i]["单位"].ToString();
                if (bStockBillEntry.GetUnitID(sConnectionString,strUnit) == 0)
                {
                    dataGridView1.Rows[i].Cells["单位"].Style.BackColor = Color.Red;
                    ierrData++;
                }

                if (bStockBillEntry.checkProductID(sConnectionString,productCode, productName) == false)
                {
                    dataGridView1.Rows[i].Cells["商品长代码"].Style.BackColor = Color.Red;
                    ierrData++;
                }
            }
            return ierrData;
        }
        #endregion

        #region Excel2DB武汉

        /// <summary>
        /// 武汉分公司导入
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private bool Excel2DB(DataTable dt)
        {
            string fBillNo = "";// 主键列
            //int fInterID = 0;//内联ID,外键列
            int fEntryID = 0;//明细表记录行号
            int successCount = 0;//成功完成导入的个数
            int iSuccessBillCount = 0;
            //string fBillNo = "";//金蝶收货单号
            bool retVal = false;
            ICStockBillService bICStockBill = new ICStockBillService();
            ICStockBillEntryService bICStockBillEntry = new ICStockBillEntryService();


            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ICStockBill mICStockBill = new ICStockBill();

                //商品代码
                int fItemID = bICStockBillEntry.GetfItemId(sConnectionString, dt.Rows[i]["商品长代码"].ToString());
                if (fItemID > 0)
                {
                    if (fBillNo != dt.Rows[i]["单据编号"].ToString())
                    {
                        //单据号                        
                        mICStockBill.FBillNo = dt.Rows[i]["单据编号"].ToString();
                        fBillNo = mICStockBill.FBillNo;
                        //内联编号
                        ///fInterID  = bICStockBill.GetMaxFInterID(sConnectionString);
                        bICStockBill.UpdateFInterID(sConnectionString);
                        mICStockBill.FInterID = bICStockBill.GetMaxFInterID(sConnectionString);

                        //fROB
                        mICStockBill.FROB = decimal.Parse(dt.Rows[i]["实收数量"].ToString()) > 0 ? 1 : -1;
                        //日期
                        mICStockBill.FDate = DateTime.Parse(dt.Rows[i]["日期"].ToString());
                        //单据类型
                        mICStockBill.FTranType = 1;
                        
                        mICStockBill.FSupplyID = 2491;//上海费列罗;
                        mICStockBill.FFManagerID = 429;
                        mICStockBill.FSManagerID = 429;
                        mICStockBill.FBillerID = 16394;


                        ///if (InsertStockBill(sConnectionString, fInterID, fBillNo, fdate, frob) == true && bStockBill.UpdateFInterID(sConnectionString) == true && bStockBill.UpdateFBillNo(sConnectionString,iBillNo + 1, "WIN+" + iBillNo.ToString().PadLeft(6, '0'), 1) == true)
                        try
                        {
                            retVal = bICStockBill.Add(sConnectionString, mICStockBill);
                            //bICStockBill.UpdateFInterID(sConnectionString);
                        }
                        catch (Exception ex)
                        {
                            MessageBoxEx.Show(this, "导入主表数据失败:" + ex.Message, "系统信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            if (retVal == true)
                            {
                                fEntryID = 1;
                                iSuccessBillCount++;
                            }
                        }
                    }

                    ICStockBillEntry mICStockBillEntry = new ICStockBillEntry();
                    //SetPropertyDefaultValue4Detail(mICStockBillEntry);
                    ///内联ID
                    mICStockBillEntry.FInterID = bICStockBill.GetMaxFInterID(sConnectionString);
                    //数量
                    mICStockBillEntry.FQty = decimal.Parse(dt.Rows[i]["基本单位实收数量"].ToString());
                    //金额
                    mICStockBillEntry.FAmount = decimal.Parse(dt.Rows[i]["金额"].ToString());
                    //价格
                    mICStockBillEntry.FPrice = mICStockBillEntry.FAmount / mICStockBillEntry.FQty;
                    //仓库ID
                    mICStockBillEntry.FDCStockID = bICStockBill.GetStockId(sConnectionString, dt.Rows[i]["收料仓库"].ToString());
                    //批号
                    mICStockBillEntry.FBatchNo = dt.Rows[i]["批号"].ToString();
                    //单位编号
                    mICStockBillEntry.FUnitID = bICStockBillEntry.GetUnitID(sConnectionString, dt.Rows[i]["单位"].ToString());//单位ID
                    //商品编号
                    mICStockBillEntry.FItemID = fItemID;
                    //输入单价
                    mICStockBillEntry.FAuxPrice = decimal.Parse(dt.Rows[i]["单价"].ToString());
                    //输入数量
                    mICStockBillEntry.FAuxQty = decimal.Parse(dt.Rows[i]["实收数量"].ToString());
                    //条目序号
                    mICStockBillEntry.FEntryID = fEntryID;
                    //销售单价
                    mICStockBillEntry.FConsignPrice = 0;
                    //销售金额
                    mICStockBillEntry.FConsignAmount = 0;

                    mICStockBillEntry.FEntrySelfB0158 = "";
                    mICStockBillEntry.FEntrySelfB0160 = "";
                    mICStockBillEntry.FEntrySelfB0161 = "";
                    mICStockBillEntry.FEntrySelfB0164 = "";
                    mICStockBillEntry.FEntrySelfB0165 = "";
                    mICStockBillEntry.FEntrySelfB0166 = "";
                    mICStockBillEntry.FEntrySelfB0167 = "";
                    mICStockBillEntry.FEntrySelfB0168 = "";
                    mICStockBillEntry.FEntrySelfB0169 = "";
                    mICStockBillEntry.FEntrySelfB0171 = "";
                    mICStockBillEntry.FEntrySelfB0170 = 0;
                    mICStockBillEntry.FEntrySelfB0162 = 0;
                    mICStockBillEntry.FEntrySelfB0172 = 0;
                    mICStockBillEntry.FEntrySelfB0173 = 0;
                    mICStockBillEntry.FEntrySelfB0157 = 0;
                    mICStockBillEntry.FEntrySelfB0163 = 0;
                    mICStockBillEntry.FEntrySelfB0159 = 0;

                    //写Detail表
                    try
                    {
                        retVal = bICStockBillEntry.Add(sConnectionString, mICStockBillEntry);
                    }
                    catch (Exception ex)
                    {
                        MessageBoxEx.Show(this, "导入明细表数据失败:" + ex.Message, "系统信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (retVal == true)
                        {
                            fEntryID++;
                            successCount++;
                        }
                    }
                }
                else
                {
                    MessageBox.Show(dt.Rows[i]["商品长代码"].ToString() + "商品编号不存在！");
                    dataGridView1.Rows[i].Selected = true;
                }
            }

            MessageBox.Show("总共有 " + dt.Rows.Count.ToString() + " 条记录," + "导入失败 " + (dt.Rows.Count - successCount).ToString() + " 条！");

            return retVal;
        }
        #endregion

        #region Excel2DB1 宜昌、襄樊、沙市

        /// <summary>
        /// 宜昌襄樊分公司导入
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private bool Excel2DB1(DataTable dt)
        {
            string fBillNo = "";// 主键列
            //int fInterID = 0;//内联ID,外键列
            int fEntryID = 0;//明细表记录行号
            int successCount = 0;//成功完成导入的个数
            int iSuccessBillCount = 0;
            //string fBillNo = "";//金蝶收货单号
            bool retVal = false;
            ICStockBillService bICStockBill = new ICStockBillService();
            ICStockBillEntryService bICStockBillEntry = new ICStockBillEntryService();


            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ICStockBill mICStockBill = new ICStockBill();

                //商品代码
                int fItemID = bICStockBillEntry.GetfItemId(sConnectionString, dt.Rows[i]["商品长代码"].ToString());
                if (fItemID > 0)
                {
                    if (fBillNo != dt.Rows[i]["单据编号"].ToString())
                    {
                        //单据号                        
                        mICStockBill.FBillNo = dt.Rows[i]["单据编号"].ToString();
                        fBillNo = mICStockBill.FBillNo;
                        //内联编号
                        ///fInterID  = bICStockBill.GetMaxFInterID(sConnectionString);
                        bICStockBill.UpdateFInterID(sConnectionString);
                        mICStockBill.FInterID = bICStockBill.GetMaxFInterID(sConnectionString);

                        //fROB
                        mICStockBill.FROB = decimal.Parse(dt.Rows[i]["实收数量"].ToString()) > 0 ? 1 : -1;
                        //日期
                        mICStockBill.FDate = DateTime.Parse(dt.Rows[i]["日期"].ToString());
                        //单据类型
                        mICStockBill.FTranType = 1;

                        if(SubCompany.ToLower()== "yichang" )
                        {
                            mICStockBill.FSupplyID = 4478;//武汉友谊
                            mICStockBill.FFManagerID = 4175;
                            mICStockBill.FSManagerID = 4174;
                            mICStockBill.FBillerID = 16394;
                        }
                        else if (SubCompany.ToLower() == "xiangfan")//
                        {
                            mICStockBill.FSupplyID = 198;//武汉友谊
                            mICStockBill.FFManagerID = 4128;
                            mICStockBill.FSManagerID = 4128;
                            mICStockBill.FBillerID = 16394;
                        }
                        else //jingzhou 
                        {
                            mICStockBill.FSupplyID = 198;//武汉友谊
                            mICStockBill.FFManagerID = 4395;
                            mICStockBill.FSManagerID = 4395;
                            mICStockBill.FBillerID = 16394;
                        }


                        ///if (InsertStockBill(sConnectionString, fInterID, fBillNo, fdate, frob) == true && bStockBill.UpdateFInterID(sConnectionString) == true && bStockBill.UpdateFBillNo(sConnectionString,iBillNo + 1, "WIN+" + iBillNo.ToString().PadLeft(6, '0'), 1) == true)
                        try
                        {
                            retVal = bICStockBill.Add(sConnectionString, mICStockBill);
                            //bICStockBill.UpdateFInterID(sConnectionString);
                        }
                        catch (Exception ex)
                        {
                            MessageBoxEx.Show(this, "导入主表数据失败:" + ex.Message, "系统信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            if (retVal == true)
                            {
                                fEntryID = 1;
                                iSuccessBillCount++;
                            }
                        }
                    }

                    ICStockBillEntry mICStockBillEntry = new ICStockBillEntry();
                    //SetPropertyDefaultValue4Detail(mICStockBillEntry);
                    ///内联ID
                    mICStockBillEntry.FInterID = bICStockBill.GetMaxFInterID(sConnectionString);
                    //数量
                    mICStockBillEntry.FQty = decimal.Parse(dt.Rows[i]["基本单位实收数量"].ToString());
                    //金额
                    mICStockBillEntry.FAmount = decimal.Parse(dt.Rows[i]["金额"].ToString());
                    //价格
                    mICStockBillEntry.FPrice = mICStockBillEntry.FAmount / mICStockBillEntry.FQty;
                    //仓库ID
                    mICStockBillEntry.FDCStockID = bICStockBill.GetStockId(sConnectionString, dt.Rows[i]["收料仓库"].ToString());
                    //批号
                    mICStockBillEntry.FBatchNo = dt.Rows[i]["批号"].ToString();
                    //单位编号
                    mICStockBillEntry.FUnitID = bICStockBillEntry.GetUnitID(sConnectionString, dt.Rows[i]["单位"].ToString());//单位ID
                    //商品编号
                    mICStockBillEntry.FItemID = fItemID;
                    //输入单价
                    mICStockBillEntry.FAuxPrice = decimal.Parse(dt.Rows[i]["单价"].ToString());
                    //输入数量
                    mICStockBillEntry.FAuxQty = decimal.Parse(dt.Rows[i]["实收数量"].ToString());
                    //条目序号
                    mICStockBillEntry.FEntryID = fEntryID;
                    //销售单价
                    mICStockBillEntry.FConsignPrice = 0;
                    //销售金额
                    mICStockBillEntry.FConsignAmount = 0;

                    ///mICStockBillEntry.FEntrySelfB0159 = 0;
                    ///mICStockBillEntry.FEntrySelfB0161 = "";
                    ///mICStockBillEntry.FEntrySelfB0162 = 0;
                    ///mICStockBillEntry.FEntrySelfB0163 = 0;
                    ///mICStockBillEntry.FEntrySelfB0166 = "";
                    ///mICStockBillEntry.FEntrySelfB0167 = "";
                    ///mICStockBillEntry.FEntrySelfB0168 = "";
                    ///mICStockBillEntry.FEntrySelfB0169 = "";
                    ///mICStockBillEntry.FEntrySelfB0170 = 0;
                    ///mICStockBillEntry.FEntrySelfB0172 = 0;
                    ///mICStockBillEntry.FEntrySelfB0171 = "";
                    ///mICStockBillEntry.FEntrySelfB0163 = 0;
                    ///mICStockBillEntry.FEntrySelfB0173 = 0;
                    ///mICStockBillEntry.FEntrySelfB0174 = 0;
                    ///mICStockBillEntry.FEntrySelfB0158 = "";
                    ///mICStockBillEntry.FEntrySelfB0163 = 0;
                    ///mICStockBillEntry.FEntrySelfB0159 = 0;

                    //写Detail表
                    try
                    {
                        retVal = bICStockBillEntry.Add(sConnectionString, mICStockBillEntry);
                    }
                    catch (Exception ex)
                    {
                        MessageBoxEx.Show(this, "导入明细表数据失败:" + ex.Message, "系统信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (retVal == true)
                        {
                            fEntryID++;
                            successCount++;
                        }
                    }
                }
                else
                {
                    MessageBox.Show(dt.Rows[i]["商品长代码"].ToString() + "商品编号不存在！");
                    dataGridView1.Rows[i].Selected = true;
                }
            }

            MessageBox.Show("总共有 " + dt.Rows.Count.ToString() + " 条记录," + "导入失败 " + (dt.Rows.Count - successCount).ToString() + " 条！");

            return retVal;
        }
        #endregion

        #endregion

        #region 无用代码

        ///// <summary>
        ///// 导入Excel文件
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void btnOpenExcel_Click(object sender, EventArgs e)
        //{
        //    Excel2DataTable common = new Excel2DataTable();
        //    ///将Excel文件转成DataTable
        //    dt = common.ExcelFile2DataTable();
        //    if (dt != null)
        //    {
        //        if (dt.Rows.Count > 0)
        //        {
        //            ///去掉统计行
        //            dt.Rows.RemoveAt(index: dt.Rows.Count - 1);
        //            dataGridView1.DataSource = dt;
        //            ///检查系统所需要的列是否存在
        //            ///2015 因四个公司条件一样，所以不区分不同的分公司
        //            //if (SubCompany.ToLower() == "wuhan")
        //            //{
        //            //    common.ChkColumnsName(new string[] {"单据编号", "日期","商品长代码","批号","单位","实收数量","单价","收料仓库","基本单位实收数量","验收","保管","制单"}, dt);
        //            //}
        //            //else if (SubCompany.ToLower() == "yichan")
        //            //{
        //            //    common.ChkColumnsName(new string[] { "单据编号", "日期", "商品长代码", "批号", "单位", "实收数量", "单价", "收料仓库", "基本单位实收数量", "验收", "保管", "制单" }, dt);
        //            //}
        //            //else if (SubCompany.ToLower() == "xiangfan")
        //            //{
        //            //    common.ChkColumnsName(new string[] { "单据编号", "日期", "商品长代码", "批号", "单位", "实收数量", "单价", "收料仓库", "基本单位实收数量", "验收", "保管", "制单" }, dt);
        //            //}
        //            //else
        //            //{ 
        //            //}

        //            ///common.ChkColumnsName(new string[] { "单据编号", "日期", "商品长代码", "批号", "单位", "实收数量", "单价", "收料仓库", "基本单位实收数量", "验收", "保管", "制单" }, dt);
        //        }
        //    }
        //}

        //private void btnImport_Click(object sender, EventArgs e)
        //{
        //    //if (SubCompany.ToLower() == "wuhan")
        //    //{
        //    //    Excel2DB(dt);
        //    //}
        //    //else
        //    //{
        //    //    Excel2DB1(dt);
        //    //}
        //}

        ///// <summary>
        ///// 验证数据是否合法
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void btnCheck_Click(object sender, EventArgs e)
        //{
        //    //checkData(dt);
        //}
        #endregion

    }
}