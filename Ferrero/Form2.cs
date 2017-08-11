﻿using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using DevComponents.DotNetBar;

using Ferrero.BLL;
using Ferrero.Model;
using Raysoft.Common;
    
namespace Ferrero
{
    public partial class Form2 :Office2007Form   
    {
        DataTable dt = new DataTable();
        public string UserName { get; set; }
        public string SubCompany { get; set; }
        string sConnectionName = "";
        bool retVal = false;


        public Form2(string userName111,string subCompany)
        {
            this.UserName = userName111;
            this.SubCompany = subCompany;
            InitializeComponent();
        }

        /// <summary>
        /// 打开Ferrero销售出库Excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpen_Click(object sender, EventArgs e)
        {
            Excel2DataTable common = new Excel2DataTable();
            ///将Excel文件转成DataTable
            dt = common.ExcelFile2DataTable();
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    ///去掉统计行
                    dt.Rows.RemoveAt(index: dt.Rows.Count - 1);
                    dataGridView1.DataSource = dt;
                    ///检查系统所需要的列是否存在
                    if (SubCompany.ToLower() == "wuhan")
                    {
                            common.ChkColumnsName(new string[] {
                        "单据编号", "日期","发货仓库","销售单价","购货单位代码","订单号","基本单位实发数量",
                        "进货单号","验收单号","订单日期","交货日期","审核日期A","产品长代码",
                        "实发数量","销售单价", "销售金额","批号","单位","F件数","F含量",
                        "F细数","F赠品","F税率","F未税金额","F价税合计","F保质日期","F生产日期",
                        "F单位","F规格","F商品","F包装","F箱数","F零头","F税额","F售价"
                        }, dt);
                    }
                    else
                    {
                        common.ChkColumnsName(new string[] {
                        "单据编号", "日期","发货仓库","销售单价","购货单位代码","订单号","基本单位实发数量",
                        "产品长代码","实发数量","销售单价", "销售金额","批号","单位"
                        }, dt);
                    }    
                }
            }
        }

        #region 私有过程


        /// <summary>
        /// 执行导入
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private void Excel2DB(DataTable dt)
        {
            int successCount = 0;//成功完成导入的个数
            string fBillNo = "";// 主键列
            //int fInterID = 0;//内联ID,外键列
            int fEntryID = 0;//明细表记录行号
            
            ICStockBillService bICStockBill = new ICStockBillService();
            ICStockBillEntryService bICStockBillEntry = new  ICStockBillEntryService();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ICStockBill mICStockBill = new ICStockBill();
                //产品代码
                int fItemID = bICStockBillEntry.GetfItemId(sConnectionName, dt.Rows[i]["产品长代码"].ToString());
                if (fItemID > 0)
                {
                    if (fBillNo != dt.Rows[i]["单据编号"].ToString())
                    {
                        //单据号
                        mICStockBill.FBillNo = dt.Rows[i]["单据编号"].ToString();
                        fBillNo = mICStockBill.FBillNo;
                        //内联编号
                        bICStockBill.UpdateFInterID(sConnectionName);
                        mICStockBill.FInterID = bICStockBill.GetMaxFInterID(sConnectionName);
                        //单据日期
                        mICStockBill.FDate = DateTime.Parse(dt.Rows[i]["日期"].ToString());
                        //红蓝字
                        //mICStockBill.FROB = decimal.Parse(dt.Rows[i]["销售单价"].ToString()) > 0 ? 1 : -1;
                        mICStockBill.FROB = decimal.Parse(dt.Rows[i]["基本单位实发数量"].ToString()) > 0 ? 1 : -1;//2015                        
                        //客户ID
                        mICStockBill.FSupplyID = bICStockBillEntry.GetSupplyID(sConnectionName, dt.Rows[i]["购货单位代码"].ToString());
                        //订单号
                        mICStockBill.FHeadSelfB0146 = dt.Rows[i]["订单号"].ToString();
                        //进货单号 
                        mICStockBill.FHeadSelfB0147 = dt.Rows[i]["进货单号"].ToString();
                        //验收单号
                        mICStockBill.FHeadSelfB0148 = dt.Rows[i]["验收单号"].ToString();
                        //订单日期
                        mICStockBill.FHeadSelfB0149 = dt.Rows[i]["订单日期"].ToString();
                        //交货日期
                        mICStockBill.FHeadSelfB0150 = dt.Rows[i]["交货日期"].ToString();
                        //审核日期A
                        mICStockBill.FHeadSelfB0151 = dt.Rows[i]["审核日期A"].ToString();
                        //系统
                        mICStockBill.FTranType = 21;
                                                
                        switch (this.SubCompany.ToLower())
                        {
                            case "wuhan":
                                mICStockBill.FFManagerID = 429;
                                mICStockBill.FSManagerID = 429;
                                mICStockBill.FBillerID = 16394;
                                break;

                            case "yichang":
                                mICStockBill.FFManagerID = 4175;
                                mICStockBill.FSManagerID = 4174;
                                //mICStockBill.FBillerID = 16395;//yachang
                                break;

                            case "xiangfan":
                                mICStockBill.FFManagerID = 4128;
                                mICStockBill.FSManagerID = 4128;
                                //mICStockBill.FBillerID = 16395;//xiangfan
                                break;

                            default:
                                break;
                        }


                        try
                        {
                             retVal =bICStockBill.Add (sConnectionName,mICStockBill);
                        }
                        catch(Exception ex)
                        {
                            MessageBoxEx.Show(this, "导入主表数据失败:" + ex.Message, "系统信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            if(retVal == true)
                            {
                                fEntryID = 1;
                            }
                        }
                    }

                    ICStockBillEntry mICStockBillEntry = new ICStockBillEntry();

                    ///内联ID
                    mICStockBillEntry.FInterID = bICStockBill.GetMaxFInterID(sConnectionName);
                    ///产品编号
                    mICStockBillEntry.FItemID = bICStockBillEntry.GetfItemId(sConnectionName, dt.Rows[i]["产品长代码"].ToString());
                    ///实发数量
                    mICStockBillEntry.FQty = decimal.Parse(dt.Rows[i]["基本单位实发数量"].ToString());
                    ///销售单价
                    mICStockBillEntry.FPrice = 0; 
                    ///销售金额
                    mICStockBillEntry.FAmount = 0;
                    ///批号
                    mICStockBillEntry.FBatchNo = dt.Rows[i]["批号"].ToString();
                    ///单位
                    mICStockBillEntry.FUnitID = bICStockBillEntry.GetUnitID(sConnectionName, dt.Rows[i]["单位"].ToString()); 
                    ///行号
                    mICStockBillEntry.FEntryID = fEntryID ;
                    //F件数
                    mICStockBillEntry.FEntrySelfB0157 = Decimal.Parse(dt.Rows[i]["F件数"].ToString());
                    ///含量
                    mICStockBillEntry.FEntrySelfB0158 = dt.Rows[i]["F含量"].ToString();
                    ///细数
                    mICStockBillEntry.FEntrySelfB0159 = Decimal.Parse(dt.Rows[i]["F细数"].ToString());
                    //赠品
                    mICStockBillEntry.FEntrySelfB0160 = dt.Rows[i]["F赠品"].ToString();
                    ///税率
                    mICStockBillEntry.FEntrySelfB0161 = dt.Rows[i]["F税率"].ToString();
                    ///未税金额
                    mICStockBillEntry.FEntrySelfB0162 = Decimal.Parse(dt.Rows[i]["F未税金额"].ToString());
                    ///介税合计
                    mICStockBillEntry.FEntrySelfB0163 = Decimal.Parse(dt.Rows[i]["F价税合计"].ToString());
                    ///保质日期
                    mICStockBillEntry.FEntrySelfB0164 = dt.Rows[i]["F保质日期"].ToString();
                    ///生产日期
                    mICStockBillEntry.FEntrySelfB0165 = dt.Rows[i]["F生产日期"].ToString();
                    ///单位
                    mICStockBillEntry.FEntrySelfB0166 = dt.Rows[i]["F单位"].ToString();
                    ///规格
                    mICStockBillEntry.FEntrySelfB0167 = dt.Rows[i]["F规格"].ToString();
                    ///商品
                    mICStockBillEntry.FEntrySelfB0168 = dt.Rows[i]["F商品"].ToString();
                    ///包装
                    mICStockBillEntry.FEntrySelfB0169 = dt.Rows[i]["F包装"].ToString();
                    ///箱数
                    mICStockBillEntry.FEntrySelfB0170 = Decimal.Parse(dt.Rows[i]["F箱数"].ToString());
                    ///零头
                    mICStockBillEntry.FEntrySelfB0171 = dt.Rows[i]["F零头"].ToString();
                    ///件数
                    mICStockBillEntry.FEntrySelfB0172 = Decimal.Parse(dt.Rows[i]["F税额"].ToString());
                    //售价
                    mICStockBillEntry.FEntrySelfB0173 = Decimal.Parse(dt.Rows[i]["F售价"].ToString());
                    //系统
                    mICStockBillEntry.FBrNo = "0";
                    //实发单价
                    mICStockBillEntry.FAuxPrice = 0;
                    //实发数量
                    mICStockBillEntry.FAuxQty = decimal.Parse(dt.Rows[i]["实发数量"].ToString());
                    //销售单价
                    mICStockBillEntry.FConsignPrice = decimal.Parse(dt.Rows[i]["销售单价"].ToString());
                    //销售金额
                    mICStockBillEntry.FConsignAmount = decimal.Parse(dt.Rows[i]["销售金额"].ToString());
                    //发货仓库
                    mICStockBillEntry.FDCStockID = bICStockBill.GetStockId(sConnectionName, dt.Rows[i]["发货仓库"].ToString());

                    //如果销售商编号在列表中也有记录
                    ///Raysoft.Common.Excel2DataTable common = new Excel2DataTable();
                    ///if (common.bIsSpecialCustom(dt.Rows[i]["购货单位代码"].ToString()) == false)
                    ///{
                    ///    //价格取列表中的价格
                    ///    mICStockBillEntry.FPrice = bICStockBillEntry.getSpecialUnitPriceByLongCode(sConnectionName, dt.Rows[i]["产品长代码"].ToString());
                    ///    //金额等于价格X数量
                    ///    mICStockBillEntry.FAmount = mICStockBillEntry.FPrice * mICStockBillEntry.FQty;
                    ///}
                    
                    ///这次交付的版本不包括启用大客户价格
                    mICStockBillEntry.FAmount = mICStockBillEntry.FPrice * mICStockBillEntry.FQty;
                    //写Detail表
                    try
                    {
                        retVal = bICStockBillEntry.Add(sConnectionName, mICStockBillEntry);
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
                    MessageBox.Show(dt.Rows[i]["产品长代码"].ToString() + "产品编号不存在！");
                    dataGridView1.Rows[i].Selected = true;
                }
            }

            MessageBox.Show(String.Format("总共有 {0} 条记录,导入失败 {1} 条！", dt.Rows.Count, dt.Rows.Count - successCount));
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private void Excel2DB1(DataTable dt)
        {
            int successCount = 0;//成功完成导入的个数
            string fBillNo = "";// 主键列
            //int fInterID = 0;//内联ID,外键列
            int fEntryID = 0;//明细表记录行号

            ICStockBillService bICStockBill = new ICStockBillService();
            ICStockBillEntryService bICStockBillEntry = new ICStockBillEntryService();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ICStockBill mICStockBill = new ICStockBill();
                //产品代码
                int fItemID = bICStockBillEntry.GetfItemId(sConnectionName, dt.Rows[i]["产品长代码"].ToString());
                if (fItemID > 0)
                {
                    if (fBillNo != dt.Rows[i]["单据编号"].ToString())
                    {
                        //单据号
                        mICStockBill.FBillNo = dt.Rows[i]["单据编号"].ToString();
                        fBillNo = mICStockBill.FBillNo;
                        //内联编号
                        bICStockBill.UpdateFInterID(sConnectionName);
                        mICStockBill.FInterID = bICStockBill.GetMaxFInterID(sConnectionName);
                        //单据日期
                        mICStockBill.FDate = DateTime.Parse(dt.Rows[i]["日期"].ToString());
                        //红蓝字
                        //mICStockBill.FROB = decimal.Parse(dt.Rows[i]["销售单价"].ToString()) > 0 ? 1 : -1;
                        mICStockBill.FROB = decimal.Parse(dt.Rows[i]["基本单位实发数量"].ToString()) > 0 ? 1 : -1;//2015                        
                        //客户ID
                        mICStockBill.FSupplyID = bICStockBillEntry.GetSupplyID(sConnectionName, dt.Rows[i]["购货单位代码"].ToString());
                        //订单号
                        mICStockBill.FHeadSelfB0146 = dt.Rows[i]["订单号"].ToString();
                        //进货单号 
                        ///mICStockBill.FHeadSelfB0148 = dt.Rows[i]["进货单号A"].ToString();
                        //验收单号
                        ///mICStockBill.FHeadSelfB0149 = dt.Rows[i]["验收单号A"].ToString();
                        //订单日期
                        ///mICStockBill.FHeadSelfB0150 = dt.Rows[i]["订单日期A"].ToString();
                        //交货日期
                        ///mICStockBill.FHeadSelfB0151 = dt.Rows[i]["交货日期A"].ToString();
                        //审核日期A
                        ///mICStockBill.FHeadSelfB0152 = dt.Rows[i]["审核日期A"].ToString();
                        //系统
                        mICStockBill.FTranType = 21;

                        switch (this.SubCompany.ToLower())
                        {
                            case "wuhan":
                                mICStockBill.FFManagerID = 429;
                                mICStockBill.FSManagerID = 429;
                                mICStockBill.FBillerID = 16394;
                                break;

                            case "yichang":
                                mICStockBill.FFManagerID = 4175;
                                mICStockBill.FSManagerID = 4174;
                                mICStockBill.FBillerID = 16394;//yachang
                                break;

                            case "xiangfan":
                                mICStockBill.FFManagerID = 4128;
                                mICStockBill.FSManagerID = 4128;
                                mICStockBill.FBillerID = 16394;//xiangfan
                                break;

                            case "jingzhou":
                                mICStockBill.FFManagerID = 4395;
                                mICStockBill.FSManagerID = 4395;
                                mICStockBill.FBillerID = 16394;//jingzhou
                                break;

                            default:
                                break;
                        }


                        try
                        {
                            retVal = bICStockBill.Add(sConnectionName, mICStockBill);
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
                            }
                        }
                    }

                    ICStockBillEntry mICStockBillEntry = new ICStockBillEntry();

                    ///内联ID
                    mICStockBillEntry.FInterID = bICStockBill.GetMaxFInterID(sConnectionName);
                    ///产品编号
                    mICStockBillEntry.FItemID = bICStockBillEntry.GetfItemId(sConnectionName, dt.Rows[i]["产品长代码"].ToString());
                    ///实发数量
                    mICStockBillEntry.FQty = decimal.Parse(dt.Rows[i]["基本单位实发数量"].ToString());
                    ///销售单价
                    mICStockBillEntry.FPrice = 0;
                    ///销售金额
                    mICStockBillEntry.FAmount = 0;
                    ///批号
                    mICStockBillEntry.FBatchNo = dt.Rows[i]["批号"].ToString();
                    ///单位
                    mICStockBillEntry.FUnitID = bICStockBillEntry.GetUnitID(sConnectionName, dt.Rows[i]["单位"].ToString());
                    ///行号
                    mICStockBillEntry.FEntryID = fEntryID;
                    //F件数
                    ///mICStockBillEntry.FEntrySelfB0158 = Decimal.Parse(dt.Rows[i]["F未税价"].ToString());
                    ///含量
                    ///mICStockBillEntry.FEntrySelfB0159 = dt.Rows[i]["F含量"].ToString();
                    ///细数
                    ///mICStockBillEntry.FEntrySelfB0160 = Decimal.Parse(dt.Rows[i]["F细数"].ToString());
                    //赠品
                    ///mICStockBillEntry.FEntrySelfB0161 = dt.Rows[i]["F赠品"].ToString();
                    ///税率
                    ///mICStockBillEntry.FEntrySelfB0162 = dt.Rows[i]["F税率"].ToString();
                    ///未税金额
                    ///mICStockBillEntry.FEntrySelfB0163 = Decimal.Parse(dt.Rows[i]["F未税金额"].ToString());
                    ///介税合计
                    ///mICStockBillEntry.FEntrySelfB0164 = Decimal.Parse(dt.Rows[i]["F价税合计"].ToString());
                    ///保质日期
                    ///mICStockBillEntry.FEntrySelfB0165 = dt.Rows[i]["F保质日期"].ToString();
                    ///生产日期
                    ///mICStockBillEntry.FEntrySelfB0166 = dt.Rows[i]["F生产日期"].ToString();
                    ///单位
                    ///mICStockBillEntry.FEntrySelfB0167 = dt.Rows[i]["F单位"].ToString();
                    ///规格
                    ///mICStockBillEntry.FEntrySelfB0168 = dt.Rows[i]["F规格"].ToString();
                    ///商品
                    ///mICStockBillEntry.FEntrySelfB0169 = dt.Rows[i]["F商品"].ToString();
                    ///包装
                    ///mICStockBillEntry.FEntrySelfB0170 = dt.Rows[i]["F包装"].ToString();
                    ///箱数
                    ///mICStockBillEntry.FEntrySelfB0171 = Decimal.Parse(dt.Rows[i]["F箱数"].ToString());
                    ///零头
                    ///mICStockBillEntry.FEntrySelfB0172 = dt.Rows[i]["F零头"].ToString();
                    ///件数
                    ///mICStockBillEntry.FEntrySelfB0173 = Decimal.Parse(dt.Rows[i]["F税额"].ToString());
                    //售价
                    ///mICStockBillEntry.FEntrySelfB0174 = Decimal.Parse(dt.Rows[i]["F售价"].ToString());
                    //系统
                    mICStockBillEntry.FBrNo = "0";
                    //实发单价
                    mICStockBillEntry.FAuxPrice = 0;
                    //实发数量
                    mICStockBillEntry.FAuxQty = decimal.Parse(dt.Rows[i]["实发数量"].ToString());
                    //销售单价
                    mICStockBillEntry.FConsignPrice = decimal.Parse(dt.Rows[i]["销售单价"].ToString());
                    //销售金额
                    mICStockBillEntry.FConsignAmount = decimal.Parse(dt.Rows[i]["销售金额"].ToString());
                    //发货仓库
                    mICStockBillEntry.FDCStockID = bICStockBill.GetStockId(sConnectionName, dt.Rows[i]["发货仓库"].ToString());

                    //如果销售商编号在列表中也有记录
                    ///Raysoft.Common.Excel2DataTable common = new Excel2DataTable();
                    ///if (common.bIsSpecialCustom(dt.Rows[i]["购货单位代码"].ToString()) == false)
                    ///{
                    ///    //价格取列表中的价格
                    ///    mICStockBillEntry.FPrice = bICStockBillEntry.getSpecialUnitPriceByLongCode(sConnectionName, dt.Rows[i]["产品长代码"].ToString());
                    ///    //金额等于价格X数量
                    ///    mICStockBillEntry.FAmount = mICStockBillEntry.FPrice * mICStockBillEntry.FQty;
                    ///}

                    ///这次交付的版本不包括启用大客户价格
                    //mICStockBillEntry.FAmount = mICStockBillEntry.FPrice * mICStockBillEntry.FQty;
                    //写Detail表
                    try
                    {
                        retVal = bICStockBillEntry.Add(sConnectionName, mICStockBillEntry);
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
                    MessageBox.Show(dt.Rows[i]["产品长代码"].ToString() + "产品编号不存在！");
                    dataGridView1.Rows[i].Selected = true;
                }
            }

            MessageBox.Show(String.Format("总共有 {0} 条记录,导入失败 {1} 条！", dt.Rows.Count, dt.Rows.Count - successCount));
        }

        private void checkData(DataTable dt)
        {
            int ierrData = 0; 
            for (int i=0; i < dt.Rows.Count; i++)
            {
                string customName = dt.Rows[i]["购货单位"].ToString();
                string customCode = dt.Rows[i]["购货单位代码"].ToString();
                string productName = dt.Rows[i]["产品名称"].ToString();
                string productCode = dt.Rows[i]["产品长代码"].ToString();
                ICStockBillService bStockBill = new  ICStockBillService();
                ICStockBillEntryService bStockBillEntry = new  ICStockBillEntryService();

                //单位编号
                string strUnit = dt.Rows[i]["单位"].ToString();
                if (bStockBillEntry.GetUnitID(sConnectionName,strUnit) == 0)
                {
                    dataGridView1.Rows[i].Cells["单位"].Style.BackColor = Color.Red;
                    ierrData++;
                }

                if (bStockBillEntry.CheckSupplyID(sConnectionName,customCode, customName) == false)
                {
                    dataGridView1.Rows[i].Cells["购货单位代码"].Style.BackColor = Color.Red;
                    ierrData++;
                }

                if (bStockBillEntry.checkProductID(sConnectionName,productCode, productName) == false)
                {
                    dataGridView1.Rows[i].Cells["产品长代码"].Style.BackColor = Color.Red;
                    ierrData++;
                }
            }
            if (ierrData > 0)
            {
                MessageBoxEx.Show(String.Format("总共测试数据 {0} 行，其中测试失败记录有 {1} 项！", dt.Rows.Count, ierrData));
            }
            else
            {
                MessageBoxEx.Show(String.Format("总共测试数据 {0} 行，没有检测到失败数据！", dt.Rows.Count));
            }
        }
        
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            //InsertStockBill(50000, "XOUT222333", new DateTime(2013,4,2),1, 8888);
            //InsertStockBillEntry(50000, 1, 6424, 100, 10, 1000, "21021122", 241, 505);
            ///检查Excel文件有关列是否存在
            if (SubCompany.ToLower() == "wuhan")
            {
                Excel2DB(dt);
            }
            else
            {
                Excel2DB1(dt);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheck_Click(object sender, EventArgs e)
        {
            checkData(dt);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form2_Load(object sender, EventArgs e)
        {            
            btnDelete.Visible = (UserName .ToLower() == "administrator")? true: false;

            if (SubCompany != "" && SubCompany != null)
            {
                sConnectionName = SubCompany;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            frmDelateConfirm frm = new frmDelateConfirm(21,sConnectionName );
            if (frm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //Delete data
            }
        }
    }
}