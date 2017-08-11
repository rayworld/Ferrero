using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;

using Ferrero.BLL;




namespace Ferrero
{
    public partial class frmDelateConfirm : Office2007Form 
    {
        public int TranType { get; set; }
        public string ConnectionName { get; set; }

        public frmDelateConfirm(int tranType, string connectionName)
        {
            TranType = tranType ;
            ConnectionName = connectionName ;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ICStockBillService bll = new  ICStockBillService();
            ICStockBillEntryService  bllEntry = new   ICStockBillEntryService ();
            int retval =0;


            if (dateTimeInput1.Value.ToString() != "0001/1/1 0:00:00" && dateTimeInput2.Value.ToString() != "0001/1/1 0:00:00")
            {
                //Get Max finterId for this period
                //maxid = bll.GetMaxFInterId(dateTimeInput1.Value, dateTimeInput2.Value);
                //Get Min FInterId For this Period
                //minid = bll.GetMinFInterId(dateTimeInput1.Value, dateTimeInput2.Value);
                //数据有效
                //if (minid > 0 && maxid > 0)
                //{
                    //Delete ICStockbill and ICStockBillEntry
                retval = bll.Delete(ConnectionName, dateTimeInput1.Value.ToShortDateString(), dateTimeInput2.Value.ToShortDateString(), TranType);
                    MessageBox.Show(retval > 0 ? "共有 " + retval + " 记录被删除。" : "没有删除任何记录!");

                //}
                //else 
                //{
                ///    MessageBox.Show("请仔细检查起止时间输入是否正确！");
                //}
            }
            else 
            {
                MessageBox.Show("请输入起始时间和结束时间！");
            }
        }

        /// <summary>
        /// 取消
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
