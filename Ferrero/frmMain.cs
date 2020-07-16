using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using DevComponents.DotNetBar;
using Ray.Framework.DBUtility;

namespace Ferrero
{
    public partial class frmMain : Office2007Form
    {
        #region 公共属性
        //用户名
        public string userName = "";
        #endregion

        #region 私有属性
        //连接名称前缀
        private static readonly string ConnectionNameProfix = "Ferrero.Dev.";
        
        //用户认证连接信息
        private static readonly string AccountConnectionString = SqlHelper.GetConnectionString(ConnectionNameProfix + "Account");
        
        //武汉分公司连接信息
        private static readonly string WhConnecitonString = ConnectionNameProfix + "";

        //宜昌分公司连接信息
        private static readonly string YcConnectionString = ConnectionNameProfix + "";
        
        //襄阳分公司连接信息
        private static readonly string XyConnectionString = ConnectionNameProfix + "";

        //荆州分公司连接信息
        private static readonly string JzConnectionString = ConnectionNameProfix + "";

        #endregion

        public frmMain()
        {
            InitializeComponent();
        }


        
        private void frmMain_Load(object sender, EventArgs e)
        {
            frmLogin login = new frmLogin(AccountConnectionString);
            login.StartPosition = FormStartPosition.CenterParent;
            if (login.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                userName = login.UserName;
                labelItem1.Text = "当前用户：" + userName;
            }
            else
            {
                this.Close();
            }
        }

        private void AdvTree1_NodeClick(object sender, DevComponents.AdvTree.TreeNodeMouseEventArgs e)
        {
            string title = "";
            if (AdvTree1.SelectedNode.Parent != null)
            {
                title = AdvTree1.SelectedNode.Text + "_" + AdvTree1.SelectedNode.Parent.Text;
            }
            if(TabIsOpen(title) == true) //过滤掉重复项目
            {
                MessageBox.Show("窗口 \"" + title + "\" 已经打开！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                switch  (AdvTree1.SelectedNode.Name.Trim().ToLower())
                {
                    case "node1":
                        AddTab("导入外购入库单_武汉分公司", new Form1(userName,"Wuhan",WhConnecitonString));
                        break;

                    case "node2":
                        AddTab("导入销售出库单_武汉分公司", new Form2(userName, "Wuhan", WhConnecitonString));
                        break;

                    case "node3":
                        AddTab("库存商品比对_武汉分公司", new Form3());
                        break;

                    case "node6":
                        AddTab("导入外购入库单_宜昌分公司", new Form1(userName,"Yichang",YcConnectionString));
                        break;

                    case "node7":
                        AddTab("导入销售出库单_宜昌分公司", new Form2(userName,"Yichang", YcConnectionString));
                        break;

                    case "node8":
                        AddTab("库存商品比对_宜昌分公司", new Form3());
                        break;

                    case "node10":
                        AddTab("导入外购入库单_襄樊分公司", new Form1(userName,"Xiangfan",XyConnectionString));
                        break;

                    case "node11":
                        AddTab("导入销售出库单_襄樊分公司", new Form2(userName,"Xiangfan", XyConnectionString));
                        break;

                    case "node12":
                        AddTab("库存商品比对_襄樊分公司", new Form3());
                        break;

                    case "node14":
                        AddTab("导入外购入库单_沙市分公司", new Form1(userName, "Jingzhou",JzConnectionString));
                        break;

                    case "node15":
                        AddTab("导入销售出库单_沙市分公司", new Form2(userName, "Jingzhou", JzConnectionString));
                        break;

                    case "node16":
                        AddTab("库存商品比对_沙市分公司", new Form3());
                        break;

                    default :
                        break;
                }
            }
        }


        #region 私有过程

        private void AddTab(string tabName, Form  frm )
        {
            //TabItem tabItem = tabControl1.CreateTab(tabName);
            TabItem tabItem = new TabItem();
            TabControlPanel tabControlPanel = new TabControlPanel();

            tabControlPanel.Dock = DockStyle.Fill;
            tabControlPanel.Name = "tabPanel";
            tabControlPanel.TabItem = tabItem;

            tabItem.AttachedControl = tabControlPanel;
            tabItem.Name = tabName;
            tabItem.Text = tabName;

            tabControl1.Controls.Add(tabControlPanel);
            tabControl1.Tabs.Add(tabItem);

            frm.FormBorderStyle = FormBorderStyle.None;
            frm.TopLevel = false;
            //frm.Parent = tabControlPanel;
            frm.MdiParent = this;
            frm.Dock = System.Windows.Forms.DockStyle.Fill;
            tabControlPanel.Controls.Add(frm);
            frm.Show();
            tabControl1.Refresh();
            //tabControl1.Controls.Add(tabControlPanel);
            tabControl1.SelectedTab = tabItem;

        }
    
        private bool TabIsOpen(string tabName)
        {
            bool RetVal  = false;
            foreach(TabItem  tabItem in tabControl1.Tabs)
            {
                if (tabItem.Text.Trim () == tabName.Trim())
                {
                    RetVal = true;
                    tabControl1.SelectedTab = tabItem;  
                    break;
                }
            }
            return RetVal;
        }

        #endregion
        
        #region 窗口样式切换
        private void mnuStyleOffice2003_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Office2007Blue;
        }

        private void mnuStyleOffice2007Silver_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Office2007Silver;
        }

        private void mnuStyleOffice2007Black_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Office2007Black;
        }

        private void mnuStyleOffice2007VistaGlass_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Office2007VistaGlass;
        }

        private void mnuStyleOffice2010Silver_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Office2010Silver;
        }

        private void mnuStyleOffice2010Blue_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Office2010Blue;
        }

        private void mnuStyleOffice2010Black_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Office2010Black;
        }

        private void mnuStyleWindows7Blue_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Windows7Blue;
        }

        private void mnuStyleVisualStudio2010Blue_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.VisualStudio2010Blue;
        }

        private void mnuStyleMetro_Click(object sender, EventArgs e)
        {
            styleManager1.ManagerStyle = eStyle.Metro;
        }

        #endregion

    }
}
