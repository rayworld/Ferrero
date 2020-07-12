using DevComponents.DotNetBar;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Ferrero
{
    public partial class FrmSheetsSelecter : Office2007Form
    {
        public string[] SheetList { get; set; }

        public string SelectedSheetName { get; set; }
        public FrmSheetsSelecter()
        {
            InitializeComponent();
        }

        private void frmSheetsSelecter_Load(object sender, EventArgs e)
        {
            listView1.View = View.Details;

            if(SheetList.Length > 0)
            {
                foreach (string ss in SheetList)
                {
                    listView1.Items.Add(ss);
                }
            }
            listView1.Columns.Add("工作簿", 300, HorizontalAlignment.Left);
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection indexes = listView1.SelectedIndices;
            if (indexes !=null && indexes.Count ==1)
            {
                //MessageBox.Show(listView1.Items[indexes[0]].Text);
                SelectedSheetName = listView1.Items[indexes[0]].Text;
                if (SelectedSheetName != "")
                {
                    this.Close();
                }
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection indexes = listView1.SelectedIndices;
            if (indexes != null && indexes.Count == 1)
            {
                //MessageBox.Show(listView1.Items[indexes[0]].Text);
                SelectedSheetName = listView1.Items[indexes[0]].Text;
                if (SelectedSheetName != "")
                {
                    this.Close();
                }
            }
        }
    }
}
