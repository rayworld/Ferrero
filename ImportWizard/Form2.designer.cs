namespace Ferrero
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.controlContainerItem2 = new DevComponents.DotNetBar.ControlContainerItem();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.controlContainerItem1 = new DevComponents.DotNetBar.ControlContainerItem();
            this.itemContainer1 = new DevComponents.DotNetBar.ItemContainer();
            this.controlContainerItem3 = new DevComponents.DotNetBar.ControlContainerItem();
            this.bar1 = new DevComponents.DotNetBar.Bar();
            this.btnOpen = new DevComponents.DotNetBar.ButtonItem();
            this.btnCheck = new DevComponents.DotNetBar.ButtonItem();
            this.btnImport = new DevComponents.DotNetBar.ButtonItem();
            this.btnDelete = new DevComponents.DotNetBar.ButtonItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bar1)).BeginInit();
            this.SuspendLayout();
            // 
            // controlContainerItem2
            // 
            this.controlContainerItem2.AllowItemResize = false;
            this.controlContainerItem2.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways;
            this.controlContainerItem2.Name = "controlContainerItem2";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 25);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(703, 321);
            this.dataGridView1.TabIndex = 1;
            // 
            // controlContainerItem1
            // 
            this.controlContainerItem1.AllowItemResize = false;
            this.controlContainerItem1.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways;
            this.controlContainerItem1.Name = "controlContainerItem1";
            // 
            // itemContainer1
            // 
            // 
            // 
            // 
            this.itemContainer1.BackgroundStyle.Class = "";
            this.itemContainer1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.itemContainer1.Name = "itemContainer1";
            this.itemContainer1.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.controlContainerItem3});
            // 
            // controlContainerItem3
            // 
            this.controlContainerItem3.AllowItemResize = false;
            this.controlContainerItem3.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways;
            this.controlContainerItem3.Name = "controlContainerItem3";
            // 
            // bar1
            // 
            this.bar1.AntiAlias = true;
            this.bar1.Dock = System.Windows.Forms.DockStyle.Top;
            this.bar1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.bar1.Items.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.btnOpen,
            this.btnCheck,
            this.btnImport,
            this.btnDelete});
            this.bar1.Location = new System.Drawing.Point(0, 0);
            this.bar1.Margin = new System.Windows.Forms.Padding(2);
            this.bar1.Name = "bar1";
            this.bar1.Size = new System.Drawing.Size(703, 25);
            this.bar1.Stretch = true;
            this.bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.bar1.TabIndex = 0;
            this.bar1.TabStop = false;
            this.bar1.Text = "bar1";
            // 
            // btnOpen
            // 
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Text = "打开(&O)";
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // btnCheck
            // 
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Text = "导入测试(&T)";
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // btnImport
            // 
            this.btnImport.Name = "btnImport";
            this.btnImport.Text = "导入销售出库单(&I)";
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Text = "删除历史数据(&D)";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(703, 346);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.bar1);
            this.DoubleBuffered = true;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bar1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private DevComponents.DotNetBar.ControlContainerItem controlContainerItem2;
        private DevComponents.DotNetBar.ControlContainerItem controlContainerItem1;
        private DevComponents.DotNetBar.ItemContainer itemContainer1;
        private DevComponents.DotNetBar.ControlContainerItem controlContainerItem3;
        private DevComponents.DotNetBar.Bar bar1;
        private DevComponents.DotNetBar.ButtonItem btnOpen;
        private DevComponents.DotNetBar.ButtonItem btnCheck;
        private DevComponents.DotNetBar.ButtonItem btnImport;
        private DevComponents.DotNetBar.ButtonItem btnDelete;
    }
}