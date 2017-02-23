namespace IndustrailInjuries
{
    partial class IndustrailInjuries
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(IndustrailInjuries));
            this.pl1 = new System.Windows.Forms.Panel();
            this.plExcel = new System.Windows.Forms.Panel();
            this.btnExcel = new System.Windows.Forms.Button();
            this.btnToExcel = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.cbxConfirm = new System.Windows.Forms.ComboBox();
            this.lblConfirm = new System.Windows.Forms.Label();
            this.dtTo = new System.Windows.Forms.DateTimePicker();
            this.dtFrom = new System.Windows.Forms.DateTimePicker();
            this.labelTo = new System.Windows.Forms.Label();
            this.btnSearch = new System.Windows.Forms.Button();
            this.lblDate = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.btnStatistics = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.txtName = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.tss0 = new System.Windows.Forms.ToolStripSeparator();
            this.txtCurrentPage = new System.Windows.Forms.ToolStripTextBox();
            this.lblPageCount = new System.Windows.Forms.ToolStripLabel();
            this.bn1 = new System.Windows.Forms.BindingNavigator(this.components);
            this.btnPrevious = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnNext = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.btnGotoPage = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.lblSeparate = new System.Windows.Forms.ToolStripLabel();
            this.lblP = new System.Windows.Forms.ToolStripLabel();
            this.tss1 = new System.Windows.Forms.ToolStripSeparator();
            this.bs1 = new System.Windows.Forms.BindingSource(this.components);
            this.dgv1 = new System.Windows.Forms.DataGridView();
            this.pl1.SuspendLayout();
            this.plExcel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bn1)).BeginInit();
            this.bn1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bs1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // pl1
            // 
            this.pl1.Controls.Add(this.plExcel);
            this.pl1.Controls.Add(this.cbxConfirm);
            this.pl1.Controls.Add(this.lblConfirm);
            this.pl1.Controls.Add(this.dtTo);
            this.pl1.Controls.Add(this.dtFrom);
            this.pl1.Controls.Add(this.labelTo);
            this.pl1.Controls.Add(this.btnSearch);
            this.pl1.Controls.Add(this.lblDate);
            this.pl1.Controls.Add(this.lblName);
            this.pl1.Controls.Add(this.btnStatistics);
            this.pl1.Controls.Add(this.btnDelete);
            this.pl1.Controls.Add(this.btnEdit);
            this.pl1.Controls.Add(this.txtName);
            this.pl1.Controls.Add(this.btnAdd);
            this.pl1.Location = new System.Drawing.Point(2, 3);
            this.pl1.Name = "pl1";
            this.pl1.Size = new System.Drawing.Size(773, 62);
            this.pl1.TabIndex = 18;
            // 
            // plExcel
            // 
            this.plExcel.Controls.Add(this.btnExcel);
            this.plExcel.Controls.Add(this.btnToExcel);
            this.plExcel.Controls.Add(this.btnImport);
            this.plExcel.Location = new System.Drawing.Point(670, 3);
            this.plExcel.Name = "plExcel";
            this.plExcel.Size = new System.Drawing.Size(88, 45);
            this.plExcel.TabIndex = 21;
            this.plExcel.MouseLeave += new System.EventHandler(this.plExcel_MouseLeave);
            // 
            // btnExcel
            // 
            this.btnExcel.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnExcel.FlatAppearance.BorderSize = 0;
            this.btnExcel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnExcel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExcel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExcel.ForeColor = System.Drawing.Color.Purple;
            this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExcel.ImageIndex = 3;
            this.btnExcel.Location = new System.Drawing.Point(3, 3);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(83, 40);
            this.btnExcel.TabIndex = 15;
            this.btnExcel.Text = "报表";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.MouseEnter += new System.EventHandler(this.btnExcel_MouseEnter);
            // 
            // btnToExcel
            // 
            this.btnToExcel.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnToExcel.FlatAppearance.BorderSize = 0;
            this.btnToExcel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnToExcel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnToExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToExcel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnToExcel.ForeColor = System.Drawing.Color.Purple;
            this.btnToExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnToExcel.ImageIndex = 3;
            this.btnToExcel.Location = new System.Drawing.Point(3, 3);
            this.btnToExcel.Name = "btnToExcel";
            this.btnToExcel.Size = new System.Drawing.Size(39, 40);
            this.btnToExcel.TabIndex = 15;
            this.btnToExcel.Text = "导出";
            this.btnToExcel.UseVisualStyleBackColor = false;
            this.btnToExcel.Visible = false;
            this.btnToExcel.Click += new System.EventHandler(this.btnToExcel_Click);
            // 
            // btnImport
            // 
            this.btnImport.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnImport.FlatAppearance.BorderSize = 0;
            this.btnImport.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnImport.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImport.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImport.ForeColor = System.Drawing.Color.Purple;
            this.btnImport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnImport.ImageIndex = 3;
            this.btnImport.Location = new System.Drawing.Point(48, 3);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(38, 40);
            this.btnImport.TabIndex = 15;
            this.btnImport.Text = "导入";
            this.btnImport.UseVisualStyleBackColor = false;
            this.btnImport.Visible = false;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // cbxConfirm
            // 
            this.cbxConfirm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxConfirm.FormattingEnabled = true;
            this.cbxConfirm.Location = new System.Drawing.Point(194, 32);
            this.cbxConfirm.Name = "cbxConfirm";
            this.cbxConfirm.Size = new System.Drawing.Size(110, 20);
            this.cbxConfirm.TabIndex = 22;
            // 
            // lblConfirm
            // 
            this.lblConfirm.AutoSize = true;
            this.lblConfirm.Location = new System.Drawing.Point(147, 35);
            this.lblConfirm.Name = "lblConfirm";
            this.lblConfirm.Size = new System.Drawing.Size(41, 12);
            this.lblConfirm.TabIndex = 21;
            this.lblConfirm.Text = "状态：";
            // 
            // dtTo
            // 
            this.dtTo.Location = new System.Drawing.Point(194, 3);
            this.dtTo.Name = "dtTo";
            this.dtTo.Size = new System.Drawing.Size(110, 21);
            this.dtTo.TabIndex = 20;
            // 
            // dtFrom
            // 
            this.dtFrom.Location = new System.Drawing.Point(58, 3);
            this.dtFrom.Name = "dtFrom";
            this.dtFrom.Size = new System.Drawing.Size(110, 21);
            this.dtFrom.TabIndex = 20;
            // 
            // labelTo
            // 
            this.labelTo.AutoSize = true;
            this.labelTo.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelTo.Location = new System.Drawing.Point(174, 6);
            this.labelTo.Name = "labelTo";
            this.labelTo.Size = new System.Drawing.Size(14, 14);
            this.labelTo.TabIndex = 0;
            this.labelTo.Text = "-";
            // 
            // btnSearch
            // 
            this.btnSearch.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnSearch.FlatAppearance.BorderSize = 0;
            this.btnSearch.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnSearch.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSearch.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSearch.ForeColor = System.Drawing.Color.Purple;
            this.btnSearch.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSearch.ImageIndex = 2;
            this.btnSearch.Location = new System.Drawing.Point(310, 7);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(56, 40);
            this.btnSearch.TabIndex = 14;
            this.btnSearch.Text = "查找(&S)";
            this.btnSearch.UseVisualStyleBackColor = false;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // lblDate
            // 
            this.lblDate.AutoSize = true;
            this.lblDate.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDate.Location = new System.Drawing.Point(3, 6);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(49, 14);
            this.lblDate.TabIndex = 0;
            this.lblDate.Text = "日期：";
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblName.Location = new System.Drawing.Point(3, 34);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(49, 14);
            this.lblName.TabIndex = 0;
            this.lblName.Text = "姓名：";
            // 
            // btnStatistics
            // 
            this.btnStatistics.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnStatistics.FlatAppearance.BorderSize = 0;
            this.btnStatistics.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnStatistics.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnStatistics.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStatistics.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnStatistics.ForeColor = System.Drawing.Color.Purple;
            this.btnStatistics.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnStatistics.ImageIndex = 3;
            this.btnStatistics.Location = new System.Drawing.Point(602, 6);
            this.btnStatistics.Name = "btnStatistics";
            this.btnStatistics.Size = new System.Drawing.Size(62, 40);
            this.btnStatistics.TabIndex = 15;
            this.btnStatistics.Text = "费用统计";
            this.btnStatistics.UseVisualStyleBackColor = false;
            this.btnStatistics.Click += new System.EventHandler(this.btnStatistics_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnDelete.FlatAppearance.BorderSize = 0;
            this.btnDelete.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnDelete.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDelete.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnDelete.ForeColor = System.Drawing.Color.Purple;
            this.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDelete.ImageIndex = 3;
            this.btnDelete.Location = new System.Drawing.Point(520, 7);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(56, 40);
            this.btnDelete.TabIndex = 15;
            this.btnDelete.Text = "删除(&D)";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnEdit.FlatAppearance.BorderSize = 0;
            this.btnEdit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnEdit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnEdit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEdit.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnEdit.ForeColor = System.Drawing.Color.Purple;
            this.btnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEdit.ImageIndex = 3;
            this.btnEdit.Location = new System.Drawing.Point(458, 7);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(56, 40);
            this.btnEdit.TabIndex = 15;
            this.btnEdit.Text = "编辑(&E)";
            this.btnEdit.UseVisualStyleBackColor = false;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(58, 32);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(82, 21);
            this.txtName.TabIndex = 1;
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnAdd.FlatAppearance.BorderSize = 0;
            this.btnAdd.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAdd.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdd.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnAdd.ForeColor = System.Drawing.Color.Purple;
            this.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAdd.ImageIndex = 2;
            this.btnAdd.Location = new System.Drawing.Point(397, 7);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(55, 40);
            this.btnAdd.TabIndex = 14;
            this.btnAdd.Text = "新增(&A)";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // tss0
            // 
            this.tss0.Name = "tss0";
            this.tss0.Size = new System.Drawing.Size(6, 25);
            // 
            // txtCurrentPage
            // 
            this.txtCurrentPage.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.txtCurrentPage.Name = "txtCurrentPage";
            this.txtCurrentPage.Size = new System.Drawing.Size(40, 25);
            // 
            // lblPageCount
            // 
            this.lblPageCount.Name = "lblPageCount";
            this.lblPageCount.Size = new System.Drawing.Size(44, 22);
            this.lblPageCount.Text = "总页数";
            // 
            // bn1
            // 
            this.bn1.AddNewItem = null;
            this.bn1.CountItem = null;
            this.bn1.DeleteItem = null;
            this.bn1.Dock = System.Windows.Forms.DockStyle.None;
            this.bn1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnPrevious,
            this.toolStripSeparator1,
            this.btnNext,
            this.tss0,
            this.toolStripSeparator3,
            this.btnGotoPage,
            this.toolStripSeparator2,
            this.txtCurrentPage,
            this.lblSeparate,
            this.lblPageCount,
            this.lblP,
            this.tss1});
            this.bn1.Location = new System.Drawing.Point(2, 594);
            this.bn1.MoveFirstItem = null;
            this.bn1.MoveLastItem = null;
            this.bn1.MoveNextItem = null;
            this.bn1.MovePreviousItem = null;
            this.bn1.Name = "bn1";
            this.bn1.PositionItem = null;
            this.bn1.Size = new System.Drawing.Size(418, 25);
            this.bn1.TabIndex = 20;
            this.bn1.Text = "bindingNavigator1";
            this.bn1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.bn1_ItemClicked);
            // 
            // btnPrevious
            // 
            this.btnPrevious.AutoToolTip = false;
            this.btnPrevious.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnPrevious.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnPrevious.Image = ((System.Drawing.Image)(resources.GetObject("btnPrevious.Image")));
            this.btnPrevious.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPrevious.Name = "btnPrevious";
            this.btnPrevious.Size = new System.Drawing.Size(63, 22);
            this.btnPrevious.Text = "上一页(&P)";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // btnNext
            // 
            this.btnNext.AutoToolTip = false;
            this.btnNext.BackColor = System.Drawing.SystemColors.ControlDark;
            this.btnNext.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(66, 22);
            this.btnNext.Text = "下一页(&N)";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // btnGotoPage
            // 
            this.btnGotoPage.AutoToolTip = false;
            this.btnGotoPage.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.btnGotoPage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnGotoPage.Image = ((System.Drawing.Image)(resources.GetObject("btnGotoPage.Image")));
            this.btnGotoPage.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnGotoPage.Name = "btnGotoPage";
            this.btnGotoPage.Size = new System.Drawing.Size(65, 22);
            this.btnGotoPage.Text = "跳转到(&G)";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // lblSeparate
            // 
            this.lblSeparate.Name = "lblSeparate";
            this.lblSeparate.Size = new System.Drawing.Size(45, 22);
            this.lblSeparate.Text = "页 / 共";
            // 
            // lblP
            // 
            this.lblP.Name = "lblP";
            this.lblP.Size = new System.Drawing.Size(20, 22);
            this.lblP.Text = "页";
            // 
            // tss1
            // 
            this.tss1.Name = "tss1";
            this.tss1.Size = new System.Drawing.Size(6, 25);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(2, 71);
            this.dgv1.MultiSelect = false;
            this.dgv1.Name = "dgv1";
            this.dgv1.RowTemplate.Height = 23;
            this.dgv1.RowTemplate.ReadOnly = true;
            this.dgv1.Size = new System.Drawing.Size(773, 520);
            this.dgv1.TabIndex = 19;
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            // 
            // IndustrailInjuries
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(792, 616);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.bn1);
            this.Controls.Add(this.pl1);
            this.Name = "IndustrailInjuries";
            this.Text = "豪美工伤信息记录";
            this.Load += new System.EventHandler(this.IndustrailInjuries_Load);
            this.pl1.ResumeLayout(false);
            this.pl1.PerformLayout();
            this.plExcel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bn1)).EndInit();
            this.bn1.ResumeLayout(false);
            this.bn1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bs1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel pl1;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.DateTimePicker dtFrom;
        private System.Windows.Forms.DateTimePicker dtTo;
        private System.Windows.Forms.Label labelTo;
        private System.Windows.Forms.BindingSource bs1;
        private System.Windows.Forms.ToolStripSeparator tss0;
        private System.Windows.Forms.ToolStripButton btnPrevious;
        private System.Windows.Forms.ToolStripTextBox txtCurrentPage;
        private System.Windows.Forms.ToolStripLabel lblPageCount;
        private System.Windows.Forms.ToolStripButton btnNext;
        private System.Windows.Forms.BindingNavigator bn1;
        private System.Windows.Forms.ToolStripLabel lblSeparate;
        private System.Windows.Forms.ToolStripButton btnGotoPage;
        private System.Windows.Forms.ToolStripSeparator tss1;
        private System.Windows.Forms.ToolStripLabel lblP;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.Button btnStatistics;
        private System.Windows.Forms.DataGridView dgv1;
        private System.Windows.Forms.Button btnToExcel;
        private System.Windows.Forms.ComboBox cbxConfirm;
        private System.Windows.Forms.Label lblConfirm;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Panel plExcel;
    }
}