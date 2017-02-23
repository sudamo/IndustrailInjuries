namespace IndustrailInjuries
{
    partial class frmEdit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmEdit));
            this.lblTitle = new System.Windows.Forms.Label();
            this.dtpAccidentTime = new System.Windows.Forms.DateTimePicker();
            this.lblName = new System.Windows.Forms.Label();
            this.lblDate = new System.Windows.Forms.Label();
            this.lblDuty = new System.Windows.Forms.Label();
            this.lblDepartment = new System.Windows.Forms.Label();
            this.lblBody = new System.Windows.Forms.Label();
            this.lblCategory = new System.Windows.Forms.Label();
            this.lblTotalCostName = new System.Windows.Forms.Label();
            this.lblProcess = new System.Windows.Forms.Label();
            this.lblReason = new System.Windows.Forms.Label();
            this.lblMeasure = new System.Windows.Forms.Label();
            this.lblRemark = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtDuty = new System.Windows.Forms.TextBox();
            this.txtBody = new System.Windows.Forms.TextBox();
            this.rtbProcess = new System.Windows.Forms.RichTextBox();
            this.rtbReason = new System.Windows.Forms.RichTextBox();
            this.rtbMeasure = new System.Windows.Forms.RichTextBox();
            this.rtbRemark = new System.Windows.Forms.RichTextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cb1 = new System.Windows.Forms.ComboBox();
            this.lblSex = new System.Windows.Forms.Label();
            this.dtpAccidentDate = new System.Windows.Forms.DateTimePicker();
            this.lblBirthday = new System.Windows.Forms.Label();
            this.lblTele = new System.Windows.Forms.Label();
            this.txtTele = new System.Windows.Forms.TextBox();
            this.lblIDNO = new System.Windows.Forms.Label();
            this.txtIDNO = new System.Windows.Forms.TextBox();
            this.lblStorer = new System.Windows.Forms.Label();
            this.lblCompCharge = new System.Windows.Forms.Label();
            this.txtStorer = new System.Windows.Forms.TextBox();
            this.txtCompCharge = new System.Windows.Forms.TextBox();
            this.btnToWord = new System.Windows.Forms.Button();
            this.lblPrincipal = new System.Windows.Forms.Label();
            this.txtPrincipal = new System.Windows.Forms.TextBox();
            this.cbxInCategory = new System.Windows.Forms.ComboBox();
            this.btnConfirm = new System.Windows.Forms.Button();
            this.pl1 = new System.Windows.Forms.Panel();
            this.rbtFemale = new System.Windows.Forms.RadioButton();
            this.rbtMale = new System.Windows.Forms.RadioButton();
            this.cbxHospital = new System.Windows.Forms.ComboBox();
            this.txtOccurredCost = new System.Windows.Forms.TextBox();
            this.txtCurrentCost = new System.Windows.Forms.TextBox();
            this.lblOccurredCost = new System.Windows.Forms.Label();
            this.lblHospital = new System.Windows.Forms.Label();
            this.lblCurrentCost = new System.Windows.Forms.Label();
            this.lblTotalCost = new System.Windows.Forms.Label();
            this.btnToExcel = new System.Windows.Forms.Button();
            this.chbNotice = new System.Windows.Forms.CheckBox();
            this.pl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("SimSun", 22F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblTitle.Location = new System.Drawing.Point(201, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(199, 30);
            this.lblTitle.TabIndex = 50;
            this.lblTitle.Text = "新增工伤信息";
            // 
            // dtpAccidentTime
            // 
            this.dtpAccidentTime.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpAccidentTime.Location = new System.Drawing.Point(121, 81);
            this.dtpAccidentTime.Name = "dtpAccidentTime";
            this.dtpAccidentTime.ShowUpDown = true;
            this.dtpAccidentTime.Size = new System.Drawing.Size(132, 21);
            this.dtpAccidentTime.TabIndex = 5;
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblName.Location = new System.Drawing.Point(56, 14);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(59, 16);
            this.lblName.TabIndex = 50;
            this.lblName.Text = "姓名：";
            // 
            // lblDate
            // 
            this.lblDate.AutoSize = true;
            this.lblDate.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDate.Location = new System.Drawing.Point(22, 83);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(93, 16);
            this.lblDate.TabIndex = 50;
            this.lblDate.Text = "工伤时间：";
            // 
            // lblDuty
            // 
            this.lblDuty.AutoSize = true;
            this.lblDuty.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDuty.Location = new System.Drawing.Point(339, 83);
            this.lblDuty.Name = "lblDuty";
            this.lblDuty.Size = new System.Drawing.Size(59, 16);
            this.lblDuty.TabIndex = 50;
            this.lblDuty.Text = "岗位：";
            // 
            // lblDepartment
            // 
            this.lblDepartment.AutoSize = true;
            this.lblDepartment.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDepartment.Location = new System.Drawing.Point(339, 48);
            this.lblDepartment.Name = "lblDepartment";
            this.lblDepartment.Size = new System.Drawing.Size(59, 16);
            this.lblDepartment.TabIndex = 50;
            this.lblDepartment.Text = "部门：";
            // 
            // lblBody
            // 
            this.lblBody.AutoSize = true;
            this.lblBody.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblBody.Location = new System.Drawing.Point(22, 113);
            this.lblBody.Name = "lblBody";
            this.lblBody.Size = new System.Drawing.Size(93, 16);
            this.lblBody.TabIndex = 50;
            this.lblBody.Text = "工伤部位：";
            // 
            // lblCategory
            // 
            this.lblCategory.AutoSize = true;
            this.lblCategory.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblCategory.Location = new System.Drawing.Point(305, 113);
            this.lblCategory.Name = "lblCategory";
            this.lblCategory.Size = new System.Drawing.Size(93, 16);
            this.lblCategory.TabIndex = 50;
            this.lblCategory.Text = "工伤类型：";
            // 
            // lblTotalCostName
            // 
            this.lblTotalCostName.AutoSize = true;
            this.lblTotalCostName.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblTotalCostName.Location = new System.Drawing.Point(322, 210);
            this.lblTotalCostName.Name = "lblTotalCostName";
            this.lblTotalCostName.Size = new System.Drawing.Size(76, 16);
            this.lblTotalCostName.TabIndex = 50;
            this.lblTotalCostName.Text = "总费用：";
            // 
            // lblProcess
            // 
            this.lblProcess.AutoSize = true;
            this.lblProcess.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblProcess.Location = new System.Drawing.Point(22, 244);
            this.lblProcess.Name = "lblProcess";
            this.lblProcess.Size = new System.Drawing.Size(93, 16);
            this.lblProcess.TabIndex = 50;
            this.lblProcess.Text = "工伤过程：";
            // 
            // lblReason
            // 
            this.lblReason.AutoSize = true;
            this.lblReason.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblReason.Location = new System.Drawing.Point(341, 244);
            this.lblReason.Name = "lblReason";
            this.lblReason.Size = new System.Drawing.Size(59, 16);
            this.lblReason.TabIndex = 50;
            this.lblReason.Text = "原因：";
            // 
            // lblMeasure
            // 
            this.lblMeasure.AutoSize = true;
            this.lblMeasure.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblMeasure.Location = new System.Drawing.Point(22, 399);
            this.lblMeasure.Name = "lblMeasure";
            this.lblMeasure.Size = new System.Drawing.Size(93, 16);
            this.lblMeasure.TabIndex = 50;
            this.lblMeasure.Text = "改善对策：";
            // 
            // lblRemark
            // 
            this.lblRemark.AutoSize = true;
            this.lblRemark.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblRemark.Location = new System.Drawing.Point(341, 398);
            this.lblRemark.Name = "lblRemark";
            this.lblRemark.Size = new System.Drawing.Size(59, 16);
            this.lblRemark.TabIndex = 50;
            this.lblRemark.Text = "备注：";
            // 
            // txtName
            // 
            this.txtName.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtName.Location = new System.Drawing.Point(121, 15);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(100, 21);
            this.txtName.TabIndex = 0;
            // 
            // txtDuty
            // 
            this.txtDuty.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDuty.Location = new System.Drawing.Point(404, 84);
            this.txtDuty.Name = "txtDuty";
            this.txtDuty.Size = new System.Drawing.Size(100, 21);
            this.txtDuty.TabIndex = 6;
            // 
            // txtBody
            // 
            this.txtBody.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtBody.Location = new System.Drawing.Point(121, 114);
            this.txtBody.Name = "txtBody";
            this.txtBody.Size = new System.Drawing.Size(132, 21);
            this.txtBody.TabIndex = 7;
            // 
            // rtbProcess
            // 
            this.rtbProcess.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtbProcess.Location = new System.Drawing.Point(25, 263);
            this.rtbProcess.Name = "rtbProcess";
            this.rtbProcess.Size = new System.Drawing.Size(259, 133);
            this.rtbProcess.TabIndex = 14;
            this.rtbProcess.Text = "";
            // 
            // rtbReason
            // 
            this.rtbReason.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtbReason.Location = new System.Drawing.Point(344, 262);
            this.rtbReason.Name = "rtbReason";
            this.rtbReason.Size = new System.Drawing.Size(259, 134);
            this.rtbReason.TabIndex = 15;
            this.rtbReason.Text = "";
            // 
            // rtbMeasure
            // 
            this.rtbMeasure.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtbMeasure.Location = new System.Drawing.Point(25, 418);
            this.rtbMeasure.Name = "rtbMeasure";
            this.rtbMeasure.Size = new System.Drawing.Size(259, 108);
            this.rtbMeasure.TabIndex = 16;
            this.rtbMeasure.Text = "";
            // 
            // rtbRemark
            // 
            this.rtbRemark.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtbRemark.Location = new System.Drawing.Point(344, 418);
            this.rtbRemark.Name = "rtbRemark";
            this.rtbRemark.Size = new System.Drawing.Size(259, 108);
            this.rtbRemark.TabIndex = 17;
            this.rtbRemark.Text = "";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(326, 652);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 22;
            this.btnSave.Text = "保存(&S)";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(529, 652);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 23;
            this.btnCancel.Text = "退出";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cb1
            // 
            this.cb1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb1.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cb1.FormattingEnabled = true;
            this.cb1.Location = new System.Drawing.Point(404, 49);
            this.cb1.Name = "cb1";
            this.cb1.Size = new System.Drawing.Size(132, 20);
            this.cb1.TabIndex = 4;
            // 
            // lblSex
            // 
            this.lblSex.AutoSize = true;
            this.lblSex.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblSex.Location = new System.Drawing.Point(339, 14);
            this.lblSex.Name = "lblSex";
            this.lblSex.Size = new System.Drawing.Size(59, 16);
            this.lblSex.TabIndex = 50;
            this.lblSex.Text = "性别：";
            // 
            // dtpAccidentDate
            // 
            this.dtpAccidentDate.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpAccidentDate.Location = new System.Drawing.Point(121, 46);
            this.dtpAccidentDate.Name = "dtpAccidentDate";
            this.dtpAccidentDate.Size = new System.Drawing.Size(132, 21);
            this.dtpAccidentDate.TabIndex = 3;
            // 
            // lblBirthday
            // 
            this.lblBirthday.AutoSize = true;
            this.lblBirthday.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblBirthday.Location = new System.Drawing.Point(22, 48);
            this.lblBirthday.Name = "lblBirthday";
            this.lblBirthday.Size = new System.Drawing.Size(93, 16);
            this.lblBirthday.TabIndex = 50;
            this.lblBirthday.Text = "工伤日期：";
            // 
            // lblTele
            // 
            this.lblTele.AutoSize = true;
            this.lblTele.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblTele.Location = new System.Drawing.Point(22, 146);
            this.lblTele.Name = "lblTele";
            this.lblTele.Size = new System.Drawing.Size(93, 16);
            this.lblTele.TabIndex = 50;
            this.lblTele.Text = "联系电话：";
            // 
            // txtTele
            // 
            this.txtTele.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtTele.Location = new System.Drawing.Point(121, 147);
            this.txtTele.Name = "txtTele";
            this.txtTele.Size = new System.Drawing.Size(132, 21);
            this.txtTele.TabIndex = 9;
            // 
            // lblIDNO
            // 
            this.lblIDNO.AutoSize = true;
            this.lblIDNO.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblIDNO.Location = new System.Drawing.Point(322, 146);
            this.lblIDNO.Name = "lblIDNO";
            this.lblIDNO.Size = new System.Drawing.Size(76, 16);
            this.lblIDNO.TabIndex = 50;
            this.lblIDNO.Text = "身份证：";
            // 
            // txtIDNO
            // 
            this.txtIDNO.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtIDNO.Location = new System.Drawing.Point(404, 147);
            this.txtIDNO.Name = "txtIDNO";
            this.txtIDNO.Size = new System.Drawing.Size(163, 21);
            this.txtIDNO.TabIndex = 10;
            // 
            // lblStorer
            // 
            this.lblStorer.AutoSize = true;
            this.lblStorer.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblStorer.Location = new System.Drawing.Point(22, 545);
            this.lblStorer.Name = "lblStorer";
            this.lblStorer.Size = new System.Drawing.Size(93, 16);
            this.lblStorer.TabIndex = 50;
            this.lblStorer.Text = "车间主任：";
            // 
            // lblCompCharge
            // 
            this.lblCompCharge.AutoSize = true;
            this.lblCompCharge.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblCompCharge.Location = new System.Drawing.Point(220, 545);
            this.lblCompCharge.Name = "lblCompCharge";
            this.lblCompCharge.Size = new System.Drawing.Size(93, 16);
            this.lblCompCharge.TabIndex = 50;
            this.lblCompCharge.Text = "安全主任：";
            // 
            // txtStorer
            // 
            this.txtStorer.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtStorer.Location = new System.Drawing.Point(121, 546);
            this.txtStorer.Name = "txtStorer";
            this.txtStorer.Size = new System.Drawing.Size(93, 21);
            this.txtStorer.TabIndex = 18;
            // 
            // txtCompCharge
            // 
            this.txtCompCharge.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCompCharge.Location = new System.Drawing.Point(319, 546);
            this.txtCompCharge.Name = "txtCompCharge";
            this.txtCompCharge.Size = new System.Drawing.Size(93, 21);
            this.txtCompCharge.TabIndex = 19;
            // 
            // btnToWord
            // 
            this.btnToWord.Location = new System.Drawing.Point(26, 652);
            this.btnToWord.Name = "btnToWord";
            this.btnToWord.Size = new System.Drawing.Size(75, 23);
            this.btnToWord.TabIndex = 24;
            this.btnToWord.Text = "导出文档";
            this.btnToWord.UseVisualStyleBackColor = true;
            this.btnToWord.Click += new System.EventHandler(this.btnToWord_Click);
            // 
            // lblPrincipal
            // 
            this.lblPrincipal.AutoSize = true;
            this.lblPrincipal.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblPrincipal.Location = new System.Drawing.Point(418, 545);
            this.lblPrincipal.Name = "lblPrincipal";
            this.lblPrincipal.Size = new System.Drawing.Size(76, 16);
            this.lblPrincipal.TabIndex = 50;
            this.lblPrincipal.Text = "责任人：";
            // 
            // txtPrincipal
            // 
            this.txtPrincipal.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtPrincipal.Location = new System.Drawing.Point(500, 546);
            this.txtPrincipal.Name = "txtPrincipal";
            this.txtPrincipal.Size = new System.Drawing.Size(93, 21);
            this.txtPrincipal.TabIndex = 20;
            // 
            // cbxInCategory
            // 
            this.cbxInCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxInCategory.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cbxInCategory.FormattingEnabled = true;
            this.cbxInCategory.Location = new System.Drawing.Point(404, 114);
            this.cbxInCategory.Name = "cbxInCategory";
            this.cbxInCategory.Size = new System.Drawing.Size(132, 20);
            this.cbxInCategory.TabIndex = 8;
            // 
            // btnConfirm
            // 
            this.btnConfirm.Location = new System.Drawing.Point(224, 652);
            this.btnConfirm.Name = "btnConfirm";
            this.btnConfirm.Size = new System.Drawing.Size(75, 23);
            this.btnConfirm.TabIndex = 26;
            this.btnConfirm.Text = "确认";
            this.btnConfirm.UseVisualStyleBackColor = true;
            this.btnConfirm.Click += new System.EventHandler(this.btnConfirm_Click);
            // 
            // pl1
            // 
            this.pl1.Controls.Add(this.rbtFemale);
            this.pl1.Controls.Add(this.rbtMale);
            this.pl1.Controls.Add(this.lblName);
            this.pl1.Controls.Add(this.dtpAccidentTime);
            this.pl1.Controls.Add(this.dtpAccidentDate);
            this.pl1.Controls.Add(this.cbxHospital);
            this.pl1.Controls.Add(this.cbxInCategory);
            this.pl1.Controls.Add(this.lblSex);
            this.pl1.Controls.Add(this.cb1);
            this.pl1.Controls.Add(this.lblDuty);
            this.pl1.Controls.Add(this.lblBody);
            this.pl1.Controls.Add(this.rtbReason);
            this.pl1.Controls.Add(this.lblTele);
            this.pl1.Controls.Add(this.rtbRemark);
            this.pl1.Controls.Add(this.lblProcess);
            this.pl1.Controls.Add(this.rtbMeasure);
            this.pl1.Controls.Add(this.lblStorer);
            this.pl1.Controls.Add(this.rtbProcess);
            this.pl1.Controls.Add(this.lblMeasure);
            this.pl1.Controls.Add(this.txtPrincipal);
            this.pl1.Controls.Add(this.lblDate);
            this.pl1.Controls.Add(this.txtCompCharge);
            this.pl1.Controls.Add(this.lblBirthday);
            this.pl1.Controls.Add(this.txtIDNO);
            this.pl1.Controls.Add(this.lblDepartment);
            this.pl1.Controls.Add(this.txtOccurredCost);
            this.pl1.Controls.Add(this.txtCurrentCost);
            this.pl1.Controls.Add(this.lblCategory);
            this.pl1.Controls.Add(this.txtStorer);
            this.pl1.Controls.Add(this.lblOccurredCost);
            this.pl1.Controls.Add(this.lblHospital);
            this.pl1.Controls.Add(this.lblCurrentCost);
            this.pl1.Controls.Add(this.lblTotalCost);
            this.pl1.Controls.Add(this.lblTotalCostName);
            this.pl1.Controls.Add(this.txtTele);
            this.pl1.Controls.Add(this.lblIDNO);
            this.pl1.Controls.Add(this.lblReason);
            this.pl1.Controls.Add(this.txtBody);
            this.pl1.Controls.Add(this.lblCompCharge);
            this.pl1.Controls.Add(this.txtDuty);
            this.pl1.Controls.Add(this.lblPrincipal);
            this.pl1.Controls.Add(this.lblRemark);
            this.pl1.Controls.Add(this.txtName);
            this.pl1.Location = new System.Drawing.Point(1, 51);
            this.pl1.Name = "pl1";
            this.pl1.Size = new System.Drawing.Size(629, 579);
            this.pl1.TabIndex = 52;
            // 
            // rbtFemale
            // 
            this.rbtFemale.AutoSize = true;
            this.rbtFemale.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtFemale.Location = new System.Drawing.Point(446, 16);
            this.rbtFemale.Name = "rbtFemale";
            this.rbtFemale.Size = new System.Drawing.Size(35, 16);
            this.rbtFemale.TabIndex = 2;
            this.rbtFemale.TabStop = true;
            this.rbtFemale.Text = "女";
            this.rbtFemale.UseVisualStyleBackColor = true;
            // 
            // rbtMale
            // 
            this.rbtMale.AutoSize = true;
            this.rbtMale.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtMale.Location = new System.Drawing.Point(404, 16);
            this.rbtMale.Name = "rbtMale";
            this.rbtMale.Size = new System.Drawing.Size(35, 16);
            this.rbtMale.TabIndex = 1;
            this.rbtMale.TabStop = true;
            this.rbtMale.Text = "男";
            this.rbtMale.UseVisualStyleBackColor = true;
            // 
            // cbxHospital
            // 
            this.cbxHospital.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxHospital.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cbxHospital.FormattingEnabled = true;
            this.cbxHospital.Location = new System.Drawing.Point(121, 178);
            this.cbxHospital.Name = "cbxHospital";
            this.cbxHospital.Size = new System.Drawing.Size(132, 20);
            this.cbxHospital.TabIndex = 11;
            // 
            // txtOccurredCost
            // 
            this.txtOccurredCost.Font = new System.Drawing.Font("SimSun", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtOccurredCost.Location = new System.Drawing.Point(404, 177);
            this.txtOccurredCost.MaxLength = 9;
            this.txtOccurredCost.Name = "txtOccurredCost";
            this.txtOccurredCost.Size = new System.Drawing.Size(132, 23);
            this.txtOccurredCost.TabIndex = 12;
            this.txtOccurredCost.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtOccurredCost_KeyPress);
            this.txtOccurredCost.Leave += new System.EventHandler(this.txtOccurredCost_Leave);
            // 
            // txtCurrentCost
            // 
            this.txtCurrentCost.Font = new System.Drawing.Font("SimSun", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCurrentCost.Location = new System.Drawing.Point(121, 211);
            this.txtCurrentCost.MaxLength = 9;
            this.txtCurrentCost.Name = "txtCurrentCost";
            this.txtCurrentCost.Size = new System.Drawing.Size(132, 23);
            this.txtCurrentCost.TabIndex = 13;
            this.txtCurrentCost.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCurrentCost_KeyPress);
            this.txtCurrentCost.Leave += new System.EventHandler(this.txtCurrentCost_Leave);
            // 
            // lblOccurredCost
            // 
            this.lblOccurredCost.AutoSize = true;
            this.lblOccurredCost.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblOccurredCost.Location = new System.Drawing.Point(288, 177);
            this.lblOccurredCost.Name = "lblOccurredCost";
            this.lblOccurredCost.Size = new System.Drawing.Size(110, 16);
            this.lblOccurredCost.TabIndex = 50;
            this.lblOccurredCost.Text = "已产生费用：";
            // 
            // lblHospital
            // 
            this.lblHospital.AutoSize = true;
            this.lblHospital.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblHospital.Location = new System.Drawing.Point(22, 177);
            this.lblHospital.Name = "lblHospital";
            this.lblHospital.Size = new System.Drawing.Size(93, 16);
            this.lblHospital.TabIndex = 50;
            this.lblHospital.Text = "医疗机构：";
            // 
            // lblCurrentCost
            // 
            this.lblCurrentCost.AutoSize = true;
            this.lblCurrentCost.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblCurrentCost.Location = new System.Drawing.Point(22, 211);
            this.lblCurrentCost.Name = "lblCurrentCost";
            this.lblCurrentCost.Size = new System.Drawing.Size(93, 16);
            this.lblCurrentCost.TabIndex = 50;
            this.lblCurrentCost.Text = "当前费用：";
            // 
            // lblTotalCost
            // 
            this.lblTotalCost.AutoSize = true;
            this.lblTotalCost.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblTotalCost.Location = new System.Drawing.Point(401, 211);
            this.lblTotalCost.Name = "lblTotalCost";
            this.lblTotalCost.Size = new System.Drawing.Size(88, 16);
            this.lblTotalCost.TabIndex = 50;
            this.lblTotalCost.Text = "计算总费用";
            // 
            // btnToExcel
            // 
            this.btnToExcel.Location = new System.Drawing.Point(122, 652);
            this.btnToExcel.Name = "btnToExcel";
            this.btnToExcel.Size = new System.Drawing.Size(75, 23);
            this.btnToExcel.TabIndex = 25;
            this.btnToExcel.Text = "导出报表";
            this.btnToExcel.UseVisualStyleBackColor = true;
            this.btnToExcel.Click += new System.EventHandler(this.btnToExcel_Click);
            // 
            // chbNotice
            // 
            this.chbNotice.AutoSize = true;
            this.chbNotice.Location = new System.Drawing.Point(416, 656);
            this.chbNotice.Name = "chbNotice";
            this.chbNotice.Size = new System.Drawing.Size(89, 16);
            this.chbNotice.TabIndex = 21;
            this.chbNotice.Text = "通知接收人";
            this.chbNotice.UseVisualStyleBackColor = true;
            // 
            // frmEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(631, 696);
            this.ControlBox = false;
            this.Controls.Add(this.chbNotice);
            this.Controls.Add(this.pl1);
            this.Controls.Add(this.btnToExcel);
            this.Controls.Add(this.btnConfirm);
            this.Controls.Add(this.btnToWord);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblTitle);
            this.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmEdit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "工伤信息";
            this.Load += new System.EventHandler(this.frmEdit_Load);
            this.pl1.ResumeLayout(false);
            this.pl1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.DateTimePicker dtpAccidentTime;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.Label lblDuty;
        private System.Windows.Forms.Label lblDepartment;
        private System.Windows.Forms.Label lblBody;
        private System.Windows.Forms.Label lblCategory;
        private System.Windows.Forms.Label lblTotalCostName;
        private System.Windows.Forms.Label lblProcess;
        private System.Windows.Forms.Label lblReason;
        private System.Windows.Forms.Label lblMeasure;
        private System.Windows.Forms.Label lblRemark;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtDuty;
        private System.Windows.Forms.TextBox txtBody;
        private System.Windows.Forms.RichTextBox rtbProcess;
        private System.Windows.Forms.RichTextBox rtbReason;
        private System.Windows.Forms.RichTextBox rtbMeasure;
        private System.Windows.Forms.RichTextBox rtbRemark;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ComboBox cb1;
        private System.Windows.Forms.Label lblSex;
        private System.Windows.Forms.DateTimePicker dtpAccidentDate;
        private System.Windows.Forms.Label lblBirthday;
        private System.Windows.Forms.Label lblTele;
        private System.Windows.Forms.TextBox txtTele;
        private System.Windows.Forms.Label lblIDNO;
        private System.Windows.Forms.TextBox txtIDNO;
        private System.Windows.Forms.Label lblStorer;
        private System.Windows.Forms.Label lblCompCharge;
        private System.Windows.Forms.TextBox txtStorer;
        private System.Windows.Forms.TextBox txtCompCharge;
        private System.Windows.Forms.Button btnToWord;
        private System.Windows.Forms.Label lblPrincipal;
        private System.Windows.Forms.TextBox txtPrincipal;
        private System.Windows.Forms.ComboBox cbxInCategory;
        private System.Windows.Forms.Button btnConfirm;
        private System.Windows.Forms.Panel pl1;
        private System.Windows.Forms.Button btnToExcel;
        private System.Windows.Forms.ComboBox cbxHospital;
        private System.Windows.Forms.Label lblHospital;
        private System.Windows.Forms.RadioButton rbtFemale;
        private System.Windows.Forms.RadioButton rbtMale;
        private System.Windows.Forms.TextBox txtOccurredCost;
        private System.Windows.Forms.TextBox txtCurrentCost;
        private System.Windows.Forms.Label lblOccurredCost;
        private System.Windows.Forms.Label lblCurrentCost;
        private System.Windows.Forms.Label lblTotalCost;
        private System.Windows.Forms.CheckBox chbNotice;
    }
}

