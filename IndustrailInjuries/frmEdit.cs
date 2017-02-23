using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using pblClass;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;


namespace IndustrailInjuries
{
    public partial class frmEdit : Form
    {
        #region Parameter Area
        Regex reg = null;                       //正则表达式
        DateTime pNow = DateTime.Now;           //当前时间
        int pNotice = 0;                        //通知状态

        sdk.WebService sms = new sdk.WebService();
        string sn = "SDK-BBX-010-22986";        //序列号
        string password = ")-df)9-4";           //密码
        string subcode = "";                    //扩展码
        string stime = "";                      //定时时间
        #endregion

        #region Property Area
        private int _InfoId;
        /// <summary>
        /// InfoId
        /// </summary>
        public int InfoId
        {
            set { _InfoId = value; }
            get { return _InfoId; }
        }

        private bool _IsConfirm;
        /// <summary>
        /// InfoId
        /// </summary>
        public bool IsConfirm
        {
            set { _IsConfirm = value; }
            get { return _IsConfirm; }
        }
        private string _mEditType;
        /// <summary>
        /// EditType
        /// </summary>
        public string EditType
        {
            set { _mEditType = value; }
            get { return _mEditType; }
        }
        #endregion

        #region Constructor Function
        public frmEdit()
        {
            InitializeComponent();
        }
        public frmEdit(int pInfoId, string pEditType)
        {
            _InfoId = pInfoId;
            _mEditType = pEditType;
            InitializeComponent();
        }
        #endregion

        #region FormLoad
        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmEdit_Load(object sender, EventArgs e)
        {
            chbNotice.Checked = true;
            FillDepartmentList();//加载下拉选项
            if (EditType == "Edit")//编辑状态
            {
                if (pblinfo.username == "") return;
                lblTitle.Text = "编辑工伤信息";
                string strSQL = @"SELECT A.[Name], A.[CreateDate], A.[AccidentDate], A.[Duty], B.PDID, A.[Body], C.[GroupId] AS GroupId1, D.[GroupId] AS GroupId2, A.OccurredCost, A.CurrentCost, A.[TotalCost], A.[Process], A.[Reason], A.[Measure], A.[Remark], A.Sex, A.Storer, A.CompCharge, A.Tele, A.IDNO, A.Birthday, A.Principal, A.IsConfirm, A.InfoId, A.Notice
                        FROM IndustrailInjuryInfo A
                        INNER JOIN productdepartment B ON A.[Department] = B.PDName
                        INNER JOIN systemCategoryGroup C ON A.InjuryCategory = C.CategoryName AND (C.CategoryId = 1 OR C.GroupId = 1)
                        INNER JOIN systemCategoryGroup D ON A.Hospital = D.CategoryName AND (D.CategoryId = 2 OR D.GroupId = 1)
                        WHERE InfoId = " + InfoId.ToString();
                System.Data.DataTable dt = null;
                try
                {
                    dt = mysql.sqltb(strSQL);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                txtName.Text = dt.Rows[0]["Name"].ToString();
                dtpAccidentTime.Value = DateTime.Parse(dt.Rows[0]["AccidentDate"].ToString());
                dtpAccidentDate.Value = DateTime.Parse(dt.Rows[0]["AccidentDate"].ToString());
                txtDuty.Text = dt.Rows[0]["Duty"].ToString();
                txtBody.Text = dt.Rows[0]["Body"].ToString();
                cbxInCategory.SelectedValue = dt.Rows[0]["GroupId1"].ToString();
                cbxHospital.SelectedValue = dt.Rows[0]["GroupId2"].ToString();
                txtOccurredCost.Text = dt.Rows[0]["OccurredCost"].ToString() == "0.00" ? "" : dt.Rows[0]["OccurredCost"].ToString();
                txtCurrentCost.Text = dt.Rows[0]["CurrentCost"].ToString() == "0.00" ? "" : dt.Rows[0]["CurrentCost"].ToString();
                lblTotalCost.Text = dt.Rows[0]["TotalCost"].ToString() == "0.00" ? "" : dt.Rows[0]["TotalCost"].ToString();
                rtbProcess.Text = dt.Rows[0]["Process"].ToString();
                rtbReason.Text = dt.Rows[0]["Reason"].ToString();
                rtbMeasure.Text = dt.Rows[0]["Measure"].ToString();
                rtbRemark.Text = dt.Rows[0]["Remark"].ToString();
                cb1.SelectedValue = dt.Rows[0]["PDID"].ToString();
                rbtMale.Checked = dt.Rows[0]["Sex"].ToString() == "男" ? true : false;
                rbtFemale.Checked = dt.Rows[0]["Sex"].ToString() == "男" ? false : true;
                txtTele.Text = dt.Rows[0]["Tele"].ToString();
                txtStorer.Text = dt.Rows[0]["Storer"].ToString();
                txtCompCharge.Text = dt.Rows[0]["CompCharge"].ToString();
                txtIDNO.Text = dt.Rows[0]["IDNO"].ToString();
                txtPrincipal.Text = dt.Rows[0]["Principal"].ToString();

                InfoId = int.Parse(dt.Rows[0]["InfoId"].ToString());
                IsConfirm = bool.Parse(dt.Rows[0]["IsConfirm"].ToString());
                pNotice = int.Parse(dt.Rows[0]["Notice"].ToString());

                if (IsConfirm)
                {
                    pl1.Enabled = false;
                    btnConfirm.Text = "激活";
                    btnSave.Enabled = false;
                    chbNotice.Enabled = false;
                }
                else
                {
                    pl1.Enabled = true;
                    btnConfirm.Text = "确认信息";
                    btnSave.Enabled = true;
                    chbNotice.Enabled = true;
                    txtName.Focus();
                    txtName.Select(0, txtName.Text.Length);
                }

                if (pblinfo.v_zz_id != 1)
                {
                    //确认权限确认
                    string strResult = string.Empty;
                    string strRight = @"DECLARE @tmp VARCHAR(4000)
                                SELECT @tmp = Contents FROM [SystemGroup] WHERE GroupId = 1
                                IF EXISTS(SELECT sub FROM dm_split(@tmp) WHERE sub = '" + pblinfo.username.Trim() + @"') SELECT 1
                                ELSE SELECT 0";
                    try
                    {
                        strResult = mysql.ExecuteScalar(strRight).ToString();
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    if (strResult == "1") btnConfirm.Enabled = true;
                    else btnConfirm.Enabled = false;

                    //文档导出权限确认
                    strResult = string.Empty;
                    strRight = @"DECLARE @tmp VARCHAR(4000)
                                SELECT @tmp = Contents FROM [SystemGroup] WHERE GroupId = 2
                                IF EXISTS(SELECT sub FROM dm_split(@tmp) WHERE sub = '" + pblinfo.username.Trim() + @"') SELECT 1
                                ELSE SELECT 0";
                    try
                    {
                        strResult = mysql.ExecuteScalar(strRight).ToString();
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    if (strResult == "1")
                    {
                        btnToWord.Enabled = true;
                        btnToExcel.Enabled = true;
                    }
                    else
                    {
                        btnToWord.Enabled = false;
                        btnToExcel.Enabled = false;
                    }
                }
            }
            else//新增状态
            {
                rbtMale.Checked = true;
                btnConfirm.Visible = false;
                btnToWord.Visible = false;
                btnToExcel.Visible = false;
                lblTotalCost.Text = "";
            }
        }
        #endregion

        #region SaveDate
        /// <summary>
        /// 保存数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            string strTick = string.Empty;
            string strPhoneNO = string.Empty;
            string strContents = string.Empty;
            if (CheckDate())
            {
                string strSQL = string.Empty;
                if (EditType == "Edit")
                {
                    strSQL = @"UPDATE IndustrailInjuryInfo
                    SET AccidentDate = '" + dtpAccidentDate.Value.GetDateTimeFormats()[4].ToString() + " " + dtpAccidentTime.Value.GetDateTimeFormats()[134].ToString() + "', Duty = '" + txtDuty.Text.Trim() + @"',
                    Hospital = '" + (cbxHospital.Text == "请选择" ? "" : cbxHospital.Text) + "', Body = '" + txtBody.Text.Trim() + "', InjuryCategory = '" + (cbxInCategory.Text == "请选择" ? "" : cbxInCategory.Text) + @"',
                    Principal = '" + txtPrincipal.Text.Trim() + "', OccurredCost = " + (txtOccurredCost.Text == "" ? "0" : txtOccurredCost.Text.Trim()) + ", CurrentCost = " + (txtCurrentCost.Text == "" ? "0" : txtCurrentCost.Text.Trim()) + @",
                    Process = '" + rtbProcess.Text.Trim() + "', Tele = '" + txtTele.Text.Trim() + "', Reason = '" + rtbReason.Text.Trim() + @"',
                    Measure = '" + rtbMeasure.Text.Trim() + "', CompCharge = '" + txtCompCharge.Text.Trim() + "', Remark = '" + rtbRemark.Text.Trim() + @"',
                    Sex = '" + (rbtMale.Checked ? "男" : "女") + "', Storer = '" + txtStorer.Text.Trim() + "', Department = '" + cb1.Text + @"',
                    IDNO = '" + txtIDNO.Text.Trim() + "', Name = '" + txtName.Text.Trim() + @"',
                    UpdateUser = '" + pblinfo.username + "', UpdateDate = '" + pNow.ToString() + @"',
                    Notice = " + (pNotice == 0 ? 1 : (pNotice == 2 ? 3 : pNotice)).ToString() + @"
                    WHERE InfoId = " + InfoId.ToString();
                }
                else
                {
                    strSQL = @"INSERT INTO IndustrailInjuryInfo(Name, AccidentDate, Duty, Department, Body, InjuryCategory, Hospital, OccurredCost, CurrentCost, Process, Reason, Measure, Remark, Sex, Storer, CompCharge, Tele, IDNO, Principal, Creator)
                            VALUES('" + txtName.Text.Trim() + "', '" + dtpAccidentDate.Value.GetDateTimeFormats()[4].ToString() + " " + dtpAccidentTime.Value.GetDateTimeFormats()[134].ToString() + "', '" + txtDuty.Text.Trim() + "', '" + cb1.Text + @"',
                            '" + txtBody.Text.Trim() + "', '" + (cbxInCategory.Text == "请选择" ? "" : cbxInCategory.Text) + "', '" + (cbxHospital.Text == "请选择" ? "" : cbxHospital.Text) + @"',
                            '" + (txtOccurredCost.Text.Length > 0 ? txtOccurredCost.Text.Trim() : "0") + @"',
                            '" + (txtCurrentCost.Text.Length > 0 ? txtCurrentCost.Text.Trim() : "0") + @"',
                            '" + rtbProcess.Text.Trim() + "', '" + rtbReason.Text.Trim() + @"',
                            '" + rtbMeasure.Text.Trim() + "', '" + rtbRemark.Text.Trim() + @"',
                            '" + (rbtMale.Checked ? "男" : "女") + "', '" + txtStorer.Text.Trim() + @"',
                            '" + txtCompCharge.Text.Trim() + "', '" + txtTele.Text.Trim() + @"',
                            '" + txtIDNO.Text.Trim() + @"',
                            '" + txtPrincipal.Text.Trim() + @"',
                            '" + pblinfo.username + "')";
                }
                try
                {
                    mysql.sqlcmd(strSQL);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }

                //短信通知
                if (chbNotice.Checked)
                {
                    strPhoneNO = GetPhoneNO("确认接收人");
                    strContents = "【豪美铝业】：已经添加新的工伤信息[" + txtName.Text.Trim() + "-" + dtpAccidentTime.Value.GetDateTimeFormats()[11].ToString() + "]，请确认！ 发送人：" + pblinfo.username;

                    if (SendMSG(strPhoneNO, strContents) == "") strTick = "保存成功并已发送信息通知接收人:[" + GetNoticeName("确认接收人") + "]，关闭此窗口？";
                    else strTick = "保存成功，关闭此窗口？";
                }
                else strTick = "保存成功，关闭此窗口？";

                if (MessageBox.Show(strTick, "存盘", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    Close();
                }
            }
        }
        #endregion

        #region Cancel
        /// <summary>
        /// 取消
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region CheckDate
        /// <summary>
        /// 数据有效性检测
        /// </summary>
        /// <returns></returns>
        private bool CheckDate()
        {
            if (txtName.Text == null || txtName.Text.Length <= 0)
            {
                MessageBox.Show("请输入姓名");
                txtName.Focus();
                return false;
            }
            //if (txtTotalCost.Text.Length > 0)
            //{
            //    reg = new Regex(@"^[0-9]\d*\.\d{0,2}$|^\d*$");
            //    if (!reg.IsMatch(txtTotalCost.Text.Trim()))
            //    {
            //        MessageBox.Show("总费用必须是浮点数");
            //        txtTotalCost.Focus();
            //        txtTotalCost.Select(0, txtTotalCost.Text.Length);
            //        return false;
            //    }
            //}
            if (rtbProcess.Text == null || rtbProcess.Text.Length <= 0)
            {
                MessageBox.Show("请输入受伤过程");
                rtbProcess.Focus();
                return false;
            }
            reg = new Regex(@"^(\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$");   //身份证正则表达式：^(\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$
            if (txtIDNO.Text.Length > 0 && txtIDNO.Text.Length != 15 && txtIDNO.Text.Length != 18)
            {
                MessageBox.Show("身份证号码的长度为15位或18位！");
                txtIDNO.Focus();
                txtIDNO.Select(0, txtIDNO.Text.Length);
                return false;
            }
            if (txtIDNO.Text.Length > 0 && !reg.IsMatch(txtIDNO.Text.Trim()))
            {
                MessageBox.Show("身份证:15位时全为数字，18位前17位为数字，最后一位是校验位，可能为数字或字符X！");
                txtIDNO.Focus();
                txtIDNO.Select(0, txtIDNO.Text.Length);
                return false;
            }
            string strSQL = "SELECT COUNT(*) FROM IndustrailInjuryInfo WHERE IsCancel = 0 AND Name = '" + txtName.Text + "' AND CONVERT(VARCHAR(10), AccidentDate, 120) = '" + dtpAccidentTime.Value.GetDateTimeFormats()[0] + "'";
            if (EditType == "Edit")
            {
                strSQL += " AND InfoId <> " + InfoId.ToString();
            }
            if (int.Parse(mysql.ExecuteScalar(strSQL).ToString()) != 0)
            {
                MessageBox.Show("[" + txtName.Text + ":" + dtpAccidentTime.Value.GetDateTimeFormats()[10] + "]这个条工伤信息已经存在，请重新输入！");
                txtName.Focus();
                txtName.Select(0, txtName.Text.Length);
                return false;
            }
            return true;
        }
        #endregion

        #region FillDepartmentCombox
        /// <summary>
        /// 填充部门下拉框
        /// </summary>
        private void FillDepartmentList()
        {
            //填充部门
            string strSQL = "SELECT PDID, PDName FROM productdepartment WHERE Status = 0";
            System.Data.DataTable dt = null;
            try
            {
                dt = mysql.sqltb(strSQL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            cb1.Items.Clear();
            cb1.DataSource = dt;
            cb1.DisplayMember = "PDName";
            cb1.ValueMember = "PDID";
            cb1.SelectedIndex = 0;

            //填充工伤类型
            strSQL = "SELECT GroupId, CategoryName FROM systemCategoryGroup WHERE CategoryId = 1";
            dt = new System.Data.DataTable();
            try
            {
                dt = mysql.sqltb(strSQL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            System.Data.DataRow dr = dt.NewRow();
            dr["CategoryName"] = "请选择";
            dr["GroupId"] = "1";
            dt.Rows.Add(dr);
            cbxInCategory.Items.Clear();
            cbxInCategory.DataSource = dt;
            cbxInCategory.DisplayMember = "CategoryName";
            cbxInCategory.ValueMember = "GroupId";
            cbxInCategory.SelectedIndex = cbxInCategory.Items.Count - 1;

            //填充医疗机构
            strSQL = "SELECT GroupId, CategoryName FROM systemCategoryGroup WHERE CategoryId = 2";
            dt = new System.Data.DataTable();
            try
            {
                dt = mysql.sqltb(strSQL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            dr = dt.NewRow();
            dr["CategoryName"] = "请选择";
            dr["GroupId"] = "1";
            dt.Rows.Add(dr);
            cbxHospital.Items.Clear();
            cbxHospital.DataSource = dt;
            cbxHospital.DisplayMember = "CategoryName";
            cbxHospital.ValueMember = "GroupId";
            cbxHospital.SelectedIndex = cbxHospital.Items.Count - 1;
        }
        #endregion

        #region Output Word
        /// <summary>
        /// Word 文档导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnToWord_Click(object sender, EventArgs e)
        {            
            //未经确认的信息不能导出文档
            string strConfirm = "SELECT IsConfirm FROM IndustrailInjuryInfo WHERE  InfoId = " + InfoId.ToString();
            try
            {
                if (!bool.Parse(mysql.ExecuteScalar(strConfirm).ToString()))
                {
                    MessageBox.Show("本工伤信息还未被确认，不能导出文档。请等工伤信息被确认后再导出！");
                    return;
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }

            //从数据库下载模板到本地
            object filePath = System.Windows.Forms.Application.StartupPath + "\\WordTemplate";
            if (!Directory.Exists(filePath.ToString()))
            {
                Directory.CreateDirectory(filePath.ToString());
            }
            if (File.Exists(filePath + "\\Int.doc"))
            {
                File.Delete(filePath + "\\Int.doc");
            }
            byte[] tempFile;
            string strGetfile = "SELECT Contents FROM Templates WHERE TempName = 'Int.doc'";
            try
            {
                tempFile = (byte[])mysql.ExecuteScalar(strGetfile);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            FileStream fs = new FileStream(filePath + "\\Int.doc", FileMode.Create, FileAccess.Write);
            fs.Write(tempFile, 0, tempFile.Length);
            fs.Close();

            filePath += "\\Int.doc";

            object savePath = "D:\\工伤报告";
            if (!Directory.Exists(savePath.ToString()))
            {
                Directory.CreateDirectory(savePath.ToString());
            }
            savePath += "\\" + dtpAccidentTime.Value.GetDateTimeFormats()[11].ToString() + txtName.Text.Trim() + "工伤信息.doc";

            //创建一个document.
            _Application oWord = new Microsoft.Office.Interop.Word.Application();
            _Document oDoc;
            object objDocType = WdDocumentType.wdTypeDocument;
            object type = WdBreakType.wdSectionBreakContinuous;
            object oMissing = Missing.Value;

            object readOnly = false;
            object isVisible = false;

            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            _Document openWord;
            openWord = oWord.Documents.Open(ref filePath, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            openWord.Select();
            openWord.Sections[1].Range.Copy();

            //插入换行符    
            oDoc.Sections[1].Range.PasteAndFormat(WdRecoveryType.wdPasteDefault);
            openWord.Close(ref oMissing, ref oMissing, ref oMissing);

            List<string> listold = new List<string>();
            List<string> listnew = new List<string>();
            string[] strListOle = { "Name", "Sex", "Birthday", "DP", "Duty", "InTime", "Body", "Tele", "IDNO", "Process", "Reason", "Measure","Principal", "Storer", "CompCharge" };
            string[] strListNew ={txtName.Text.Trim(), rbtMale.Checked ? "男" : "女","",cb1.Text,txtDuty.Text.Trim(),dtpAccidentTime.Value.GetDateTimeFormats()[47].ToString(),txtBody.Text.Trim(),
                                    txtTele.Text.Trim(),txtIDNO.Text,rtbProcess.Text.Trim(), rtbReason.Text.Trim(),rtbMeasure.Text.Trim(),txtPrincipal.Text.Trim(),txtStorer.Text.Trim(),txtCompCharge.Text.Trim()};
            for (int i = 0; i < strListOle.Length;i++)
            {
                listold.Add(strListOle[i]);
                listnew.Add(strListNew[i]);
            }

            //保存文件时处理同名情况
            bool pSave = false;
            int pInt = 1;
            while (pSave == false)
            {
                if (File.Exists(savePath.ToString()))
                {
                    if (savePath.ToString().LastIndexOf(")") == savePath.ToString().Length - 5 && savePath.ToString().LastIndexOf(")") != 0)
                        savePath = savePath.ToString().Remove(savePath.ToString().LastIndexOf("(") + 1, savePath.ToString().Length - savePath.ToString().LastIndexOf("(") - 1) + pInt.ToString() + ").xls";
                    else savePath = savePath.ToString().Remove(savePath.ToString().Length - 4, 4) + "(" + pInt.ToString() + ").xls";
                    pInt++;
                }
                else
                    pSave = true;
            }

            ReplaceWordDocAndSave(oDoc, savePath, listold, listnew);

            //关闭wordDoc文档对象     
            oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            //关闭wordApp组件对象     
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);

            string strUpdate = "UPDATE IndustrailInjuryInfo SET LastToWordUser = '" + pblinfo.username + "', LastToWordDate = '" + pNow.ToString() + "' WHERE InfoId = " + InfoId.ToString();
            try
            {
                mysql.sqlcmd(strUpdate);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            MessageBox.Show("文件导出成功，保存在[" + savePath + "]");
        }
        #endregion

        #region Fill Word
        /// <summary>
        /// 填充文档
        /// </summary>
        /// <param name="oDoc"></param>
        /// <param name="savePath"></param>
        /// <param name="findText"></param>
        /// <param name="replaceText"></param>
        protected void ReplaceWordDocAndSave(_Document oDoc, object savePath, List<string> findText, List<string> replaceText)
        {
            object format = WdSaveFormat.wdFormatDocument;
            string[] newStr = replaceText.ToArray();
            int i = 0;
            object miss = Missing.Value;
            object FindText, ReplaceWith, Replace;
            //object MissingValue = Type.Missing;

            foreach (string str in findText)
            {
                oDoc.Content.Find.Text = str;
                //要查找的文本 
                FindText = str;
                //替换文本 
                ReplaceWith = newStr[i];
                i++;
                Replace = WdReplace.wdReplaceAll;
                //移除Find的搜索文本和段落格式设置 
                oDoc.Content.Find.ClearFormatting();

                if (oDoc.Content.Find.Execute(ref FindText, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref ReplaceWith, ref Replace, ref miss, ref miss, ref miss, ref miss))
                {
                    //Response.Write("替换成功！");
                }
                else
                {
                //    Response.Write("没有相关要替换的：（" + str + "）字符");
                }
            }
            oDoc.SaveAs(ref savePath, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
        }
        #endregion

        #region Confirm
        /// <summary>
        /// btnConfirm_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfirm_Click(object sender, EventArgs e)
        {
            string strTick = string.Empty;
            string strPhoneNO = string.Empty;
            string strContents = string.Empty;
            if (CheckDate())
            {
                string strSQLConfirm = string.Empty;
                if (IsConfirm)
                {
                    strSQLConfirm = "UPDATE [IndustrailInjuryInfo] SET IsConfirm = 0 WHERE InfoId = " + InfoId.ToString();
                    IsConfirm = false;
                    pl1.Enabled = true;
                    btnSave.Enabled = true;
                    chbNotice.Enabled = true;
                    txtName.Focus();
                    txtName.Select(0, txtName.Text.Length);
                }
                else
                {
                    strSQLConfirm = "UPDATE [IndustrailInjuryInfo] SET IsConfirm = 1 WHERE InfoId = " + InfoId.ToString();
                    IsConfirm = true;
                    pl1.Enabled = false;
                    btnSave.Enabled = false;
                    chbNotice.Enabled = false;
                }
                try
                {
                    mysql.sqlcmd(strSQLConfirm);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (IsConfirm)
                {
                    btnConfirm.Text = "激活";
                    string strSQL = string.Empty;
                    if (EditType == "Edit")
                    {
                        strSQL = @"UPDATE IndustrailInjuryInfo
                            SET AccidentDate = '" + dtpAccidentDate.Value.GetDateTimeFormats()[4].ToString() + " " + dtpAccidentTime.Value.GetDateTimeFormats()[134].ToString() + "', Duty = '" + txtDuty.Text.Trim() + @"',
                            Hospital = '" + (cbxHospital.Text == "请选择" ? "" : cbxHospital.Text) + "', Body = '" + txtBody.Text.Trim() + "', InjuryCategory = '" + (cbxInCategory.Text == "请选择" ? "" : cbxInCategory.Text) + @"',
                            Principal = '" + txtPrincipal.Text.Trim() + "', OccurredCost = " + (txtOccurredCost.Text == "" ? "0" : txtOccurredCost.Text.Trim()) + ", CurrentCost = " + (txtCurrentCost.Text == "" ? "0" : txtCurrentCost.Text.Trim()) + @",
                            Process = '" + rtbProcess.Text.Trim() + "', Tele = '" + txtTele.Text.Trim() + "', Reason = '" + rtbReason.Text.Trim() + @"',
                            Measure = '" + rtbMeasure.Text.Trim() + "', CompCharge = '" + txtCompCharge.Text.Trim() + "', Remark = '" + rtbRemark.Text.Trim() + @"',
                            Sex = '" + (rbtMale.Checked ? "男" : "女") + "', Storer = '" + txtStorer.Text.Trim() + "', Department = '" + cb1.Text + @"',
                            IDNO = '" + txtIDNO.Text.Trim() + "', Name = '" + txtName.Text.Trim() + @"',
                            UpdateUser = '" + pblinfo.username + "', UpdateDate = '" + pNow.ToString() + @"',
                            Notice = " + (pNotice == 0 ? 2 : (pNotice == 1 ? 3 : pNotice)).ToString() + @"
                            WHERE InfoId = " + InfoId.ToString();
                    }
//                    else
//                    {
//                        strSQL = @"INSERT INTO IndustrailInjuryInfo(Name, AccidentDate, Duty, Department, Body, Category, TotalCost, Process, Reason, Measure, Remark, Sex, Storer, CompCharge, Tele, IDNO, Birthday, Principal)
//                            VALUES('" + txtName.Text.Trim() + "', '" + dtPicker.Value.ToString() + "', '" + txtDuty.Text.Trim() + "', '" + cb1.Text + @"',
//                            '" + txtBody.Text.Trim() + "', '" + (cbxInCategory.Text == "请选择" ? "" : cbxInCategory.Text) + @"',
//                            '" + (txtTotalCost.Text.Length > 0 ? txtTotalCost.Text.Trim() : "0") + @"',
//                            '" + rtbProcess.Text.Trim() + "', '" + rtbReason.Text.Trim() + @"',
//                            '" + rtbMeasure.Text.Trim() + "', '" + rtbRemark.Text.Trim() + @"',
//                            '" + txtSex.Text.Trim() + "', '" + txtStorer.Text.Trim() + @"',
//                            '" + txtCompCharge.Text.Trim() + "', '" + txtTele.Text.Trim() + @"',
//                            '" + txtIDNO.Text.Trim() + "', '" + dtpBirthday.Value.ToString() + @"',
//                            '" + txtPrincipal.Text.Trim() + "')";
//                    }
                    try
                    {
                        mysql.sqlcmd(strSQL);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }

                    if (chbNotice.Checked)
                    {
                        strPhoneNO = GetPhoneNO("报告接收人");
                        strContents = "【豪美铝业】：工伤信息[" + txtName.Text.Trim() + "-" + dtpAccidentTime.Value.GetDateTimeFormats()[11].ToString() + "]，已经确认！ 发送人：" + pblinfo.username;

                        if (SendMSG(strPhoneNO, strContents) == "") strTick = "保存成功并已发送信息通知接收人:[" + GetNoticeName("确认接收人") + "]，关闭此窗口？";
                        else strTick = "保存成功，关闭此窗口？";
                    }
                    else strTick = "保存成功，关闭此窗口？";

                    if (MessageBox.Show("信息已确认并保存,您要退出编辑吗？", "信息确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        Close();
                }
                else
                {
                    btnConfirm.Text = "确认信息";
                    //MessageBox.Show("已激活！");
                }
            }
        }
        #endregion

        #region Output Excel
        /// <summary>
        /// ToExcel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            //未经确认的信息不能导出文档
            string strConfirm = "SELECT IsConfirm FROM IndustrailInjuryInfo WHERE  InfoId = " + InfoId.ToString();
            try
            {
                if (!bool.Parse(mysql.ExecuteScalar(strConfirm).ToString()))
                {
                    MessageBox.Show("本工伤信息还未被确认，不能导出工伤申报表。请等工伤信息被确认后再导出！");
                    return;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            object missing = Missing.Value;
            string savePath = "D:\\工伤报告\\";
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }

            string strGetInfo = @"SELECT Name, CreateDate, AccidentDate, AccidentCategory, Duty, Department, Body, InjuryCategory, Hospital, TotalCost, Process, Reason, Measure, Remark, Sex, Storer, CompCharge, Tele, IDNO, Birthday, Principal, IsConfirm, InfoId
                                FROM IndustrailInjuryInfo
                                WHERE InfoId = " + InfoId.ToString();
            System.Data.DataTable dt = null;
            try
            {
                dt = mysql.sqltb(strGetInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            System.Data.DataRow dr = dt.Rows[0];

            Excel.Application myApp = new Excel.Application();
            Excel.Workbook myBook = myApp.Workbooks.Add(missing);
            Excel.Worksheet mySheet = myBook.Worksheets[1] as Excel.Worksheet;
            myApp.DisplayAlerts = false;

            (mySheet.Columns["A:A", Type.Missing] as Excel.Range).ColumnWidth = "15";
            (mySheet.Columns["B:B", Type.Missing] as Excel.Range).ColumnWidth = "15";
            (mySheet.Columns["C:C", Type.Missing] as Excel.Range).ColumnWidth = "15";
            (mySheet.Columns["D:D", Type.Missing] as Excel.Range).ColumnWidth = "18";
            (mySheet.Columns["E:E", Type.Missing] as Excel.Range).ColumnWidth = "27";

            ((Excel.Range)mySheet.Rows[1, missing]).RowHeight = "19.50";
            ((Excel.Range)mySheet.Rows[2, missing]).RowHeight = "30";
            ((Excel.Range)mySheet.Rows[3, missing]).RowHeight = "30";
            ((Excel.Range)mySheet.Rows[4, missing]).RowHeight = "21";
            ((Excel.Range)mySheet.Rows[5, missing]).RowHeight = "24.75";
            ((Excel.Range)mySheet.Rows[6, missing]).RowHeight = "25.5";
            ((Excel.Range)mySheet.Rows[7, missing]).RowHeight = "24";
            ((Excel.Range)mySheet.Rows[8, missing]).RowHeight = "79.50";

            mySheet.Cells[1, 1] = "编号：        ";
            Excel.Range range = mySheet.Range[mySheet.Cells[1, 1], mySheet.Cells[1, 5]];
            range.Merge();
            range.Font.Size = 11;
            range.Font.Name = "宋体";
            range.HorizontalAlignment = -4152;//左对齐：-4131；右对齐：-4152；居中：-4108

            range = mySheet.Range[mySheet.Cells[2, 1], mySheet.Cells[2, 5]];
            range.Value2 = "清远市工伤事故申报表";
            range.Merge();
            range.Font.Size = 20;
            range.Font.Name = "宋体";
            range.HorizontalAlignment = -4108;

            mySheet.Cells[3, 1] = "参保单位（公章）：";
            mySheet.Cells[3, 1].Font.Size = 14;
            range = mySheet.Range[mySheet.Cells[1, 1], mySheet.Cells[3, 5]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            mySheet.Cells[4, 1] = "工伤人员姓名";
            mySheet.Cells[4, 2] = dr["Name"].ToString();
            mySheet.Cells[4, 3] = "身份证号码";
            mySheet.Cells[4, 4] = "'" + dr["IDNO"].ToString();
            mySheet.Range[mySheet.Cells[4, 4], mySheet.Cells[4, 5]].Merge();

            mySheet.Cells[5, 1] = "工伤事故时间";
            mySheet.Cells[5, 2] = dr["AccidentDate"].ToString();
            range = mySheet.Range[mySheet.Cells[5, 2], mySheet.Cells[5, 3]];
            range.Merge();
            range.HorizontalAlignment = -4131;
            mySheet.Cells[5, 4] = "事故类别";
            mySheet.Cells[5, 5] = dr["AccidentCategory"].ToString() == "" ? "因工受伤" : dr["AccidentCategory"].ToString();

            mySheet.Cells[6, 1] = "就医机构名称";
            mySheet.Cells[6, 2] = dr["Hospital"].ToString();
            mySheet.Cells[6, 3] = "科室：          床号：";
            mySheet.Range[mySheet.Cells[6, 3], mySheet.Cells[6, 4]].Merge();
            mySheet.Cells[6, 5] = "□住院  □ 门诊";

            mySheet.Cells[7, 1] = "工伤事故经过及受伤情况摘要";
            range = mySheet.Range[mySheet.Cells[7, 1], mySheet.Cells[7, 5]];
            range.Merge();
            range.Font.Size = 14;
            range.Font.Name = "宋体";
            range.Font.Bold = true;
            range.HorizontalAlignment = -4108;

            mySheet.Cells[8, 1] = "  " + dr["Process"].ToString();
            mySheet.Range[mySheet.Cells[8, 1], mySheet.Cells[8, 5]].Merge();
            mySheet.Range[mySheet.Cells[8, 1], mySheet.Cells[8, 5]].WrapText = true;

            range = mySheet.Range[mySheet.Cells[4, 1], mySheet.Cells[8, 5]];
            range.Borders.LineStyle = 1;

            mySheet.Cells[9, 1] = "单位申办人：";
            mySheet.Cells[9, 5] = "社保受理人：";
            mySheet.Cells[10, 1] = "申办时间：";
            mySheet.Cells[10, 2] = "   年   月   日";
            mySheet.Range[mySheet.Cells[10, 2], mySheet.Cells[10, 3]].Merge();
            mySheet.Cells[10, 5] = "受理时间：     年   月   日";
            mySheet.Cells[11, 1] = "说明：";
            mySheet.Range[mySheet.Cells[11, 1], mySheet.Cells[11, 5]].Merge();
            mySheet.Cells[12, 1] = "1、本表一式三份，以黑色水笔填写（或打印），医保科、稽查科、参保单位各执一份；";
            mySheet.Cells[13, 1] = "2、参保职工发生工伤事故24小时内，必须填写本表报社保部门备案；";
            mySheet.Cells[14, 1] = "3、未在发生工伤24小时填写本表申报的，社保部门不予受理，不予支付工伤保险待遇；";
            mySheet.Cells[15, 1] = "4、工伤职工必须到市内工伤定点医疗机构就医；因伤情严重的可先到就近医疗机构进行抢救治疗，";
            mySheet.Cells[16, 1] = "   伤情稳定后及时转市内定点医疗机构继续治疗，并报社保部门备案； ";
            mySheet.Cells[17, 1] = "5、工伤职工需转市外就医的，由本地工伤定点医疗机构提出并填写“清远市工伤转院申请表”，";
            mySheet.Cells[18, 1] = "   由其单位同意并报社保部门批准方可转出（伤情危重的可先转出，二个工作日内须到";
            mySheet.Cells[19, 1] = "   社保部门补办相关转院手续）， 擅自到市外就医发生的工伤医疗费用，社保部门不予支付；";
            mySheet.Cells[20, 1] = "6、工伤职工就医时应随身携带身份证（或其他有效证件），以便社保部门工作人员稽查时身份核对；";
            mySheet.Cells[21, 1] = "7、事故类别：①因工受伤、②因工失踪、③因工死亡；";
            mySheet.Cells[22, 1] = "8、各地社保局工伤经办联系及传真电话： 市直： 3376210、3383483（传真）、清新：5817290、5812319（传真）、";
            mySheet.Cells[23, 1] = "   城区：3379635（传真）、连州：6677101、6677033(传真)、英德;2233923、2236883（传真）、连山：8716722、8716028（传真）、";
            mySheet.Cells[24, 1] = "   阳山：7803160、7801528（传真） 佛冈：4284017、4281608（传真）。";

            mySheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;//横向打印预览
            myBook.Windows[1].DisplayGridlines = false; //隐藏网格线

            myBook.Saved = true;
            savePath += DateTime.Parse(dr["AccidentDate"].ToString()).GetDateTimeFormats()[11].ToString() + dr["Name"].ToString() + "工伤事故申报表.xls";

            //保存文件时处理同名情况
            bool pSave = false;
            int pInt = 1;
            while (pSave == false)
            {
                if (File.Exists(savePath))
                {
                    if (savePath.LastIndexOf(")") == savePath.Length - 5 && savePath.LastIndexOf(")") != 0)
                        savePath = savePath.Remove(savePath.LastIndexOf("(") + 1, savePath.Length - savePath.LastIndexOf("(") - 1) + pInt.ToString() + ").xls";
                    else savePath = savePath.Remove(savePath.Length - 4, 4) + "(" + pInt.ToString() + ").xls";
                    pInt++;
                }
                else
                    pSave = true;
            }

            //文件保存
            try
            {
                myBook.SaveAs(savePath, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing);
                myBook.Close(null, null, null);
                myApp.Workbooks.Close();
                myApp.Application.Quit();
                myApp.Quit();
                mySheet = null;
                myBook = null;
                myApp = null;

                GC.Collect();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            string strUpdate = "UPDATE IndustrailInjuryInfo SET LastToExcelUser = '" + pblinfo.username + "', LastToExcelDate = '" + pNow.ToString() + "' WHERE InfoId = " + InfoId.ToString();
            try
            {
                mysql.sqlcmd(strUpdate);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            MessageBox.Show("报表导出成功：[" + savePath + "]");
        }
        #endregion

        #region Cost Events

        /// <summary>
        /// OccurredCost_KeyPress
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtOccurredCost_KeyPress(object sender, KeyPressEventArgs e)
        {
            reg = new Regex(@"^[1-9]\d*|0$");//^[1-9]\d*|0$   //^(?:[1-9]+\d*?|0)(\.\d+)?$
            if (e.KeyChar == 46 && txtOccurredCost.Text.IndexOf(".") != -1)
            {
                e.Handled = true;
            }
            if (e.KeyChar != '\b' && e.KeyChar != 46)
            {
                if (!reg.IsMatch(e.KeyChar.ToString()))
                {
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// OccurredCost_Leave
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtOccurredCost_Leave(object sender, EventArgs e)
        {
            lblTotalCost.Text = (float.Parse(txtOccurredCost.Text == "" ? "0" : txtOccurredCost.Text) + float.Parse(txtCurrentCost.Text == "" ? "0" : txtCurrentCost.Text)).ToString();
        }

        /// <summary>
        /// CurrentCost_KeyPress
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtCurrentCost_KeyPress(object sender, KeyPressEventArgs e)
        {
            reg = new Regex(@"^[1-9]\d*|0$");//^[1-9]\d*|0$   //^(?:[1-9]+\d*?|0)(\.\d+)?$
            if (e.KeyChar == 46 && txtCurrentCost.Text.IndexOf(".") != -1)
            {
                e.Handled = true;
            }
            if (e.KeyChar != '\b' && e.KeyChar != 46)
            {
                if (!reg.IsMatch(e.KeyChar.ToString()))
                {
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// CurrentCost_Leave
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtCurrentCost_Leave(object sender, EventArgs e)
        {
            lblTotalCost.Text = (float.Parse(txtOccurredCost.Text == "" ? "0" : txtOccurredCost.Text) + float.Parse(txtCurrentCost.Text == "" ? "0" : txtCurrentCost.Text)).ToString();
        }
        #endregion

        #region Get Recipient Name
        /// <summary>
        /// 获取接收人姓名
        /// </summary>
        /// <param name="pType"></param>
        /// <returns></returns>
        private string GetNoticeName(string pType)
        {
            string strReturn = string.Empty;
            string strSQL = string.Empty;
            if (pType == "确认接收人") strSQL = "SELECT Contents AS Name FROM SystemGroup WHERE GroupId = 4";
            else if (pType == "报告接收人") strSQL = "SELECT Contents AS Name FROM SystemGroup WHERE GroupId = 5";
            try
            {
                strReturn = mysql.ExecuteScalar(strSQL).ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;
        }
        #endregion

        #region Get Recipient Telephone
        /// <summary>
        /// 获取接收人电话
        /// </summary>
        /// <param name="pType"></param>
        /// <returns></returns>
        private string GetPhoneNO(string pType)
        {
            string strReturn = string.Empty;
            string strSQL = string.Empty;
            if (pType == "确认接收人") strSQL = "SELECT SUBSTRING(Remark, 4, 11) AS Tel FROM SystemGroup WHERE GroupId = 4";
            else if (pType == "报告接收人") strSQL = "SELECT SUBSTRING(Remark, 4, 11) AS Tel FROM SystemGroup WHERE GroupId = 5";
            try
            {
                strReturn = mysql.ExecuteScalar(strSQL).ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strReturn;
        }
        #endregion

        #region Send Message
        /// <summary>
        /// 发送短信通知
        /// </summary>
        /// <param name="pPhoneNO"></param>
        /// <param name="pContents"></param>
        /// <returns></returns>
        private string SendMSG(string pPhoneNO, string pContents)
        {
            if (pPhoneNO.Length != 11) return "通知发送失败！";
            string pwd = getMD5(sn + password);
            long startTime, endTime;
            startTime = DateTime.Now.Ticks;//记录当前时刻
            //发送短信
            //定时时间，扩展码和rrid可以为空，当rrid(唯一标识串)为空时，返回系统指定rrid
            string result = sms.mdsmssend(sn, pwd, pPhoneNO, pContents, subcode, stime, "", "");

            endTime = DateTime.Now.Ticks;
            TimeSpan sendSpan = new TimeSpan(endTime - startTime);


            if (result.StartsWith("-") || result.Equals(""))
            {
                MessageBox.Show("通知发送失败：" + result, "系统提示!");
                return "通知发送失败！";
            }
            else
            {
                //MessageBox.Show("短信已成功发送至责任人！", "系统提示");
                return "";
            }
        }
        #endregion

        #region Analysis MD5
        /// <summary>
        /// MD5验证
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public string getMD5(string source)
        {
            string result = "";
            try
            {
                MD5 getmd5 = new MD5CryptoServiceProvider();
                byte[] targetStr = getmd5.ComputeHash(Encoding.UTF8.GetBytes(source));
                result = BitConverter.ToString(targetStr).Replace("-", "");
                return result;
            }
            catch (Exception)
            {
                return "0";
            }
        }
        #endregion
    }
}