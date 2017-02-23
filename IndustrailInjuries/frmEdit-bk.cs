using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using pblClass;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;


namespace IndustrailInjuries
{
    public partial class frmEdit : Form
    {
        #region Parameter Area
        Regex reg = null;
        #endregion

        #region 属性字段区
        private string _mName;
        /// <summary>
        /// PDName
        /// </summary>
        public string IName
        {
            set { this._mName = value; }
            get { return this._mName; }
        }

        private DateTime _mAccidentDate;
        /// <summary>
        /// Principal
        /// </summary>
        public DateTime AccidentDate
        {
            set { this._mAccidentDate = value; }
            get { return this._mAccidentDate; }
        }
        private string _mEditType;
        /// <summary>
        /// EditType
        /// </summary>
        public string EditType
        {
            set { this._mEditType = value; }
            get { return this._mEditType; }
        }
        #endregion

        #region 构造函数
        public frmEdit()
        {
            InitializeComponent();
        }
        public frmEdit(string pName, DateTime pAcci, string pEditType)
        {
            this._mName = pName;
            this._mAccidentDate = pAcci;
            this._mEditType = pEditType;
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
            FillDepartmentList();
            if (this.EditType == "Edit")
            {
                this.lblTitle.Text = "编辑工伤信息";
                string strSQL = @"SELECT A.[Name], A.[CreateDate], A.[AccidentDate], A.[Duty], B.PDID, A.[Body], C.[GroupId], A.[Grade], A.[Cost], A.[Process], A.[Reason], A.[Measure], A.[Remark], A.Sex, A.Storer, A.CompCharge, A.Tele, A.IDNO, A.Birthday, A.Principal
                        FROM IndustrailInjuryInfo A
                        INNER JOIN productdepartment B ON A.[Department] = B.PDName
                        INNER JOIN systemCategoryGroup C ON A.Category = C.CategoryName
                        WHERE Name = '" + this.IName + "' AND CONVERT(VARCHAR(10), AccidentDate, 120) = '" + this.AccidentDate.GetDateTimeFormats()[0] + "'";
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
                this.txtName.Text = dt.Rows[0]["Name"].ToString();
                this.dtPicker.Value = DateTime.Parse(dt.Rows[0]["AccidentDate"].ToString());
                this.txtDuty.Text = dt.Rows[0]["Duty"].ToString();
                this.txtBody.Text = dt.Rows[0]["Body"].ToString();
                this.cbxInCategory.SelectedValue = dt.Rows[0]["GroupId"].ToString();
                this.txtGrade.Text = dt.Rows[0]["Grade"].ToString();
                this.txtCost.Text = dt.Rows[0]["Cost"].ToString();
                this.rtbProcess.Text = dt.Rows[0]["Process"].ToString();
                this.rtbReason.Text = dt.Rows[0]["Reason"].ToString();
                this.rtbMeasure.Text = dt.Rows[0]["Measure"].ToString();                
                this.rtbRemark.Text = dt.Rows[0]["Remark"].ToString();
                this.cb1.SelectedValue = dt.Rows[0]["PDID"].ToString();
                this.txtSex.Text = dt.Rows[0]["Sex"].ToString();
                this.txtTele.Text = dt.Rows[0]["Tele"].ToString();
                this.txtStorer.Text = dt.Rows[0]["Storer"].ToString();
                this.txtCompCharge.Text = dt.Rows[0]["CompCharge"].ToString();
                this.txtIDNO.Text = dt.Rows[0]["IDNO"].ToString();
                this.dtpBirthday.Value = DateTime.Parse(dt.Rows[0]["Birthday"].ToString());
                this.txtPrincipal.Text = dt.Rows[0]["Principal"].ToString();

                this.txtName.ReadOnly = true;
            }
            //else
            //{

            //}
            this.txtName.Focus();
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
            if (CheckDate())
            {
                string strSQL = string.Empty;
                if (this.EditType == "Edit")
                {
                    strSQL = @"UPDATE IndustrailInjuryInfo
                    SET AccidentDate = '" + this.dtPicker.Value.ToString() + "', Duty = '" + this.txtDuty.Text.Trim() + @"',
                    Body = '" + this.txtBody.Text.Trim() + "', Category = '" + this.cbxInCategory.Text + @"',
                    Grade = '" + this.txtGrade.Text.Trim() + "', Principal = '" + this.txtPrincipal.Text.Trim() + "', Cost = '" + this.txtCost.Text.Trim() + @"',
                    Process = '" + this.rtbProcess.Text.Trim() + "', Tele = '" + this.txtTele.Text.Trim() + "', Reason = '" + this.rtbReason.Text.Trim() + @"',
                    Measure = '" + this.rtbMeasure.Text.Trim() + "', CompCharge = '" + this.txtCompCharge.Text.Trim() + "', Remark = '" + this.rtbRemark.Text.Trim() + @"',
                    Sex = '" + this.txtSex.Text.Trim() + "', Storer = '" + this.txtStorer.Text.Trim() + "', Department = '" + this.cb1.Text + @"',
                    IDNO = '" + this.txtIDNO.Text.Trim() + "', Birthday = '" + this.dtpBirthday.Value.ToString() + @"'
                    WHERE Name = '" + this.txtName.Text + "' AND CONVERT(VARCHAR(10), AccidentDate, 120) = '" + this.AccidentDate.GetDateTimeFormats()[0] + "'";
                }
                else
                {
                    strSQL = @"INSERT INTO IndustrailInjuryInfo(Name, AccidentDate, Duty, Department, Body, Category, Grade, Cost, Process, Reason, Measure, Remark, Sex, Storer, CompCharge, Tele, IDNO, Birthday, Principal)
                            VALUES('" + this.txtName.Text.Trim() + "', '" + this.dtPicker.Value.ToString() + "', '" + this.txtDuty.Text.Trim() + "', '" + this.cb1.Text + @"',
                            '" + this.txtBody.Text.Trim() + "', '" + this.cbxInCategory.Text + @"',
                            '" + this.txtGrade.Text.Trim() + "', '" + this.txtCost.Text.Trim() + @"',
                            '" + this.rtbProcess.Text.Trim() + "', '" + this.rtbReason.Text.Trim() + @"',
                            '" + this.rtbMeasure.Text.Trim() + "', '" + this.rtbRemark.Text.Trim() + @"',
                            '" + this.txtSex.Text.Trim() + "', '" + this.txtStorer.Text.Trim() + @"',
                            '" + this.txtCompCharge.Text.Trim() + "', '" + this.txtTele.Text.Trim() + @"',
                            '" + this.txtIDNO.Text.Trim() + "', '" + this.dtpBirthday.Value.ToString() + @"',
                            '" + this.txtPrincipal.Text.Trim() + "')";
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
                if (MessageBox.Show("保存成功，关闭此窗口？", "存盘", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    this.Close();
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
            this.Close();
        }
        #endregion

        #region CheckDate
        /// <summary>
        /// 数据有效性检测
        /// </summary>
        /// <returns></returns>
        private bool CheckDate()
        {
            if(this.txtSex.Text.Length<1)
            {
                MessageBox.Show("请输入员工性别！");
                this.txtSex.Focus();
                return false;
            }
            if (this.txtSex.Text != "男" && this.txtSex.Text != "女")
            {
                MessageBox.Show("性别有误，请重新填写！");
                this.txtSex.Focus();
                this.txtSex.Select(0, this.txtSex.Text.Length);
                return false;
            }
            //Regex regex = new Regex(@"^[0-9]\d*\.\d{0,2}$|^\d*$");
            if (this.txtName.Text == null || this.txtName.Text.Length <= 0)
            {
                MessageBox.Show("请输入姓名");
                this.txtName.Focus();
                return false;
            }
            if (this.txtCost.Text == null || this.txtCost.Text.Length <= 0)
            {
                MessageBox.Show("请输入费用");
                this.txtCost.Focus();
                return false;
            }
            reg = new Regex(@"^[0-9]\d*\.\d{0,2}$|^\d*$");
            if (!reg.IsMatch(this.txtCost.Text.Trim()))
            {
                MessageBox.Show("费用必须是浮点数");
                this.txtCost.Focus();
                this.txtCost.Select(0, this.txtCost.Text.Length);
                return false;
            }
            if (this.rtbProcess.Text == null || this.rtbProcess.Text.Length <= 0)
            {
                MessageBox.Show("请输受伤过程");
                this.rtbProcess.Focus();
                return false;
            }
            reg = new Regex(@"^(\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$");//     /^(\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$/ /(^\d{15}$/)|(\d{17}(?:\d|x|X)$/
                                                                       //     /^(\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$/
            if (this.txtIDNO.Text.Length > 0 && this.txtIDNO.Text.Length != 15 && this.txtIDNO.Text.Length != 18)
            {
                MessageBox.Show("身份证号码的长度为15位或18位！");
                this.txtIDNO.Focus();
                this.txtIDNO.Select(0, this.txtIDNO.Text.Length);
                return false;
            }
            if (this.txtIDNO.Text.Length > 0 && !reg.IsMatch(this.txtIDNO.Text.Trim()))
            {
                MessageBox.Show("身份证:15位时全为数字，18位前17位为数字，最后一位是校验位，可能为数字或字符X！");
                this.txtIDNO.Focus();
                this.txtIDNO.Select(0, this.txtIDNO.Text.Length);
                return false;
            }
            string strSQL = "SELECT COUNT(*) FROM IndustrailInjuryInfo WHERE Name = '" + this.txtName.Text + "' AND CONVERT(VARCHAR(10), AccidentDate, 120) = '" + this.dtPicker.Value.GetDateTimeFormats()[0] + "'";
            if (this.EditType == "Edit")
            {
                strSQL += " AND (Name <> '" + this.Name + "' AND CONVERT(VARCHAR(10), AccidentDate, 120) <> '" + this.AccidentDate.GetDateTimeFormats()[0] + "')";
            }
            if (int.Parse(mysql.ExecuteScalar(strSQL).ToString()) != 0)
            {
                MessageBox.Show("[" + this.txtName.Text + ":" + this.dtPicker.Value.GetDateTimeFormats()[10] + "]这个条工伤信息已经存在，请重新输入！");
                this.txtName.Focus();
                this.txtName.Select(0, this.txtName.Text.Length);
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

            this.cb1.Items.Clear();
            this.cb1.DataSource = dt;
            this.cb1.DisplayMember = "PDName";
            this.cb1.ValueMember = "PDID";
            this.cb1.SelectedIndex = 0;

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

            this.cbxInCategory.Items.Clear();
            this.cbxInCategory.DataSource = dt;
            this.cbxInCategory.DisplayMember = "CategoryName";
            this.cbxInCategory.ValueMember = "GroupId";
            this.cbxInCategory.SelectedIndex = 0;
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
            //MessageBox.Show("Bug 修复中....."); return;
            // object varFileName = "c:\\temp\\doc.docx";
            //object varFalseValue = false;
            //object varTrueValue = true;
            //object varMissing = Type.Missing;
            //string varText;

            //// Create a reference to Microsoft Word application
            //Microsoft.Office.Interop.Word.Application varWord = new Microsoft.Office.Interop.Word.Application();
            //Microsoft.Office.Interop.Word.Document varDoc;
            //varDoc = varWord.Documents.Open(ref varFileName, ref varMissing, ref varFalseValue, ref varMissing, ref varMissing,
            // ref varMissing, ref varMissing, ref varMissing, ref varMissing, ref varMissing,
            // ref varMissing, ref varMissing, ref varMissing, ref varMissing, ref varMissing, ref varMissing);

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; //endofdoc是预定义的bookmark
            
            object filePath = System.Windows.Forms.Application.StartupPath + @"\WordTemplate\Int.doc";
            object savePath = "D:\\工伤报告";
            //Create Directory
            if (!Directory.Exists(savePath.ToString()))
            {
                Directory.CreateDirectory(savePath.ToString());
            }

            savePath += "\\" + this.dtPicker.Value.GetDateTimeFormats()[11].ToString() + this.txtName.Text.Trim() + "工伤信息.doc";

            //创建一个document.
            Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word._Document oDoc;
            //oWord.Visible = true;

            //oDoc = oWord.Documents.Open(ref filePath, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            //oDoc = copyWordDoc(filePath);
            object objDocType = WdDocumentType.wdTypeDocument;
            object type = WdBreakType.wdSectionBreakContinuous;

            //Word应用程序变量    
            //Microsoft.Office.Interop.Word.Application wordApp;
            //Word文档变量 
            //Document newWordDoc;

            object readOnly = false;
            object isVisible = false;

            //初始化 
            //由于使用的是COM库，因此有许多变量需要用Missing.Value代替 
            //wordApp = new Microsoft.Office.Interop.Word.Application();

            //Object Nothing = System.Reflection.Missing.Value;

            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            Microsoft.Office.Interop.Word._Document openWord;
            openWord = oWord.Documents.Open(ref filePath, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            openWord.Select();
            openWord.Sections[1].Range.Copy();

            object start = 0;
            Range newRang = oDoc.Range(ref start, ref start);

            //插入换行符    
            //newWordDoc.Sections[1].Range.InsertBreak(ref type); 
            oDoc.Sections[1].Range.PasteAndFormat(WdRecoveryType.wdPasteDefault);
            openWord.Close(ref oMissing, ref oMissing, ref oMissing);


            List<string> listold = new List<string>();
            List<string> listnew = new List<string>();
            string[] strListOle = { "Name", "Sex", "Birthday", "DP", "Duty", "InTime", "Body", "Tele", "IDNO", "Process", "Reason", "Measure","Principal", "Storer", "CompCharge" };
            string[] strListNew ={this.txtName.Text.Trim(), this.txtSex.Text,this.dtpBirthday.Value.GetDateTimeFormats()[0].ToString(),this.cb1.Text,this.txtDuty.Text.Trim(),this.dtPicker.Value.GetDateTimeFormats()[0].ToString(),this.txtBody.Text.Trim(),
                                    this.txtTele.Text.Trim(),this.txtIDNO.Text,this.rtbProcess.Text.Trim(), this.rtbReason.Text.Trim(),this.rtbMeasure.Text.Trim(),this.txtPrincipal.Text.Trim(),this.txtStorer.Text.Trim(),this.txtCompCharge.Text.Trim()};
            for (int i = 0; i < strListOle.Length;i++)
            {
                listold.Add(strListOle[i]);
                listnew.Add(strListNew[i]);
            }

            ReplaceWordDocAndSave(oDoc, savePath, listold, listnew);

            //oWord.Visible = true;            

            //关闭wordDoc文档对象     
            oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            //关闭wordApp组件对象     
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);

            MessageBox.Show("文件导出成功，保存在<D:\\工伤报告>文件夹。");
            #region mark1
            ////在document的开始部分添加一个paragraph.
            //Microsoft.Office.Interop.Word.Paragraph oPara1;
            //oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            //oPara1.Range.Text = "Heading 1";
            //oPara1.Range.Font.Bold = 1;
            //oPara1.Format.SpaceAfter = 24;        //24 pt 行间距
            //oPara1.Range.InsertParagraphAfter();

            ////在当前document的最后添加一个paragraph
            //Microsoft.Office.Interop.Word.Paragraph oPara2;
            //object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara2.Range.Text = "Heading 2";
            //oPara2.Format.SpaceAfter = 6;
            //oPara2.Range.InsertParagraphAfter();

            ////接着添加一个paragraph
            //Microsoft.Office.Interop.Word.Paragraph oPara3;
            //oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:";
            //oPara3.Range.Font.Bold = 0;
            //oPara3.Format.SpaceAfter = 24;
            //oPara3.Range.InsertParagraphAfter();

            ////添加一个3行5列的表格，填充数据，并且设定第一行的样式
            //Microsoft.Office.Interop.Word.Table oTable;
            //Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oTable = oDoc.Tables.Add(wrdRng, 3, 5, ref oMissing, ref oMissing);
            //oTable.Range.ParagraphFormat.SpaceAfter = 6;
            //int r, c;
            //string strText;
            //for (r = 1; r <= 3; r++)
            //    for (c = 1; c <= 5; c++)
            //    {
            //        strText = "r" + r + "c" + c;
            //        oTable.Cell(r, c).Range.Text = strText;
            //    }
            //oTable.Rows[1].Range.Font.Bold = 1;
            //oTable.Rows[1].Range.Font.Italic = 1;

            ////接着添加一些文字
            //Microsoft.Office.Interop.Word.Paragraph oPara4;
            //oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara4.Range.InsertParagraphBefore();
            //oPara4.Range.Text = "And here's another table:";
            //oPara4.Format.SpaceAfter = 24;
            //oPara4.Range.InsertParagraphAfter();

            ////添加一个5行2列的表，填充数据并且改变列宽
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
            //oTable.Range.ParagraphFormat.SpaceAfter = 6;
            //for (r = 1; r <= 5; r++)
            //    for (c = 1; c <= 2; c++)
            //    {
            //        strText = "r" + r + "c" + c;
            //        oTable.Cell(r, c).Range.Text = strText;
            //    }
            //oTable.Columns[1].Width = oWord.InchesToPoints(2); //设置列宽
            //oTable.Columns[2].Width = oWord.InchesToPoints(3);

            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            //object oPos;
            //double dPos = oWord.InchesToPoints(7);
            //oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            //do
            //{
            //    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //    wrdRng.ParagraphFormat.SpaceAfter = 6;
            //    wrdRng.InsertAfter("A line of text");
            //    wrdRng.InsertParagraphAfter();
            //    oPos = wrdRng.get_Information
            //                               (Microsoft.Office.Interop.Word.WdInformation.wdVerticalPositionRelativeToPage);
            //}
            //while (dPos >= Convert.ToDouble(oPos));
            //object oCollapseEnd = Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd;
            //object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
            //wrdRng.Collapse(ref oCollapseEnd);
            //wrdRng.InsertBreak(ref oPageBreak);
            //wrdRng.Collapse(ref oCollapseEnd);
            //wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            //wrdRng.InsertParagraphAfter();

            ////添加一个chart
            //Microsoft.Office.Interop.Word.InlineShape oShape;
            //object oClassType = "MSGraph.Chart.8";
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            ////Demonstrate use of late bound oChart and oChartApp objects to
            ////manipulate the chart object with MSGraph.
            //object oChart;
            //object oChartApp;
            //oChart = oShape.OLEFormat.Object;
            //oChartApp = oChart.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oChart, null);

            ////Change the chart type to Line.
            //object[] Parameters = new Object[1];
            //Parameters[0] = 4; //xlLine = 4
            //oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty, null, oChart, Parameters);

            ////Update the chart image and quit MSGraph.
            //oChartApp.GetType().InvokeMember("Update", BindingFlags.InvokeMethod, null, oChartApp, null);
            //oChartApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, oChartApp, null);
            ////... If desired, you can proceed from here using the Microsoft Graph
            ////Object model on the oChart and oChartApp objects to make additional
            ////changes to the chart.

            ////Set the width of the chart.
            //oShape.Width = oWord.InchesToPoints(6.25f);
            //oShape.Height = oWord.InchesToPoints(3.57f);

            //Add text after the chart.
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng.InsertParagraphAfter();
            //wrdRng.InsertAfter("THE END.");
            #endregion
        }
        #endregion

        #region CopyDoc
        ///// <summary> 
        ///// 从源DOC文档复制内容返回一个Document类 
        ///// </summary> 
        ///// <param name="sorceDocPath">源DOC文档路径</param> 
        ///// <returns>Document</returns> 
        //protected Document copyWordDoc(object sorceDocPath)
        //{
            //object objDocType = WdDocumentType.wdTypeDocument;
            //object type = WdBreakType.wdSectionBreakContinuous;

            ////Word应用程序变量    
            //Microsoft.Office.Interop.Word.Application wordApp;
            ////Word文档变量 
            //Document newWordDoc;

            //object readOnly = false;
            //object isVisible = false;

            ////初始化 
            ////由于使用的是COM库，因此有许多变量需要用Missing.Value代替 
            //wordApp = new Microsoft.Office.Interop.Word.Application();

            //Object Nothing = System.Reflection.Missing.Value;

            //newWordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //Document openWord;
            //openWord = wordApp.Documents.Open(ref sorceDocPath, ref Nothing, ref readOnly, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref isVisible, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //openWord.Select();
            //openWord.Sections[1].Range.Copy();

            //object start = 0;
            //Range newRang = newWordDoc.Range(ref start, ref start);

            ////插入换行符    
            ////newWordDoc.Sections[1].Range.InsertBreak(ref type); 
            //newWordDoc.Sections[1].Range.PasteAndFormat(WdRecoveryType.wdPasteDefault);
            //openWord.Close(ref Nothing, ref Nothing, ref Nothing);
            //return newWordDoc;
        //}
        #endregion

        #region Fill Word
        /// <summary>
        /// 填充文档
        /// </summary>
        /// <param name="oDoc"></param>
        /// <param name="savePath"></param>
        /// <param name="findText"></param>
        /// <param name="replaceText"></param>
        protected void ReplaceWordDocAndSave(Microsoft.Office.Interop.Word._Document oDoc, object savePath, List<string> findText, List<string> replaceText)
        {
            object format = WdSaveFormat.wdFormatDocument;
            object readOnly = false;
            object isVisible = false;

            //string strOldText = "{WORD}";
            //string strNewText = "替换后的文本";

            List<string> IListOldStr = findText;
            List<string> IListNewStr = replaceText;

            string[] newStr = IListNewStr.ToArray();
            int i = 0;

            Object Nothing = System.Reflection.Missing.Value;

            //Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            //Microsoft.Office.Interop.Word.Document oDoc = docObject;

            object FindText, ReplaceWith, Replace;
            object MissingValue = Type.Missing;

            foreach (string str in IListOldStr)
            {
                oDoc.Content.Find.Text = str;
                //要查找的文本 
                FindText = str;
                //替换文本 
                //ReplaceWith = strNewText;
                ReplaceWith = newStr[i];
                i++;

                //wdReplaceAll - 替换找到的所有项。 
                //wdReplaceNone - 不替换找到的任何项。 
                //wdReplaceOne - 替换找到的第一项。 
                Replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                //移除Find的搜索文本和段落格式设置 
                oDoc.Content.Find.ClearFormatting();

                if (oDoc.Content.Find.Execute(ref FindText, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref ReplaceWith, ref Replace, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    //this.Response.Write("替换成功！");
                    //Response.Write("<br>");
                }
                else
                {
                //    Response.Write("没有相关要替换的：（" + str + "）字符");
                //    Response.Write("<br>");
                }
            }
            oDoc.SaveAs(ref savePath, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
        }
        #endregion
    }
}
