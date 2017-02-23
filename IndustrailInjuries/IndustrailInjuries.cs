using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using pblClass;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace IndustrailInjuries
{
    public partial class IndustrailInjuries : Form
    {
        #region Parameter Area
        int pageSize = 25;              //每页显示行数        
        int pageCount = 0;              //页数＝总记录数/每页显示行数
        int recordCount = 0;            //总记录数
        int currentPage = 0;            //当前页号
        //int currentRow = 0;             //当前记录行
        string strSQL = string.Empty;   //SQL执行语句
        Regex reg;                      //正则表达式
        DataTable dt = new DataTable();
        #endregion       

        #region Initialize Area
        /// <summary>
        /// 初始化
        /// </summary>
        public IndustrailInjuries()
        {
            InitializeComponent();
        }
        #endregion

        #region Form Load
        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IndustrailInjuries_Load(object sender, EventArgs e)
        {
            //MessageBox.Show("维护中.....");
            //pl1.Enabled = false;

            dtTo.Value = DateTime.Now;
            dtFrom.Value = dtTo.Value.AddYears(-3);
            pl1.Dock = DockStyle.Top;
            bn1.Dock = DockStyle.Bottom;
            dgv1.Dock = DockStyle.Fill;

            FillCBXConfirm();
            GetDataSource("");

            if (pblinfo.v_zz_id != 1)
            {
                //文档导出权限确认
                string strResult = string.Empty;
                string strRight = @"DECLARE @tmp VARCHAR(4000)
                                SELECT @tmp = Contents FROM [SystemGroup] WHERE GroupId = 2
                                IF EXISTS(SELECT sub FROM dm_split(@tmp) WHERE sub = '" + pblClass.pblinfo.username + @"') SELECT 1
                                ELSE SELECT 0";
                try
                {
                    strResult = mysql.ExecuteScalar(strRight).ToString();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (strResult == "0")
                {
                    plExcel.Visible = false;
                    btnStatistics.Visible = false;
                }
            }
        }
        #endregion

        #region Page Button Control
        /// <summary>
        /// 分页按钮设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bn1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.Text == "上一页(&P)")
            {
                if (currentPage - 1 <= 0)
                {
                    MessageBox.Show("已经是第一页，请点击“下一页”查看！");
                    return;
                }
                else
                {
                    currentPage--;
                }
            }
            if (e.ClickedItem.Text == "下一页(&N)")
            {
                if (currentPage + 1 > pageCount)
                {
                    MessageBox.Show("已经是最后一页，请点击“上一页”查看！");
                    return;
                }
                else
                {
                    currentPage++;
                }
            }
            if (e.ClickedItem.Text == "跳转到(&G)")
            {
                reg = new Regex(@"^[0-9]*[1-9][0-9]*$");

                if (!reg.IsMatch(txtCurrentPage.Text))
                {
                    MessageBox.Show("输入的页码格式不正确！");
                    txtCurrentPage.Focus();
                    txtCurrentPage.Text = pageCount.ToString();
                    txtCurrentPage.Select(0, txtCurrentPage.Text.Length);
                    return;
                }
                if (int.Parse(txtCurrentPage.Text) > pageCount)
                {
                    MessageBox.Show("跳转页超过了总页数！");
                    return;
                }
                currentPage = int.Parse(txtCurrentPage.Text);
            }
            GetDataSource("Navi");
        }
        #endregion

        #region Get Date Source
        /// <summary>
        /// 数据源
        /// </summary>
        private void GetDataSource(string pType)
        {
            strSQL = @"SELECT Name, CONVERT(VARCHAR(20), AccidentDate, 120) AccidentDate, Sex, Department, Duty, Tele, IDNO, Body, InjuryCategory,
	                    OccurredCost, CurrentCost, TotalCost, Process, Reason, Measure, Principal, Storer, CompCharge, Remark, CASE IsConfirm WHEN 0 THEN '未确认' ELSE '已确认' END IsConfirm, InfoId, Creator
                    FROM IndustrailInjuryInfo WHERE IsCancel = 0";
            if (pType == "Search" || pType == "Navi")
            {
                if (dtFrom.Value > dtTo.Value)
                {
                    MessageBox.Show("起始日期不能大于结束日期！");
                    return;
                }
                if (cbxConfirm.SelectedValue.ToString() != "-1")
                    strSQL += " AND IsConfirm = " + cbxConfirm.SelectedValue.ToString();
                if (txtName.TextLength > 0) strSQL += " AND Name LIKE '%" + txtName.Text + "%'";
            }
            strSQL += " AND DATEADD(DD, DATEDIFF(DD, 0, AccidentDate), 0) BETWEEN '" + dtFrom.Value + "' AND '" + dtTo.Value + @"'
                        ORDER BY AccidentDate DESC";
            try
            {
                dt = mysql.sqltb(strSQL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            //计算总记录数
            recordCount = dt.Rows.Count;
            if (recordCount == 0)
            {
                dgv1.DataSource = null;
                return;
            }
            //计算出总页数
            pageCount = (recordCount / pageSize);
            if ((recordCount % pageSize) > 0) pageCount++;

            if (pType == "" || pType == "Search")
            {
                currentPage = 1;    //首次加载，当前页数从1开始
            }
            
            //设置分页
            int nStartPos = 0;   //当前页面开始记录行
            int nEndPos = 0;     //当前页面结束记录行
            DataTable dtTemp = dt.Clone();   //克隆DataTable结构框架

            if (currentPage == pageCount)
            {
                nEndPos = recordCount;
            }
            else
            {
                nEndPos = pageSize * currentPage;
            }

            nStartPos = pageSize * (currentPage - 1);
            lblPageCount.Text = pageCount.ToString();
            txtCurrentPage.Text = currentPage.ToString();

            //从元数据源复制记录行
            for (int i = nStartPos; i < nEndPos; i++)
            {
                dtTemp.ImportRow(dt.Rows[i]);
                //currentRow++;
            }
            bs1.DataSource = dtTemp;
            bn1.BindingSource = bs1;
            dgv1.DataSource = bs1;

            dgv1.EnableHeadersVisualStyles = false;
            dgv1.Columns[0].HeaderText = "姓名";
            dgv1.Columns[1].HeaderText = "工伤日期";
            dgv1.Columns[2].HeaderText = "性别";
            dgv1.Columns[3].HeaderText = "部门";
            dgv1.Columns[4].HeaderText = "岗位";
            dgv1.Columns[5].HeaderText = "电话";
            dgv1.Columns[6].HeaderText = "身份证";
            dgv1.Columns[7].HeaderText = "受伤部位";
            dgv1.Columns[8].HeaderText = "工伤类型";
            dgv1.Columns[9].HeaderText = "已产生费用";
            dgv1.Columns[10].HeaderText = "当前费用";
            dgv1.Columns[11].HeaderText = "总费用";
            dgv1.Columns[12].HeaderText = "工伤过程";
            dgv1.Columns[13].HeaderText = "原因";
            dgv1.Columns[14].HeaderText = "改善对策";
            dgv1.Columns[15].HeaderText = "负责人";
            dgv1.Columns[16].HeaderText = "车间主任";
            dgv1.Columns[17].HeaderText = "安全主任";
            dgv1.Columns[18].HeaderText = "备注";
            dgv1.Columns[19].HeaderText = "确认";
            dgv1.Columns[20].HeaderText = "ID";
            dgv1.Columns[21].HeaderText = "创建者";
            dgv1.Columns[20].Visible = false;
            dgv1.Columns[21].Visible = false;
            dgv1.ColumnHeadersHeight = 35;
            dgv1.Columns[1].Width = 120;
            dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            //区别已经确认的工伤信息
            int displayCount = dgv1.Rows.Count - 1 > pageSize ? pageSize : dgv1.Rows.Count - 1;
            for (int i = 0; i < displayCount; i++)
                if (dgv1.Rows[i].Cells[19].Value.ToString() == "未确认")
                    dgv1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(245, 200, 160);
            //dgv1.Columns[1].DefaultCellStyle.BackColor = System.Drawing.Color.DeepSkyBlue;
        }
        #endregion

        #region Search Message
        /// <summary>
        /// 查找
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch_Click(object sender, EventArgs e)
        {
            GetDataSource("Search");
        }
        #endregion

        #region Delete Message
        /// <summary>
        /// 删除数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            //删除权限判断
            if (pblinfo.v_zz_id != 1)
            {
                string strRight = @"DECLARE @tmp VARCHAR(4000)
                                SELECT @tmp = Contents FROM [SystemGroup] WHERE GroupId = 1
                                IF EXISTS(SELECT sub FROM dm_split(@tmp) WHERE sub = '" + pblClass.pblinfo.username + @"') SELECT 1
                                ELSE SELECT 0";
                try
                {
                    strRight = mysql.ExecuteScalar(strRight).ToString();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (strRight == "0" && pblinfo.username != dgv1.CurrentRow.Cells[21].Value.ToString())   //没有管理权限并且不是信息的创建人
                {
                    MessageBox.Show("非管理员不能删除其他人创建的工伤信息！");
                    return;
                }
            }

            if (dgv1.RowCount > 0 & dgv1.CurrentRow != null)
            {
                string pName = dgv1.CurrentRow.Cells[0].Value.ToString();
                string pInfoId = dgv1.CurrentRow.Cells[20].Value.ToString();
                DateTime pAcci = DateTime.Parse(dgv1.CurrentRow.Cells[1].Value.ToString());

                if (MessageBox.Show("你确定要删除[" + pName + ":" + pAcci.GetDateTimeFormats()[10] + "]的工伤信息吗？", "删除部门", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    string strDel = "UPDATE IndustrailInjuryInfo SET IsCancel = 1, UpdateUser = '" + pblinfo.username + "', UpdateDate = '" + System.DateTime.Now.ToString() + "' WHERE InfoId = " + pInfoId;
                    try
                    {
                        mysql.sqlcmd(strDel);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                    if (recordCount == 1) { dgv1.DataSource = null; return; }//只有一条数据时
                    if (currentPage == pageCount && (recordCount % pageSize) == 1)//当前是最后一页且本页只有一条数据时
                    {
                        currentPage--;
                    }
                    //else if (currentPage == pageCount && currentRow == dgv1.Rows.Count - 1)//删除排最后的一条数据（非本页最后一条数据）
                    //{
                    //    currentRow--;
                    //}
                }

                GetDataSource("Delete");
            }
        }
        #endregion

        #region Add Message
        /// <summary>
        /// 新增
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAdd_Click(object sender, EventArgs e)
        {
            frmEdit addFrom = new frmEdit();
            addFrom.ShowDialog();
            GetDataSource("");
        }
        #endregion

        #region Edit Message
        /// <summary>
        /// 编辑
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (dgv1.RowCount > 0 & dgv1.CurrentRow != null)
            {
                string pInfoId = dgv1.CurrentRow.Cells[20].Value.ToString();
                frmEdit editFrom = new frmEdit(int.Parse(pInfoId), "Edit");
                editFrom.ShowDialog();
            }
            GetDataSource("Search");
        }
        #endregion

        #region DoubleClick Message
        /// <summary>
        /// 双击信息进入编辑窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            if (dgv1.RowCount > 0)
            {
                if (dgv1.CurrentRow.Cells[0].Value.ToString() != "")
                {
                    btnEdit_Click(sender, e);
                }
            }
        }
        #endregion

        #region FillStatus
        /// <summary>
        /// 填充类型下拉框
        /// </summary>
        private void FillCBXConfirm()
        {
            DataTable dtFill = new DataTable();
            DataRow dr = null;
            dtFill.Columns.Add("Confirm");
            dtFill.Columns.Add("index");

            dr = dtFill.NewRow();
            dr["Confirm"] = "全部";
            dr["index"] = "-1";
            dtFill.Rows.Add(dr);

            dr = dtFill.NewRow();
            dr["Confirm"] = "未确认";
            dr["index"] = "0";
            dtFill.Rows.Add(dr);

            dr = dtFill.NewRow();
            dr["Confirm"] = "已确认";
            dr["index"] = "1";
            dtFill.Rows.Add(dr);

            cbxConfirm.Items.Clear();
            cbxConfirm.DataSource = dtFill;
            cbxConfirm.DisplayMember = "Confirm";
            cbxConfirm.ValueMember = "index";
            cbxConfirm.SelectedValue = "-1";
        }
        #endregion

        #region Cost Statistics
        /// <summary>
        /// 工伤费用统计信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStatistics_Click(object sender, EventArgs e)
        {
            frmStatistics pFrom = new frmStatistics();
            pFrom.ShowDialog();

            //导入工伤明细信息
//            string filePath = string.Empty;
//            string strError = string.Empty;
//            OpenFileDialog fileDialog = new OpenFileDialog();
//            fileDialog.Multiselect = false;
//            fileDialog.Title = "请选择Excel文件";
//            fileDialog.Filter = "Excel报表(*.xls)|*.xls|Excel其他文件(*.xlsx)|*.xlsx";
//            if (fileDialog.ShowDialog() == DialogResult.OK)
//            {
//                filePath = fileDialog.FileName;
//            }

//            object missing = System.Type.Missing;
//            Excel.Application myApp = new Excel.Application();
//            myApp.DisplayAlerts = false;
//            Excel.Workbook workBook = myApp.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
//            Excel.Worksheet worksheet = workBook.Worksheets[1] as Excel.Worksheet;
//            myApp.Visible = false;

//            string pAccidentDate = string.Empty;
//            string pName = string.Empty;
//            string pSex = string.Empty;
//            string pDp = string.Empty;
//            string pDuty = string.Empty;
//            string pBody = string.Empty;
//            string pCategory = string.Empty;
//            float pOccurredCost = 0;
//            string pProc = string.Empty;
//            string pRe = string.Empty;
//            string pMe = string.Empty;

//            string strInsert = string.Empty;

//            for (int i = 3; i <= worksheet.UsedRange.Rows.Count; i++)
//            {
//                try
//                {
//                    pAccidentDate = worksheet.Cells[i, 1].Text;
//                    pName = worksheet.Cells[i, 2].Value2;
//                    pSex = worksheet.Cells[i, 3].Value2;
//                    pDp = worksheet.Cells[i, 4].Value2;
//                    pDuty = worksheet.Cells[i, 5].Value2;
//                    pBody = worksheet.Cells[i, 6].Value2;
//                    pCategory = worksheet.Cells[i, 7].Value2;
//                    pOccurredCost = (float)worksheet.Cells[i, 8].Value2;
//                    pProc = worksheet.Cells[i, 11].Value2;
//                    pRe = worksheet.Cells[i, 12].Value2;
//                    pMe = worksheet.Cells[i, 13].Value2;

//                    strInsert = @"INSERT INTO IndustrailInjuryInfo(AccidentDate, Name, Sex, Duty, Department, InjuryCategory, Body, OccurredCost,
//	                                Process, Reason, Measure, Creator, CreateDate, IsConfirm, IsCancel)
//                                VALUES('" + pAccidentDate + "', '" + pName + "', '" + pSex + "', '" + pDuty + "', '" + pDp + "', '" + pCategory + "', '" + pBody + "', " + pOccurredCost.ToString() + @",
//                                        '" + pProc + "', '" + pRe + "', '" + pMe + "', 'Administrator', GETDATE(), 1, 0)";

//                    mysql.sqlcmd(strInsert);
//                }
//                catch (Exception ex)
//                {
//                    worksheet.Cells[i, 16] = ex.Message;
//                    strError += ex.Message;
//                    continue;
//                }
//                worksheet.Cells[i, 16] = "导入成功";
//            }
//            if (strError != "")
//            {
//                MessageBox.Show("有些信息导入出错，请查看最后一列的错误提示！");
//            }
//            else
//            {
//                MessageBox.Show("全部费用已经成功导入！");
//            }
//            myApp.Visible = true;
        }
        #endregion
        
        #region 从数据源读取记录到DataTable中
        /// <summary>
        /// 数据源
        /// </summary>
        //private void getinfo()
        //{
        //string strConn = "SERVER=127.0.0.1;DATABASE=NORTHWIND;UID=SA;PWD=ULTRATEL";   //数据库连接字符串
        //SqlConnection conn = new SqlConnection(strConn);
        //conn.Open();
        //string strSql = "SELECT * FROM CUSTOMERS";
        //SqlDataAdapter sda = new SqlDataAdapter(strSql, conn);
        //sda.Fill(ds, "ds");
        //conn.Close();
        //dt = ds.Tables[0];
        //InitDataSet();

        //string strSQL = "SELECT Name, AccidentDate, Duty, Department, Body, Category, Grade, Cost, Process, Reason, Measure, Remark FROM IndustrailInjuryInfo";
        //try
        //{
        //    dt = mysql.sqltb(strSQL);
        //}
        //catch (Exception ex)
        //{
        //    MessageBox.Show(ex.Message);
        //}
        //InitDataSet();
        //}
        #endregion

        #region Export Excel
        /// <summary>
        /// 导出工伤明显表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count == 0)
            {
                return;
            }

            Excel.Application xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            Excel.Workbook xlBook = xlApp.Workbooks.Add(System.Type.Missing);
            Excel.Worksheet worksheet = xlBook.Worksheets[1] as Excel.Worksheet;

            //标题
            string strTitle = string.Empty;
            //if (dtFrom.Value.Year == dtTo.Value.Year)
            //{
            //    strTitle = dtFrom.Value.Year.ToString() + "年";
            //    if (dtFrom.Value.Month == dtTo.Value.Month)
            //        strTitle += dtFrom.Value.Month.ToString() + "月 ";
            //    else strTitle += dtFrom.Value.Month.ToString() + " - " + dtTo.Value.Month.ToString() + "月 ";
            //}
            //else
            //{
            //    strTitle = dtFrom.Value.GetDateTimeFormats()[158].ToString() + " - " + dtTo.Value.GetDateTimeFormats()[158].ToString() + " ";
            //}
            strTitle += "工伤明细表";

            //生成表头
            (worksheet.Columns["A:A", Type.Missing] as Excel.Range).ColumnWidth = "18";      //工伤日期
            (worksheet.Columns["B:B", Type.Missing] as Excel.Range).ColumnWidth = "9";       //姓名
            (worksheet.Columns["C:C", Type.Missing] as Excel.Range).ColumnWidth = "5";       //性别
            (worksheet.Columns["D:D", Type.Missing] as Excel.Range).ColumnWidth = "12.25";   //部门
            (worksheet.Columns["E:E", Type.Missing] as Excel.Range).ColumnWidth = "9";       //岗位
            (worksheet.Columns["F:F", Type.Missing] as Excel.Range).ColumnWidth = "9";       //受伤部位
            (worksheet.Columns["G:G", Type.Missing] as Excel.Range).ColumnWidth = "9";       //工伤类型
            (worksheet.Columns["H:H", Type.Missing] as Excel.Range).ColumnWidth = "10";      //已产生费用
            (worksheet.Columns["I:I", Type.Missing] as Excel.Range).ColumnWidth = "10";      //当前费用
            (worksheet.Columns["J:J", Type.Missing] as Excel.Range).ColumnWidth = "10";      //总费用
            (worksheet.Columns["K:K", Type.Missing] as Excel.Range).ColumnWidth = "35";      //工伤过程
            (worksheet.Columns["L:L", Type.Missing] as Excel.Range).ColumnWidth = "35";      //原因
            (worksheet.Columns["M:M", Type.Missing] as Excel.Range).ColumnWidth = "35";      //改善对策
            (worksheet.Columns["N:N", Type.Missing] as Excel.Range).ColumnWidth = "15";      //备注
            (worksheet.Columns["O:O", Type.Missing] as Excel.Range).ColumnWidth = "8";       //确认

            worksheet.Cells[2, 1] = "工伤日期";
            worksheet.Cells[2, 2] = "姓名";
            worksheet.Cells[2, 3] = "性别";
            worksheet.Cells[2, 4] = "部门";
            worksheet.Cells[2, 5] = "岗位";
            worksheet.Cells[2, 6] = "受伤部位";
            worksheet.Cells[2, 7] = "工伤类型";
            worksheet.Cells[2, 8] = "已产生费用";
            worksheet.Cells[2, 9] = "当前费用";
            worksheet.Cells[2, 10] = "总费用";
            worksheet.Cells[2, 11] = "工伤过程";
            worksheet.Cells[2, 12] = "原因";
            worksheet.Cells[2, 13] = "改善对策";
            worksheet.Cells[2, 14] = "备注";
            worksheet.Cells[2, 15] = "确认";

            (worksheet.Columns["K:K", Type.Missing] as Excel.Range).WrapText = true;     //工伤过程
            (worksheet.Columns["L:L", Type.Missing] as Excel.Range).WrapText = true;     //原因
            (worksheet.Columns["M:M", Type.Missing] as Excel.Range).WrapText = true;     //改善对策
            (worksheet.Columns["N:N", Type.Missing] as Excel.Range).WrapText = true;     //备注

            worksheet.Rows.RowHeight = "85.50";
            worksheet.Cells[1, 1] = strTitle;
            Excel.Range rTitle=worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 15]];
            rTitle.Merge();
            rTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rTitle.Interior.Color = Color.FromArgb(26, 180, 240);
            rTitle.Font.Bold = true;
            rTitle.Font.Size = 24;
            rTitle.Font.Name = "宋体";
            
            Excel.Range rCTitle = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[2, 15]];
            rCTitle.Interior.Color = Color.FromArgb(135, 165, 175);

            ((Excel.Range)worksheet.Rows[1, Type.Missing]).RowHeight = "50";
            ((Excel.Range)worksheet.Rows[2, Type.Missing]).RowHeight = "25";
            ((Excel.Range)worksheet.Columns[2, Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            (worksheet.Columns["A:A", System.Type.Missing] as Excel.Range).NumberFormat = "yyyy-MM-dd hh:mm";

            //填充数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    worksheet.Cells[i + 3, 1] = dt.Rows[i]["AccidentDate"].ToString();
                    worksheet.Cells[i + 3, 2] = dt.Rows[i]["Name"].ToString();
                    worksheet.Cells[i + 3, 3] = dt.Rows[i]["Sex"].ToString();
                    worksheet.Cells[i + 3, 4] = dt.Rows[i]["Department"].ToString();
                    worksheet.Cells[i + 3, 5] = dt.Rows[i]["Duty"].ToString();
                    worksheet.Cells[i + 3, 6] = dt.Rows[i]["Body"].ToString();
                    worksheet.Cells[i + 3, 7] = dt.Rows[i]["InjuryCategory"].ToString();
                    worksheet.Cells[i + 3, 8] = dt.Rows[i]["OccurredCost"].ToString();
                    worksheet.Cells[i + 3, 9] = dt.Rows[i]["CurrentCost"].ToString();
                    worksheet.Cells[i + 3, 10] = "=SUM(RC[-1],RC[-2])";
                    worksheet.Cells[i + 3, 11] = dt.Rows[i]["Process"].ToString();
                    worksheet.Cells[i + 3, 12] = dt.Rows[i]["Reason"].ToString();
                    worksheet.Cells[i + 3, 13] = dt.Rows[i]["Measure"].ToString();
                    worksheet.Cells[i + 3, 14] = dt.Rows[i]["Remark"].ToString();
                    worksheet.Cells[i + 3, 15] = dt.Rows[i]["IsConfirm"].ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出错误：" + ex.Message);
                }
            }

            xlApp.Visible = true;
        }
        #endregion

        #region ChangecbxConfirm
        /// <summary>
        /// cbxConfirm_SelectedValueChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxConfirm_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetDataSource("Search");
        }
        #endregion

        #region Control Import Excel Button

        /// <summary>
        /// MouseEnter
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcel_MouseEnter(object sender, EventArgs e)
        {
            btnExcel.Visible = false;
            btnToExcel.Visible = true;
            btnImport.Visible = true;
        }

        /// <summary>
        /// MouseLeave
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void plExcel_MouseLeave(object sender, EventArgs e)
        {
            btnToExcel.Visible = false; ;
            btnImport.Visible = false;
            btnExcel.Visible = true;
        }
        #endregion

        #region Import Excel
        /// <summary>
        /// 导入工伤费用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string strError = string.Empty;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Title = "请选择Excel文件";
            fileDialog.Filter = "Excel报表(*.xls)|*.xls|Excel其他文件(*.xlsx)|*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = fileDialog.FileName;
            }
            if (filePath.Length <= 0) return;

            //导入费用
            object missing = Type.Missing;
            Excel.Application myApp = new Excel.Application();
            myApp.DisplayAlerts = false;
            Excel.Workbook workBook = myApp.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Excel.Worksheet worksheet = workBook.Worksheets[1] as Excel.Worksheet;
            myApp.Visible = false;

            if (worksheet.Cells[1, 1].Value2 != "工伤明细表")
            {
                MessageBox.Show("请选择[工伤明细表]导入");
                return;
            }
            string strAccidentDate = string.Empty;
            string strName = string.Empty;
            float pOccurredCost;
            float pCurrentCost;

            for (int i = 3; i <= worksheet.UsedRange.Rows.Count; i++)
            {
                try
                {
                    pOccurredCost = 0;
                    pCurrentCost = 0;
                    
                    if (worksheet.Cells[i, 1].Value2 != null)
                    {
                        strAccidentDate = worksheet.Cells[i, 1].Text;
                        strAccidentDate = strAccidentDate.Substring(0, 10);
                        pOccurredCost = (float)worksheet.Cells[i, 8].Value2;
                        pCurrentCost = (float)worksheet.Cells[i, 9].Value2;
                    }
                    else break;
                    strName = worksheet.Cells[i, 2].Value2;
                    string strSQL = @"UPDATE IndustrailInjuryInfo
                                SET OccurredCost = " + pOccurredCost.ToString() + ", CurrentCost = " + pCurrentCost.ToString() + @",
                                UpdateUser = '" + pblinfo.username + "', UpdateDate = '" + System.DateTime.Now.ToString() + @"'
                                WHERE CONVERT(VARCHAR(10), AccidentDate, 120) = '" + strAccidentDate + "' AND Name = '" + strName + "' AND IsCancel = 0";

                    mysql.sqlcmd(strSQL);
                }
                catch (Exception ex)
                {
                    worksheet.Cells[i, 16] = ex.Message;
                    strError += ex.Message;
                    continue;
                }
                worksheet.Cells[i, 16] = "导入成功";
            }
            if (strError != "")
            {
                MessageBox.Show("有些信息导入出错，请查看最后一列的错误提示！");
            }
            else
            {
                MessageBox.Show("全部费用已经成功导入！");
            }
            myApp.Visible = true;
        }
        #endregion
    }
}
