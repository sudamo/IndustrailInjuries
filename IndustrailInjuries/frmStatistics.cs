using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using pblClass;
using Excel = Microsoft.Office.Interop.Excel;

namespace IndustrailInjuries
{
    public partial class frmStatistics : Form
    {
        #region Parameter Area
        string strSQL = string.Empty;       //SQL语句
        int pDT = DateTime.Now.Year;        //当前年份
        int pMax = 5;                       //设定年最多只能查找多少数之内的数据
        int pYear = DateTime.Now.Year;
        #endregion

        #region Initialize Area
        /// <summary>
        /// 构造函数
        /// </summary>
        public frmStatistics()
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
        private void frmStatistics_Load(object sender, EventArgs e)
        {
            if (DateTime.Now.Month <= 3) pYear--;
            FillCbx();
            pl1.Dock = DockStyle.Top;
            dgv1.Dock = DockStyle.Fill;
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            GetDataSource();
        }
        #endregion

        #region Get Date Source
        /// <summary>
        /// 获取数据源
        /// </summary>
        private void GetDataSource()
        {
            DataTable dt = null;

            if (strSQL == string.Empty)
            {
                strSQL = @"WITH cte AS
                        (
	                        SELECT YEAR([AccidentDate]) [Year],
			                        CASE WHEN MONTH([AccidentDate]) = 1 THEN SUM(TotalCost) ELSE 0 END Jan,
			                        CASE WHEN MONTH([AccidentDate]) = 2 THEN SUM(TotalCost) ELSE 0 END Feb,
			                        CASE WHEN MONTH([AccidentDate]) = 3 THEN SUM(TotalCost) ELSE 0 END Mar,
			                        CASE WHEN MONTH([AccidentDate]) = 4 THEN SUM(TotalCost) ELSE 0 END Apr,
			                        CASE WHEN MONTH([AccidentDate]) = 5 THEN SUM(TotalCost) ELSE 0 END May,
			                        CASE WHEN MONTH([AccidentDate]) = 6 THEN SUM(TotalCost) ELSE 0 END Jun,
			                        CASE WHEN MONTH([AccidentDate]) = 7 THEN SUM(TotalCost) ELSE 0 END Jul,
			                        CASE WHEN MONTH([AccidentDate]) = 8 THEN SUM(TotalCost) ELSE 0 END Aug,
			                        CASE WHEN MONTH([AccidentDate]) = 9 THEN SUM(TotalCost) ELSE 0 END Sep,
			                        CASE WHEN MONTH([AccidentDate]) = 10 THEN SUM(TotalCost) ELSE 0 END Oct,
			                        CASE WHEN MONTH([AccidentDate]) = 11 THEN SUM(TotalCost) ELSE 0 END Nov,
			                        CASE WHEN MONTH([AccidentDate]) = 12 THEN SUM(TotalCost) ELSE 0 END [Dec]
	                        FROM IndustrailInjuryInfo
                            WHERE IsCancel = 0
	                        GROUP BY YEAR([AccidentDate]), MONTH([AccidentDate])
                        )
                        SELECT [Year], SUM(Jan) Jan, SUM(Feb) Feb, SUM(Mar) Mar, SUM(Apr) Apr, SUM(May) May, SUM(Jun) Jun, SUM(Jul) Jul, SUM(Aug) Aug,
	                        SUM(Sep) Sep, SUM(Oct) Oct, SUM(Nov) Nov, SUM([Dec]) [Dec],
                            SUM(Jan + Feb + Mar + Apr + May + Jun + Jul + Aug + Sep + Oct + Nov + [Dec]) Total
                        FROM cte
                        WHERE [Year] > " + (pDT - pMax).ToString() + @"
                        GROUP BY [Year] ORDER BY [Year]";
                try
                {
                    dt = mysql.sqltb(strSQL);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                dgv1.DataSource = dt;
                dgv1.EnableHeadersVisualStyles = false;
                dgv1.Columns[0].HeaderText = "年份";
                dgv1.Columns[1].HeaderText = "一月";
                dgv1.Columns[2].HeaderText = "二月";
                dgv1.Columns[3].HeaderText = "三月";
                dgv1.Columns[4].HeaderText = "四月";
                dgv1.Columns[5].HeaderText = "五月";
                dgv1.Columns[6].HeaderText = "六月";
                dgv1.Columns[7].HeaderText = "七月";
                dgv1.Columns[8].HeaderText = "八月";
                dgv1.Columns[9].HeaderText = "九月";
                dgv1.Columns[10].HeaderText = "十月";
                dgv1.Columns[11].HeaderText = "十一月";
                dgv1.Columns[12].HeaderText = "十二月";
                dgv1.Columns[13].HeaderText = "总计";
                dgv1.ColumnHeadersHeight = 35;
                dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            }
            else
            {
                try
                {
                    dt = mysql.sqltb(strSQL);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                dgv1.DataSource = null;
                dgv1.DataSource = dt;
                dgv1.EnableHeadersVisualStyles = false;
                dgv1.Columns[0].HeaderText = "年份";
                dgv1.Columns[1].HeaderText = "部门";
                dgv1.Columns[2].HeaderText = "一月";
                dgv1.Columns[3].HeaderText = "二月";
                dgv1.Columns[4].HeaderText = "三月";
                dgv1.Columns[5].HeaderText = "四月";
                dgv1.Columns[6].HeaderText = "五月";
                dgv1.Columns[7].HeaderText = "六月";
                dgv1.Columns[8].HeaderText = "七月";
                dgv1.Columns[9].HeaderText = "八月";
                dgv1.Columns[10].HeaderText = "九月";
                dgv1.Columns[11].HeaderText = "十月";
                dgv1.Columns[12].HeaderText = "十一月";
                dgv1.Columns[13].HeaderText = "十二月";
                dgv1.Columns[14].HeaderText = "总计";
                dgv1.ColumnHeadersHeight = 35;
                dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            }
        }
        #endregion

        #region Fill ComboBox
        /// <summary>
        /// 填充下拉框
        /// </summary>
        private void FillCbx()
        {
            DataTable dt = null;
            DataRow dr = null;

            //填充年份下拉框
            dt = new DataTable();
            dt.Columns.Add("Year");
            dt.Columns.Add("index");

            dr = dt.NewRow();
            dr["Year"] = "所有年份";
            dr["index"] = "0";
            dt.Rows.Add(dr);
            for (int i = 1; i <= pMax; i++)
            {
                dr = dt.NewRow();
                dr["Year"] = (pDT - i + 1).ToString();
                dr["index"] = i.ToString();
                dt.Rows.Add(dr);
            }
            cbxYear.DataSource = dt;
            cbxYear.DisplayMember = "Year";
            cbxYear.ValueMember = "index";

            //填充部门下拉框
            string strSQLD = "SELECT PDID, PDName FROM productdepartment WHERE Status = 0";
            dt = new DataTable();
            try
            {
                dt = mysql.sqltb(strSQLD);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            dr = dt.NewRow();
            dr["PDName"] = "所有部门";
            dr["PDID"] = "0";
            dt.Rows.Add(dr);
            //cbxCondition.Items.Clear();
            cbxDepartment.DataSource = dt;
            cbxDepartment.DisplayMember = "PDName";
            cbxDepartment.ValueMember = "PDID";
            cbxDepartment.SelectedValue = "0";
        }
        #endregion

        #region Set Date Source
        /// <summary>
        /// 配置数据源SQL语句
        /// </summary>
        private void SetDataSource()
        {
            strSQL = @"SELECT A.[Year], A.PDName,
	                        SUM(ISNULL(Jan, 0)) Jan, SUM(ISNULL(Feb, 0)) Feb, SUM(ISNULL(Mar, 0)) Mar, SUM(ISNULL(Apr, 0)) Apr, SUM(ISNULL(May, 0)) May,
	                        SUM(ISNULL(Jun, 0)) Jun, SUM(ISNULL(Jul, 0)) Jul, SUM(ISNULL(Aug, 0)) Aug, SUM(ISNULL(Sep, 0)) Sep, SUM(ISNULL(Oct, 0)) Oct,
	                        SUM(ISNULL(Nov, 0)) Nov, SUM(ISNULL([Dec], 0)) [Dec],
	                        ISNULL(SUM(Jan + Feb + Mar + Apr + May + Jun + Jul + Aug + Sep + Oct + Nov + [Dec]), 0) Total
                        FROM
                        (
	                        SELECT YEAR([AccidentDate]) [Year], D.PDName
	                        FROM IndustrailInjuryInfo C
	                        CROSS JOIN ProductDepartment D
	                        GROUP BY YEAR([AccidentDate]), D.PDName
                        ) A
                        LEFT JOIN
                        (
	                        SELECT YEAR([AccidentDate]) [Year], Department,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 1 THEN TotalCost ELSE 0 END) Jan,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 2 THEN TotalCost ELSE 0 END) Feb,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 3 THEN TotalCost ELSE 0 END) Mar,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 4 THEN TotalCost ELSE 0 END) Apr,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 5 THEN TotalCost ELSE 0 END) May,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 6 THEN TotalCost ELSE 0 END) Jun,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 7 THEN TotalCost ELSE 0 END) Jul,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 8 THEN TotalCost ELSE 0 END) Aug,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 9 THEN TotalCost ELSE 0 END) Sep,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 10 THEN TotalCost ELSE 0 END) Oct,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 11 THEN TotalCost ELSE 0 END) Nov,
			                        SUM(CASE WHEN MONTH([AccidentDate]) = 12 THEN TotalCost ELSE 0 END) [Dec]
	                        FROM IndustrailInjuryInfo
	                        WHERE IsCancel = 0
	                        GROUP BY YEAR([AccidentDate]), Department
                        ) B ON A.[Year] = B.[Year] AND A.PDName = B.Department";

            if (cbxYear.SelectedIndex == 0 && cbxDepartment.SelectedIndex == cbxDepartment.Items.Count - 1)
                strSQL += " WHERE A.[Year] >= " + (pDT - pMax).ToString();
            else if (cbxYear.SelectedIndex == 0 && cbxDepartment.SelectedIndex != cbxDepartment.Items.Count - 1)
                strSQL += " WHERE A.[Year] >= " + (pDT - pMax).ToString() + " AND A.PDName = '" + cbxDepartment.Text + "'";
            else if (cbxYear.SelectedIndex != 0 && cbxDepartment.SelectedIndex == cbxDepartment.Items.Count - 1)
                strSQL += " WHERE A.[Year] = " + cbxYear.Text;
            else
                strSQL += " WHERE A.[Year] = " + cbxYear.Text + " AND A.PDName = '" + cbxDepartment.Text + "'";
            strSQL += " GROUP BY A.[Year], A.PDName ORDER BY A.[Year], A.PDName";
            GetDataSource();
        }
        #endregion

        #region ToExcel
        /// <summary>
        /// 导出报表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("请稍等，功能实现中....."); return;

            string savePath = @"D:\工伤报告\";
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }

            //从数据库下载模板到本地
            object filePath = Application.StartupPath + "\\WordTemplate";
            if (!Directory.Exists(filePath.ToString()))
            {
                Directory.CreateDirectory(filePath.ToString());
            }
            //删除原来的模版
            if (File.Exists(filePath + "\\InjuryTemplate.xls"))
            {
                File.Delete(filePath + "\\InjuryTemplate.xls");               
            }
            byte[] tempFile;
            string strGetfile = "SELECT Contents FROM Templates WHERE TempName = 'InjuryTemplate.xls'";
            try
            {
                tempFile = (byte[])mysql.ExecuteScalar(strGetfile);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            FileStream fs = new FileStream(filePath + "\\InjuryTemplate.xls", FileMode.Create, FileAccess.Write);
            fs.Write(tempFile, 0, tempFile.Length);
            fs.Close();

            string fileName = pYear.ToString() + "年每月工伤分析情况.xls";
            object miss = Type.Missing;

            Excel.Application xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            Excel.Workbook xlBook = xlApp.Workbooks.Open(filePath + "\\InjuryTemplate.xls", miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss);
            Excel.Worksheet worksheet = xlBook.Worksheets[1] as Excel.Worksheet;

            DataTable dt = null;
            string strGetInfo = string.Empty;


            //月度工伤费用对比
            strGetInfo = @"WITH cte AS
                            (
	                            SELECT YEAR([AccidentDate]) [Year],
			                            CASE WHEN MONTH([AccidentDate]) = 1 THEN SUM(TotalCost) ELSE 0 END Jan,
			                            CASE WHEN MONTH([AccidentDate]) = 2 THEN SUM(TotalCost) ELSE 0 END Feb,
			                            CASE WHEN MONTH([AccidentDate]) = 3 THEN SUM(TotalCost) ELSE 0 END Mar,
			                            CASE WHEN MONTH([AccidentDate]) = 4 THEN SUM(TotalCost) ELSE 0 END Apr,
			                            CASE WHEN MONTH([AccidentDate]) = 5 THEN SUM(TotalCost) ELSE 0 END May,
			                            CASE WHEN MONTH([AccidentDate]) = 6 THEN SUM(TotalCost) ELSE 0 END Jun,
			                            CASE WHEN MONTH([AccidentDate]) = 7 THEN SUM(TotalCost) ELSE 0 END Jul,
			                            CASE WHEN MONTH([AccidentDate]) = 8 THEN SUM(TotalCost) ELSE 0 END Aug,
			                            CASE WHEN MONTH([AccidentDate]) = 9 THEN SUM(TotalCost) ELSE 0 END Sep,
			                            CASE WHEN MONTH([AccidentDate]) = 10 THEN SUM(TotalCost) ELSE 0 END Oct,
			                            CASE WHEN MONTH([AccidentDate]) = 11 THEN SUM(TotalCost) ELSE 0 END Nov,
			                            CASE WHEN MONTH([AccidentDate]) = 12 THEN SUM(TotalCost) ELSE 0 END [Dec]
	                            FROM IndustrailInjuryInfo
                                WHERE IsCancel = 0
	                            GROUP BY YEAR([AccidentDate]), MONTH([AccidentDate])
                            )
                            SELECT B.[Year],
	                            CONVERT(DECIMAL(9, 0), SUM(ISNULL(Jan, 0))) Jan, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Feb, 0))) Feb, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Mar, 0))) Mar, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Apr, 0))) Apr, CONVERT(DECIMAL(9, 0), SUM(ISNULL(May, 0))) May,
	                            CONVERT(DECIMAL(9, 0), SUM(ISNULL(Jun, 0))) Jun, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Jul, 0))) Jul, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Aug, 0))) Aug, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Sep, 0))) Sep, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Oct, 0))) Oct,
	                            CONVERT(DECIMAL(9, 0), SUM(ISNULL(Nov, 0))) Nov, CONVERT(DECIMAL(9, 0), SUM(ISNULL([Dec], 0))) [Dec]
                            FROM cte A
                            RIGHT JOIN
                            (
	                            SELECT " + pYear.ToString() + @" [Year]
	                            UNION
	                            SELECT " + pYear.ToString() + @" - 1 [Year]
                            ) B ON A.[Year] = B.[Year]
                            WHERE B.[Year] >= " + pYear.ToString() + @" - 1
                            GROUP BY B.[Year]
                            ORDER BY B.[Year]";
            try
            {
                dt = mysql.sqltb(strGetInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                xlBook.Close(null, null, null);
                xlApp.Workbooks.Close();
                xlApp.Application.Quit();
                xlApp.Quit();
                worksheet = null;
                xlBook = null;
                xlApp = null;
                return;
            }

            //标题
            string strTitle1 = string.Empty;
            strTitle1 = DateTime.Now.Year.ToString() + "年月度工伤费用对比图";

            //填充数据
            for (int i = 0; i < 2; i++)
            {
                try
                {
                    worksheet.Cells[i + 3, 1] = dt.Rows[i]["Year"].ToString() + "年";
                    worksheet.Cells[i + 3, 2] = dt.Rows[i]["Jan"].ToString();
                    worksheet.Cells[i + 3, 3] = dt.Rows[i]["Feb"].ToString();
                    worksheet.Cells[i + 3, 4] = dt.Rows[i]["Mar"].ToString();
                    worksheet.Cells[i + 3, 5] = dt.Rows[i]["Apr"].ToString();
                    worksheet.Cells[i + 3, 6] = dt.Rows[i]["May"].ToString();
                    worksheet.Cells[i + 3, 7] = dt.Rows[i]["Jun"].ToString();
                    worksheet.Cells[i + 3, 8] = dt.Rows[i]["Jul"].ToString();
                    worksheet.Cells[i + 3, 9] = dt.Rows[i]["Aug"].ToString();
                    worksheet.Cells[i + 3, 10] = dt.Rows[i]["Sep"].ToString();
                    worksheet.Cells[i + 3, 11] = dt.Rows[i]["Oct"].ToString();
                    worksheet.Cells[i + 3, 12] = dt.Rows[i]["Nov"].ToString();
                    worksheet.Cells[i + 3, 13] = dt.Rows[i]["Dec"].ToString();
                    worksheet.Cells[i + 3, 14] = "=SUM(RC[-12]:RC[-1])";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出错误：" + ex.Message);
                }
            }

            //月度工伤件数对比
            strGetInfo = @"WITH cte AS
                        (
	                        SELECT YEAR([AccidentDate]) [Year],
			                        CASE WHEN MONTH([AccidentDate]) = 1 THEN COUNT(1) ELSE 0 END Jan,
			                        CASE WHEN MONTH([AccidentDate]) = 2 THEN COUNT(1) ELSE 0 END Feb,
			                        CASE WHEN MONTH([AccidentDate]) = 3 THEN COUNT(1) ELSE 0 END Mar,
			                        CASE WHEN MONTH([AccidentDate]) = 4 THEN COUNT(1) ELSE 0 END Apr,
			                        CASE WHEN MONTH([AccidentDate]) = 5 THEN COUNT(1) ELSE 0 END May,
			                        CASE WHEN MONTH([AccidentDate]) = 6 THEN COUNT(1) ELSE 0 END Jun,
			                        CASE WHEN MONTH([AccidentDate]) = 7 THEN COUNT(1) ELSE 0 END Jul,
			                        CASE WHEN MONTH([AccidentDate]) = 8 THEN COUNT(1) ELSE 0 END Aug,
			                        CASE WHEN MONTH([AccidentDate]) = 9 THEN COUNT(1) ELSE 0 END Sep,
			                        CASE WHEN MONTH([AccidentDate]) = 10 THEN COUNT(1) ELSE 0 END Oct,
			                        CASE WHEN MONTH([AccidentDate]) = 11 THEN COUNT(1) ELSE 0 END Nov,
			                        CASE WHEN MONTH([AccidentDate]) = 12 THEN COUNT(1) ELSE 0 END [Dec]
	                        FROM IndustrailInjuryInfo
                            WHERE IsCancel = 0
	                        GROUP BY YEAR([AccidentDate]), MONTH([AccidentDate])
                        )
                        SELECT B.[Year],
	                        SUM(ISNULL(Jan, 0)) Jan, SUM(ISNULL(Feb, 0)) Feb, SUM(ISNULL(Mar, 0)) Mar, SUM(ISNULL(Apr, 0)) Apr, SUM(ISNULL(May, 0)) May,
	                        SUM(ISNULL(Jun, 0)) Jun, SUM(ISNULL(Jul, 0)) Jul, SUM(ISNULL(Aug, 0)) Aug, SUM(ISNULL(Sep, 0)) Sep, SUM(ISNULL(Oct, 0)) Oct,
	                        SUM(ISNULL(Nov, 0)) Nov, SUM(ISNULL([Dec], 0)) [Dec]
                        FROM cte A
                        RIGHT JOIN
                        (
	                        SELECT " + pYear.ToString() + @" [Year]
	                        UNION
	                        SELECT " + pYear.ToString() + @" - 1 [Year]
                        ) B ON A.[Year] = B.[Year]
                        WHERE B.[Year] >= " + pYear.ToString() + @" - 1
                        GROUP BY B.[Year]
                        ORDER BY B.[Year]";
            try
            {
                dt = mysql.sqltb(strGetInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                xlBook.Close(null, null, null);
                xlApp.Workbooks.Close();
                xlApp.Application.Quit();
                xlApp.Quit();
                worksheet = null;
                xlBook = null;
                xlApp = null;
                return;
            }

            //标题
            strTitle1 = DateTime.Now.Year.ToString() + "年月度工伤件数对比";
            worksheet.Cells[27, 1] = strTitle1;

            //填充数据
            for (int i = 0; i < 2; i++)
            {
                try
                {
                    worksheet.Cells[i + 29, 1] = dt.Rows[i]["Year"].ToString() + "年";
                    worksheet.Cells[i + 29, 2] = dt.Rows[i]["Jan"].ToString();
                    worksheet.Cells[i + 29, 3] = dt.Rows[i]["Feb"].ToString();
                    worksheet.Cells[i + 29, 4] = dt.Rows[i]["Mar"].ToString();
                    worksheet.Cells[i + 29, 5] = dt.Rows[i]["Apr"].ToString();
                    worksheet.Cells[i + 29, 6] = dt.Rows[i]["May"].ToString();
                    worksheet.Cells[i + 29, 7] = dt.Rows[i]["Jun"].ToString();
                    worksheet.Cells[i + 29, 8] = dt.Rows[i]["Jul"].ToString();
                    worksheet.Cells[i + 29, 9] = dt.Rows[i]["Aug"].ToString();
                    worksheet.Cells[i + 29, 10] = dt.Rows[i]["Sep"].ToString();
                    worksheet.Cells[i + 29, 11] = dt.Rows[i]["Oct"].ToString();
                    worksheet.Cells[i + 29, 12] = dt.Rows[i]["Nov"].ToString();
                    worksheet.Cells[i + 29, 13] = dt.Rows[i]["Dec"].ToString();
                    worksheet.Cells[i + 29, 14] = "=SUM(RC[-12]:RC[-1])";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出错误：" + ex.Message);
                }
            }

            //年度工伤类型分析
            strGetInfo = @"GetIndustrialInjuries";
            try
            {
                dt = mysql.sqltb(strGetInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                xlBook.Close(null, null, null);
                xlApp.Workbooks.Close();
                xlApp.Application.Quit();
                xlApp.Quit();
                worksheet = null;
                xlBook = null;
                xlApp = null;
                return;
            }

            //标题
            strTitle1 = DateTime.Now.Year.ToString() + "工伤类型分析表";
            worksheet.Cells[52, 1] = strTitle1;

            //填充数据
            for (int i = 0; i < 12; i++)
            {
                try
                {
                    worksheet.Cells[i + 55, 1] = (i + 1).ToString() + "月份";
                    worksheet.Cells[i + 55, 2] = dt.Rows[i]["物体打击"].ToString();
                    worksheet.Cells[i + 55, 3] = dt.Rows[i]["机械伤害"].ToString();
                    worksheet.Cells[i + 55, 4] = dt.Rows[i]["起重伤害"].ToString();
                    worksheet.Cells[i + 55, 5] = dt.Rows[i]["摔伤"].ToString();
                    worksheet.Cells[i + 55, 6] = dt.Rows[i]["灼伤"].ToString();
                    worksheet.Cells[i + 55, 7] = dt.Rows[i]["车辆伤害"].ToString();
                    worksheet.Cells[i + 55, 8] = dt.Rows[i]["碰伤"].ToString();
                    worksheet.Cells[i + 55, 9] = dt.Rows[i]["其他伤害"].ToString();
                    worksheet.Cells[i + 55, 10] = "=SUM(RC[-8]:RC[-1])";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出错误：" + ex.Message);
                }
            }

            //各车间工伤件数及费用

            worksheet.ChartObjects(miss);
            xlBook.Saved = true;
            savePath += fileName;

            //保存文件时处理同名情况
            bool pSave = false;
            int pInt = 1;
            while (pSave == false)
            {
                if (File.Exists(savePath))
                {
                    if (savePath.LastIndexOf(")") == savePath.Length - 5 && savePath.LastIndexOf(")") != 0)
                        savePath = savePath.Remove(savePath.LastIndexOf("(") + 1, savePath.Length - savePath.LastIndexOf("(") - 1) + pInt.ToString() + ").xls";
                    else savePath = savePath.Remove(savePath.Length - 4, 4) +"("+ pInt.ToString() + ").xls";
                    pInt++;
                }
                else
                    pSave = true;
            }
            //文件保存
            try
            {
                xlBook.SaveAs(savePath, miss, miss, miss, miss, miss, Excel.XlSaveAsAccessMode.xlExclusive, miss, miss, miss, miss);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                xlBook.Close(null, null, null);
                xlApp.Workbooks.Close();
                xlApp.Application.Quit();
                xlApp.Quit();
                worksheet = null;
                xlBook = null;
                xlApp = null;

                GC.Collect();
            }
            MessageBox.Show("报表导出成功：[" + savePath + "]");

            //Excel.Application xlApp = new Excel.Application();
            //xlApp.DisplayAlerts = false;
            //Excel.Workbook xlBook = xlApp.Workbooks.Add(miss);
            //Excel.Worksheet worksheet = xlBook.Worksheets[1] as Excel.Worksheet;

            //DataTable dt = null;
            //string strGetInfo = string.Empty;

            ////月度工伤费用对比
            //strGetInfo = @"WITH cte AS
            //                (
            //                 SELECT YEAR([AccidentDate]) [Year],
            //                   CASE WHEN MONTH([AccidentDate]) = 1 THEN SUM(TotalCost) ELSE 0 END Jan,
            //                   CASE WHEN MONTH([AccidentDate]) = 2 THEN SUM(TotalCost) ELSE 0 END Feb,
            //                   CASE WHEN MONTH([AccidentDate]) = 3 THEN SUM(TotalCost) ELSE 0 END Mar,
            //                   CASE WHEN MONTH([AccidentDate]) = 4 THEN SUM(TotalCost) ELSE 0 END Apr,
            //                   CASE WHEN MONTH([AccidentDate]) = 5 THEN SUM(TotalCost) ELSE 0 END May,
            //                   CASE WHEN MONTH([AccidentDate]) = 6 THEN SUM(TotalCost) ELSE 0 END Jun,
            //                   CASE WHEN MONTH([AccidentDate]) = 7 THEN SUM(TotalCost) ELSE 0 END Jul,
            //                   CASE WHEN MONTH([AccidentDate]) = 8 THEN SUM(TotalCost) ELSE 0 END Aug,
            //                   CASE WHEN MONTH([AccidentDate]) = 9 THEN SUM(TotalCost) ELSE 0 END Sep,
            //                   CASE WHEN MONTH([AccidentDate]) = 10 THEN SUM(TotalCost) ELSE 0 END Oct,
            //                   CASE WHEN MONTH([AccidentDate]) = 11 THEN SUM(TotalCost) ELSE 0 END Nov,
            //                   CASE WHEN MONTH([AccidentDate]) = 12 THEN SUM(TotalCost) ELSE 0 END [Dec]
            //                 FROM IndustrailInjuryInfo
            //                    WHERE IsCancel = 0
            //                 GROUP BY YEAR([AccidentDate]), MONTH([AccidentDate])
            //                )
            //                SELECT B.[Year],
            //                 CONVERT(DECIMAL(9, 0), SUM(ISNULL(Jan, 0))) Jan, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Feb, 0))) Feb, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Mar, 0))) Mar, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Apr, 0))) Apr, CONVERT(DECIMAL(9, 0), SUM(ISNULL(May, 0))) May,
            //                 CONVERT(DECIMAL(9, 0), SUM(ISNULL(Jun, 0))) Jun, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Jul, 0))) Jul, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Aug, 0))) Aug, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Sep, 0))) Sep, CONVERT(DECIMAL(9, 0), SUM(ISNULL(Oct, 0))) Oct,
            //                 CONVERT(DECIMAL(9, 0), SUM(ISNULL(Nov, 0))) Nov, CONVERT(DECIMAL(9, 0), SUM(ISNULL([Dec], 0))) [Dec]
            //                FROM cte A
            //                RIGHT JOIN
            //                (
            //                 SELECT " + pYear.ToString() + @" [Year]
            //                 UNION
            //                 SELECT " + pYear.ToString() + @" - 1 [Year]
            //                ) B ON A.[Year] = B.[Year]
            //                WHERE B.[Year] >= " + pYear.ToString() + @" - 1
            //                GROUP BY B.[Year]
            //                ORDER BY B.[Year]";
            //try
            //{
            //    dt = mysql.sqltb(strGetInfo);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    xlBook.Close(null, null, null);
            //    xlApp.Workbooks.Close();
            //    xlApp.Application.Quit();
            //    xlApp.Quit();
            //    worksheet = null;
            //    xlBook = null;
            //    xlApp = null;
            //    return;
            //}

            ////标题
            //string strTitle1 = string.Empty;
            //strTitle1 = DateTime.Now.Year.ToString() + "年月度工伤费用对比图";

            ////生成表头
            //(worksheet.Columns["A:A", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["B:B", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["C:C", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["D:D", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["E:E", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["F:F", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["G:G", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["H:H", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["I:I", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["J:J", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["K:K", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["L:L", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["M:M", Type.Missing] as Excel.Range).ColumnWidth = "8.50";
            //(worksheet.Columns["N:N", Type.Missing] as Excel.Range).ColumnWidth = "8.50";

            //worksheet.Cells[2, 1] = "年/月";
            //worksheet.Cells[2, 2] = "1月";
            //worksheet.Cells[2, 3] = "2月";
            //worksheet.Cells[2, 4] = "3月";
            //worksheet.Cells[2, 5] = "4月";
            //worksheet.Cells[2, 6] = "5月";
            //worksheet.Cells[2, 7] = "6月";
            //worksheet.Cells[2, 8] = "7月";
            //worksheet.Cells[2, 9] = "8月";
            //worksheet.Cells[2, 10] = "9月";
            //worksheet.Cells[2, 11] = "10月";
            //worksheet.Cells[2, 12] = "11月";
            //worksheet.Cells[2, 13] = "12月";
            //worksheet.Cells[2, 14] = "总计";

            //worksheet.Cells[1, 1] = strTitle1;
            //Excel.Range rTitle = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 14]];
            //rTitle.Merge();
            //rTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ////rTitle.Interior.Color = Color.FromArgb(26, 180, 240);
            //rTitle.Font.Bold = true;
            //rTitle.Font.Size = 20;
            //rTitle.Font.Name = "宋体";
            //((Excel.Range)worksheet.Rows[1, Type.Missing]).RowHeight = "42";

            ////((Excel.Range)worksheet.Columns[2, Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ////(worksheet.Columns["A:A", Type.Missing] as Excel.Range).NumberFormat = "yyyy-MM-dd hh:mm";

            ////填充数据
            //for (int i = 0; i < 2; i++)
            //{
            //    try
            //    {
            //        worksheet.Cells[i + 3, 1] = dt.Rows[i]["Year"].ToString() + "年";
            //        worksheet.Cells[i + 3, 2] = dt.Rows[i]["Jan"].ToString();
            //        worksheet.Cells[i + 3, 3] = dt.Rows[i]["Feb"].ToString();
            //        worksheet.Cells[i + 3, 4] = dt.Rows[i]["Mar"].ToString();
            //        worksheet.Cells[i + 3, 5] = dt.Rows[i]["Apr"].ToString();
            //        worksheet.Cells[i + 3, 6] = dt.Rows[i]["May"].ToString();
            //        worksheet.Cells[i + 3, 7] = dt.Rows[i]["Jun"].ToString();
            //        worksheet.Cells[i + 3, 8] = dt.Rows[i]["Jul"].ToString();
            //        worksheet.Cells[i + 3, 9] = dt.Rows[i]["Aug"].ToString();
            //        worksheet.Cells[i + 3, 10] = dt.Rows[i]["Sep"].ToString();
            //        worksheet.Cells[i + 3, 11] = dt.Rows[i]["Oct"].ToString();
            //        worksheet.Cells[i + 3, 12] = dt.Rows[i]["Nov"].ToString();
            //        worksheet.Cells[i + 3, 13] = dt.Rows[i]["Dec"].ToString();
            //        worksheet.Cells[i + 3, 14] = "=SUM(RC[-12]:RC[-1])";
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("导出错误：" + ex.Message);
            //    }
            //}
            //((Excel.Range)worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[4, 14]]).Borders.LineStyle = 1;

            //Excel.Range pDataRange = (Excel.Range)worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[4, 13]];
            ////Excel.Series pSeries;

            ////Excel.Chart newChart = (Excel.Chart)((Excel.Workbook)worksheet.Parent).Charts.Add(miss, miss, miss, miss);
            //Excel.Chart newChart = xlBook.Charts.Add(miss, miss, miss, miss);
            //newChart.ChartWizard(pDataRange, Excel.XlChartType.xlColumnClustered, miss, miss, miss, miss, miss, miss, miss);

            //newChart.HasTitle = true;
            //newChart.ChartTitle.Text = "月度工伤费用统计图";
            //newChart.Location(Excel.XlChartLocation.xlLocationAsObject, "Sheet1");
            ////xlBook.ActiveChart.PlotArea.Interior.ColorIndex = 15;
            //xlBook.ActiveChart.PlotArea.Interior.Color = Color.FromArgb(192, 192, 192);

            //pDataRange = (Excel.Range)worksheet.Rows.get_Item(6, miss);
            ////xlBook.ActiveChart.Legend.Top = (double)pDataRange.Top;
            //worksheet.Shapes.Item("Chart 1").Top = (float)pDataRange.Top;
            //pDataRange = (Excel.Range)worksheet.Columns.get_Item(2, miss);
            ////xlBook.ActiveChart.Legend.Left = (double)pDataRange.Left;
            //worksheet.Shapes.Item("Chart 1").Left = (float)pDataRange.Left;
            //worksheet.Shapes.Item("Chart 1").Width = 680;
            //worksheet.Shapes.Item("Chart 1").Height = 250;

            ////newChart.Legend.Width = 150;
            ////newChart.Legend.Height = 8;
            ////newChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;

            ////Excel.ChartObject pCO = (Excel.ChartObject)worksheet.ChartObjects(worksheet.Application.ActiveChart.Name.Replace(worksheet.Name + "", ""));
            ////Excel.ChartObject pCO = worksheet.ChartObjects(miss);
            ////pCO.Left = Single.Parse(pDataRange.Left.ToString());
            ////pCO.Top = Single.Parse(pDataRange.Top.ToString()) + Single.Parse(pDataRange.Height.ToString());
            ////pCO.RoundedCorners = true;            

            ////xlBook.Charts.Add(miss, miss, 1, miss);
            ////newChart.ChartType = Excel.XlChartType.xlColumnClustered;
            ////newChart.SetSourceData((Excel.Range)worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[4, 13]], miss);

            ////月度工伤件数对比
            //strGetInfo = @"WITH cte AS
            //            (
            //             SELECT YEAR([AccidentDate]) [Year],
            //               CASE WHEN MONTH([AccidentDate]) = 1 THEN COUNT(1) ELSE 0 END Jan,
            //               CASE WHEN MONTH([AccidentDate]) = 2 THEN COUNT(1) ELSE 0 END Feb,
            //               CASE WHEN MONTH([AccidentDate]) = 3 THEN COUNT(1) ELSE 0 END Mar,
            //               CASE WHEN MONTH([AccidentDate]) = 4 THEN COUNT(1) ELSE 0 END Apr,
            //               CASE WHEN MONTH([AccidentDate]) = 5 THEN COUNT(1) ELSE 0 END May,
            //               CASE WHEN MONTH([AccidentDate]) = 6 THEN COUNT(1) ELSE 0 END Jun,
            //               CASE WHEN MONTH([AccidentDate]) = 7 THEN COUNT(1) ELSE 0 END Jul,
            //               CASE WHEN MONTH([AccidentDate]) = 8 THEN COUNT(1) ELSE 0 END Aug,
            //               CASE WHEN MONTH([AccidentDate]) = 9 THEN COUNT(1) ELSE 0 END Sep,
            //               CASE WHEN MONTH([AccidentDate]) = 10 THEN COUNT(1) ELSE 0 END Oct,
            //               CASE WHEN MONTH([AccidentDate]) = 11 THEN COUNT(1) ELSE 0 END Nov,
            //               CASE WHEN MONTH([AccidentDate]) = 12 THEN COUNT(1) ELSE 0 END [Dec]
            //             FROM IndustrailInjuryInfo
            //                WHERE IsCancel = 0
            //             GROUP BY YEAR([AccidentDate]), MONTH([AccidentDate])
            //            )
            //            SELECT B.[Year],
            //             SUM(ISNULL(Jan, 0)) Jan, SUM(ISNULL(Feb, 0)) Feb, SUM(ISNULL(Mar, 0)) Mar, SUM(ISNULL(Apr, 0)) Apr, SUM(ISNULL(May, 0)) May,
            //             SUM(ISNULL(Jun, 0)) Jun, SUM(ISNULL(Jul, 0)) Jul, SUM(ISNULL(Aug, 0)) Aug, SUM(ISNULL(Sep, 0)) Sep, SUM(ISNULL(Oct, 0)) Oct,
            //             SUM(ISNULL(Nov, 0)) Nov, SUM(ISNULL([Dec], 0)) [Dec]
            //            FROM cte A
            //            RIGHT JOIN
            //            (
            //             SELECT " + pYear.ToString() + @" [Year]
            //             UNION
            //             SELECT " + pYear.ToString() + @" - 1 [Year]
            //            ) B ON A.[Year] = B.[Year]
            //            WHERE B.[Year] >= " + pYear.ToString() + @" - 1
            //            GROUP BY B.[Year]
            //            ORDER BY B.[Year]";
            //try
            //{
            //    dt = mysql.sqltb(strGetInfo);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    xlBook.Close(null, null, null);
            //    xlApp.Workbooks.Close();
            //    xlApp.Application.Quit();
            //    xlApp.Quit();
            //    worksheet = null;
            //    xlBook = null;
            //    xlApp = null;
            //    return;
            //}

            ////标题
            //strTitle1 = DateTime.Now.Year.ToString() + "年月度工伤件数对比";
            //worksheet.Cells[27, 1] = strTitle1;
            //rTitle = worksheet.Range[worksheet.Cells[27, 1], worksheet.Cells[27, 14]];
            //rTitle.Merge();
            //rTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //rTitle.Font.Bold = true;
            //rTitle.Font.Size = 20;
            //rTitle.Font.Name = "宋体";
            //((Excel.Range)worksheet.Rows[27, Type.Missing]).RowHeight = "42";

            ////生成表头
            //worksheet.Cells[28, 1] = "年/月";
            //worksheet.Cells[28, 2] = "1月";
            //worksheet.Cells[28, 3] = "2月";
            //worksheet.Cells[28, 4] = "3月";
            //worksheet.Cells[28, 5] = "4月";
            //worksheet.Cells[28, 6] = "5月";
            //worksheet.Cells[28, 7] = "6月";
            //worksheet.Cells[28, 8] = "7月";
            //worksheet.Cells[28, 9] = "8月";
            //worksheet.Cells[28, 10] = "9月";
            //worksheet.Cells[28, 11] = "10月";
            //worksheet.Cells[28, 12] = "11月";
            //worksheet.Cells[28, 13] = "12月";
            //worksheet.Cells[28, 14] = "总计";

            ////填充数据
            //for (int i = 0; i < 2; i++)
            //{
            //    try
            //    {
            //        worksheet.Cells[i + 29, 1] = dt.Rows[i]["Year"].ToString() + "年";
            //        worksheet.Cells[i + 29, 2] = dt.Rows[i]["Jan"].ToString();
            //        worksheet.Cells[i + 29, 3] = dt.Rows[i]["Feb"].ToString();
            //        worksheet.Cells[i + 29, 4] = dt.Rows[i]["Mar"].ToString();
            //        worksheet.Cells[i + 29, 5] = dt.Rows[i]["Apr"].ToString();
            //        worksheet.Cells[i + 29, 6] = dt.Rows[i]["May"].ToString();
            //        worksheet.Cells[i + 29, 7] = dt.Rows[i]["Jun"].ToString();
            //        worksheet.Cells[i + 29, 8] = dt.Rows[i]["Jul"].ToString();
            //        worksheet.Cells[i + 29, 9] = dt.Rows[i]["Aug"].ToString();
            //        worksheet.Cells[i + 29, 10] = dt.Rows[i]["Sep"].ToString();
            //        worksheet.Cells[i + 29, 11] = dt.Rows[i]["Oct"].ToString();
            //        worksheet.Cells[i + 29, 12] = dt.Rows[i]["Nov"].ToString();
            //        worksheet.Cells[i + 29, 13] = dt.Rows[i]["Dec"].ToString();
            //        worksheet.Cells[i + 29, 14] = "=SUM(RC[-12]:RC[-1])";
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("导出错误：" + ex.Message);
            //    }
            //}
            //((Excel.Range)worksheet.Range[worksheet.Cells[28, 1], worksheet.Cells[30, 14]]).Borders.LineStyle = 1;

            //pDataRange = (Excel.Range)worksheet.Range[worksheet.Cells[28, 1], worksheet.Cells[30, 13]];

            //newChart = (Excel.Chart)((Excel.Workbook)worksheet.Parent).Charts.Add(miss, miss, miss, miss);
            //newChart.ChartWizard(pDataRange, Excel.XlChartType.xlColumnClustered, miss, miss, miss, miss, miss, miss, miss);

            //newChart.HasTitle = true;
            //newChart.ChartTitle.Text = "月度工伤件数统计图";
            //newChart.Location(Excel.XlChartLocation.xlLocationAsObject, "Sheet1");
            //xlBook.ActiveChart.PlotArea.Interior.ColorIndex = 15;
            ////xlBook.ActiveChart.PlotArea.Interior.Color = Color.FromArgb(192, 192, 192);

            //pDataRange = (Excel.Range)worksheet.Rows.get_Item(32, miss);
            //worksheet.Shapes.Item("Chart 2").Top = (float)(double)pDataRange.Top;
            //pDataRange = (Excel.Range)worksheet.Columns.get_Item(2, miss);
            //worksheet.Shapes.Item("Chart 2").Left = (float)(double)pDataRange.Left;
            //worksheet.Shapes.Item("Chart 2").Width = 680;
            //worksheet.Shapes.Item("Chart 2").Height = 250;

            ////年度工伤类型分析
            //strGetInfo = @"GetIndustrialInjuries";
            //try
            //{
            //    dt = mysql.sqltb(strGetInfo);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    xlBook.Close(null, null, null);
            //    xlApp.Workbooks.Close();
            //    xlApp.Application.Quit();
            //    xlApp.Quit();
            //    worksheet = null;
            //    xlBook = null;
            //    xlApp = null;
            //    return;
            //}

            ////标题
            //strTitle1 = DateTime.Now.Year.ToString() + "工伤类型分析表";
            //worksheet.Cells[52, 1] = strTitle1;
            //rTitle = worksheet.Range[worksheet.Cells[52, 1], worksheet.Cells[52, 10]];
            //rTitle.Merge();
            //rTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //rTitle.Font.Bold = true;
            //rTitle.Font.Size = 20;
            //rTitle.Font.Name = "宋体";
            //((Excel.Range)worksheet.Rows[52, Type.Missing]).RowHeight = "42";

            ////生成表头
            //worksheet.Cells[53, 1] = "月份/类型";
            //((Excel.Range)worksheet.Range[worksheet.Cells[53, 1], worksheet.Cells[54, 1]]).Merge();
            //worksheet.Cells[53, 2] = "类型";
            //Excel.Range rTmp = worksheet.Range[worksheet.Cells[53, 2], worksheet.Cells[53, 9]];
            //rTmp.Merge();
            //rTmp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //worksheet.Cells[53, 10] = "合计";
            //((Excel.Range)worksheet.Range[worksheet.Cells[53, 10], worksheet.Cells[54, 10]]).Merge();

            //worksheet.Cells[54, 2] = "物体打击";
            //worksheet.Cells[54, 3] = "机械伤害";
            //worksheet.Cells[54, 4] = "起重伤害";
            //worksheet.Cells[54, 5] = "摔伤";
            //worksheet.Cells[54, 6] = "灼伤";
            //worksheet.Cells[54, 7] = "车辆伤害";
            //worksheet.Cells[54, 8] = "碰伤";
            //worksheet.Cells[54, 9] = "其他伤害";

            ////填充数据
            //for (int i = 0; i < 12; i++)
            //{
            //    try
            //    {
            //        worksheet.Cells[i + 55, 1] = (i + 1).ToString() + "月份";
            //        worksheet.Cells[i + 55, 2] = dt.Rows[i]["物体打击"].ToString();
            //        worksheet.Cells[i + 55, 3] = dt.Rows[i]["机械伤害"].ToString();
            //        worksheet.Cells[i + 55, 4] = dt.Rows[i]["起重伤害"].ToString();
            //        worksheet.Cells[i + 55, 5] = dt.Rows[i]["摔伤"].ToString();
            //        worksheet.Cells[i + 55, 6] = dt.Rows[i]["灼伤"].ToString();
            //        worksheet.Cells[i + 55, 7] = dt.Rows[i]["车辆伤害"].ToString();
            //        worksheet.Cells[i + 55, 8] = dt.Rows[i]["碰伤"].ToString();
            //        worksheet.Cells[i + 55, 9] = dt.Rows[i]["其他伤害"].ToString();
            //        worksheet.Cells[i + 55, 10] = "=SUM(RC[-8]:RC[-1])";
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("导出错误：" + ex.Message);
            //    }
            //}
            //worksheet.Cells[67, 1] = "合计";
            //worksheet.Cells[67, 2] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 3] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 4] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 5] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 6] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 7] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 8] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 9] = "=SUM(R[-12]C:R[-1]C)";
            //worksheet.Cells[67, 10] = "=SUM(R[-12]C:R[-1]C)";

            //worksheet.Cells[68, 1] = "比率";
            //worksheet.Cells[68, 2] = "=R[-1]C/R[-1]C[8]";
            //worksheet.Cells[68, 3] = "=R[-1]C/R[-1]C[7]";
            //worksheet.Cells[68, 4] = "=R[-1]C/R[-1]C[6]";
            //worksheet.Cells[68, 5] = "=R[-1]C/R[-1]C[5]";
            //worksheet.Cells[68, 6] = "=R[-1]C/R[-1]C[4]";
            //worksheet.Cells[68, 7] = "=R[-1]C/R[-1]C[3]";
            //worksheet.Cells[68, 8] = "=R[-1]C/R[-1]C[2]";
            //worksheet.Cells[68, 9] = "=R[-1]C/R[-1]C[1]";
            //worksheet.Cells[68, 10] = "=SUM(RC[-8]:RC[-1])";

            //((Excel.Range)worksheet.Range[worksheet.Cells[53, 1], worksheet.Cells[68, 10]]).Borders.LineStyle = 1;
            //((Excel.Range)worksheet.Rows[68, Type.Missing]).NumberFormatLocal = "0.00%";

            //pDataRange = (Excel.Range)worksheet.Range[worksheet.Cells[54, 1], worksheet.Cells[66, 9]];

            //newChart = (Excel.Chart)((Excel.Workbook)worksheet.Parent).Charts.Add(miss, miss, miss, miss);
            //newChart.ChartWizard(pDataRange, Excel.XlChartType.xlColumnStacked, miss, miss, miss, miss, miss, miss, miss);

            //newChart.HasTitle = true;
            //newChart.ChartTitle.Text = "月度工伤类型统计图";
            //newChart.Location(Excel.XlChartLocation.xlLocationAsObject, "Sheet1");
            //xlBook.ActiveChart.PlotArea.Interior.ColorIndex = 19;
            ////newChart.ShowDataLabelsOverMaximum = true;
            ////newChart.PlotArea.Interior.Color = Color.FromArgb(192, 192, 192);
            ////newChart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowValue, true, miss, miss, miss, miss, miss, miss, miss, miss);

            //pDataRange = (Excel.Range)worksheet.Rows.get_Item(70, miss);
            //worksheet.Shapes.Item("Chart 3").Top = (float)(double)pDataRange.Top;
            //pDataRange = (Excel.Range)worksheet.Columns.get_Item(2, miss);
            //worksheet.Shapes.Item("Chart 3").Left = (float)(double)pDataRange.Left;
            //worksheet.Shapes.Item("Chart 3").Width = 500;
            //worksheet.Shapes.Item("Chart 3").Height = 300;

            ////各车间工伤件数及费用

            //worksheet.ChartObjects(miss);
            ////xlApp.Visible = true;
            ////xlApp.DisplayFullScreen = true;
            //xlBook.Saved = true;
            //savePath += fileName;

            ////保存文件时处理同名情况
            //bool pSave = false;
            //int pInt = 1;
            //while (pSave == false)
            //{
            //    if (File.Exists(savePath))
            //    {
            //        if (pInt == 1) savePath = savePath.Remove(savePath.Length - 4, 4) + "(" + pInt.ToString() + ").xls";
            //        else savePath = savePath.Remove(savePath.Length - 6, 6) + pInt.ToString() + ").xls";
            //        pInt++;
            //    }
            //    else
            //        pSave = true;
            //}
            ////文件保存
            //try
            //{
            //    xlBook.SaveAs(savePath, miss, miss, miss, miss, miss, Excel.XlSaveAsAccessMode.xlExclusive, miss, miss, miss, miss);                
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
            //finally
            //{
            //    xlBook.Close(null, null, null);
            //    xlApp.Workbooks.Close();
            //    xlApp.Application.Quit();
            //    xlApp.Quit();
            //    worksheet = null;
            //    xlBook = null;
            //    xlApp = null;

            //    GC.Collect();
            //}
            //MessageBox.Show("报表导出成功：[" + savePath + "]");
        }
        #endregion

        #region Events Area
        /// <summary>
        /// cbxYear_SelectedValueChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxYear_SelectedValueChanged(object sender, EventArgs e)
        {
            if (strSQL != "")
                SetDataSource();
        }

        /// <summary>
        /// cbxDepartment_SelectedValueChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxDepartment_SelectedValueChanged(object sender, EventArgs e)
        {
            if (strSQL != "")
                SetDataSource();
        }

        #region CreateChart
        /*
        private void CreateChart(Excel._Workbook m_Book, Excel._Worksheet m_Sheet, int num)
        {
            Excel.Range oResizeRange;
            Excel.Series oSeries;

            m_Book.Charts.Add(Missing.Value, Missing.Value, 1, Missing.Value);
            m_Book.ActiveChart.ChartType = Excel.XlChartType.xlLine;//设置图形

            //设置数据取值范围
            m_Book.ActiveChart.SetSourceData(m_Sheet.get_Range("A2", "C" + num.ToString()), Excel.XlRowCol.xlColumns);
            //m_Book.ActiveChart.Location(Excel.XlChartLocation.xlLocationAutomatic, title);
            //以下是给图表放在指定位置
            m_Book.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, m_Sheet.Name);
            oResizeRange = (Excel.Range)m_Sheet.Rows.get_Item(10, Missing.Value);
            m_Sheet.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;  //调图表的位置上边距
            oResizeRange = (Excel.Range)m_Sheet.Columns.get_Item(6, Missing.Value);  //调图表的位置左边距
            // m_Sheet.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
            m_Sheet.Shapes.Item("Chart 1").Width = 400;   //调图表的宽度
            m_Sheet.Shapes.Item("Chart 1").Height = 250;  //调图表的高度

            m_Book.ActiveChart.PlotArea.Interior.ColorIndex = 19;  //设置绘图区的背景色 
            m_Book.ActiveChart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;//设置绘图区边框线条
            m_Book.ActiveChart.PlotArea.Width = 400;   //设置绘图区宽度
            //m_Book.ActiveChart.ChartArea.Interior.ColorIndex = 10; //设置整个图表的背影颜色
            //m_Book.ActiveChart.ChartArea.Border.ColorIndex = 8;// 设置整个图表的边框颜色
            m_Book.ActiveChart.ChartArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;//设置边框线条
            m_Book.ActiveChart.HasDataTable = false;


            //设置Legend图例的位置和格式
            m_Book.ActiveChart.Legend.Top = 20.00; //具体设置图例的上边距
            m_Book.ActiveChart.Legend.Left = 60.00;//具体设置图例的左边距
            m_Book.ActiveChart.Legend.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            m_Book.ActiveChart.Legend.Width = 150;
            m_Book.ActiveChart.Legend.Font.Size = 9.5;
            //m_Book.ActiveChart.Legend.Font.Bold = true;
            m_Book.ActiveChart.Legend.Font.Name = "宋体";
            //m_Book.ActiveChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;//设置图例的位置
            m_Book.ActiveChart.Legend.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;//设置图例边框线条



            //设置X轴的显示
            Excel.Axis xAxis = (Excel.Axis)m_Book.ActiveChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            xAxis.MajorGridlines.Border.LineStyle = Excel.XlLineStyle.xlDot;
            xAxis.MajorGridlines.Border.ColorIndex = 1;//gridLine横向线条的颜色
            xAxis.HasTitle = false;
            xAxis.MinimumScale = 1500;
            xAxis.MaximumScale = 6000;
            xAxis.TickLabels.Font.Name = "宋体";
            xAxis.TickLabels.Font.Size = 9;



            //设置Y轴的显示
            Excel.Axis yAxis = (Excel.Axis)m_Book.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            yAxis.TickLabelSpacing = 30;
            yAxis.TickLabels.NumberFormat = "M月D日";
            yAxis.TickLabels.Orientation = Excel.XlTickLabelOrientation.xlTickLabelOrientationHorizontal;//Y轴显示的方向,是水平还是垂直等
            yAxis.TickLabels.Font.Size = 8;
            yAxis.TickLabels.Font.Name = "宋体";

            //m_Book.ActiveChart.Floor.Interior.ColorIndex = 8;  
            /***以下是设置标题*****
            //m_Book.ActiveChart.HasTitle=true;
            //m_Book.ActiveChart.ChartTitle.Text = "净值指数";
            //m_Book.ActiveChart.ChartTitle.Shadow = true;
            //m_Book.ActiveChart.ChartTitle.Border.LineStyle = Excel.XlLineStyle.xlContinuous;
            

            oSeries = (Excel.Series)m_Book.ActiveChart.SeriesCollection(1);
            oSeries.Border.ColorIndex = 45;
            oSeries.Border.Weight = Excel.XlBorderWeight.xlThick;
            oSeries = (Excel.Series)m_Book.ActiveChart.SeriesCollection(2);
            oSeries.Border.ColorIndex = 9;
            oSeries.Border.Weight = Excel.XlBorderWeight.xlThick;

        }*/
        #endregion

        #endregion
    }
}
