using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Collections.Generic;
using System.Globalization;
using NPOI.SS.Formula.Eval;
using System.Text;

namespace SqlToExcel
{
    /// <summary>
    /// 
    /// </summary>
    public partial class SqlToExcel : Form
    {
        /// <summary>
        /// 
        /// </summary>
        public SqlToExcel()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SqlToExcel_Load(object sender, EventArgs e)
        {

            TreeNode rootNode = new TreeNode();
            rootNode.Name = "root";
            rootNode.Text = "全选";
            TreeNode treeNodeArea;
            string strConnection = ConfigurationManager.ConnectionStrings["Portal"] != null
                ? ConfigurationManager.ConnectionStrings["Portal"].ConnectionString
                : "";
            SqlConnection objConnection = new SqlConnection(strConnection);
            try
            {
                objConnection.Open();

                string sqlString = "SELECT * FROM ptlCompany ORDER BY ProjectName";
                SqlCommand cmd = new SqlCommand(sqlString, objConnection);
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = sqlString;

                SqlDataReader sqlDataReader = cmd.ExecuteReader();
                DataTable dataTable = new DataTable();
                dataTable.Load(sqlDataReader);
                sqlDataReader.Dispose();
                cmd.Dispose();

                sqlString = "SELECT * FROM ptlOrganization WHERE OType = 1 AND Dr = 0";
                cmd = new SqlCommand(sqlString, objConnection);
                cmd.CommandType = CommandType.Text;
                sqlDataReader = cmd.ExecuteReader();
                cmd.Dispose();
                DataTable dataTableOrg = new DataTable();
                dataTableOrg.Load(sqlDataReader);
                sqlDataReader.Dispose();

                string defaultProjectId = ConfigurationManager.AppSettings["DefaultProjectId"];
                foreach (DataRow dataRow in dataTableOrg.Rows)
                {
                    treeNodeArea = new TreeNode();
                    treeNodeArea.Text = dataRow["OrgName"].ToString();
                    treeNodeArea.Name = dataRow["MId"].ToString();

                    sqlString = @"SELECT B.MId,B.ProjectName FROM ptlOrganization A
                            JOIN ptlCompany B ON A.MId = B.OrgId
                            WHERE B.Dr = 0 AND A.Dr = 0 AND A.FullId LIKE @OrgId";
                    cmd = new SqlCommand(sqlString, objConnection);
                    cmd.Parameters.Add(new SqlParameter("@OrgId", string.Format("%-{0}-%", dataRow["MId"])));
                    cmd.CommandType = CommandType.Text;
                    sqlDataReader = cmd.ExecuteReader();
                    cmd.Dispose();
                    DataTable dataTableProj = new DataTable();
                    dataTableProj.Load(sqlDataReader);
                    sqlDataReader.Dispose();
                    foreach (DataRow dataRowProj in dataTableProj.Rows)
                    {
                        TreeNode treeNodeProj = new TreeNode();
                        treeNodeProj.Name = dataRowProj["MId"].ToString();
                        treeNodeProj.Text = dataRowProj["ProjectName"].ToString();

                        if (("," + defaultProjectId + ",").IndexOf("," + dataRowProj["MId"] + ",", StringComparison.Ordinal) >= 0)
                        {
                            treeNodeProj.Checked = true;
                        }

                        treeNodeArea.Nodes.Add(treeNodeProj);
                    }
                    if (dataTableProj.Rows.Count > 0)
                    {
                        rootNode.Nodes.Add(treeNodeArea);
                        treeNodeArea.ExpandAll();
                    }
                }

                sqlDataReader.Close();
                objConnection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                objConnection.Dispose();
            }
            treeNodeArea = new TreeNode {Name = "Other", Text = "其他系统"};

            foreach (ConnectionStringSettings connString in ConfigurationManager.ConnectionStrings)
            {
                TreeNode treeNodeProj = new TreeNode();
                treeNodeProj.Name = connString.Name;
                treeNodeProj.Text = connString.Name;
                treeNodeArea.Nodes.Add(treeNodeProj);
            }
            rootNode.Nodes.Add(treeNodeArea);

            trvConnection.Nodes.Add(rootNode);
            rootNode.ExpandAll();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExec_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                txtFilePath.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            if (string.IsNullOrWhiteSpace(txtSql.Text))
            {
                MessageBox.Show("请输入SQL.");
                return;
            }

            string fileName = string.Empty;
            string pathName = string.Empty;
            if (ckbExcel.Checked)
            {
                pathName = txtFilePath.Text;
            }
            else
            {
                fileName = txtFilePath.Text;
            }
            if (!string.IsNullOrEmpty(pathName) && !Directory.Exists(pathName))
            {
                Directory.CreateDirectory(pathName);
            }
            FileStream fileStream = null;
            XSSFWorkbook xssfWorkbook = null;

            if (!ckbExcel.Checked)
            {
                if ((fileName.Length > 6 &&
                    !fileName.Substring(fileName.Length - 5).Equals(".xlsx") &&
                    !fileName.Substring(fileName.Length - 4).Equals(".xls"))
                    || fileName.Length <= 6)
                {
                    if (!Directory.Exists(fileName))
                    {
                        Directory.CreateDirectory(fileName);
                    }
                    string excelName = DateTime.Now.ToString("yyyyMMddHHmmsss");
                    if (!string.IsNullOrEmpty(txtExcelName.Text))
                    {
                        excelName = txtExcelName.Text;
                    }

                    if (fileName[fileName.Length - 1].Equals("\\"))
                    {
                        fileName = fileName + excelName + ".xlsx";
                    }
                    else
                    {
                        fileName = fileName + "\\" + excelName + ".xlsx";
                    }
                }
                fileStream = new FileStream(fileName, FileMode.OpenOrCreate);
                xssfWorkbook = new XSSFWorkbook();
            }


            string strConnection = ConfigurationManager.ConnectionStrings["Portal"] == null
                ? ""
                : ConfigurationManager.ConnectionStrings["Portal"].ConnectionString;
                SqlConnection objConnection = new SqlConnection(strConnection);

                Dictionary<string, string> dicProjectConn = new Dictionary<string, string>();
            try
            {


                using (objConnection)
                {
                    objConnection.Open();
                    string sqlString = "SELECT * FROM ptlCompany ORDER BY ProjectName";
                    SqlCommand cmd = new SqlCommand(sqlString, objConnection);
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = sqlString;
                    SqlDataReader sqlDataReader = cmd.ExecuteReader();
                    cmd.Dispose();
                    DataTable dataTable = new DataTable();
                    dataTable.Load(sqlDataReader);

                    foreach (DataRow dataRowProj in dataTable.Rows)
                    {
                        dicProjectConn.Add(Convert.ToString(dataRowProj["MId"]),
                            dataRowProj["ConnectionString"].ToString());
                    }

                }
            } catch(Exception) {
                // ignored
            }
            foreach (ConnectionStringSettings connString in ConfigurationManager.ConnectionStrings)
            {
                dicProjectConn.Add(connString.Name, connString.ConnectionString);
            }

            foreach (TreeNode treeNodeArea in trvConnection.Nodes[0].Nodes)
            {
                foreach (TreeNode treeNodeProj in treeNodeArea.Nodes)
                {
                    if (!treeNodeProj.Checked)
                    {
                        continue;
                    }

                    string connectionString = dicProjectConn[treeNodeProj.Name];
                    if (string.IsNullOrEmpty(connectionString))
                    {
                        continue;
                    }

                    objConnection = new SqlConnection(connectionString);
                    try
                    {
                        using (objConnection)
                        {
                            objConnection.Open();
                            objConnection.Close();
                        }
                    }
                    catch
                    {
                        continue;
                    }

                    using (objConnection)
                    {
                        objConnection.ConnectionString = connectionString;
                        objConnection.Open();
                        string sql = txtSql.Text;
                        SqlCommand sqlCommand = new SqlCommand(sql, objConnection);
                        sqlCommand.CommandType = CommandType.Text;
                        sqlCommand.CommandText = sql;
                        SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                        DataTable dataTable = new DataTable();
                        dataTable.Load(sqlDataReader);
                        sqlDataReader.Dispose();
                        sqlDataReader.Close();
                        sqlCommand.Dispose();
                        objConnection.Close();

                        if (dataTable.Rows.Count == 0)
                        {
                            continue;
                        }

                        if (ckbExcel.Checked)
                        {
                            fileStream = new FileStream(
                                pathName + "\\" + treeNodeProj.Text + ".xlsx", FileMode.OpenOrCreate);

                            xssfWorkbook = new XSSFWorkbook();
                        }

                        XSSFSheet xssfSheet;
                        if (ckbSheet.Checked)
                        {
                            xssfSheet = (XSSFSheet) xssfWorkbook.GetSheet(treeNodeProj.Text);
                            if (xssfSheet == null)
                            {
                                xssfSheet = (XSSFSheet) xssfWorkbook.CreateSheet(treeNodeProj.Text);
                            }
                        }
                        else
                        {
                            if (xssfWorkbook.NumberOfSheets > 0)
                            {
                                xssfSheet = (XSSFSheet) xssfWorkbook.GetSheetAt(0);
                            }
                            else
                            {
                                xssfSheet = (XSSFSheet) xssfWorkbook.CreateSheet();
                            }
                        }
                        //写入数据
                        FillDataFromTable(xssfWorkbook, xssfSheet, "A" + xssfSheet.LastRowNum + 1, dataTable);

                        if (ckbExcel.Checked)
                        {
                            xssfWorkbook.Write(fileStream);
                            if (fileStream.CanWrite)
                            {
                                fileStream.Flush();
                            }
                            fileStream.Close();
                        }
                    }

                }
            }
            if (!ckbExcel.Checked)
            {
                xssfWorkbook.Write(fileStream);
                if (fileStream.CanWrite)
                {
                    fileStream.Flush();
                }
                fileStream.Close();
            }
            DialogResult dialogResult = MessageBox.Show("执行成功.");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel文件(*.xlsx,*.xls)|*.xlsx;*.xls";
            openFileDialog1.Multiselect = false;
            openFileDialog1.AddExtension = false;
            openFileDialog1.ValidateNames = false;
            openFileDialog1.CheckFileExists = false;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.FileName = "D:\\";

            DialogResult dialogResult;
            if (!ckbExcel.Checked)
            {
                dialogResult = openFileDialog1.ShowDialog();
            }
            else
            {
                dialogResult = folderBrowserDialog1.ShowDialog();
            }
            if (dialogResult == DialogResult.OK)
            {
                if (!ckbExcel.Checked)
                {
                    txtFilePath.Text = openFileDialog1.FileName;
                }
                else
                {
                    txtFilePath.Text = folderBrowserDialog1.SelectedPath;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startAddress"></param>
        /// <param name="dataTable"></param>
        /// <param name="iWorkbook"></param>
        public void FillDataFromTable(XSSFWorkbook iWorkbook, XSSFSheet sheet, string startAddress, DataTable dataTable)
        {
            XSSFCellStyle headerCellStyle = (XSSFCellStyle)iWorkbook.CreateCellStyle();
            CT_Color ctColor = new CT_Color();
            ctColor.SetRgb(112, 173, 71);
            XSSFColor xssfColor = new XSSFColor(ctColor);
            headerCellStyle.SetFillBackgroundColor(xssfColor);
            headerCellStyle.SetFillForegroundColor(xssfColor);

            XSSFFont hssfFont = iWorkbook.CreateFont() as XSSFFont;
            hssfFont.FontHeightInPoints = 10;
            hssfFont.FontName = "宋体";
            hssfFont.Boldweight = 700;
            headerCellStyle.SetFont(hssfFont);

            XSSFCellStyle contentCellStyle = (XSSFCellStyle)iWorkbook.CreateCellStyle();
            XSSFFont contentHssfFont = iWorkbook.CreateFont() as XSSFFont;
            contentHssfFont.FontHeightInPoints = 10;
            contentHssfFont.FontName = "宋体";
            contentCellStyle.SetFont(contentHssfFont);

            string rowIndexStr = string.Empty;
            string cellIndexStr = string.Empty;
            int cellIndex = 0;
            for (int i = 0; i < startAddress.Length; i++)
            {
                int tempNum;
                if (int.TryParse(startAddress[i].ToString(), out tempNum))
                {
                    rowIndexStr += "" + tempNum;
                }
                else
                {
                    cellIndexStr += "" + startAddress[i];
                }
            }
            var rowIndex = Convert.ToInt32(rowIndexStr);

            for (int i = cellIndexStr.Length - 1; i >= 0; i--)
            {
                if (i == cellIndexStr.Length - 1)
                {
                    cellIndex += cellIndexStr[i] - 65;
                }
                else
                {
                    cellIndex += (cellIndexStr[i] - 64) * 26;
                }
            }

            cellIndex = 0;

            //textBox1.Text += "\r\n 共有数据:" + _DataTable.Rows.Count;

            int tempCellIndex = cellIndex;
            try
            {
                //sheet分开包含表头
                if (!ckbSheet.Checked)
                {
                    rowIndex = sheet.LastRowNum;
                    if (rowIndex != 0)
                    {
                        rowIndex++;
                    }
                }

                //是否包含表头
                XSSFRow excelRow;
                XSSFCell excelCell;
                if (ckbIsIncludeHeader.Checked && rowIndex <= 1)
                {
                    excelRow = sheet.GetRow(rowIndex) as XSSFRow ?? sheet.CreateRow(rowIndex) as XSSFRow;
                    excelRow.HeightInPoints = 20;
                    foreach (DataColumn dataColumn in dataTable.Columns)
                    {
                        excelCell = excelRow.GetCell(tempCellIndex) as XSSFCell;
                        if (excelCell == null)
                        {
                            excelCell = excelRow.CreateCell(tempCellIndex) as XSSFCell;
                        }
                        if (string.IsNullOrEmpty(dataColumn.ColumnName))
                        {
                            excelCell.SetCellType(CellType.Blank);
                            excelCell.CellStyle = headerCellStyle;
                            tempCellIndex++;
                        }
                        else
                        {
                            excelCell.SetCellType(CellType.String);
                            excelCell.SetCellValue(Convert.ToString(dataColumn.ColumnName));
                            excelCell.CellStyle = headerCellStyle;
                            tempCellIndex++;
                        }
                    }
                    rowIndex++;
                    tempCellIndex = cellIndex;
                }


                //填充数据
                foreach (DataRow dataRow in dataTable.Rows)
                {
                    excelRow = sheet.GetRow(rowIndex) as XSSFRow ?? sheet.CreateRow(rowIndex) as XSSFRow;
                    excelRow.HeightInPoints = 20;
                    foreach (DataColumn dataColumn in dataTable.Columns)
                    {
                        excelCell = excelRow.GetCell(tempCellIndex) as XSSFCell;
                        if (excelCell == null)
                        {
                            excelCell = excelRow.CreateCell(tempCellIndex) as XSSFCell;
                        }

                        if (dataRow[dataColumn] == DBNull.Value || string.IsNullOrEmpty(Convert.ToString(dataRow[dataColumn])))
                        {
                            excelCell.SetCellType(CellType.Blank);
                            excelCell.CellStyle = contentCellStyle;
                            tempCellIndex++;
                            continue;
                        }
                        if (dataRow[dataColumn] is decimal || dataRow[dataColumn] is int)
                        {
                            excelCell.SetCellType(CellType.Numeric);
                            excelCell.SetCellValue(Convert.ToDouble(dataRow[dataColumn]));
                            excelCell.CellStyle = contentCellStyle;
                        }
                        else
                        {
                            excelCell.SetCellType(CellType.String);
                            excelCell.SetCellValue(Convert.ToString(dataRow[dataColumn]));
                            excelCell.CellStyle = contentCellStyle;
                        }

                        tempCellIndex++;
                    }
                    tempCellIndex = cellIndex;
                    rowIndex++;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("行" + rowIndex + "列" + tempCellIndex + "_" + ex.Message, ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ckbExcel_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbExcel.Checked)
            {
                lblExcelName.Visible = false;
                txtExcelName.Visible = false;
            }
            else
            {
                lblExcelName.Visible = true;
                txtExcelName.Visible = true;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvConnection_AfterCheck(object sender, TreeViewEventArgs e)
        {
            foreach (TreeNode treeNode in e.Node.Nodes)
            {
                treeNode.Checked = e.Node.Checked;
            }
        }
        
        /// <summary>
        /// 将制定sheet中的数据导出到datatable中
        /// </summary>
        /// <param name="sheet">需要导出的sheet</param>
        /// <param name="headerRowIndex">列头所在行号，-1表示没有列头</param>
        /// <param name="needHeader"></param>
        /// <returns></returns>
        static DataTable ImportDt(ISheet sheet, int headerRowIndex, bool needHeader) {
            DataTable table = new DataTable();
            IRow headerRow;
            int cellCount;
            if(headerRowIndex < 0 || !needHeader) {
                headerRow = sheet.GetRow(0);
                cellCount = headerRow.LastCellNum;

                for(int i = headerRow.FirstCellNum; i <= cellCount; i++) {
                    DataColumn column = new DataColumn(Convert.ToString(i));
                    table.Columns.Add(column);
                }
            } else {
                headerRow = sheet.GetRow(headerRowIndex);
                cellCount = headerRow.LastCellNum;

                for(int i = headerRow.FirstCellNum; i <= cellCount; i++) {
                    if(headerRow.GetCell(i) == null) {
                        if(table.Columns.IndexOf(Convert.ToString(i)) > 0) {
                            DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        } else {
                            DataColumn column = new DataColumn(Convert.ToString(i));
                            table.Columns.Add(column);
                        }

                    } else if(table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0) {
                        DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                        table.Columns.Add(column);
                    } else {
                        DataColumn column = new DataColumn(headerRow.GetCell(i).ToString());
                        table.Columns.Add(column);
                    }
                }
            }
            for(int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++) {
                IRow row;
                if(sheet.GetRow(i) == null) {
                    row = sheet.CreateRow(i);
                } else {
                    row = sheet.GetRow(i);
                }

                DataRow dataRow = table.NewRow();

                for(int j = row.FirstCellNum; j <= cellCount; j++) {
                    if(row.GetCell(j) != null) {
                        switch(row.GetCell(j).CellType) {
                            case CellType.String:
                                string str = row.GetCell(j).StringCellValue;
                                if(str != null && str.Length > 0) {
                                    dataRow[j] = str;
                                } else {
                                    dataRow[j] = null;
                                }
                                break;
                            case CellType.Numeric:
                                if(DateUtil.IsCellDateFormatted(row.GetCell(j))) {
                                    dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                } else {
                                    dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                }
                                break;
                            case CellType.Boolean:
                                dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                break;
                            case CellType.Error:
                                dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                break;
                            case CellType.Formula:
                                switch(row.GetCell(j).CachedFormulaResultType) {
                                    case CellType.String:
                                        string strFormula = row.GetCell(j).StringCellValue;
                                        if(!string.IsNullOrEmpty(strFormula)) {
                                            dataRow[j] = strFormula;
                                        } else {
                                            dataRow[j] = null;
                                        }
                                        break;
                                    case CellType.Numeric:
                                        dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue,
                                            CultureInfo.InvariantCulture);
                                        break;
                                    case CellType.Boolean:
                                        dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                        break;
                                    case CellType.Error:
                                        dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                        break;
                                    default:
                                        dataRow[j] = "";
                                        break;
                                }
                                break;
                            default:
                                dataRow[j] = "";
                                break;
                        }
                    }

                }
                table.Rows.Add(dataRow);
            }
            return table;
        }
    }
}