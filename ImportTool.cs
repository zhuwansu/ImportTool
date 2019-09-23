using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelCompareTool;
using OfficeOpenXml;

namespace ImportTool
{
    public partial class ImportTool : Form
    {
        public Data Data { get; set; }

        public ImportTool()
        {
            InitializeComponent();
            Data = Data.Get();
            textBox1.Text = Data?.ConnectionString ?? "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fileName = "";
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                var res = dialog.ShowDialog();
                if (res != DialogResult.OK) return;

                fileName = dialog.FileName;
                DataTable dt = GetDataTable(dialog.OpenFile());
                var sbSql = new StringBuilder();
                sbSql.Append($"CREATE TABLE {dt.TableName}(");
                for (int i = 0 ; i < dt.Columns.Count - 1 ; i++)
                {
                    sbSql.Append("\r" + dt.Columns[i].ColumnName.Trim() + " VARCHAR(100) NULL, ");
                }
                sbSql.Append("\r" + dt.Columns[dt.Columns.Count - 1].ColumnName.Trim() + " VARCHAR(100) NULL) ");

                var sqlConnection = new SqlConnection(textBox1.Text);
                sqlConnection.Open();
                var cmd = sqlConnection.CreateCommand();
                cmd.CommandText = sbSql.ToString();
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnection))
                {
                    bulkCopy.BatchSize = 10000; //一次更新的行数
                    bulkCopy.BulkCopyTimeout = 60; //操作超时的时间
                    bulkCopy.DestinationTableName = dt.TableName; //目标表
                    for (int i = 0 ; i < dt.Columns.Count ; i++)
                        bulkCopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName); //添加源和目标表之间的列映射。
                    bulkCopy.DestinationTableName = dt.TableName;
                    bulkCopy.WriteToServer(dt);
                }
                if (sqlConnection.State != ConnectionState.Closed)
                {
                    sqlConnection.Close();
                }
                Data = Data ?? new Data();
                Data.ConnectionString = sqlConnection.ConnectionString;
                Data.Serializer();
                richTextBox1.AppendText(Environment.NewLine);
                richTextBox1.AppendText($"【成功】 {DateTime.Now.ToShortTimeString() } 【表名字】： {dt.TableName}");
                MessageBox.Show("成功");
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText(Environment.NewLine);
                richTextBox1.AppendText($"【失败】 {DateTime.Now.ToShortTimeString() } 【文件】： {fileName}");
                richTextBox1.AppendText(Environment.NewLine);
                richTextBox1.AppendText(ex.Message);

                MessageBox.Show(ex.Message);
            }
        }

        private static DataTable GetDataTable(Stream stream)
        {
            DataTable tbl = new DataTable();
            using (stream)
            using (ExcelPackage package = new ExcelPackage(stream))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                tbl.TableName = ws.Name + DateTime.Now.ToString("_yyyyMMddhhmmss");
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(firstRowCell.Text);
                }
                var startRow = 2;
                for (int rowNum = startRow ; rowNum <= ws.Dimension.End.Row ; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
            }

            return tbl;
        }

    }
}
