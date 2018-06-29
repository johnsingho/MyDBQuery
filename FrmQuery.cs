using Common;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace MyDBQuery
{
    public partial class FrmQuery : Form
    {
        private string mOracleConnStr = string.Empty;
        private DataTable mLastQuery = null;
        public FrmQuery()
        {
            InitializeComponent();
            mOracleConnStr = ConfigurationManager.ConnectionStrings["OracleConnStr"].ConnectionString;
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            ClearGrid();
            DoQuery(textSql.Text.Trim());
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            ClearGrid();
            textSql.Clear();
            textSql.Focus();
        }
        private void ClearGrid()
        {
            gridData.DataSource = null;
        }
        private void DoQuery(string sSql)
        {
            if (string.IsNullOrEmpty(sSql)) { return; }
            mLastQuery = new DataTable();
            try
            {
                using (var conn = new OracleConnection(mOracleConnStr))
                {
                    using (var dr = OraClientHelper.ExecuteReader(conn, CommandType.Text, sSql, null))
                    {
                        mLastQuery.Load(dr);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "连接Oracle失败");
                return;
            }
            
            gridData.DataSource = mLastQuery;
            setRowNumber(gridData);
        }

        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
            dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if(null== mLastQuery)
            {
                ClearGrid();
                DoQuery(textSql.Text.Trim());
            }
            if (null != mLastQuery)
            {
                var bys = EPPExcelHelper.BuilderExcel(mLastQuery);
                if(null!=bys && bys.Length > 0)
                {
                    SaveFileDialog saveFileDlg = new SaveFileDialog();

                    saveFileDlg.Filter = "Excel files (*.xlsx)|*.xlsx";
                    saveFileDlg.FilterIndex = 1;
                    saveFileDlg.RestoreDirectory = true;

                    if (saveFileDlg.ShowDialog() == DialogResult.OK)
                    {
                        Stream myStream;
                        if ((myStream = saveFileDlg.OpenFile()) != null)
                        {
                            myStream.Write(bys, 0, bys.Length);
                            myStream.Close();
                            MessageBox.Show("导出完成！");
                        }
                    }
                }
            }
        }
    }
}
