using Common;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace MyDBQuery
{
    public partial class FrmQuery : Form
    {
        private DataTable mLastQuery = null;
        private TConnectType mConnectType=null;
        public FrmQuery()
        {
            InitializeComponent();
            //mOracleConnStr = ConfigurationManager.ConnectionStrings["OracleConnStr"].ConnectionString;
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
            if (string.IsNullOrEmpty(sSql) || null== mConnectType) {
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            mLastQuery = new DataTable();
            if (DbType.SQLServer == mConnectType.Type)
            {
                try
                {
                    var dt = SqlServerHelper.ExecuteQuery(mConnectType.ConnStr, sSql);
                    mLastQuery = dt;
                }
                catch (Exception ex)
                {
                    this.Cursor = Cursors.Default;
                    MessageBox.Show(ex.Message, "连接SqlServer失败");
                    return;
                }
            }
            else if (DbType.Oracle == mConnectType.Type)
            {
                try
                {
                    using (var conn = new OracleConnection(mConnectType.ConnStr))
                    {
                        using (var dr = OraClientHelper.ExecuteReader(conn, CommandType.Text, sSql, null))
                        {
                            mLastQuery.Load(dr);
                        }
                    }
                }
                catch (Exception ex)
                {
                    this.Cursor = Cursors.Default;
                    MessageBox.Show(ex.Message, "连接Oracle失败");
                    return;
                }
            }            
            
            gridData.DataSource = mLastQuery;
            setRowNumber(gridData);
            this.Cursor = Cursors.Default;
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

        private void btnConnectType_Click(object sender, EventArgs e)
        {
            FrmConnectType frmConn = new FrmConnectType();
            if(DialogResult.OK == frmConn.ShowDialog())
            {
                mConnectType = frmConn.ConnectType;
                btnQuery.Enabled = true;
                btnExport.Enabled = true;

                var sTitle = string.Format("{0} [{1}]", "Query Test", mConnectType.Type.ToString());
                this.Text = sTitle;
            }
        }
    }
}
