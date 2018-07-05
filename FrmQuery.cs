using Common;
using MyDBQuery.common;
using MySql.Data.MySqlClient;
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
            btnQuery.Enabled = false;
            btnExport.Enabled = false;
            btnImport.Enabled = false;

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

            switch (mConnectType.Type)
            {
                case DbType.SQLServer:
                    QuerySqlServer(sSql);
                    break;
                case DbType.Oracle:
                    QueryOracle(sSql);
                    break;
                case DbType.MySql:
                    QueryMySql(sSql);
                    break;
                default:break;
            } 
            
            gridData.DataSource = mLastQuery;
            setRowNumber(gridData);
            this.Cursor = Cursors.Default;
        }

        #region 相关查询方法
        private bool QuerySqlServer(string sSql)
        {
            var bRet = false;
            try
            {
                var dt = SqlServerHelper.ExecuteQuery(mConnectType.ConnStr, sSql);
                mLastQuery = dt;
                bRet = true;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message, "连接SqlServer失败");                
            }
            return bRet;
        }

        private bool QueryOracle(string sSql)
        {
            var bRet = false;
            try
            {
                using (var conn = new OracleConnection(mConnectType.ConnStr))
                {
                    using (var dr = OraClientHelper.ExecuteReader(conn, CommandType.Text, sSql, null))
                    {
                        mLastQuery.Load(dr);
                    }
                }
                bRet = true;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message, "连接Oracle失败");
            }
            return bRet;
        }

        private bool QueryMySql(string sSql)
        {
            var bRet = false;
            try
            {
                using (var conn = new MySqlConnection(mConnectType.ConnStr))
                {
                    using (var dr = MySqlClientHelper.ExecuteReader(conn, CommandType.Text, sSql, null))
                    {
                        mLastQuery.Load(dr);
                    }
                }
                bRet = true;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message, "连接Mysql失败");
            }
            return bRet;
        }
        #endregion

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
            }//if

        }

        private void btnConnectType_Click(object sender, EventArgs e)
        {
            FrmConnectType frmConn = new FrmConnectType();
            if(DialogResult.OK == frmConn.ShowDialog())
            {
                mConnectType = frmConn.ConnectType;
                btnQuery.Enabled = true;
                btnExport.Enabled = true;
                btnImport.Enabled = true;

                var sTitle = string.Format("{0} [{1}]", "MyDBQuery", mConnectType.Type.ToString());
                this.Text = sTitle;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (mConnectType == null || mConnectType.Type != DbType.SQLServer)
            {
                MessageBox.Show("暂时只实现了SQLServer的");
                return;
            }

            FrmImport frmConn = new FrmImport();
            if (DialogResult.OK != frmConn.ShowDialog())
            {
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            var sTarTable = frmConn.mTarTable;
            var sXlsFile = frmConn.mXlsFile;
            switch (mConnectType.Type)
            {
                case DbType.SQLServer:
                    ImportSQLServer(sTarTable, sXlsFile);
                    break;
                default:break;
            }
            this.Cursor = Cursors.Default;
        }


        #region 导入
        private void ImportSQLServer(string sTarTable, string sXlsFile)
        {
            try
            {
                var dt = EPPExcelHelper.ReadExcel(new FileInfo(sXlsFile));
                var sErr = string.Empty;
                var nItems = SqlServerHelper.BulkToDB(mConnectType.ConnStr, dt, sTarTable, out sErr);
                if (!string.IsNullOrEmpty(sErr))
                {
                    MessageBox.Show(sErr);
                }
                else
                {
                    var str = string.Format("导入完成， {0}条记录", nItems);
                    MessageBox.Show(str);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Import error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
    }
}
