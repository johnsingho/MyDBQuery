using Common;
using MyDBQuery.common;
using MySql.Data.MySqlClient;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;
using System.Data.SQLite;
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
            mLastQuery = null;
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
                case DbType.Sqlite:
                    QuerySqlite(sSql);
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
                var ds = SqlServerHelper.ExecuteQuery(mConnectType.ConnStr, sSql);
                if ( !DataTableHelper.IsEmptyDataSet(ds))
                {
                    mLastQuery = ds.Tables[0];
                }
                else
                {
                    mLastQuery = null;
                }
                
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


        private bool QuerySqlite(string sSql)
        {
            var bRet = false;
            try
            {
                using (var conn = new SQLiteConnection(mConnectType.ConnStr))
                {
                    using (var dr = SqliteHelper.ExecuteReader(conn, CommandType.Text, sSql, null))
                    {
                        mLastQuery.Load(dr);
                    }
                }
                bRet = true;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message, "连接Sqlite失败");
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

        private string GetSaveFileName()
        {
            SaveFileDialog saveFileDlg = new SaveFileDialog();

            saveFileDlg.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDlg.FilterIndex = 1;
            saveFileDlg.RestoreDirectory = true;

            if (saveFileDlg.ShowDialog() == DialogResult.OK)
            {
                return saveFileDlg.FileName;
            }
            return string.Empty;
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            //if(null== mLastQuery)
            //{
            //    ClearGrid();
            //    DoQuery(textSql.Text.Trim());
            //}

            var sfn = GetSaveFileName();
            if (string.IsNullOrEmpty(sfn)) { return; }
            if (null != mLastQuery)
            {                
                var bys = EPPExcelHelper.BuilderExcel(mLastQuery);                
                using (var fs=new FileStream(sfn, FileMode.Create))
                {
                    fs.Write(bys, 0, bys.Length);
                }                
            }
            else
            {
                this.Cursor = Cursors.WaitCursor;
                BuildExpByDataReader(sfn, textSql.Text.Trim());
                this.Cursor = Cursors.Default;
            }
        }

        private bool BuildExpByDataReader(string sfn, string sql)
        {
            var bRet = false;
            switch (mConnectType.Type)
            {
                case DbType.SQLServer:
                    {
                        using (var dr = SqlServerHelper.ExecuteReader(mConnectType.ConnStr, sql, null))
                        {
                            bRet = EPPExcelHelper.BuilderExcel(sfn, dr);
                        }
                    }
                    break;
                case DbType.Oracle:
                    {
                        using (var conn = new OracleConnection(mConnectType.ConnStr))
                        {
                            using (var dr = OraClientHelper.ExecuteReader(conn, CommandType.Text, sql, null))
                            {
                                bRet = EPPExcelHelper.BuilderExcel(sfn, dr);
                            }
                        }
                    }
                    break;
                case DbType.MySql:
                    {
                        using (var conn = new MySqlConnection(mConnectType.ConnStr))
                        {
                            using (var dr = MySqlClientHelper.ExecuteReader(conn, CommandType.Text, sql, null))
                            {
                                bRet = EPPExcelHelper.BuilderExcel(sfn, dr);
                            }
                        }
                    }
                    break;
                default: break;
            }
            return bRet;
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
            if (mConnectType == null)
            {
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
/*TODO
                case DbType.Sqlite:
                    ImportSqlite(sTarTable, mLastQuery);
                    break;
*/
                default:
                    MessageBox.Show("暂时只实现了SQLServer、Sqlite的");
                    break;
            }
            this.Cursor = Cursors.Default;
        }


        #region 导入
        private bool ImpSQLServerHandler(DataTable dt, ref long nHandles, out string sErr)
        {
            var nItems = SqlServerHelper.BulkToDB(mConnectType.ConnStr, dt, dt.TableName, out sErr);
            nHandles = nHandles + nItems;
            return string.IsNullOrEmpty(sErr);
        }

        private void ImportSQLServer(string sTarTable, string sXlsFile)
        {
            try
            {
                var sErr = string.Empty;
                var nItems = EPPExcelHelper.ReadExcel(new FileInfo(sXlsFile), sTarTable, ImpSQLServerHandler, out sErr);                
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

        private void ImportSqlite(string sTarTable, DataTable dt)
        {
            try
            {
                var sErr = string.Empty;
                var nItems = SqliteHelper.BulkToDB(mConnectType.ConnStr, dt, sTarTable, out sErr);
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Import error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        #endregion
    }
}
