using MyDBQuery.common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyDBQuery
{
    public partial class FrmConnectType : Form
    {
        public TConnectType ConnectType = new TConnectType();
        //private BindingSource bsLog = new BindingSource();

        public FrmConnectType()
        {
            InitializeComponent();
            cmbType.DataSource = Enum.GetValues(typeof(DbType));
            //cmbType.SelectedIndex = -1;
            //gvHis.DataSource = bsLog;
        }

        private void btnSetConn_Click(object sender, EventArgs e)
        {
            var sDbType = cmbType.SelectedItem.ToString();
            var sConn = txtConnStr.Text.Trim();
            if (string.IsNullOrEmpty(sConn))
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
                return;
            }
            List<KeyValuePair> logs = null;
            try
            {
                logs = LoadLogs(sDbType);
            }
            catch (Exception)
            {
            }
            bool bExist = false;
            if (null!=logs)
            {
                foreach (var log in logs)
                {
                    if (Md5Helper.VerifyMd5(sConn, log.Key))
                    {
                        bExist = true;
                        break;
                    }
                }
            }

            if (!bExist)
            {
                var oIni = new IniParser(AppCommon.IniFile);
                var sMd5 = Md5Helper.CalcMd5(sConn);
                oIni.AddSetting(sDbType, sMd5, sConn);
                oIni.SaveSettings();
            }
            ConnectType.ConnStr = sConn;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private List<KeyValuePair> LoadLogs(string sDbType)
        {
            var oIni = new IniParser(AppCommon.IniFile);
            return oIni.EnumSectionAndValues(sDbType);
        }
        private void OnSelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox senderComboBox = (ComboBox)sender;
            gvHis.DataSource = null; //clear all

            var sel = senderComboBox.SelectedItem;
            if (null!=sel)
            {
                var oDbType = (DbType)sel;
                ConnectType.Type = oDbType;
                var sDbType = oDbType.ToString();
                try
                {
                    var logs = LoadLogs(sDbType);
                    if (null!=logs || logs.Count>0)
                    {
                        var dt = logs.ToDataTable<KeyValuePair>();
                        //bsLog.DataSource = dt;
                        dt.Columns.Remove("Key");
                        gvHis.DataSource = dt;
                        gvHis.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "配置文件有问题");
                }

            }
        }

        private void OnCellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            var ind = e.RowIndex;
            if (ind >= 0)
            {
                var gv = (DataGridView)sender;
                var r = gv.Rows[ind];
                var txt = r.Cells[0].Value as string;
                txtConnStr.Text = txt;
            }
        }
    }
}
