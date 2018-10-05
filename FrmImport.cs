using MyDBQuery.common;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MyDBQuery
{
    public partial class FrmImport : Form
    {
        public string mTarTable { get; set; }
        public string mXlsFile { get; set; }
        private List<KeyValuePair> mLstImpHIs {get;set;}

        public FrmImport()
        {
            InitializeComponent();
            InitTarTableAutoComplete();
            //cmbTarTab.Text = "TB_tmpIDCard";
        }

        private void InitTarTableAutoComplete()
        {
            cmbTarTab.AutoCompleteMode = AutoCompleteMode.Suggest;
            cmbTarTab.AutoCompleteSource = AutoCompleteSource.CustomSource;

            var namesCollection = new AutoCompleteStringCollection();
            var oIni = new IniParser(AppCommon.IniFile);
            mLstImpHIs = oIni.EnumSectionAndValues(AppCommon.SEC_IMPTABHIS);
            mLstImpHIs.ForEach(p =>
            {
                namesCollection.Add(p.Val);
            });
            cmbTarTab.AutoCompleteCustomSource = namesCollection;
        }

        public void SetCheckLastQuery(bool bcheck)
        {
            chkUseLastQuery.Checked = bcheck;
            txtXlsFile.Enabled = !bcheck;
            btnBrowse.Enabled = !bcheck;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            dlgOpen.Filter = "xlsx Files|*.xlsx|excel2003 Files|*.xls";
            dlgOpen.Title = "Select a Excel File";
            if (dlgOpen.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtXlsFile.Text = dlgOpen.FileName;
            }
        }

        private void btnDoImport_Click(object sender, EventArgs e)
        {
            if (!chkUseLastQuery.Checked)
            {
                mTarTable = cmbTarTab.Text.Trim();
                mXlsFile = txtXlsFile.Text.Trim();
                if (string.IsNullOrEmpty(mTarTable) || string.IsNullOrEmpty(mXlsFile))
                {
                    MessageBox.Show("请提供目标表和excel表格");
                    return;
                }
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void FrmImport_FormClosed(object sender, FormClosedEventArgs e)
        {
            var sLast = cmbTarTab.Text.Trim();
            if (string.IsNullOrEmpty(sLast))
            {
                return;
            }
            if(null == mLstImpHIs.Find(x =>
            {
                return 0==string.Compare(x.Val, sLast, true);
            }))
            {
                var oIni = new IniParser(AppCommon.IniFile);
                var ds = DateTime.Now.ToString("MMddHHmmss");
                oIni.AddSetting(AppCommon.SEC_IMPTABHIS, ds, sLast);
                oIni.SaveSettings();
            }
        }
    }
}
