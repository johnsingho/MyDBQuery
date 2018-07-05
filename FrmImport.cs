using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyDBQuery
{
    public partial class FrmImport : Form
    {
        public string mTarTable { get; set; }
        public string mXlsFile { get; set; }

        public FrmImport()
        {
            InitializeComponent();
            txtTar.Text = "TB_tmpIDCard";
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
            mTarTable = txtTar.Text.Trim();
            mXlsFile = txtXlsFile.Text.Trim();
            if(string.IsNullOrEmpty(mTarTable) || string.IsNullOrEmpty(mXlsFile))
            {
                MessageBox.Show("请提供目标表和excel表格");
                return;
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
