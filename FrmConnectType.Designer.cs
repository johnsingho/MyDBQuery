namespace MyDBQuery
{
    partial class FrmConnectType
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.txtConnStr = new System.Windows.Forms.TextBox();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.cmbType = new System.Windows.Forms.ComboBox();
            this.btnSetConn = new System.Windows.Forms.Button();
            this.gvHis = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvHis)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.txtConnStr);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(699, 326);
            this.splitContainer1.SplitterDistance = 301;
            this.splitContainer1.TabIndex = 0;
            // 
            // txtConnStr
            // 
            this.txtConnStr.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtConnStr.Location = new System.Drawing.Point(0, 0);
            this.txtConnStr.Multiline = true;
            this.txtConnStr.Name = "txtConnStr";
            this.txtConnStr.Size = new System.Drawing.Size(301, 326);
            this.txtConnStr.TabIndex = 0;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.cmbType);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.btnSetConn);
            this.splitContainer2.Panel2.Controls.Add(this.gvHis);
            this.splitContainer2.Size = new System.Drawing.Size(394, 326);
            this.splitContainer2.SplitterDistance = 45;
            this.splitContainer2.TabIndex = 0;
            // 
            // cmbType
            // 
            this.cmbType.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.cmbType.BackColor = System.Drawing.Color.Wheat;
            this.cmbType.FormattingEnabled = true;
            this.cmbType.Location = new System.Drawing.Point(3, 11);
            this.cmbType.Name = "cmbType";
            this.cmbType.Size = new System.Drawing.Size(166, 21);
            this.cmbType.TabIndex = 0;
            this.cmbType.SelectedIndexChanged += new System.EventHandler(this.OnSelectedIndexChanged);
            // 
            // btnSetConn
            // 
            this.btnSetConn.Location = new System.Drawing.Point(307, 254);
            this.btnSetConn.Name = "btnSetConn";
            this.btnSetConn.Size = new System.Drawing.Size(75, 23);
            this.btnSetConn.TabIndex = 1;
            this.btnSetConn.Text = "Ok";
            this.btnSetConn.UseVisualStyleBackColor = true;
            this.btnSetConn.Click += new System.EventHandler(this.btnSetConn_Click);
            // 
            // gvHis
            // 
            this.gvHis.AllowUserToAddRows = false;
            this.gvHis.AllowUserToDeleteRows = false;
            this.gvHis.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvHis.Dock = System.Windows.Forms.DockStyle.Top;
            this.gvHis.Location = new System.Drawing.Point(0, 0);
            this.gvHis.Name = "gvHis";
            this.gvHis.ReadOnly = true;
            this.gvHis.Size = new System.Drawing.Size(394, 250);
            this.gvHis.TabIndex = 0;
            this.gvHis.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.OnCellDoubleClick);
            // 
            // FrmConnectType
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(699, 326);
            this.Controls.Add(this.splitContainer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "FrmConnectType";
            this.Text = "选择连接的数据库类型";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gvHis)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TextBox txtConnStr;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.ComboBox cmbType;
        private System.Windows.Forms.Button btnSetConn;
        private System.Windows.Forms.DataGridView gvHis;
    }
}