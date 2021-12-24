
namespace Ontap
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.grbthongtin = new System.Windows.Forms.GroupBox();
            this.txtdkt = new System.Windows.Forms.TextBox();
            this.txtdqt = new System.Windows.Forms.TextBox();
            this.txtmssv = new System.Windows.Forms.TextBox();
            this.txthoten = new System.Windows.Forms.TextBox();
            this.dgvthongke = new System.Windows.Forms.DataGridView();
            this.clnhoten = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clnmssv = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clndqt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clndkt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clntongket = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnthem = new System.Windows.Forms.Button();
            this.btnsua = new System.Windows.Forms.Button();
            this.btnluu = new System.Windows.Forms.Button();
            this.btnxoa = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tệpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.thêmMớiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mởToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lưuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.thoátToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.chứcNăngToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lưuDạngExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mởFileExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton3 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton4 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton5 = new System.Windows.Forms.ToolStripButton();
            this.lblnote = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.grbthongtin.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvthongke)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(272, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "QUẢN LÝ DANH SÁCH SINH VIÊN";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 17);
            this.label2.TabIndex = 1;
            this.label2.Text = "Họ tên";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(31, 87);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 17);
            this.label3.TabIndex = 1;
            this.label3.Text = "MSSV";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(39, 129);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(38, 17);
            this.label4.TabIndex = 1;
            this.label4.Text = "ĐQT";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(157, 129);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(36, 17);
            this.label5.TabIndex = 1;
            this.label5.Text = "ĐKT";
            // 
            // grbthongtin
            // 
            this.grbthongtin.Controls.Add(this.txtdkt);
            this.grbthongtin.Controls.Add(this.txtdqt);
            this.grbthongtin.Controls.Add(this.txtmssv);
            this.grbthongtin.Controls.Add(this.txthoten);
            this.grbthongtin.Controls.Add(this.label2);
            this.grbthongtin.Controls.Add(this.label5);
            this.grbthongtin.Controls.Add(this.label3);
            this.grbthongtin.Controls.Add(this.label4);
            this.grbthongtin.Location = new System.Drawing.Point(33, 88);
            this.grbthongtin.Name = "grbthongtin";
            this.grbthongtin.Size = new System.Drawing.Size(288, 171);
            this.grbthongtin.TabIndex = 2;
            this.grbthongtin.TabStop = false;
            // 
            // txtdkt
            // 
            this.txtdkt.Location = new System.Drawing.Point(199, 126);
            this.txtdkt.Name = "txtdkt";
            this.txtdkt.Size = new System.Drawing.Size(64, 22);
            this.txtdkt.TabIndex = 2;
            // 
            // txtdqt
            // 
            this.txtdqt.Location = new System.Drawing.Point(87, 126);
            this.txtdqt.Name = "txtdqt";
            this.txtdqt.Size = new System.Drawing.Size(64, 22);
            this.txtdqt.TabIndex = 2;
            // 
            // txtmssv
            // 
            this.txtmssv.Location = new System.Drawing.Point(87, 84);
            this.txtmssv.Name = "txtmssv";
            this.txtmssv.Size = new System.Drawing.Size(100, 22);
            this.txtmssv.TabIndex = 2;
            // 
            // txthoten
            // 
            this.txthoten.Location = new System.Drawing.Point(87, 48);
            this.txthoten.Name = "txthoten";
            this.txthoten.Size = new System.Drawing.Size(100, 22);
            this.txthoten.TabIndex = 2;
            this.txthoten.TextChanged += new System.EventHandler(this.txthoten_TextChanged);
            // 
            // dgvthongke
            // 
            this.dgvthongke.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvthongke.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvthongke.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clnhoten,
            this.clnmssv,
            this.clndqt,
            this.clndkt,
            this.clntongket});
            this.dgvthongke.Location = new System.Drawing.Point(369, 97);
            this.dgvthongke.Name = "dgvthongke";
            this.dgvthongke.RowHeadersWidth = 51;
            this.dgvthongke.RowTemplate.Height = 24;
            this.dgvthongke.Size = new System.Drawing.Size(679, 197);
            this.dgvthongke.TabIndex = 3;
            this.dgvthongke.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvthongke_CellClick);
            this.dgvthongke.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvthongke_CellContentClick);
            // 
            // clnhoten
            // 
            this.clnhoten.HeaderText = "Họ tên";
            this.clnhoten.MinimumWidth = 6;
            this.clnhoten.Name = "clnhoten";
            // 
            // clnmssv
            // 
            this.clnmssv.HeaderText = "MSSV";
            this.clnmssv.MinimumWidth = 6;
            this.clnmssv.Name = "clnmssv";
            // 
            // clndqt
            // 
            this.clndqt.HeaderText = "ĐQT";
            this.clndqt.MinimumWidth = 6;
            this.clndqt.Name = "clndqt";
            // 
            // clndkt
            // 
            this.clndkt.HeaderText = "ĐKT";
            this.clndkt.MinimumWidth = 6;
            this.clndkt.Name = "clndkt";
            // 
            // clntongket
            // 
            this.clntongket.HeaderText = "Tổng kết";
            this.clntongket.MinimumWidth = 6;
            this.clntongket.Name = "clntongket";
            // 
            // btnthem
            // 
            this.btnthem.Location = new System.Drawing.Point(39, 271);
            this.btnthem.Name = "btnthem";
            this.btnthem.Size = new System.Drawing.Size(75, 23);
            this.btnthem.TabIndex = 4;
            this.btnthem.Text = "Thêm";
            this.btnthem.UseVisualStyleBackColor = true;
            this.btnthem.Click += new System.EventHandler(this.btnthem_Click);
            // 
            // btnsua
            // 
            this.btnsua.Location = new System.Drawing.Point(39, 315);
            this.btnsua.Name = "btnsua";
            this.btnsua.Size = new System.Drawing.Size(75, 23);
            this.btnsua.TabIndex = 4;
            this.btnsua.Text = "Sửa";
            this.btnsua.UseVisualStyleBackColor = true;
            this.btnsua.Click += new System.EventHandler(this.btnsua_Click);
            // 
            // btnluu
            // 
            this.btnluu.Location = new System.Drawing.Point(180, 271);
            this.btnluu.Name = "btnluu";
            this.btnluu.Size = new System.Drawing.Size(75, 23);
            this.btnluu.TabIndex = 4;
            this.btnluu.Text = "Lưu";
            this.btnluu.UseVisualStyleBackColor = true;
            this.btnluu.Click += new System.EventHandler(this.btnluu_Click);
            // 
            // btnxoa
            // 
            this.btnxoa.Location = new System.Drawing.Point(180, 315);
            this.btnxoa.Name = "btnxoa";
            this.btnxoa.Size = new System.Drawing.Size(75, 23);
            this.btnxoa.TabIndex = 4;
            this.btnxoa.Text = "Xóa";
            this.btnxoa.UseVisualStyleBackColor = true;
            this.btnxoa.Click += new System.EventHandler(this.btnxoa_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tệpToolStripMenuItem,
            this.chứcNăngToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1070, 28);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tệpToolStripMenuItem
            // 
            this.tệpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.thêmMớiToolStripMenuItem,
            this.mởToolStripMenuItem,
            this.lưuToolStripMenuItem,
            this.toolStripSeparator1,
            this.thoátToolStripMenuItem});
            this.tệpToolStripMenuItem.Name = "tệpToolStripMenuItem";
            this.tệpToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.tệpToolStripMenuItem.Size = new System.Drawing.Size(48, 24);
            this.tệpToolStripMenuItem.Text = "&Tệp";
            // 
            // thêmMớiToolStripMenuItem
            // 
            this.thêmMớiToolStripMenuItem.Name = "thêmMớiToolStripMenuItem";
            this.thêmMớiToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.thêmMớiToolStripMenuItem.Size = new System.Drawing.Size(216, 26);
            this.thêmMớiToolStripMenuItem.Text = "Thêm mới ";
            // 
            // mởToolStripMenuItem
            // 
            this.mởToolStripMenuItem.Name = "mởToolStripMenuItem";
            this.mởToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.mởToolStripMenuItem.Size = new System.Drawing.Size(216, 26);
            this.mởToolStripMenuItem.Text = "Mở";
            // 
            // lưuToolStripMenuItem
            // 
            this.lưuToolStripMenuItem.Name = "lưuToolStripMenuItem";
            this.lưuToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.lưuToolStripMenuItem.Size = new System.Drawing.Size(216, 26);
            this.lưuToolStripMenuItem.Text = "Lưu";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(213, 6);
            // 
            // thoátToolStripMenuItem
            // 
            this.thoátToolStripMenuItem.Name = "thoátToolStripMenuItem";
            this.thoátToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.thoátToolStripMenuItem.Size = new System.Drawing.Size(216, 26);
            this.thoátToolStripMenuItem.Text = "Thoát";
            // 
            // chứcNăngToolStripMenuItem
            // 
            this.chứcNăngToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lưuDạngExcelToolStripMenuItem,
            this.mởFileExcelToolStripMenuItem});
            this.chứcNăngToolStripMenuItem.Name = "chứcNăngToolStripMenuItem";
            this.chứcNăngToolStripMenuItem.Size = new System.Drawing.Size(93, 24);
            this.chứcNăngToolStripMenuItem.Text = "Chức năng";
            // 
            // lưuDạngExcelToolStripMenuItem
            // 
            this.lưuDạngExcelToolStripMenuItem.Name = "lưuDạngExcelToolStripMenuItem";
            this.lưuDạngExcelToolStripMenuItem.Size = new System.Drawing.Size(192, 26);
            this.lưuDạngExcelToolStripMenuItem.Text = "Lưu dạng Excel";
            this.lưuDạngExcelToolStripMenuItem.Click += new System.EventHandler(this.lưuDạngExcelToolStripMenuItem_Click);
            // 
            // mởFileExcelToolStripMenuItem
            // 
            this.mởFileExcelToolStripMenuItem.Name = "mởFileExcelToolStripMenuItem";
            this.mởFileExcelToolStripMenuItem.Size = new System.Drawing.Size(192, 26);
            this.mởFileExcelToolStripMenuItem.Text = "Mở file Excel";
            this.mởFileExcelToolStripMenuItem.Click += new System.EventHandler(this.mởFileExcelToolStripMenuItem_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1,
            this.toolStripButton2,
            this.toolStripButton3,
            this.toolStripButton4,
            this.toolStripButton5});
            this.toolStrip1.Location = new System.Drawing.Point(0, 28);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1070, 27);
            this.toolStrip1.TabIndex = 7;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(29, 24);
            this.toolStripButton1.Text = "toolStripButton1";
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(29, 24);
            this.toolStripButton2.Text = "toolStripButton2";
            this.toolStripButton2.Click += new System.EventHandler(this.toolStripButton2_Click);
            // 
            // toolStripButton3
            // 
            this.toolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton3.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton3.Image")));
            this.toolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton3.Name = "toolStripButton3";
            this.toolStripButton3.Size = new System.Drawing.Size(29, 24);
            this.toolStripButton3.Text = "toolStripButton3";
            this.toolStripButton3.Click += new System.EventHandler(this.toolStripButton3_Click);
            // 
            // toolStripButton4
            // 
            this.toolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton4.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton4.Image")));
            this.toolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton4.Name = "toolStripButton4";
            this.toolStripButton4.Size = new System.Drawing.Size(29, 24);
            this.toolStripButton4.Text = "toolStripButton4";
            // 
            // toolStripButton5
            // 
            this.toolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton5.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton5.Image")));
            this.toolStripButton5.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton5.Name = "toolStripButton5";
            this.toolStripButton5.Size = new System.Drawing.Size(29, 24);
            this.toolStripButton5.Text = "toolStripButton5";
            // 
            // lblnote
            // 
            this.lblnote.AutoSize = true;
            this.lblnote.Location = new System.Drawing.Point(366, 297);
            this.lblnote.Name = "lblnote";
            this.lblnote.Size = new System.Drawing.Size(48, 17);
            this.lblnote.TabIndex = 1;
            this.lblnote.Text = "Chú ý:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1070, 450);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.btnxoa);
            this.Controls.Add(this.btnluu);
            this.Controls.Add(this.btnsua);
            this.Controls.Add(this.lblnote);
            this.Controls.Add(this.btnthem);
            this.Controls.Add(this.dgvthongke);
            this.Controls.Add(this.grbthongtin);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.grbthongtin.ResumeLayout(false);
            this.grbthongtin.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvthongke)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox grbthongtin;
        private System.Windows.Forms.TextBox txtdkt;
        private System.Windows.Forms.TextBox txtdqt;
        private System.Windows.Forms.TextBox txtmssv;
        private System.Windows.Forms.TextBox txthoten;
        private System.Windows.Forms.DataGridView dgvthongke;
        private System.Windows.Forms.DataGridViewTextBoxColumn clnhoten;
        private System.Windows.Forms.DataGridViewTextBoxColumn clnmssv;
        private System.Windows.Forms.DataGridViewTextBoxColumn clndqt;
        private System.Windows.Forms.DataGridViewTextBoxColumn clndkt;
        private System.Windows.Forms.DataGridViewTextBoxColumn clntongket;
        private System.Windows.Forms.Button btnthem;
        private System.Windows.Forms.Button btnsua;
        private System.Windows.Forms.Button btnluu;
        private System.Windows.Forms.Button btnxoa;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tệpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem thêmMớiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mởToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem lưuToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem thoátToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem chứcNăngToolStripMenuItem;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripButton toolStripButton2;
        private System.Windows.Forms.ToolStripButton toolStripButton3;
        private System.Windows.Forms.ToolStripButton toolStripButton4;
        private System.Windows.Forms.ToolStripButton toolStripButton5;
        private System.Windows.Forms.Label lblnote;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem lưuDạngExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mởFileExcelToolStripMenuItem;
    }
}

