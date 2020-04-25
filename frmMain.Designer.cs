namespace btlquanlycuahanginternet
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.danhMụcToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuPhong = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMayTinh = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuThueMay = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuChonMay = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuBaoTri = new System.Windows.Forms.ToolStripMenuItem();
            this.báoCáoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuChiPhiBaoTri = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuDoanhThu = new System.Windows.Forms.ToolStripMenuItem();
            this.trợGiúpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuTKMT = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuNhanVien = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuthoat = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuDMNV = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.danhMụcToolStripMenuItem,
            this.báoCáoToolStripMenuItem,
            this.trợGiúpToolStripMenuItem,
            this.mnuthoat});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(739, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // danhMụcToolStripMenuItem
            // 
            this.danhMụcToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuPhong,
            this.mnuMayTinh,
            this.mnuThueMay,
            this.mnuBaoTri,
            this.mnuDMNV});
            this.danhMụcToolStripMenuItem.Name = "danhMụcToolStripMenuItem";
            this.danhMụcToolStripMenuItem.Size = new System.Drawing.Size(74, 20);
            this.danhMụcToolStripMenuItem.Text = "Danh mục";
            // 
            // mnuPhong
            // 
            this.mnuPhong.Name = "mnuPhong";
            this.mnuPhong.Size = new System.Drawing.Size(180, 22);
            this.mnuPhong.Text = "Phòng";
            // 
            // mnuMayTinh
            // 
            this.mnuMayTinh.Name = "mnuMayTinh";
            this.mnuMayTinh.Size = new System.Drawing.Size(180, 22);
            this.mnuMayTinh.Text = "Máy tính";
            this.mnuMayTinh.Click += new System.EventHandler(this.mnuMayTinh_Click);
            // 
            // mnuThueMay
            // 
            this.mnuThueMay.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuChonMay});
            this.mnuThueMay.Name = "mnuThueMay";
            this.mnuThueMay.Size = new System.Drawing.Size(180, 22);
            this.mnuThueMay.Text = "Thuê máy";
            // 
            // mnuChonMay
            // 
            this.mnuChonMay.Name = "mnuChonMay";
            this.mnuChonMay.Size = new System.Drawing.Size(129, 22);
            this.mnuChonMay.Text = "Chọn Máy";
            // 
            // mnuBaoTri
            // 
            this.mnuBaoTri.Name = "mnuBaoTri";
            this.mnuBaoTri.Size = new System.Drawing.Size(180, 22);
            this.mnuBaoTri.Text = "Bảo Trì";
            // 
            // báoCáoToolStripMenuItem
            // 
            this.báoCáoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuChiPhiBaoTri,
            this.mnuDoanhThu});
            this.báoCáoToolStripMenuItem.Name = "báoCáoToolStripMenuItem";
            this.báoCáoToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.báoCáoToolStripMenuItem.Text = "Báo cáo";
            // 
            // mnuChiPhiBaoTri
            // 
            this.mnuChiPhiBaoTri.Name = "mnuChiPhiBaoTri";
            this.mnuChiPhiBaoTri.Size = new System.Drawing.Size(149, 22);
            this.mnuChiPhiBaoTri.Text = "Chi phí bảo trì";
            this.mnuChiPhiBaoTri.Click += new System.EventHandler(this.mnuChiPhiBaoTri_Click);
            // 
            // mnuDoanhThu
            // 
            this.mnuDoanhThu.Name = "mnuDoanhThu";
            this.mnuDoanhThu.Size = new System.Drawing.Size(149, 22);
            this.mnuDoanhThu.Text = "Doanh thu";
            // 
            // trợGiúpToolStripMenuItem
            // 
            this.trợGiúpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuTKMT,
            this.mnuNhanVien});
            this.trợGiúpToolStripMenuItem.Name = "trợGiúpToolStripMenuItem";
            this.trợGiúpToolStripMenuItem.Size = new System.Drawing.Size(69, 20);
            this.trợGiúpToolStripMenuItem.Text = "Tìm Kiếm";
            // 
            // mnuTKMT
            // 
            this.mnuTKMT.Name = "mnuTKMT";
            this.mnuTKMT.Size = new System.Drawing.Size(129, 22);
            this.mnuTKMT.Text = "Máy Tính";
            this.mnuTKMT.Click += new System.EventHandler(this.mnuTKMT_Click);
            // 
            // mnuNhanVien
            // 
            this.mnuNhanVien.Name = "mnuNhanVien";
            this.mnuNhanVien.Size = new System.Drawing.Size(129, 22);
            this.mnuNhanVien.Text = "Nhân Viên";
            this.mnuNhanVien.Click += new System.EventHandler(this.mnuNhanVien_Click);
            // 
            // mnuthoat
            // 
            this.mnuthoat.Name = "mnuthoat";
            this.mnuthoat.Size = new System.Drawing.Size(49, 20);
            this.mnuthoat.Text = "Thoát";
            this.mnuthoat.Click += new System.EventHandler(this.mnuthoat_Click);
            // 
            // mnuDMNV
            // 
            this.mnuDMNV.Name = "mnuDMNV";
            this.mnuDMNV.Size = new System.Drawing.Size(180, 22);
            this.mnuDMNV.Text = "Nhân Viên";
            this.mnuDMNV.Click += new System.EventHandler(this.mnuDMNV_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(739, 354);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmMain";
            this.Text = "Chương Trình Quản Lý Internet";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem danhMụcToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mnuPhong;
        private System.Windows.Forms.ToolStripMenuItem mnuMayTinh;
        private System.Windows.Forms.ToolStripMenuItem báoCáoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem trợGiúpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mnuChiPhiBaoTri;
        private System.Windows.Forms.ToolStripMenuItem mnuDoanhThu;
        private System.Windows.Forms.ToolStripMenuItem mnuTKMT;
        private System.Windows.Forms.ToolStripMenuItem mnuNhanVien;
        private System.Windows.Forms.ToolStripMenuItem mnuthoat;
        private System.Windows.Forms.ToolStripMenuItem mnuThueMay;
        private System.Windows.Forms.ToolStripMenuItem mnuChonMay;
        private System.Windows.Forms.ToolStripMenuItem mnuBaoTri;
        private System.Windows.Forms.ToolStripMenuItem mnuDMNV;
    }
}