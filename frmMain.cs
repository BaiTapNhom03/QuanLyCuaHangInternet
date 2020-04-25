using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace btlquanlycuahanginternet
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            Class.functions.Connect();
        }

        private void mnuthoat_Click(object sender, EventArgs e)
        {
            Class.functions.DisConnect();
            Application.Exit();
        }

        private void mnuTKMT_Click(object sender, EventArgs e)
        {
            frmTKMayTinh f = new frmTKMayTinh();
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
        }

        private void mnuMayTinh_Click(object sender, EventArgs e)
        {
            frmMayTinh f = new frmMayTinh();
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
        }

        private void mnuChiPhiBaoTri_Click(object sender, EventArgs e)
        {
            frmBCCPBaoTri f = new frmBCCPBaoTri();
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
        }

        private void mnuNhanVien_Click(object sender, EventArgs e)
        {
            frmTKNV f = new frmTKNV();
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
        }

        private void mnuDMNV_Click(object sender, EventArgs e)
        {
            frmDMNV f = new frmDMNV();
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
        }
    }
}
