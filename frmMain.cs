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
        }
    }
}
