using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using btlquanlycuahanginternet.Class;
using COMExcel = Microsoft.Office.Interop.Excel;

namespace btlquanlycuahanginternet
{
    public partial class frmBCCPBaoTri : Form
    {
        DataTable tableBaoCao;
        public frmBCCPBaoTri()
        {
            InitializeComponent();
        }

        private void frmBCCPBaoTri_Load(object sender, EventArgs e)
        {
            Class.functions.Connect();
            txtmamay.ReadOnly = true;
            txtmaphong.ReadOnly = true;
            txttenmay.ReadOnly = true;
            txtthanhten.ReadOnly = true;
            txtTongTien.ReadOnly = true;
            txtTongTien.Text = "0";
            cboThang.Items.Add("1");
            cboThang.Items.Add("2");
            cboThang.Items.Add("3");
            cboThang.Items.Add("4");
            cboThang.Items.Add("5");
            cboThang.Items.Add("6");
            cboThang.Items.Add("7");
            cboThang.Items.Add("8");
            cboThang.Items.Add("9");
            cboThang.Items.Add("10");
            cboThang.Items.Add("11");
            cboThang.Items.Add("12");
            cboQuy.Items.Add("1");
            cboQuy.Items.Add("2");
            cboQuy.Items.Add("3");
            cboQuy.Items.Add("4");
            functions.FillCombo("SELECT MaPhong, TenPhong FROM Phong", cboMaPhong, "MaPhong", "MaPhong");
            cboMaPhong.SelectedIndex = -1;
            loadDataToGridView();
            dataGridView_CPBT.DataSource = null;
        }
        private void loadDataToGridView()
        {
            String sql;
            sql = "select MaPhong, MayTinh.MaMay, TenMay,NgayBaoTri,ThanhTien from MayTinh join ChiTietBaoTri on MayTinh.MaMay=ChiTietBaoTri.MaMay join BaoTri on ChiTietBaoTri.MaBaoTri=BaoTri.MaBaoTri";
            tableBaoCao = Class.functions.GetDataToTable(sql);
            //str = "SELECT TongTien FROM tblHDBan WHERE MaHDBan = N'" + txtMaHDBan.Text + "'";
            //txtTongTien.Text = Functions.GetFieldValues(str);
            //lblBangChu.Text = "Bằng chữ: " + functions.ChuyenSoSangChu(txtTongTien.Text);
            dataGridView_CPBT.DataSource = tableBaoCao;
            dataGridView_CPBT.Columns[0].HeaderText = "Mã Phòng";
            dataGridView_CPBT.Columns[1].HeaderText = "Mã Máy";
            dataGridView_CPBT.Columns[2].HeaderText = "Tên Máy";
            dataGridView_CPBT.Columns[3].HeaderText = "Ngày Bảo Trì";
            dataGridView_CPBT.Columns[4].HeaderText = "Thành Tiền";
            dataGridView_CPBT.Columns[0].Width = 150;
            dataGridView_CPBT.Columns[1].Width = 100;
            dataGridView_CPBT.Columns[2].Width = 100;
            dataGridView_CPBT.Columns[3].Width = 100;
            dataGridView_CPBT.Columns[4].Width = 100;
            dataGridView_CPBT.AllowUserToAddRows = false;
            dataGridView_CPBT.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string sql = "SELECT MaPhong, MayTinh.MaMay, TenMay,NgayBaoTri,ThanhTien from MayTinh join ChiTietBaoTri on MayTinh.MaMay=ChiTietBaoTri.MaMay join BaoTri on ChiTietBaoTri.MaBaoTri=BaoTri.MaBaoTri WHERE 1=1";
            if (cboMaPhong.Text != "")
                sql = sql + " AND MaPhong Like '%" + cboMaPhong.Text + "%' ";
            if (cboThang.Text != "")
                sql = sql + " AND MONTH(NgayBaoTri) Like '%" + cboThang.Text + "%' ";
            if (cboQuy.Text != "")
                sql = sql + " AND DATEPART(quarter, NgayBaoTri) Like '%" + cboQuy.Text + "%'";
            if (txtNam.Text != "")
                sql = sql + "AND Year(NgayBaoTri) Like '%" + txtNam.Text + "%'";
            tableBaoCao = functions.GetDataToTable(sql);
            if (tableBaoCao.Rows.Count == 0)
            {
                MessageBox.Show("Không có bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Có " +tableBaoCao.Rows.Count + " bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            tableBaoCao = Class.functions.GetDataToTable(sql);
            dataGridView_CPBT.DataSource = tableBaoCao;
            ResetValues();
        }
        private void ResetValues()
        {
            txtNam.Text = "";
            cboThang.Text = "";
            cboQuy.Text = "";
            txtTongTien.Text = "0";
            cboThang.Focus();
        }

        private void btnTimLai_Click(object sender, EventArgs e)
        {
            ResetValues();
            dataGridView_CPBT.DataSource = null;
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            // Khởi động chương trình Excel
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook; //Trong 1 chương trình Excel có nhiều Workbook
            COMExcel.Worksheet exSheet; //Trong 1 Workbook có nhiều Worksheet
            COMExcel.Range exRange;
            int phong = 0, cot = 0;
            exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            exSheet = exBook.Worksheets[1];
            // Định dạng chung
            exRange = exSheet.Cells[1, 1];
            exRange.Range["A1:Z300"].Font.Name = "Times new roman"; //Font chữ
            exRange.Range["A1:B3"].Font.Size = 10;
            exRange.Range["A1:B3"].Font.Bold = true;
            exRange.Range["A1:B3"].Font.ColorIndex = 5; //Màu xanh da trời
            exRange.Range["A1:A1"].ColumnWidth = 7;
            exRange.Range["B1:B1"].ColumnWidth = 15;
            exRange.Range["A1:B1"].MergeCells = true;
            exRange.Range["A1:B1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A1:B1"].Value = "Cửa Hàng Internet03";
            exRange.Range["A2:B2"].MergeCells = true;
            exRange.Range["A2:B2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:B2"].Value = "Xuân Mai - Hà Nội";
            exRange.Range["A3:B3"].MergeCells = true;
            exRange.Range["A3:B3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:B3"].Value = "Điện thoại: (04)39641582";
            exRange.Range["C2:E2"].Font.Size = 16;
            exRange.Range["C2:E2"].Font.Bold = true;
            exRange.Range["C2:E2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:E2"].MergeCells = true;
            exRange.Range["C2:E2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:E2"].Value = "Báo Cáo Tổng Phí Bảo Trì";
            //Tạo dòng tiêu đề bảng
            exRange.Range["A6:F6"].Font.Bold = true;
            exRange.Range["A6:F6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C6:F6"].ColumnWidth = 12;
            exRange.Range["A6:A6"].Value = "STT";
            exRange.Range["B6:B6"].Value = "Mã Phòng";
            exRange.Range["C6:C6"].Value = "Mã Máy";
            exRange.Range["D6:D6"].Value = "Tên Máy";
            exRange.Range["E6:E6"].Value = "Ngày Bảo Trì";
            exRange.Range["F6:F6"].Value = "Thành Tiền";
            for (phong = 0; phong < tableBaoCao.Rows.Count; phong++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 12
                exSheet.Cells[1][phong + 7] = phong + 1;
                for (cot = 0; cot < tableBaoCao.Columns.Count; cot++)
                //Điền thông tin hàng từ cột thứ 2, dòng 7
                {
                    exSheet.Cells[cot + 2][phong + 7] = tableBaoCao.Rows[phong][cot].ToString();
                    if (cot == 3) exSheet.Cells[cot + 2][phong + 7] = tableBaoCao.Rows[phong][cot].ToString();
                }
            }
            exApp.Visible = true;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
