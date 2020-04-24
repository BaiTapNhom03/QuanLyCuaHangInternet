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
using COMExcel=Microsoft.Office.Interop.Excel;

namespace btlquanlycuahanginternet
{
    public partial class frmTKMayTinh : Form
    {
        DataTable tableTKMT;
        public frmTKMayTinh()
        {
            InitializeComponent();
        }

        private void frmTKMayTinh_Load(object sender, EventArgs e)
        {
            Class.functions.Connect();
            txtmamay.ReadOnly = true;
            txtMaPhong.ReadOnly = true;
            txtTenMay.ReadOnly = true;
            txtTinhTrang.ReadOnly = true;
            functions.FillCombo("SELECT DISTINCT TinhTrang FROM MayTinh", cboTinhTrang, "TinhTrang", "TinhTrang");
            cboTinhTrang.SelectedIndex = -1;
            functions.FillCombo("SELECT MaPhong, TenPhong FROM Phong", cboMaPhong, "MaPhong", "MaPhong");
            cboMaPhong.SelectedIndex = -1;
            loadDataToGridView();
            dataGridView_TKMT.DataSource = null;

        }
        private void loadDataToGridView()
        {
            String sql;
            sql = "select*from MayTinh";
            tableTKMT = Class.functions.GetDataToTable(sql);
            dataGridView_TKMT.DataSource = tableTKMT;
            dataGridView_TKMT.Columns[0].HeaderText = "Mã Máy";
            dataGridView_TKMT.Columns[1].HeaderText = "Tên Máy";
            dataGridView_TKMT.Columns[2].HeaderText = "Mã Phòng";
            dataGridView_TKMT.Columns[3].HeaderText = "Tình Trạng";
            dataGridView_TKMT.Columns[0].Width = 150;
            dataGridView_TKMT.Columns[1].Width = 100;
            dataGridView_TKMT.Columns[2].Width = 100;
            dataGridView_TKMT.Columns[3].Width = 100;
            dataGridView_TKMT.AllowUserToAddRows = false;
            dataGridView_TKMT.EditMode = DataGridViewEditMode.EditProgrammatically;

        }
        private void dataGridView_TKMT_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtmamay.Text = dataGridView_TKMT.CurrentRow.Cells["MaMay"].Value.ToString();
            txtTenMay.Text = dataGridView_TKMT.CurrentRow.Cells["TenMay"].Value.ToString();
            txtMaPhong.Text = dataGridView_TKMT.CurrentRow.Cells["MaPhong"].Value.ToString();
            txtTinhTrang.Text = dataGridView_TKMT.CurrentRow.Cells["TinhTrang"].Value.ToString();
        }
        private void ResetValues()
        {
            foreach (Control Ctl in this.Controls)
                if (Ctl is TextBox)
                    Ctl.Text = "";
            cboMaPhong.Focus();
        }
        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string sql;
            if ((cboMaPhong.Text == "") && (cboTinhTrang.Text == ""))
            {
                MessageBox.Show("Hãy nhập một điều kiện tìm kiếm!!!", "Yêu cầu ...", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            sql = "SELECT * FROM MayTinh WHERE 1=1";
            if (cboMaPhong.Text != "")
                sql = sql + " AND MaPhong Like '%" + cboMaPhong.Text + "%' ";
            if (cboTinhTrang.Text != "")
                sql = sql + " AND TinhTrang Like '%" + cboTinhTrang.Text + "%'";
            DataTable tblMT = functions.GetDataToTable(sql);
            if (tblMT.Rows.Count == 0)
            {
                MessageBox.Show("Không có bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Có " + tblMT.Rows.Count + " bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            tableTKMT = Class.functions.GetDataToTable(sql);
            dataGridView_TKMT.DataSource = tableTKMT;
            ResetValues();
        }
        private void btnTimLai_Click(object sender, EventArgs e)
        {
            ResetValues();
            dataGridView_TKMT.DataSource = null;
        }

        private void btnHienThi_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn hiển thị thông tin chi tiết?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)

            {
                string sql;
                sql = "SELECT * FROM MayTinh";
                DataTable tblMT = Class.functions.GetDataToTable(sql);
                dataGridView_TKMT.DataSource = tblMT;
            }
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            // Khởi động chương trình Excel
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook; //Trong 1 chương trình Excel có nhiều Workbook
            COMExcel.Worksheet exSheet; //Trong 1 Workbook có nhiều Worksheet
            COMExcel.Range exRange;
            int may = 0, cot = 0;
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
            exRange.Range["C2:E2"].Value = "Danh Sách Máy Tính";
            //Tạo dòng tiêu đề bảng
            exRange.Range["A6:P6"].Font.Bold = true;
            exRange.Range["A6:P6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C6:P6"].ColumnWidth = 12;
            exRange.Range["A6:A6"].Value = "STT";
            exRange.Range["B6:B6"].Value = "Mã Máy";
            exRange.Range["C6:C6"].Value = "Tên Máy";
            exRange.Range["D6:D6"].Value = "Mã Phòng";
            exRange.Range["E6:E6"].Value = "Mã Ổ Cứng";
            exRange.Range["F6:F6"].Value = "Mã DLượng";
            exRange.Range["G6:G6"].Value = "Mã Chip";
            exRange.Range["H6:H6"].Value = "Mã Ram";
            exRange.Range["I6:I6"].Value = "Mã Tốc Độ";
            exRange.Range["J6:J6"].Value = "Mã MH";
            exRange.Range["K6:K6"].Value = "Mã Size MH";
            exRange.Range["L6:L6"].Value = "Mã Chuột";
            exRange.Range["M6:M6"].Value = "Mã Bàn Phím";
            exRange.Range["N6:N6"].Value = "Mã Ổ Đĩa";
            exRange.Range["O6:P6"].Value = "Mã Loa";
            exRange.Range["P6:P6"].Value = "Tình Trạng";
            for (may = 0; may < tableTKMT.Rows.Count; may++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 12
                exSheet.Cells[1][may + 7] = may + 1;
                for (cot = 0; cot < tableTKMT.Columns.Count; cot++)
                //Điền thông tin hàng từ cột thứ 2, dòng 7
                {
                    exSheet.Cells[cot + 2][may + 7] = tableTKMT.Rows[may][cot].ToString();
                    if (cot == 3) exSheet.Cells[cot + 2][may + 7] = tableTKMT.Rows[may][cot].ToString();
                }
            }
            exApp.Visible = true;
        }

       
    }
}
