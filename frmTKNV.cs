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
    public partial class frmTKNV : Form
    {
        DataTable tableTKNV;
        public frmTKNV()
        {
            InitializeComponent();
        }

        private void frmTKNV_Load(object sender, EventArgs e)
        {
            Class.functions.Connect();
            txtmanv.ReadOnly = true;
            txtTenNVien.ReadOnly = true;
            txtmaca.ReadOnly = true;
            txtgioitinh.ReadOnly = true;
            txtnamsinh.ReadOnly = true;
            txtsdt.ReadOnly = true;
            functions.FillCombo("SELECT DISTINCT GioiTinh FROM NhanVien", cbogioitinh, "GioiTinh", "GioiTinh");
            cbogioitinh.SelectedIndex = -1;
            functions.FillCombo("SELECT DISTINCT MaCa FROM NhanVien", cbomaca, "MaCa", "MaCa");
            cbomaca.SelectedIndex = -1;
            loadDataToGridView();
            dataGridView_TKNV.DataSource = null;
        }
        private void loadDataToGridView()
        {
            String sql;
            sql = "select*from NhanVien";
            tableTKNV = Class.functions.GetDataToTable(sql);
            dataGridView_TKNV.DataSource = tableTKNV;
            dataGridView_TKNV.Columns[0].HeaderText = "Mã NV";
            dataGridView_TKNV.Columns[1].HeaderText = "Tên NV";
            dataGridView_TKNV.Columns[2].HeaderText = "Mã Ca";
            dataGridView_TKNV.Columns[3].HeaderText = "Giới Tính";
            dataGridView_TKNV.Columns[4].HeaderText = "Năm Sinh";
            dataGridView_TKNV.Columns[5].HeaderText = "SĐT";
            dataGridView_TKNV.Columns[2].HeaderText = "Địa Chỉ";
            dataGridView_TKNV.Columns[0].Width = 50;
            dataGridView_TKNV.Columns[1].Width = 250;
            dataGridView_TKNV.Columns[2].Width = 50;
            dataGridView_TKNV.Columns[3].Width = 50;
            dataGridView_TKNV.Columns[4].Width = 80;
            dataGridView_TKNV.Columns[5].Width = 80;
            dataGridView_TKNV.Columns[6].Width = 80;
            dataGridView_TKNV.AllowUserToAddRows = false;
            dataGridView_TKNV.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void dataGridView_TKNV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtmanv.Text = dataGridView_TKNV.CurrentRow.Cells["MaNV"].Value.ToString();
            txtTenNVien.Text = dataGridView_TKNV.CurrentRow.Cells["TenNV"].Value.ToString();
            txtmaca.Text = dataGridView_TKNV.CurrentRow.Cells["MaCa"].Value.ToString();
            txtgioitinh.Text = dataGridView_TKNV.CurrentRow.Cells["GioiTinh"].Value.ToString();
            txtnamsinh.Text = dataGridView_TKNV.CurrentRow.Cells["NamSinh"].Value.ToString();
            txtsdt.Text = dataGridView_TKNV.CurrentRow.Cells["SDT"].Value.ToString();
        }
        private void ResetValues()
        {
            txttenNV.Text = "";
            cbogioitinh.Text = "";
            cbomaca.Text = "";
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string sql;
            if ((txttenNV.Text == "") && (cbomaca.Text == "")&&(cbogioitinh.Text==""))
            {
                MessageBox.Show("Hãy nhập một điều kiện tìm kiếm!!!", "Yêu cầu ...", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            sql = "SELECT * FROM NhanVien WHERE 1=1";
            if (txttenNV.Text != "")
                sql = sql + " AND TenNV Like '%" + txttenNV.Text + "%' ";
            if (cbogioitinh.Text != "")
                sql = sql + " AND GioiTinh Like '%" + cbogioitinh.Text + "%'";
            if (cbomaca.Text != "")
                sql = sql + " AND MaCa Like '%" + cbomaca.Text + "%'";
            tableTKNV = functions.GetDataToTable(sql);
            if (tableTKNV.Rows.Count == 0)
            {
                MessageBox.Show("Không có bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Có " + tableTKNV.Rows.Count + " bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            tableTKNV = Class.functions.GetDataToTable(sql);
            dataGridView_TKNV.DataSource = tableTKNV;
            ResetValues();
        }

        private void btnTimLai_Click(object sender, EventArgs e)
        {
            ResetValues();
            dataGridView_TKNV.DataSource = null;
        }

        private void btnHienThi_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn hiển thị thông tin chi tiết?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)

            {
                string sql;
                sql = "SELECT * FROM NhanVien";
                DataTable tblMT = Class.functions.GetDataToTable(sql);
                dataGridView_TKNV.DataSource = tblMT;
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
            exRange.Range["C2:E2"].Value = "Danh Sách Nhân Viên";
            //Tạo dòng tiêu đề bảng
            exRange.Range["A6:H6"].Font.Bold = true;
            exRange.Range["A6:H6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C6:H6"].ColumnWidth = 12;
            exRange.Range["A6:A6"].Value = "STT";
            exRange.Range["B6:B6"].Value = "Mã NV";
            exRange.Range["C6:C6"].Value = "Tên NV";
            exRange.Range["D6:D6"].Value = "Mã Ca";
            exRange.Range["E6:E6"].Value = "Năm Sinh";
            exRange.Range["F6:F6"].Value = "Giới Tính";
            exRange.Range["G6:G6"].Value = "Địa Chỉ";
            exRange.Range["H6:H6"].Value = "SĐT";
            for (may = 0; may < tableTKNV.Rows.Count; may++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 12
                exSheet.Cells[1][may + 7] = may + 1;
                for (cot = 0; cot < tableTKNV.Columns.Count; cot++)
                //Điền thông tin hàng từ cột thứ 2, dòng 7
                {
                    exSheet.Cells[cot + 2][may + 7] = tableTKNV.Rows[may][cot].ToString();
                    if (cot == 3) exSheet.Cells[cot + 2][may + 7] = tableTKNV.Rows[may][cot].ToString();
                }
            }
            exApp.Visible = true;
        }
    }
}
