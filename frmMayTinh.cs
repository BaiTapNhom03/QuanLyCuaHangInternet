using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using btlquanlycuahanginternet.Class;

namespace btlquanlycuahanginternet
{
    public partial class frmMayTinh : Form
    {
        public frmMayTinh()
        {
            InitializeComponent();
        }

        private void frmMayTinh_Load(object sender, EventArgs e)
        {
            Class.functions.Connect();
            txtmamay.Enabled = false;
            btnLuu.Enabled = false;
            btnHuy.Enabled = false;
            functions.FillCombo("SELECT MaOCung, TenOCung FROM O_Cung", cboOCung, "MaOCung", "TenOCung");
            cboOCung.SelectedIndex = -1;
            functions.FillCombo("SELECT MaDLuong, TenDLuong FROM Dung_Luong", cboDLuong, "MaDLuong", "TenDLuong");
            cboDLuong.SelectedIndex = -1;
            functions.FillCombo("SELECT MaChip, TenChip FROM Chip", cboChip, "MaChip", "TenChip");
            cboChip.SelectedIndex = -1;
            functions.FillCombo("SELECT MaRam, TenRam FROM Ram", cboRam, "MaRam", "TenRam");
            cboRam.SelectedIndex = -1;
            functions.FillCombo("SELECT MaTocDo, TenTocDo FROM Toc_Do", cboTocDo, "MaTocDo", "TenTocDo");
            cboTocDo.SelectedIndex = -1;
            functions.FillCombo("SELECT MaManHinh, TenManHinh FROM Man_Hinh", cboMH, "MaManHinh", "TenManHinh");
            cboMH.SelectedIndex = -1;
            functions.FillCombo("SELECT MaSizeMH, TenSizeMH FROM SizeMH", cboSizeMH, "MaSizeMH", "TenSizeMH");
            cboSizeMH.SelectedIndex = -1;
            functions.FillCombo("SELECT MaChuot, TenChuot FROM Chuot", cboChuot, "MaChuot", "TenChuot");
            cboChuot.SelectedIndex = -1;
            functions.FillCombo("SELECT MaBanPhim, TenBanPhim FROM BanPhim", cboBPhim, "MaBanPhim", "TenBanPhim");
            cboBPhim.SelectedIndex = -1;
            functions.FillCombo("SELECT MaODia, TenODia FROM O_Dia", cboODia, "MaODia", "TenODia");
            cboODia.SelectedIndex = -1;
            functions.FillCombo("SELECT MaLoa, TenLoa FROM Loa", cboLoa, "MaLoa", "TenLoa");
            cboLoa.SelectedIndex = -1;
            loadDataToGridView();
            ResetValues();
        }
        private void loadDataToGridView()
        {
            string sql = "select * from MayTinh";
            DataTable tableMT = Class.functions.GetDataToTable(sql);
            dataGridView_maytinh.DataSource = tableMT;

        }

        private void dataGridView_maytinh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (btnThem.Enabled == false)
            {
                MessageBox.Show("Đang ở chế độ thêm mới!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtmamay.Focus();
                return;
            }
            txtmamay.Text = dataGridView_maytinh.CurrentRow.Cells["MaMay"].Value.ToString();
            txttenmt.Text = dataGridView_maytinh.CurrentRow.Cells["TenMay"].Value.ToString();
            txtmaphong.Text = dataGridView_maytinh.CurrentRow.Cells["MaPhong"].Value.ToString();
            cboOCung.Text = dataGridView_maytinh.CurrentRow.Cells["MaOCung"].Value.ToString();
            cboDLuong.Text = dataGridView_maytinh.CurrentRow.Cells["MaDLuong"].Value.ToString();
            cboChip.Text = dataGridView_maytinh.CurrentRow.Cells["MaChip"].Value.ToString();
            cboRam.Text = dataGridView_maytinh.CurrentRow.Cells["MaRam"].Value.ToString();
            cboTocDo.Text = dataGridView_maytinh.CurrentRow.Cells["MaTocDo"].Value.ToString();
            cboMH.Text = dataGridView_maytinh.CurrentRow.Cells["MaManHinh"].Value.ToString();
            cboSizeMH.Text = dataGridView_maytinh.CurrentRow.Cells["MaSizeMH"].Value.ToString();
            cboChuot.Text = dataGridView_maytinh.CurrentRow.Cells["MaChuot"].Value.ToString();
            cboBPhim.Text = dataGridView_maytinh.CurrentRow.Cells["MaBanPhim"].Value.ToString();
            cboODia.Text = dataGridView_maytinh.CurrentRow.Cells["MaODia"].Value.ToString();
            cboLoa.Text = dataGridView_maytinh.CurrentRow.Cells["MaLoa"].Value.ToString();
            txttinhtrang.Text = dataGridView_maytinh.CurrentRow.Cells["TinhTrang"].Value.ToString();
            txtghichu.Text = dataGridView_maytinh.CurrentRow.Cells["GhiChu"].Value.ToString();
            btnSua.Enabled = true;
            btnXoa.Enabled = true;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnHuy.Enabled = true;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;
            txtmamay.Enabled = true;
            txtmamay.Focus();
            ResetValues();
            loadDataToGridView();
        }
        private void ResetValues()
        {
            txtmamay.Text = "";
            txttenmt.Text = "";
            txtmaphong.Text = "";
            txttinhtrang.Text = "";
            cboOCung.Text = "";
            cboDLuong.Text = "";
            cboChip.Text = "";
            cboRam.Text = "";
            cboTocDo.Text = "";
            cboMH.Text = "";
            cboSizeMH.Text = "";
            cboChuot.Text = "";
            cboBPhim.Text = "";
            cboODia.Text = "";
            cboLoa.Text = "";
            txtghichu.Text = "";
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (txtmamay.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã máy", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtmamay.Focus();
                return;
            }
            if (txttenmt.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên máy", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttenmt.Focus();
                return;
            }
            if (txtmaphong.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã phòng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtmaphong.Focus();
                return;
            }
            if (txttinhtrang.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tình trạng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttinhtrang.Focus();
                return;
            }
            if (cboOCung.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã ổ cứng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboOCung.Focus();
                return;
            }
            if (cboDLuong.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã dung lượng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboDLuong.Focus();
                return;
            }
            if (cboChip.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã chip", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboChip.Focus();
                return;
            }
            if (cboRam.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã ram", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttinhtrang.Focus();
                return;
            }
            if (cboTocDo.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã tốc độ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboTocDo.Focus();
                return;
            }
            if (cboMH.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã màn hình", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMH.Focus();
                return;
            }
            if (cboSizeMH.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã size màn hình", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboSizeMH.Focus();
                return;
            }
            if (cboBPhim.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã bàn phím", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboBPhim.Focus();
                return;
            }
            if (cboODia.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã ổ đĩa", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboODia.Focus();
                return;
            }
            if (cboChuot.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã chuột", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboDLuong.Focus();
                return;
            }
            if (cboLoa.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã loa", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboLoa.Focus();
                return;
            }
            sql = "SELECT MaMay FROM MayTinh WHERE MaMay=N'" + txtmamay.Text.Trim() + "'";
            if (functions.CheckKey(sql))
            {
                MessageBox.Show("Mã máy này đã tồn tại, bạn phải chọn mã máy khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtmamay.Focus();
                return;
            }
            sql = "INSERT INTO MayTinh(MaMay,TenMay,MaPhong,MaOCung,MaDLuong,MaChip,MaRam,MaTocDo,MaManHinh,MaSizeMH,MaChuot,MaBanPhim,MaODia,MaLoa,TinhTrang,Ghichu)" +
                " VALUES('" + txtmamay.Text.Trim() + "','" + txttenmt.Text.Trim() + "','" + txtmaphong.Text.Trim() + "','" + cboOCung.SelectedValue.ToString() + "'," +
                "'" + cboDLuong.SelectedValue.ToString() + "','" + cboChip.SelectedValue.ToString() + "','" + cboRam.SelectedValue.ToString() + "'," +
                "'" + cboTocDo.SelectedValue.ToString() + "','" + cboMH.SelectedValue.ToString() + "','" + cboSizeMH.SelectedValue.ToString() + "'," +
                "'" + cboChuot.SelectedValue.ToString() + "','" + cboBPhim.SelectedValue.ToString() + "','" + cboODia.SelectedValue.ToString() + "'," +
                "'" + cboLoa.SelectedValue.ToString() + "','" + txttinhtrang.Text.Trim() + "','" + txtghichu.Text.Trim() + "')";
            functions.RunSQL(sql);
            loadDataToGridView();
            ResetValues();
            btnXoa.Enabled = true;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnHuy.Enabled = false;
            btnLuu.Enabled = false;
            txtmamay.Enabled = false;
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (txtmamay.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã máy", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtmamay.Focus();
                return;
            }
            if (txttenmt.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên máy", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttenmt.Focus();
                return;
            }
            if (txtmaphong.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã phòng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtmaphong.Focus();
                return;
            }
            if (txttinhtrang.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tình trạng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttinhtrang.Focus();
                return;
            }
            if (cboOCung.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã ổ cứng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboOCung.Focus();
                return;
            }
            if (cboDLuong.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã dung lượng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboDLuong.Focus();
                return;
            }
            if (cboChip.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã chip", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboChip.Focus();
                return;
            }
            if (cboRam.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã ram", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttinhtrang.Focus();
                return;
            }
            if (cboTocDo.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã tốc độ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboTocDo.Focus();
                return;
            }
            if (cboMH.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã màn hình", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMH.Focus();
                return;
            }
            if (cboSizeMH.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã size màn hình", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboSizeMH.Focus();
                return;
            }
            if (cboBPhim.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã bàn phím", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboBPhim.Focus();
                return;
            }
            if (cboODia.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã ổ đĩa", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboODia.Focus();
                return;
            }
            if (cboChuot.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã chuột", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboDLuong.Focus();
                return;
            }
            if (cboLoa.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã loa", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboLoa.Focus();
                return;
            }
            string sql = "update MayTinh set TenMay='" + txttenmt.Text.Trim().ToString()+ "',MaPhong='" + txtmaphong.Text.Trim().ToString() + "',MaOCung='" + cboOCung.Text+ "'," +
                 "MaDLuong='" + cboDLuong.Text + "',MaChip='" + cboChip.Text + "',MaRam='" + cboRam.Text + "'," +
                 "MaTocDo='" + cboTocDo.Text + "',MaManHinh='" + cboMH.Text+ "',MaSizeMH='" + cboSizeMH.Text+ "'," +
                 "MaChuot='" + cboChuot.Text + "',MaBanPhim='" + cboBPhim.Text + "',MaODia='" + cboODia.Text+ "'," +
                 "MaLoa='" + cboLoa.Text+ "',TinhTrang='" + txttinhtrang.Text + "',GhiChu='" + txtghichu.Text + "'Where MaMay='" + txtmamay.Text + "'";
            functions.RunSQL(sql);
            loadDataToGridView();
            ResetValues();
            btnHuy.Enabled = false;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (txtmamay.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (MessageBox.Show("Bạn có muốn xoá bản ghi này không?", "Thông báo",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                sql = "DELETE MayTinh WHERE MaMay='" + txtmamay.Text + "'";
                functions.RunSqlDel(sql);
                loadDataToGridView();
                ResetValues();
            }
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            ResetValues();
            btnHuy.Enabled = false;
            btnThem.Enabled = true;
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
            btnLuu.Enabled = false;
            txtmamay.Enabled = false;
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
