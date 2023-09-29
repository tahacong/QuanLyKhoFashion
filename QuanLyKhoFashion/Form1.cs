//Data Source=10.7.23.8;Initial Catalog=fashionwh;Persist Security Info=True;User ID=sa;Password=Tanhatminh@123

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Security.Cryptography;
using System.Threading;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using static OfficeOpenXml.ExcelErrorValue;
using QuanLyKhoFashion.fashionwhDataSetTableAdapters;

namespace QuanLyKhoFashion
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //Khai báo biến toàn phần
        List<string> listItemFontVitri = new List<string>();
        //Tự động điểu chỉnh độ rộng cột dgv
        private void AdjustColumnWidths(DataGridView dgv)
        {
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; // Đặt auto size cho tất cả các cột
            }
        }
        private void CauHinhBanDau()
        {
            //Tạo cột cho dgvLayHang
            dgvLayHang.Columns.Add("vitri_layhang", "Vị Trí");
            dgvLayHang.Columns.Add("mahang_layhang", "Mã Hàng");
            dgvLayHang.Columns.Add("ten_layhang", "Tên Hàng");
            dgvLayHang.Columns.Add("bienthe_layhang", "Biến Thể");
            dgvLayHang.Columns.Add("soluong_layhang", "Số Lượng");

            //Tạo cột cho dgvHangThieu
            dgvHangthieu.Columns.Add("mahang_thieu", "Mã Hàng");
            dgvHangthieu.Columns.Add("ten_layhang", "Tên Hàng");
            dgvHangthieu.Columns.Add("bienthe_thieu", "Biến Thể");
            dgvHangthieu.Columns.Add("soluong_thieu", "Số Lượng Thiếu");
        }
        //Cấu hình ẩn hiện các button trong tab 2 ( lấy hàng)
        private void HienNutChonListLayHang()
        {
            btnChonListLayHang_t2.Enabled = true;
            btnChonListLayHang_t2.BackColor=Color.Green;
        }
        private void AnNutChonListLayHang()
        {
            btnChonListLayHang_t2.Enabled = false;
            btnChonListLayHang_t2.BackColor = Color.WhiteSmoke;
        }
        private void HienNutXemTonKho()
        {
            btnXemTonKho_t2.Enabled = true;
            btnXemTonKho_t2.BackColor = Color.Green;
        }
        private void AnNutXemTonKho()
        {
            btnXemTonKho_t2.Enabled = false;
            btnXemTonKho_t2.BackColor = Color.WhiteSmoke;
        }
        private void HienNutTinhToan()
        {
            btnTinhToan_t2.Enabled = true;
            btnTinhToan_t2.BackColor = Color.Green;
        }
        private void AnNutTinhToan()
        {
            btnTinhToan_t2.Enabled = false;
            btnTinhToan_t2.BackColor = Color.WhiteSmoke;
        }
        private void HienNutXacNhan()
        {
            btnXacNhanXuatHang_t2.Enabled = true;
            btnXacNhanXuatHang_t2.BackColor = Color.Green;
        }
        private void AnNutXacNhan()
        {
            btnXacNhanXuatHang_t2.Enabled = false;
            btnXacNhanXuatHang_t2.BackColor = Color.WhiteSmoke;
        }
        private void HienNutLayHang()
        {
            btnLayHang_t2.Enabled = true;
            btnLayHang_t2.BackColor = Color.Green;
        }
        private void AnNutLayhang()
        {
            btnLayHang_t2.Enabled = false;
            btnLayHang_t2.BackColor = Color.WhiteSmoke;
        }
        //Cấu hình ẩn hiện các button trong tab 3 ( Xử lý file từ sàn)
        private void HienNutChonFileBill_Tiktok()
        {
            btnChonFileBill_t3.Enabled = true;
            btnChonFileBill_t3.BackColor = Color.Green;
        }
        private void AnNutChonFileBill_Tiktok()
        {
            btnChonFileBill_t3.Enabled = false;
            btnChonFileBill_t3.BackColor = Color.WhiteSmoke;
        }
        private void HienNutNhapVaoHeThong_Tiktok()
        {
            btn_NhapvaohethongTiktok_t3.Enabled = true;
            btn_NhapvaohethongTiktok_t3.BackColor = Color.Green;
        }
        private void AnNutNhapVaoHeThong_Tiktok()
        {
            btn_NhapvaohethongTiktok_t3.Enabled = false;
            btn_NhapvaohethongTiktok_t3.BackColor = Color.WhiteSmoke;
        }
        private void HienNutChonFileBill_Shopee()
        {
            btnChonFileBill_t3_Shopee.Enabled = true;
            btnChonFileBill_t3_Shopee.BackColor = Color.Green;
        }
        private void AnNutChonFileBill_Shopee()
        {
            btnChonFileBill_t3_Shopee.Enabled = false;
            btnChonFileBill_t3_Shopee.BackColor = Color.WhiteSmoke;
        }
        private void HienNutNhapVaoHeThong_Shopee()
        {
            btn_Nhapvaohethong_t3_Shopee.Enabled = true;
            btn_Nhapvaohethong_t3_Shopee.BackColor = Color.Green;
        }
        private void AnNutNhapVaoHeThong_Shopee()
        {
            btn_Nhapvaohethong_t3_Shopee.Enabled = false;
            btn_Nhapvaohethong_t3_Shopee.BackColor = Color.WhiteSmoke;
        }

        /// <summary>
        /// PHẦN NHẬP HÀNG T1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            //Ẩn nút trong tab2
            AnNutChonListLayHang();
            AnNutXemTonKho();
            AnNutTinhToan();
            AnNutXacNhan();

            //Ẩn nút trong tab3
            AnNutNhapVaoHeThong_Tiktok();
            AnNutNhapVaoHeThong_Shopee();

            CauHinhBanDau();
            TruyvanSQL.HienThiLenCBO("mavitri", cboVitri_t1, "bangvitri");

            XemTonKho(dgvTonKho);
            AdjustColumnWidths(dgvTonKho);
            XemTonKho(dgvTonKho_t2);
            AdjustColumnWidths(dgvTonKho_t2);
            //Tab 4 cài đặt
            LoadVitri();
            LoadMatHang();
        }
        private void btnNhaphang_t1_Click(object sender, EventArgs e)
        {
            //TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("bangvitri", "mavitri",cboVitri_t1.Text)
            if (txtMaSKU_t1.Text != "" && txtSoLuong_t1.Text != "" && cboVitri_t1.Text != "")
            {
                if (TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("bangvitri", "mavitri", cboVitri_t1.Text))
                {
                    if (TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("mathang", "mahang", txtMaSKU_t1.Text))
                    {
                        TruyvanSQL.NhapHang(cboVitri_t1.Text, txtMaSKU_t1.Text, Convert.ToInt32(txtSoLuong_t1.Text), Convert.ToInt32(txtGiaNhap_t1.Text));
                        MessageBox.Show($"Nhập thành công {txtSoLuong_t1.Text} {txtTenHang_t1.Text} (mã {txtMaSKU_t1.Text}) -  {txtBienThe_t1.Text} vào vị trí {cboVitri_t1.Text}");
                        ThayDoiTrangThaiStatus(true);
                    }
                    else { MessageBox.Show("Mã hàng SKU không tồn tại"); ThayDoiTrangThaiStatus(false); }
                }
                else { MessageBox.Show("Vị trí không tồn tại"); ThayDoiTrangThaiStatus(false); }
            }
            else { MessageBox.Show("Vui lòng điền đầy đủ thông tin các ô: Mã SKU, Biến thể, Số lượng, Vị Trí"); ThayDoiTrangThaiStatus(false); }
        }
        private void cboVitri_t1_TextChanged(object sender, EventArgs e)
        {
            lblSoluongTaiViTri.Text = Convert.ToString(TruyvanSQL.HangTonKho("vitri='" + cboVitri_t1.Text + "'"));
            TruyvanSQL.LoadDataDGV(dgvVitri, $"vitri = '{cboVitri_t1.Text}'");
            AdjustColumnWidths(dgvVitri);
        }
        private void ThayDoiTrangThaiStatus(bool trueorfalse)
        {
            if (trueorfalse)
            {
                lblStt.Text = $"Nhập thành công {txtSoLuong_t1.Text} - mã {txtMaSKU_t1.Text} - vị trí {cboVitri_t1.Text} ";
                lblStt.ForeColor = Color.Green;
            }else
            {
                lblStt.Text = "Nhập không thành công, Vui lòng kiểm tra lại";
                lblStt.ForeColor = Color.Red;
            }    
        }
        private void txtMaSKU_t1_Leave(object sender, EventArgs e)
        {
            //Hiển thị ảnh minh họa
            picMau.Image = null;
            ThaoTac.LoadImageFromURL(picMau, TruyvanSQL.LoadSqltoString($"SELECT [linkanh] FROM [mathang] WHERE mahang='{txtMaSKU_t1.Text}'"));
            //Hiển thị tên sản phẩm
            txtTenHang_t1.Text = TruyvanSQL.LoadSqltoString($"SELECT [tenhang] FROM [mathang] WHERE mahang='{txtMaSKU_t1.Text}'");
            //Hiển thị Biến thể
            txtBienThe_t1.Text = TruyvanSQL.LoadSqltoString($"SELECT [bienthe] FROM [mathang] WHERE mahang='{txtMaSKU_t1.Text}'");
            //truy vấn hiển thị dữ liệu lên dgvHang
            TruyvanSQL.LoadDataDGV(dgvHang, $"lichsu.mahang='{txtMaSKU_t1.Text}'");
            AdjustColumnWidths(dgvHang);
            //truy vấn hiển thị các vị trí đang chứa cùng loại hàng
            string str = "";
            List<string> listVitri = TruyvanSQL.LoadSQLtoList_LocTrung("SELECT lichsu.vitri FROM lichsu WHERE lichsu.mahang='" + txtMaSKU_t1.Text + "'");
            foreach (string vitri in listVitri)
            {
                str = str + " | " + vitri;
            }
            txtVitrihientai_t1.Text = str;
            listItemFontVitri = listVitri;
            //Sắp xếp lại combobox, đẩy các vị trí đang có hàng cùng loại lên đầu
            ThaoTac.SapXepComboBox(cboVitri_t1, listVitri);
            //Hiển thị số lượng mã hàng đó đang có trong kho
            lblSLHangCungMa.Text = Convert.ToString(TruyvanSQL.HangTonKho("mahang='" + txtMaSKU_t1.Text + "'"));
        }
        private void cboVitri_t1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index >= 0)
            {
                // Lấy phần tử hiện tại từ ComboBox
                var item = cboVitri_t1.Items[e.Index].ToString();

                // Kiểm tra xem phần tử có tồn tại trong danh sách hay không
                if (listItemFontVitri.Contains(item))
                {
                    // Đặt font chữ khác cho phần tử nếu nó tồn tại trong danh sách
                    e.Graphics.DrawString(item, new Font(cboVitri_t1.Font, FontStyle.Bold), Brushes.Red, e.Bounds);
                }
                else
                {
                    e.DrawBackground();
                    e.Graphics.DrawString(item, cboVitri_t1.Font, Brushes.Black, e.Bounds);
                }
            }
        }
        private void btnXemTonKho_t1_Click(object sender, EventArgs e)//Xem tồn kho
        {
            XemTonKho(dgvTonKho);
        }
        private void XemTonKho(DataGridView dgv)
        {
            SqlConnection Ketnoi = TruyvanSQL.TaoKetNoi();
            try
            {
                Ketnoi.Open();
                //string command = $"SELECT lichsu.mahang,mathang.tenhang,lichsu.size,lichsu.mamau,SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE -soluong END) AS soluong_tonkho,lichsu.vitri FROM lichsu LEFT JOIN mathang ON lichsu.mahang=mathang.mahang  GROUP BY lichsu.mahang,mathang.tenhang,lichsu.size,lichsu.vitri,lichsu.mamau";
                /*
                string command = $"SELECT lichsu.mahang, mathang.tenhang, lichsu.bienthe, " +
                 $"SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE -soluong END) AS soluong_tonkho, " +
                 $"lichsu.vitri " +
                 $"FROM lichsu " +
                 $"LEFT JOIN mathang ON lichsu.mahang = mathang.mahang " +
                 $"GROUP BY lichsu.mahang, mathang.tenhang, lichsu.bienthe, lichsu.vitri" +
                 $"HAVING SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE -soluong END) <> 0"; // Thêm điều kiện HAVING
                */
                string command = "SELECT lichsu.mahang, mathang.tenhang, mathang.bienthe, " +
                    "SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE -soluong END) AS soluong_tonkho, " +
                    "lichsu.vitri " +
                    "FROM lichsu " +
                    "LEFT JOIN mathang ON lichsu.mahang = mathang.mahang " +
                    "GROUP BY lichsu.mahang, mathang.tenhang, mathang.bienthe, lichsu.vitri " +
                    "HAVING SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE -soluong END) <> 0";
                SqlCommand cmd = new SqlCommand(command, Ketnoi);
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                //tai du lieu len girdView
                dgv.DataSource = dt;
                List<string> tencot = new List<string> { "vitri", "mahang", "tenhang", "bienthe", "soluong_tonkho" };
                List<string> newHeader = new List<string> { "Vị Trí", "SKU", "Tên", "Biến Thể", "Số lượng" };
                ThaoTac.DoiTenCot(dgv, tencot, newHeader);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
            Ketnoi.Close();
            AdjustColumnWidths(dgvTonKho);
        }
        private void dgvTonKho_SelectionChanged(object sender, EventArgs e)
        {
            // List<string> tencot = new List<string> { "vitri", "mahang", "tenhang", "bienthe", "soluong_tonkho" };
            if (dgvTonKho.SelectedRows.Count > 0) // Kiểm tra xem có dòng được chọn không
            {
                DataGridViewRow selectedRow = dgvTonKho.SelectedRows[0]; // Lấy dòng đầu tiên trong các dòng được chọn
                // Gán dữ liệu từ các ô trong dòng được chọn vào các TextBox tương ứng
                txtMaSKU_t1.Text = selectedRow.Cells["mahang"].Value.ToString();
                txtMaSKU_t1.TextChanged += txtMaSKU_t1_Leave;
            }
        }
        /// <summary>
        /// PHẦN XUẤT HÀNG T2
        /// </summary> 
        private void LoadFileLayHang ()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog.Title = "Select an Excel File";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (var package = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy worksheet đầu tiên

                    List<HangHoa> hangHoas = new List<HangHoa>();

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        HangHoa hangHoa = new HangHoa
                        {
                            MaHang = worksheet.Cells[row, 1].Value.ToString(),
                            TenHang = worksheet.Cells[row, 2].Value.ToString(),
                            BienThe = worksheet.Cells[row, 3].Value.ToString(),
                            SoLuong = int.Parse(worksheet.Cells[row, 4].Value.ToString()),
                            VanDon = worksheet.Cells[row, 5].Value.ToString()
                        };
                        hangHoas.Add(hangHoa);
                    }

                    // Tính tổng theo mã hàng và lưu vào danh sách HangHoaTong
                    var tongSoLuongTheoMaHang = hangHoas.GroupBy(x => x.MaHang)
                        .Select(group => new HangHoaTong
                        {
                            MaHang = group.Key,
                            TenHang = group.First().TenHang,
                            BienThe = group.First().BienThe,
                            TongSoLuong = group.Sum(item => item.SoLuong)
                        }).ToList();

                    // Nạp dữ liệu vào DataGridView
                    dgvListCanLay_t2.DataSource = tongSoLuongTheoMaHang;
                }
                AnNutChonListLayHang();
                HienNutXemTonKho();
            }
        }
        private void TinhToanSoLuongLayHang(DataGridView dgvCanLay, DataGridView dgvTonKho)
        {
            foreach (DataGridViewRow row in dgvCanLay.Rows)
            {
                string maHang = row.Cells["Mahang"].Value?.ToString();
                string bienthe = row.Cells["BienThe"].Value?.ToString();
                // Kiểm tra nếu maHang hoặc bienthe là null thì bỏ qua vòng lặp
                if (maHang == null)
                {
                    continue;
                }

                int soLuongCanLay = Convert.ToInt32(row.Cells["TongSoLuong"].Value);

                /*
                var availableRows = dgvTonKho.Rows.Cast<DataGridViewRow>()
                    .Where(r => r.Cells["mahang"].Value?.ToString() == maHang && r.Cells["bienthe"].Value?.ToString() == bienthe )
                    .OrderBy(r => Convert.ToInt32(r.Cells["soluong_tonkho"].Value))
                    .ToList();
                */
                var availableRows = dgvTonKho.Rows.Cast<DataGridViewRow>()
                    .Where(r => r.Cells["mahang"].Value?.ToString() == maHang)
                    .OrderBy(r => Convert.ToInt32(r.Cells["soluong_tonkho"].Value))
                    .ToList();

                int remainingQuantity = soLuongCanLay;

                foreach (DataGridViewRow availRow in availableRows)
                {
                    int soluongTonKho = Convert.ToInt32(availRow.Cells["soluong_tonkho"].Value);
                    if (remainingQuantity > 0 && soluongTonKho > 0)
                    {
                        int takenQuantity = Math.Min(remainingQuantity, soluongTonKho);
                        availRow.Cells["soluong_tonkho"].Value = soluongTonKho - takenQuantity;
                        remainingQuantity -= takenQuantity;
                        string tenHang = availRow.Cells["tenhang"].Value.ToString();
                        
                        dgvLayHang.Rows.Add(availRow.Cells["vitri"].Value?.ToString(), maHang,tenHang, bienthe, takenQuantity);
                    }
                }
                if(remainingQuantity>0)
                {
                    dgvHangthieu.Rows.Add(maHang, TruyvanSQL.LayGiaTriText("mathang", "tenhang", "mahang", maHang), bienthe,  remainingQuantity);
                }    
            }
        }
        private void btnLayHang_t2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Xác nhận việc bắt đầu lấy hàng?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                HienNutChonListLayHang();
                AnNutLayhang();
                dgvLayHang.DataSource = null;
                dgvListCanLay_t2.DataSource = null;
                dgvHangthieu.DataSource = null;
            }
        }
        private void btnXemTonKho_t2_Click(object sender, EventArgs e)
        {
            XemTonKho(dgvTonKho_t2);
            AdjustColumnWidths(dgvTonKho_t2);
            AnNutXemTonKho();
            HienNutTinhToan();
        }
        private void btnChonListLayHang_t2_Click(object sender, EventArgs e)
        {
            LoadFileLayHang();
            AdjustColumnWidths(dgvListCanLay_t2);
        }
        private void btnTinhToan_t2_Click(object sender, EventArgs e)
        {
            dgvLayHang.Rows.Clear();
            dgvHangthieu.Rows.Clear();
            TinhToanSoLuongLayHang(dgvListCanLay_t2, dgvTonKho_t2);
            AdjustColumnWidths(dgvLayHang);
            AdjustColumnWidths(dgvHangthieu);
            dgvLayHang.Sort(dgvLayHang.Columns["vitri_layhang"], ListSortDirection.Ascending);
            dgvHangthieu.Sort(dgvHangthieu.Columns["mahang_thieu"], ListSortDirection.Ascending);
            //dgvLayHang.Sort(dgvLayHang.Columns["mahang_layhang"], ListSortDirection.Ascending);
            AnNutTinhToan();
            HienNutXacNhan();
        }
        private void btnXacNhanXuatHang_t2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có muốn xác nhận việc xuất hàng?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                TruyvanSQL.UpdateLayHangtoLichSu(dgvLayHang);
                ThaoTac.XuatExceltuDGB(dgvLayHang, "LayHang");
                ThaoTac.XuatExceltuDGBKhongMo(dgvHangthieu, "HangThieu");
                AnNutXacNhan();
                HienNutLayHang();
            }
            
            
        }
        private void btnRestart_t2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có muốn xác nhận việc nhập hàng lại từ đầu?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                HienNutLayHang();
                AnNutChonListLayHang();
                AnNutXemTonKho();
                AnNutTinhToan();
                AnNutXacNhan();
                dgvLayHang.DataSource = null;
                dgvListCanLay_t2.DataSource= null;
                dgvHangthieu.DataSource = null;
                
            }
        }
        /// <summary>
        /// Tab 4 cài đặt
        /// </summary>
        private void LoadVitri()
        {
            TruyvanSQL.SQLtoDGV(dgvBangVitri_t4, "SELECT * FROM bangvitri");
            List<string> tencot = new List<string> { "mavitri", "chitiet" };
            List<string> newHeader = new List<string> { "Mã Vị Trí", "Chi Tiết" };
            ThaoTac.DoiTenCot(dgvBangVitri_t4, tencot, newHeader);
            AdjustColumnWidths(dgvBangVitri_t4);
        }
        private void LoadMatHang()
        {
            TruyvanSQL.SQLtoDGV(dgvMatHang_t4, "SELECT * FROM mathang");
            List<string> tencot = new List<string> { "mahang", "tenhang","bienthe","linkanh" };
            List<string> newHeader = new List<string> { "Mã Hàng", "Tên Hàng", "Biến Thể", "Link Ảnh" };
            ThaoTac.DoiTenCot(dgvMatHang_t4, tencot, newHeader);
            AdjustColumnWidths(dgvMatHang_t4);
        }
        private void LoadFileMatHang(string excelFilePath, DataGridView dgv)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy worksheet đầu tiên

                    DataTable dt = new DataTable();
                    int totalCols = worksheet.Dimension.End.Column;

                    // Tạo các cột cho DataTable dựa trên tên cột trong Excel
                    for (int col = 1; col <= totalCols; col++)
                    {
                        dt.Columns.Add(worksheet.Cells[1, col].Value.ToString());
                    }

                    // Đọc dữ liệu từ Excel và thêm vào DataTable
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        DataRow dataRow = dt.NewRow();
                        for (int col = 1; col <= totalCols; col++)
                        {
                            dataRow[col - 1] = worksheet.Cells[row, col].Value;
                        }
                        dt.Rows.Add(dataRow);
                    }

                    // Nạp DataTable vào DataGridView
                    dgv.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        private void dgvBangVitri_t4_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvBangVitri_t4.SelectedRows.Count > 0) // Kiểm tra xem có dòng được chọn không
            {
                DataGridViewRow selectedRow = dgvBangVitri_t4.SelectedRows[0]; // Lấy dòng đầu tiên trong các dòng được chọn
                // Gán dữ liệu từ các ô trong dòng được chọn vào các TextBox tương ứng
                txtMaViTri_t4.Text = selectedRow.Cells["mavitri"].Value.ToString();
                txtChiTietViTri_t4.Text = selectedRow.Cells["chitiet"].Value.ToString();
            }
        }
        private void btnXoa_t4_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd= new SqlCommand($"DELETE FROM [bangvitri] WHERE [mavitri]='{txtMaViTri_t4.Text}'");
                TruyvanSQL.ThemSuaXoa(cmd);
                MessageBox.Show($"Xóa thành công mã vị trí {txtMaViTri_t4.Text}");
            }catch(Exception)
            {
                MessageBox.Show($"Lỗi chưa xóa được mã vị trí :{txtMaViTri_t4.Text}");
            }
            LoadVitri();
        }
        private void btnThemSua_t4_Click(object sender, EventArgs e)
        {
            if(TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("bangvitri","mavitri",txtMaViTri_t4.Text))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("UPDATE bangvitri SET [chitiet]=@chitiet WHERE [mavitri]=@mavitri");
                    cmd.Parameters.AddWithValue("@chitiet", txtChiTietViTri_t4.Text);
                    cmd.Parameters.AddWithValue("@mavitri", txtMaViTri_t4.Text);
                    TruyvanSQL.ThemSuaXoa(cmd);
                    MessageBox.Show($"Sửa thành công mã vị trí {txtMaViTri_t4.Text}");
                }
                catch (Exception)
                {
                    MessageBox.Show($"Lỗi chưa sửa được mã vị trí :{txtMaViTri_t4.Text}");
                }
            }else
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO bangvitri (mavitri, chitiet) VALUES (@mavitri, @chitiet)");
                    cmd.Parameters.AddWithValue("@chitiet", txtChiTietViTri_t4.Text);
                    cmd.Parameters.AddWithValue("@mavitri", txtMaViTri_t4.Text);
                    TruyvanSQL.ThemSuaXoa(cmd);
                    MessageBox.Show($"Thêm thành công mã vị trí {txtMaViTri_t4.Text}");
                }
                catch (Exception)
                {
                    //MessageBox.Show(ex.ToString());
                    MessageBox.Show($"Lỗi chưa thêm được mã vị trí :{txtMaViTri_t4.Text}");
                }
            }
            LoadVitri();
        }
        private void dgvMatHang_t4_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvMatHang_t4.SelectedRows.Count > 0) // Kiểm tra xem có dòng được chọn không
            {
                DataGridViewRow selectedRow = dgvMatHang_t4.SelectedRows[0]; // Lấy dòng đầu tiên trong các dòng được chọn
                // Gán dữ liệu từ các ô trong dòng được chọn vào các TextBox tương ứng
                txtMaHang_t4.Text = selectedRow.Cells["mahang"].Value.ToString();
                txtTenHang_t4.Text = selectedRow.Cells["tenhang"].Value.ToString();
                txtBienThe_t4.Text = selectedRow.Cells["bienthe"].Value.ToString();
                txtLinkAnh_t4.Text = selectedRow.Cells["linkanh"].Value.ToString();

                //Hiển thị ảnh minh họa
                picAnh_Img_t4.Image = null;
                ThaoTac.LoadImageFromURL(picAnh_Img_t4, selectedRow.Cells["linkanh"].Value.ToString());
            }
        }
        private void btnXoaMaHang_t4_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand($"DELETE FROM [mathang] WHERE [mahang]='{txtMaHang_t4.Text}'");
                TruyvanSQL.ThemSuaXoa(cmd);
                MessageBox.Show($"Xóa thành công mã màu {txtMaHang_t4.Text}");
            }
            catch (Exception)
            {
                MessageBox.Show($"Lỗi chưa xóa được mã hàng :{txtMaHang_t4.Text}");
            }
            LoadMatHang();
        }
        private void btn_ThemSuaMaHang_t4_Click(object sender, EventArgs e)
        {
            if (TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("mathang", "mahang", txtMaHang_t4.Text))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("UPDATE mathang SET [tenhang]=@tenhang,[bienthe]=@bienthe,[linkanh]=@linkanh WHERE [mahang]=@mahang");
                    cmd.Parameters.AddWithValue("@tenhang", txtTenHang_t4.Text);
                    cmd.Parameters.AddWithValue("@mahang", txtMaHang_t4.Text);
                    cmd.Parameters.AddWithValue("bienthe", txtBienThe_t4.Text);
                    cmd.Parameters.AddWithValue("linkanh", txtLinkAnh_t4.Text);
                    TruyvanSQL.ThemSuaXoa(cmd);
                    MessageBox.Show($"Sửa thành công mã hàng {txtMaHang_t4.Text}");
                }
                catch (Exception)
                {
                    MessageBox.Show($"Lỗi chưa sửa được mã hàng :{txtMaHang_t4.Text}");
                }
            }
            else
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO mathang (mahang, tenhang, bienthe, linkanh) VALUES (@mahang, @tenhang, @bienthe, @linkanh)");
                    cmd.Parameters.AddWithValue("@tenhang", txtTenHang_t4.Text);
                    cmd.Parameters.AddWithValue("@mahang", txtMaHang_t4.Text);
                    cmd.Parameters.AddWithValue("bienthe", txtBienThe_t4.Text);
                    cmd.Parameters.AddWithValue("linkanh",txtLinkAnh_t4.Text );
                    TruyvanSQL.ThemSuaXoa(cmd);
                    MessageBox.Show($"Thêm thành công mã hàng {txtMaHang_t4.Text}");
                }
                catch (Exception )
                {
                    MessageBox.Show($"Lỗi chưa thêm được mã hàng :{txtMaViTri_t4.Text}");
                }
            }
            LoadMatHang();
        }
        private void btn_CSHL_LoadFile_t4_Click(object sender, EventArgs e)
        {
            // Chọn tệp Excel để nạp dữ liệu
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFilePath = openFileDialog.FileName;
                LoadFileMatHang(excelFilePath,dgv_ChinhsuaMatHang_HangLoat_t4);
            }
        }
        private void txtLinkAnh_t4_Leave(object sender, EventArgs e)
        {
            picAnh_Img_t4.Image = null;
            ThaoTac.LoadImageFromURL(picAnh_Img_t4, txtLinkAnh_t4.Text);
        }
        private void btn_CSHL_ThemSua_t4_Click(object sender, EventArgs e)
        {
            TruyvanSQL.UpdateMatHangtoBangMatHang(dgv_ChinhsuaMatHang_HangLoat_t4);
        }
        /// <summary>
        /// Tab 3 Xử lý bill
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //Nhập đơn tiktok
        private void btnChonFileBill_t3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Load data from Excel file
                using (var package = new ExcelPackage(new System.IO.FileInfo(openFileDialog.FileName)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Tạo bảng và add columns
                    DataTable dt = new DataTable();
                    // Add column title từ file excel
                    
                    foreach (int columnIndex in new int[] { 7, 8, 9, 10, 34, 36, 37 })
                    {
                        dt.Columns.Add(worksheet.Cells[1, columnIndex].Text, typeof(string));
                    }
                    dt.Columns.Add("Noiban");
                    // Loop through rows in the Excel worksheet and add data to DataTable
                    for (int row = worksheet.Dimension.Start.Row + 2; row <= worksheet.Dimension.End.Row; row++)
                    {

                        DataRow newRow = dt.NewRow();

                        foreach (int columnIndex in new int[] { 7, 8, 9, 10, 34, 36, 37 })
                        {
                            string cellValue = worksheet.Cells[row, columnIndex].Text;

                            // Check if the cell has data
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                newRow[worksheet.Cells[1, columnIndex].Text] = cellValue;
                            }
                            else
                            {
                                //newRow[worksheet.Cells[1, columnIndex].Text] = "N/A"; // Điền N/A nếu dữ liệu trong ô tính excel là trống.
                            }
                        }
                        newRow["Noiban"] = "Tiktok";
                        dt.Rows.Add(newRow);
                    }
                    // Gán DataTable to DataGridView
                    dgvXuLyBill_t3.DataSource = dt;
                }
                HienNutNhapVaoHeThong_Tiktok();
                AnNutChonFileBill_Tiktok();
            }
            AdjustColumnWidths(dgvXuLyBill_t3);  
        }
        private void btn_NhapvaohethongTiktok_t3_Click(object sender, EventArgs e)
        {
            TruyvanSQL.UpdateLayHangtoDongGoi(dgvXuLyBill_t3);
            ThaoTac.XuatExcelTiktok(dgvXuLyBill_t3, "Tiktok");
            AnNutNhapVaoHeThong_Tiktok();
            HienNutChonFileBill_Tiktok();
            dgvXuLyBill_t3.DataSource = null;
        }
        //Nhập đơn shoppe
        private void btnChonFileBill_t3_Shopee_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Load data from Excel file
                using (var package = new ExcelPackage(new System.IO.FileInfo(openFileDialog.FileName)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Tạo bảng và add columns
                    DataTable dt = new DataTable();
                    // Add column title từ file excel

                    foreach (int columnIndex in new int[] { 19, 16, 20, 26, 6, 7, 5 })
                    {
                        dt.Columns.Add(worksheet.Cells[1, columnIndex].Text, typeof(string));
                    }
                    dt.Columns.Add("Noiban");
                    // Loop through rows in the Excel worksheet and add data to DataTable
                    for (int row = worksheet.Dimension.Start.Row + 2; row <= worksheet.Dimension.End.Row; row++)
                    {

                        DataRow newRow = dt.NewRow();

                        foreach (int columnIndex in new int[] { 19, 16, 20, 26, 6, 7, 5 })
                        {
                            string cellValue = worksheet.Cells[row, columnIndex].Text;

                            // Check if the cell has data
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                newRow[worksheet.Cells[1, columnIndex].Text] = cellValue;
                            }
                            else
                            {
                                //newRow[worksheet.Cells[1, columnIndex].Text] = "N/A"; // Điền N/A nếu dữ liệu trong ô tính excel là trống.
                            }
                        }
                        newRow["Noiban"] = "Shoppe";
                        dt.Rows.Add(newRow);
                    }
                    //Đổi tên cột về đúng chuẩn tiktok
                    dt.Columns["SKU phân loại hàng"].ColumnName = "Seller SKU";
                    dt.Columns["Tên sản phẩm"].ColumnName = "Product Name";
                    dt.Columns["Tên phân loại hàng"].ColumnName = "Variation";
                    dt.Columns["Số lượng"].ColumnName = "Quantity";
                    dt.Columns["Mã vận đơn"].ColumnName = "Tracking ID";
                    dt.Columns["Đơn Vị Vận Chuyển"].ColumnName = "Shipping Provider Name";
                    dt.Columns["Nhận xét từ Người mua"].ColumnName = "Buyer Message";
                    // Gán DataTable to DataGridView
                    dgvXuLyBill_t3.DataSource = dt;
                }
                HienNutNhapVaoHeThong_Shopee();
                AnNutChonFileBill_Shopee();
            }
        }
        private void btn_Nhapvaohethong_t3_Shopee_Click(object sender, EventArgs e)
        {
            TruyvanSQL.UpdateLayHangtoDongGoi(dgvXuLyBill_t3);
            ThaoTac.XuatExcelTiktok(dgvXuLyBill_t3, "Shoppe");
            AnNutNhapVaoHeThong_Shopee();
            HienNutChonFileBill_Shopee();
            dgvXuLyBill_t3.DataSource = null;   
        }
        /// <summary>
        /// Tab 5 Nhập hàng loạt
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void btn_ChonFile_t5_Click(object sender, EventArgs e)
        {
            // Chọn tệp Excel để nạp dữ liệu
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFilePath = openFileDialog.FileName;
                LoadFileMatHang(excelFilePath, dgvNhapHangLoat_t5);
            }
        }
        private void btn_Nhap_t5_Click(object sender, EventArgs e)
        {
            TruyvanSQL.NhapHangLoat(dgvNhapHangLoat_t5);
            dgvNhapHangLoat_t5.DataSource = null;
        }

        private void dgvHang_MouseClick(object sender, MouseEventArgs e)
        {
            // List<string> tencot = new List<string> { "vitri", "mahang", "tenhang", "bienthe", "soluong_tonkho" };
            if (dgvHang.SelectedRows.Count > 0) // Kiểm tra xem có dòng được chọn không
            {
                DataGridViewRow selectedRow = dgvHang.SelectedRows[0]; // Lấy dòng đầu tiên trong các dòng được chọn
                // Gán dữ liệu từ các ô trong dòng được chọn vào các TextBox tương ứng
                txtMaSKU_t1.Text = selectedRow.Cells["mahang"].Value.ToString();
                txtMaSKU_t1.TextChanged += txtMaSKU_t1_Leave;
            }
        }

        private void dgvVitri_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // List<string> tencot = new List<string> { "vitri", "mahang", "tenhang", "bienthe", "soluong_tonkho" };
            if (dgvVitri.SelectedRows.Count > 0) // Kiểm tra xem có dòng được chọn không
            {
                DataGridViewRow selectedRow = dgvVitri.SelectedRows[0]; // Lấy dòng đầu tiên trong các dòng được chọn
                // Gán dữ liệu từ các ô trong dòng được chọn vào các TextBox tương ứng
                txtMaSKU_t1.Text = selectedRow.Cells["mahang"].Value.ToString();
                txtMaSKU_t1.TextChanged += txtMaSKU_t1_Leave;
            }
        }
    }
}
