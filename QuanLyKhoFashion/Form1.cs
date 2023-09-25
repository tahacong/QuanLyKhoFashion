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
        /// PHẦN NHẬP HÀNG
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
            TruyvanSQL.HienThiLenCBO("bienthe", cboBienThe_t1, "bienthe");

            XemTonKho(dgvTonKho);
            AdjustColumnWidths(dgvTonKho);
            XemTonKho(dgvTonKho_t2);
            AdjustColumnWidths(dgvTonKho_t2);
            //Tab 4 cài đặt
            LoadVitri();
            LoadMatHang();
            LoadLinkAnh();
            TruyvanSQL.HienThiLenCBO("mahang", cboAnh_Mahang_t4, "mathang");
            TruyvanSQL.HienThiLenCBO("bienthe", cboAnh_BienThe_t4, "bienthe");
        }
        private void btnNhaphang_t1_Click(object sender, EventArgs e)
        {
            //TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("bangvitri", "mavitri",cboVitri_t1.Text)
            if (txtMaSKU_t1.Text != "" && cboBienThe_t1.Text != "" && txtSoLuong_t1.Text != "" && cboVitri_t1.Text != "")
            {
                if (TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("bangvitri", "mavitri", cboVitri_t1.Text))
                {
                    if (TruyvanSQL.KiemTraTonTaiGiaTriTrongBang("mathang", "mahang", txtMaSKU_t1.Text))
                    {
                        if (TruyvanSQL.KiemTraTonTaiGiaTriTrongBangPlus("bienthe", "bienthe=N'" + cboBienThe_t1.Text + "' AND mahang='" + txtMaSKU_t1.Text + "'")) 
                        {
                            TruyvanSQL.NhapHang(cboVitri_t1.Text, txtMaSKU_t1.Text, cboBienThe_t1.Text, Convert.ToInt32(txtSoLuong_t1.Text));
                            MessageBox.Show($"Nhập thành công {txtSoLuong_t1.Text} {txtTenHang_t1.Text} (mã {txtMaSKU_t1.Text}) size {cboBienThe_t1.Text} vào vị trí {cboVitri_t1.Text}");
                            ThayDoiTrangThaiStatus(true);
                        }
                        else { MessageBox.Show("biến không có trong bảng"); ThayDoiTrangThaiStatus(false); }
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
        private void cboBienThe_t1_TextChanged(object sender, EventArgs e)
        {
            lblSLBienThe.Text = Convert.ToString(TruyvanSQL.HangTonKho("mahang='" + txtMaSKU_t1.Text + "' AND bienthe='" + cboBienThe_t1.Text + "'"));
            picMau.Image = null;
            ThaoTac.LoadImageFromURL(picMau, TruyvanSQL.LoadSqltoString($"SELECT [linkanh] FROM [bienthe] WHERE mahang='{txtMaSKU_t1.Text}' AND bienthe=N'{cboBienThe_t1.Text}'"));
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
            //Hiển thị các biến thể đang có của mã hàng
            TruyvanSQL.HienThiLenCBO_CoDieuKien("bienthe", cboBienThe_t1, "bienthe", "mahang='" + txtMaSKU_t1.Text + "'");
            //Hiển thị ảnh minh họa
            picMau.Image = null;
            ThaoTac.LoadImageFromURL(picMau, TruyvanSQL.LoadSqltoString($"SELECT [linkanh] FROM [bienthe] WHERE mahang='{txtMaSKU_t1.Text}' AND bienthe=N'{cboBienThe_t1.Text}'"));
            //Hiển thị tên sản phẩm
            txtTenHang_t1.Text = TruyvanSQL.LoadSqltoString($"SELECT [tenhang] FROM [mathang] WHERE mahang='{txtMaSKU_t1.Text}'");

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
                string command = "SELECT lichsu.mahang, mathang.tenhang, lichsu.bienthe, " +
                    "SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE -soluong END) AS soluong_tonkho, " +
                    "lichsu.vitri " +
                    "FROM lichsu " +
                    "LEFT JOIN mathang ON lichsu.mahang = mathang.mahang " +
                    "GROUP BY lichsu.mahang, mathang.tenhang, lichsu.bienthe, lichsu.vitri " +
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
        /// <summary>
        /// PHẦN XUẤT HÀNG
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
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    DataTable dt = new DataTable();
                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        dt.Columns.Add(firstRowCell.Text);
                    }

                    for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                    {
                        var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                        var newRow = dt.NewRow();
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column - 1] = cell.Text;
                        }
                        dt.Rows.Add(newRow);
                    }

                    // Tính tổng theo mã và size và hiển thị kết quả lên DataGridView
                    var groupedData = dt.AsEnumerable()
                                        .GroupBy(row => new { MaHang = row.Field<string>("mahang"), BienThe = row.Field<string>("bienthe") })
                                        .Select(group => new
                                        {
                                            MaHang = group.Key.MaHang,
                                            BienThe = group.Key.BienThe,
                                            TongSoLuong = group.Sum(row => Convert.ToInt32(row["soluong"]))
                                        })
                                        .ToList();
                    // Sắp xếp theo mã hàng và size theo thứ tự ưu tiên
                    var sortedData = groupedData.OrderBy(item => item.MaHang)
                                                .ToList();

                    dgvListCanLay_t2.DataSource = sortedData;
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
                if (maHang == null || bienthe == null)
                {
                    continue;
                }

                int soLuongCanLay = Convert.ToInt32(row.Cells["TongSoLuong"].Value);

                var availableRows = dgvTonKho.Rows.Cast<DataGridViewRow>()
                    .Where(r => r.Cells["mahang"].Value?.ToString() == maHang && r.Cells["bienthe"].Value?.ToString() == bienthe )
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
            DialogResult result = MessageBox.Show("Bạn có muốn xác nhận việc nhập hàng?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

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
            List<string> tencot = new List<string> { "mahang", "tenhang" };
            List<string> newHeader = new List<string> { "Mã Hàng", "Tên Hàng" };
            ThaoTac.DoiTenCot(dgvMatHang_t4, tencot, newHeader);
            AdjustColumnWidths(dgvMatHang_t4);
        }
        private void LoadLinkAnh()
        {
            TruyvanSQL.SQLtoDGV(dgvAnh_BienThe_t4, "SELECT bienthe.mahang,mathang.tenhang,bienthe.bienthe,bienthe.linkanh FROM bienthe LEFT JOIN mathang ON bienthe.mahang=mathang.mahang");
            List<string> tencot = new List<string> { "mahang", "tenhang", "bienthe", "linkanh" };
            List<string> newHeader = new List<string> { "Mã Hàng", "Tên Hàng","Biến Thể","Link Ảnh" };
            ThaoTac.DoiTenCot(dgvAnh_BienThe_t4, tencot, newHeader);
            AdjustColumnWidths(dgvAnh_BienThe_t4);
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
                    SqlCommand cmd = new SqlCommand("UPDATE mathang SET [tenhang]=@tenhang WHERE [mahang]=@mahang");
                    cmd.Parameters.AddWithValue("@tenhang", txtTenHang_t4.Text);
                    cmd.Parameters.AddWithValue("@mahang", txtMaHang_t4.Text);
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
                    SqlCommand cmd = new SqlCommand("INSERT INTO mathang (mahang, tenhang) VALUES (@mahang, @tenhang)");
                    cmd.Parameters.AddWithValue("@tenhang", txtTenHang_t4.Text);
                    cmd.Parameters.AddWithValue("@mahang", txtMaHang_t4.Text);
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
        private void dgvAnh_Link_t4_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvAnh_BienThe_t4.SelectedRows.Count > 0) // Kiểm tra xem có dòng được chọn không
            {
                DataGridViewRow selectedRow = dgvAnh_BienThe_t4.SelectedRows[0]; // Lấy dòng đầu tiên trong các dòng được chọn
                // Gán dữ liệu từ các ô trong dòng được chọn vào các TextBox tương ứng
                cboAnh_Mahang_t4.Text = selectedRow.Cells["mahang"].Value.ToString();
                cboAnh_BienThe_t4.Text = selectedRow.Cells["bienthe"].Value.ToString();
                txtAnh_Link_t4.Text = selectedRow.Cells["linkanh"].Value.ToString();
                txtAnh_tenHang_t4.Text= selectedRow.Cells["tenhang"].Value.ToString();
                //Hiển thị ảnh minh họa
                picAnh_Img_t4.Image = null;
                ThaoTac.LoadImageFromURL(picAnh_Img_t4, selectedRow.Cells["linkanh"].Value.ToString());
            }
        }
        private void btnAnh_ThemSua_t4_Click(object sender, EventArgs e)
        {
            if (TruyvanSQL.KiemTraTonTaiGiaTriTrongBangN("bienthe", "mahang", cboAnh_Mahang_t4.Text) && TruyvanSQL.KiemTraTonTaiGiaTriTrongBangN("bienthe", "bienthe", cboAnh_BienThe_t4.Text))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("UPDATE bienthe SET [linkanh]=@linkanh WHERE [mahang]=@mahang AND [bienthe]=@bienthe");
                    cmd.Parameters.AddWithValue("@linkanh", txtAnh_Link_t4.Text);
                    cmd.Parameters.AddWithValue("@mahang", cboAnh_Mahang_t4.Text);
                    cmd.Parameters.AddWithValue("@bienthe", cboAnh_BienThe_t4.Text);
                    TruyvanSQL.ThemSuaXoa(cmd);
                    MessageBox.Show($"Sửa thành công biến thể {cboAnh_Mahang_t4.Text} {cboAnh_BienThe_t4.Text}");
                }
                catch (Exception)
                {
                    MessageBox.Show($"Lỗi chưa sửa được biến thể {cboAnh_Mahang_t4.Text} {cboAnh_BienThe_t4.Text}");
                }
            }
            else
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO bienthe (mahang,bienthe, linkanh) VALUES (@mahang,@bienthe, @linkanh)");
                    cmd.Parameters.AddWithValue("@mahang", cboAnh_Mahang_t4.Text);
                    cmd.Parameters.AddWithValue("@bienthe", cboAnh_BienThe_t4.Text);
                    cmd.Parameters.AddWithValue("@linkanh", txtAnh_Link_t4.Text);
                    TruyvanSQL.ThemSuaXoa(cmd);
                    MessageBox.Show($"Thêm thành công bien thể:  {cboAnh_Mahang_t4.Text} {cboAnh_BienThe_t4.Text}");
                }
                catch (Exception )
                {
                    MessageBox.Show($"Lỗi chưa thêm được biến thể {cboAnh_Mahang_t4.Text} {cboAnh_BienThe_t4.Text}");
                }
            }
            LoadLinkAnh();
        }
        private void btnAnh_Xoa_t4_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand($"DELETE FROM [bienthe] WHERE [mahang]='{cboAnh_Mahang_t4.Text}' AND [bienthe]=N'{cboAnh_BienThe_t4.Text}'");
                TruyvanSQL.ThemSuaXoa(cmd);
                MessageBox.Show($"Xóa thành công biến thể: {cboAnh_Mahang_t4.Text} {cboAnh_BienThe_t4.Text}");
            }
            catch (Exception)
            {
                MessageBox.Show($"Lỗi chưa xóa được biến thể: {cboAnh_Mahang_t4.Text} {cboAnh_BienThe_t4.Text}");
            }
            LoadLinkAnh();
        }
        private void txtAnh_Link_t4_Leave(object sender, EventArgs e)
        {
            picAnh_Img_t4.Image = null;
            ThaoTac.LoadImageFromURL(picAnh_Img_t4, txtAnh_Link_t4.Text);
        }
        /// <summary>
        /// Tab 3 Xử lý bill
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                    
                    foreach (int columnIndex in new int[] { 7, 8, 9, 10, 34 })
                    {
                        dt.Columns.Add(worksheet.Cells[1, columnIndex].Text, typeof(string));
                    }
                    dt.Columns.Add("Noiban");
                    // Loop through rows in the Excel worksheet and add data to DataTable
                    for (int row = worksheet.Dimension.Start.Row + 2; row <= worksheet.Dimension.End.Row; row++)
                    {

                        DataRow newRow = dt.NewRow();

                        foreach (int columnIndex in new int[] { 7, 8, 9, 10, 34 })
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
        }


    }

}
