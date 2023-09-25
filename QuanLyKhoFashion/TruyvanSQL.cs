using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

namespace QuanLyKhoFashion
{
    class TruyvanSQL
    {
        // Ket noi den CSDL
        //private static string DuongDan = Properties.Resources.duongdan; //"Data Source=10.7.23.8;Initial Catalog=fashionwh;Persist Security Info=True;User ID=sa;Password=Tanhatminh@123";
        private static string DuongDan = Properties.Settings.Default.connectString+";Persist Security Info = True; User ID = sa; Password=Tanhatminh@123";
        public static SqlConnection TaoKetNoi()
        {
            return new SqlConnection(DuongDan);
        }

        public static DataTable TaoBang(string Sqlcm)   //Truy vấn SELECT và trả kết quả là 1 bảng data
        {
            SqlConnection ketnoi = TaoKetNoi();
            ketnoi.Open();
            SqlCommand cm = new SqlCommand();
            cm.CommandText = Sqlcm;
            cm.Connection = ketnoi;
            SqlDataAdapter da = new SqlDataAdapter(cm);
            DataTable dt = new DataTable();
            da.Fill(dt);
            ketnoi.Close();
            return dt;


        }
        // Truy vấn ( Thêm, Sửa , Xóa) đối với SQL
        public static void ThemSuaXoa(SqlCommand Lenh)
        {
            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            Lenh.Connection = Ketnoi;
            Lenh.ExecuteNonQuery();
            Ketnoi.Close();
        }

        //Load dữ liệu từ 1 cột mssql lên combobox
        public static void HienThiLenCBO(string columnName, ComboBox cbobox, string table)
        {
            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            SqlCommand cmd = new SqlCommand($"SELECT [{columnName}] FROM [{table}]", Ketnoi);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                cbobox.Items.Add(dr[0].ToString());
            }
            Ketnoi.Close();
        }
        //Load dữ liệu lên CBO có điều kiện
        public static void HienThiLenCBO_CoDieuKien(string columnName, ComboBox cbobox, string table, string dieukien)
        {
            cbobox.Items.Clear();
            cbobox.Text = "";
            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            SqlCommand cmd = new SqlCommand($"SELECT [{columnName}] FROM [{table}] WHERE {dieukien}", Ketnoi);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                cbobox.Items.Add(dr[0].ToString());
            }
            Ketnoi.Close();
            if(cbobox.Items.Count>0){cbobox.SelectedIndex= 0;}    
        }
        //Load dữ liệu từ 1 cột mssql lên 1 list
        public static List<string> LoadSQLtoList_LocTrung(string command)
        {
            List<string> list = new List<string>();
            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            SqlCommand cmd = new SqlCommand(command, Ketnoi);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                list.Add(dr[0].ToString());
            }
            Ketnoi.Close();
            List<string> lst = list.Distinct().ToList();
            return lst;
        }
        //Load dữ liệu từ 1 cột mssql lên thành 1 dòng textbox có ngăn cách bằng dấu "|"
        public static void LoadSQLtoText(TextBox txt, string command)
        {
            List<string> list = new List<string>();
            string str = "";
            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            SqlCommand cmd = new SqlCommand(command, Ketnoi);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                list.Add(dr[0].ToString());
            }
            List<string> lst = list.Distinct().ToList();
            foreach (string item in lst)
            {
                str = str + " | " + item;
            }
            Ketnoi.Close();
            txt.Text = str;
        }
        //Tính tổng 1 cột trong mssql và cho kết quả
        public static int TinhTongSoLuong1Cot(string columnName, string table, string dieukien)
        {
            int tongsoluong = 0;
            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            SqlCommand cmd = new SqlCommand($"SELECT [{columnName}] FROM [{table}] WHERE {dieukien}", Ketnoi);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                tongsoluong = tongsoluong + Convert.ToInt32(dr[0].ToString());
            }
            Ketnoi.Close();
            return tongsoluong;

        }
        //Tính tổng số lượng hàng tồn kho theo điều kiện
        public static int HangTonKho( string dieukien)
        {
            int tongsoluong = 0;
            string columnName = "soluong_tonkho"; // tên cột tính toán tồn kho
            //string dieukien = "mahang = 'test01' AND size = 'S'"; // điều kiện lọc dữ liệu

            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            SqlCommand cmd = new SqlCommand($"SELECT SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE - soluong END) AS {columnName} FROM lichsu WHERE {dieukien}", Ketnoi);
            SqlDataReader dr = cmd.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    tongsoluong = Convert.ToInt32(dr[columnName].ToString());
                }
            }catch (Exception) { }
            Ketnoi.Close();
            return tongsoluong;
        }

        //Truy vấn 1 bảng và hiển thị dữ liệu lên datagridView
        public static void LoadDataDGV(DataGridView dgv, string dieukien)
        {
            SqlConnection Ketnoi = TaoKetNoi();
            try
            {
                Ketnoi.Open();
                string command = $"SELECT lichsu.vitri,lichsu.mahang,mathang.tenhang,lichsu.bienthe,SUM(CASE WHEN trangthai = 'in' THEN soluong ELSE -soluong END) AS soluong_tonkho FROM lichsu LEFT JOIN mathang ON lichsu.mahang=mathang.mahang  WHERE {dieukien} GROUP BY lichsu.vitri,lichsu.mahang,mathang.tenhang,lichsu.bienthe";
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
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
            }
            Ketnoi.Close();
        }
        //Lấy dữ liệu 1 ô trong DGV sang textbox
        public static void CellDgvtoTextBox(TextBox txt, DataGridView dgv, int x, int y)
        {
            if (dgv.Rows.Count > y && dgv.Columns.Count > x) // Đảm bảo rằng có đủ hàng
            {
                string cellValue = dgv.Rows[y].Cells[x].Value.ToString();
                txt.Text = cellValue;
            }

        }
        //Load dữ liệu SQL thành 1 string
        public static string LoadSqltoString(string command)
        {
            string str = "";
            SqlConnection Ketnoi = TaoKetNoi();
            Ketnoi.Open();
            SqlCommand cmd = new SqlCommand(command, Ketnoi);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                str=dr[0].ToString();
            }
            Ketnoi.Close();
            return str;
        }
        //Nhập hàng
        public static void NhapHang(string vitri, string mahang, string bienthe, int soluong)
        {
            SqlCommand cmd = new SqlCommand($"INSERT INTO lichsu (vitri, mahang, bienthe, soluong, thoigian, trangthai) VALUES (@vitri, @mahang,@bienthe,@soluong',GETDATE(), 'in')");
            cmd.Parameters.AddWithValue("@vitri", vitri);
            cmd.Parameters.AddWithValue("@mahang", mahang);
            cmd.Parameters.AddWithValue("@bienthe", bienthe);
            cmd.Parameters.AddWithValue("@soluong", soluong);
            ThemSuaXoa(cmd);
        }
        //Xuất hàng
        public static void XuatHang(string vitri, string mahang, string bienthe, int soluong)
        {
            SqlCommand cmd = new SqlCommand($"INSERT INTO lichsu (vitri, mahang, bienthe, soluong, thoigian, trangthai) VALUES ('{vitri}', '{mahang}','{bienthe}','{soluong}',GETDATE(), 'out')");
            ThemSuaXoa(cmd);
        }
        //Kiểm tra sự tồn tại của giá trị khi có điều kiện
        public static bool KiemTraTonTaiGiaTriTrongBang(string bang,string cotgiatri,string giatri)
        {
            bool ketqua;
            SqlConnection ketnoi = TaoKetNoi();
            ketnoi.Open();
            string query = $"IF EXISTS (SELECT 1 FROM [{bang}] WHERE [{cotgiatri}] = '{giatri}') SELECT 1 ELSE SELECT 0";

            using (SqlCommand command = new SqlCommand(query, ketnoi))
            {
                int result = Convert.ToInt32(command.ExecuteScalar());
                ketqua =Convert.ToBoolean(result);
            }
            ketnoi.Close();
            return ketqua;
        }
        public static bool KiemTraTonTaiGiaTriTrongBangN(string bang, string cotgiatri, string giatri)
        {
            bool ketqua;
            SqlConnection ketnoi = TaoKetNoi();
            ketnoi.Open();
            string query = $"IF EXISTS (SELECT 1 FROM [{bang}] WHERE [{cotgiatri}] = N'{giatri}') SELECT 1 ELSE SELECT 0";

            using (SqlCommand command = new SqlCommand(query, ketnoi))
            {
                int result = Convert.ToInt32(command.ExecuteScalar());
                ketqua = Convert.ToBoolean(result);
            }
            ketnoi.Close();
            return ketqua;
        }
        public static bool KiemTraTonTaiGiaTriTrongBangPlus(string bang, string dieukien)
        {
            bool ketqua;
            SqlConnection ketnoi = TaoKetNoi();
            ketnoi.Open();
            string query = $"IF EXISTS (SELECT 1 FROM [{bang}] WHERE {dieukien}) SELECT 1 ELSE SELECT 0";

            using (SqlCommand command = new SqlCommand(query, ketnoi))
            {
                int result = Convert.ToInt32(command.ExecuteScalar());
                ketqua = Convert.ToBoolean(result);
            }
            ketnoi.Close();
            return ketqua;
        }
        //Lấy giá trị text từ mssql
        public static string LayGiaTriText(string bang, string cotgiatri, string cotdieukien, string dieukien)
        {
            string ketqua;
            SqlConnection ketnoi = TaoKetNoi();
            ketnoi.Open();
            string query = $"SELECT [{cotgiatri}] FROM [{bang}] WHERE [{cotdieukien}] = '{dieukien}'";

            using (SqlCommand command = new SqlCommand(query, ketnoi))
            {
                ketqua = Convert.ToString(command.ExecuteScalar());
            }
            ketnoi.Close();
            return ketqua;
        }
        //Đẩy dữ liệu từ DGV Lấy hàng insert lên bảng lịch sử MSSQL
        public static void UpdateLayHangtoLichSu(DataGridView dgv)
        {
            SqlConnection ketnoi = TaoKetNoi();
            ketnoi.Open();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Cells["vitri_layhang"].Value != null && row.Cells["mahang_layhang"].Value != null)
                {
                    string vitri = row.Cells["vitri_layhang"].Value.ToString();
                    string mahang = row.Cells["mahang_layhang"].Value.ToString();
                    string bienthe = row.Cells["bienthe_layhang"].Value.ToString();
                    int soluong = Convert.ToInt32(row.Cells["soluong_layhang"].Value.ToString());

                    SqlCommand query = new SqlCommand("INSERT INTO lichsu (vitri, mahang, bienthe,soluong, thoigian, trangthai) VALUES (@vitri, @mahang,@bienthe,@soluong,GETDATE(), 'out');");
                    query.Parameters.AddWithValue("@vitri", vitri);
                    query.Parameters.AddWithValue("@mahang", mahang);
                    query.Parameters.AddWithValue("@bienthe", bienthe);
                    query.Parameters.AddWithValue("@soluong", soluong);
                    ThemSuaXoa(query);
                    
                }
            }
            ketnoi.Close();
            MessageBox.Show("Cập nhật dữ liệu lấy hàng thành công!");
        }
        //Đẩy dữ liệu từ DGV Lấy hàng insert lên bảng đóng gói MSSQL
        public static void UpdateLayHangtoDongGoi(DataGridView dgv)
        {
            SqlConnection ketnoi = TaoKetNoi();
            ketnoi.Open();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Cells["Tracking ID"].Value != null && row.Cells["Seller SKU"].Value != null)
                {
                    string vandon = row.Cells["Tracking ID"].Value.ToString();
                    string mahang = row.Cells["Seller SKU"].Value.ToString();
                    string bienthe = row.Cells["Variation"].Value.ToString();
                    int soluong = Convert.ToInt32(row.Cells["Quantity"].Value.ToString());

                    SqlCommand query = new SqlCommand("INSERT INTO donggoi (vandon,mahang,bienthe,soluong) VALUES (@vandon,@mahang,@bienthe,@soluong)");
                    query.Parameters.AddWithValue("@vandon", vandon);
                    query.Parameters.AddWithValue("@mahang", mahang);
                    query.Parameters.AddWithValue("@bienthe", bienthe);
                    query.Parameters.AddWithValue("@soluong", soluong);
                    ThemSuaXoa(query);

                }
            }
            ketnoi.Close();
            MessageBox.Show("Cập nhật dữ liệu lấy hàng thành công!");
        }
        //Load bảng mã màu lên DGV
        public static void SQLtoDGV(DataGridView dgv, string command)
        {
            SqlConnection Ketnoi = TaoKetNoi();
            try
            {
                Ketnoi.Open();
                SqlCommand cmd = new SqlCommand(command, Ketnoi);
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                //tai du lieu len girdView
                dgv.DataSource = dt;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
            }
            Ketnoi.Close();
        }
    }
}
