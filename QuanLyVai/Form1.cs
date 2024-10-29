using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyVai
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // Thay đổi màu nền dòng tiêu đề
            dgvNhapVai.EnableHeadersVisualStyles = false; // Bỏ qua phong cách mặc định
            dgvNhapVai.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvNhapVai.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(35, 91, 159); //
            dgvNhapVai.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; // Màu chữ trắng (RGB: 255, 255, 255)
            //dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 12, FontStyle.Regular); // Thay đổi font chữ nếu cần
            // Thay đổi màu đường kẻ
            dgvNhapVai.GridColor = Color.FromArgb(217, 220, 229); // Màu đen (RGB: 0, 0, 0)

            // Tùy chọn: Thay đổi kiểu đường viền
            dgvNhapVai.CellBorderStyle = DataGridViewCellBorderStyle.Single; // Kiểu đường viền đơn
                                                                                //dataGridView1.Columns["Column2"].ReadOnly = true; // Ví dụ cột không cho phép nhập
                                                                                // Thay đổi màu sắc cho các cột
            foreach (DataGridViewColumn column in dgvNhapVai.Columns)
            {
                if (column.ReadOnly)
                {
                    column.DefaultCellStyle.BackColor = Color.FromArgb(242, 242, 240); // Màu xám
                }
                else
                {
                    column.DefaultCellStyle.BackColor = Color.FromArgb(253,250,183); // Màu vàng
                }
            }
        }

        private void btnXemTonKho_t1_Click(object sender, EventArgs e)
        {
            // Thêm dòng trống
            int newRowIndex = dgvNhapVai.Rows.Add();

            // Đặt con trỏ chuột vào dòng mới thêm
            dgvNhapVai.CurrentCell = dgvNhapVai.Rows[newRowIndex].Cells[0];


        }

        private void btnNhaphang_t1_Click(object sender, EventArgs e)
        {
            dgvNhapVai.Rows.Clear();
        }

        private void PasteClipboardData()
        {
            try
            {
                // Lấy dữ liệu từ clipboard
                string s = Clipboard.GetText();
                string[] lines = s.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                int rowIndex = dgvNhapVai.CurrentCell.RowIndex;
                int colIndex = dgvNhapVai.CurrentCell.ColumnIndex;

                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line)) continue;

                    string[] cells = line.Split('\t');
                    int tempColIndex = colIndex;

                    // Thêm hàng mới nếu không đủ hàng trong DataGridView
                    if (rowIndex >= dgvNhapVai.Rows.Count)
                    {
                        dgvNhapVai.Rows.Add();
                    }

                    // Thêm dữ liệu vào các ô
                    for (int i = 0; i < cells.Length; i++)
                    {
                        // Kiểm tra xem cột có cho phép nhập liệu hay không
                        if (tempColIndex < dgvNhapVai.Columns.Count && !dgvNhapVai.Columns[tempColIndex].ReadOnly)
                        {
                            dgvNhapVai[tempColIndex, rowIndex].Value = cells[i];
                        }
                        tempColIndex++;
                    }
                    rowIndex++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra khi dán dữ liệu: " + ex.Message);
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                PasteClipboardData();
                e.Handled = true;
            }

        }
    }
}
