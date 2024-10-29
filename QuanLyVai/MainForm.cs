using QuanLyVai;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyVai
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        
        private void MainForm_Load(object sender, EventArgs e)
        {
            newfrm = new NhapVai();
            MoformMoi();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            
        }

        private Form newfrm;

        // Phương thức xử lý khi nhấn nút
        private void btnInsertRow_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem newfrm có khác null không trước khi gọi phương thức
            if (newfrm is NhapVai nhapVaiForm)
            {
                nhapVaiForm.ThemDongDGV(); // Gọi phương thức thêm dòng vào DataGridView
            }
        }

        // Phương thức xử lý khi double-click vào node trong TreeView
        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            spCMain.Panel2.Controls.Clear();
            switch (e.Node.Text)
            {
                case "Nhập kho":
                    // Khởi tạo Form NhapVai
                    newfrm = new NhapVai();
                    MoformMoi();
                    break;
                case "Xuất kho":
                    // Khởi tạo Form XuatVai
                    newfrm = new XuatVai();
                    MoformMoi();
                    break;
            }
        }

        // Phương thức mở Form mới
        private void MoformMoi()
        {
            // Mở Form con
            if (newfrm != null)
            {
                newfrm.TopLevel = false;
                newfrm.FormBorderStyle = FormBorderStyle.None;
                newfrm.Dock = DockStyle.Fill;
                spCMain.Panel2.Controls.Add(newfrm);
                newfrm.Show(); // Hiển thị Form
                newfrm.FormClosed += (s, args) => newfrm = null; // Reset newfrm khi Form đã đóng
            }
        }
    }
}
// Tạo một interface để định nghĩa phương thức ThemDongDGV
public interface IThemDongDGV
{
    void ThemDongDGV();
}

// Đảm bảo cả NhapVai và XuatVai đều thực hiện interface này
public class NhapVai : Form, IThemDongDGV
{
    public void ThemDongDGV()
    {
        // Logic để thêm dòng vào DataGridView cho NhapVai
    }
}

public class XuatVai : Form, IThemDongDGV
{
    public void ThemDongDGV()
    {
        // Logic để thêm dòng vào DataGridView cho XuatVai
    }
}
