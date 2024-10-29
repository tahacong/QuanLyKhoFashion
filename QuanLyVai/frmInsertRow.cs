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
    public partial class frmInsertRow : Form
    {
        public frmInsertRow()
        {
            InitializeComponent();
        }

        private void frmInsertRow_Load(object sender, EventArgs e)
        {
            txtSoDong.Text = "1";
        }

        private void txtSoDong_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void txtSoDong_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Kiểm tra xem ký tự nhập vào có phải là số hoặc phím điều khiển (Backspace) hay không
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                // Nếu không phải, hủy sự kiện để không cho nhập ký tự này
                e.Handled = true;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            NhapVai.RowNumInsert = 0;
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            NhapVai.RowNumInsert = Convert.ToInt32(txtSoDong.Text);
            this.Close();
        }

        private void frmInsertRow_Shown(object sender, EventArgs e)
        {
            txtSoDong.Focus();
        }
    }
}
