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
    public partial class XuatVai : Form, IThemDongDGV
    {
        public XuatVai()
        {
            InitializeComponent();
        }

        private void XuatVai_Load(object sender, EventArgs e)
        {

        }
        public void ThemDongDGV()
        {
            button1.BackColor = Color.Red;
        }
    }
}
