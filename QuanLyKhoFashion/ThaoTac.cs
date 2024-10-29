using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using OfficeOpenXml;
using System.Diagnostics;

namespace QuanLyKhoFashion
{
    internal class ThaoTac
    {
        public static void LoadImageFromURL(PictureBox pic,string imageURL)
        {
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    byte[] data = webClient.DownloadData(imageURL); // Tải dữ liệu hình ảnh từ URL
                    using (MemoryStream stream = new MemoryStream(data))
                    {
                        Image image = Image.FromStream(stream); // Tạo đối tượng hình ảnh từ dữ liệu
                        pic.SizeMode = PictureBoxSizeMode.Zoom; // Đặt chế độ co dãn hình ảnh
                        pic.Image = image; // Gán hình ảnh cho PictureBox
                    }
                }
            }
            catch (Exception)
            {
                
            }
        }
        //Đổi tên cột trong DGV
        public static void DoiTenCot(DataGridView dgv,List<string> listTenCot, List<string> newHeader)
        {
            for (int i = 0; i < listTenCot.Count; i++)
            {
                dgv.Columns[listTenCot[i].ToString()].HeaderText = newHeader[i].ToString();
            }
        }
        //Xóa tất cả phần tử trong combobox
        public static void SapXepComboBox(ComboBox cbo,List<string> listUse)
        {
            //list1 = list1.Except(list2).ToList();
            List<string> Oldlist = new List<string>();
            foreach (string item in cbo.Items) { Oldlist.Add(item); }
            cbo.Items.Clear();
            Oldlist=Oldlist.Except(listUse).ToList();
            foreach(string item in listUse) { cbo.Items.Add(item); }
            foreach (string item in Oldlist) { cbo.Items.Add(item); }
        }
        //Xuất File Excel từ DGV
        public static void XuatExceltuDGB(DataGridView dgv,string TenThuMuc)
        {
            string filePath = Application.StartupPath+"/"+TenThuMuc+"/";
            string fileName = "LH" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            FileInfo newFile = new FileInfo(filePath+fileName);
            using (ExcelPackage excelPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                // Đổ tiêu đề cột vào Excel
                for (int col = 1; col <= dgv.Columns.Count; col++)
                {
                    worksheet.Cells[1, col].Value = dgv.Columns[col - 1].HeaderText;
                }

                // Đổ dữ liệu từ DataGridView vào Excel
                for (int row = 1; row <= dgv.Rows.Count; row++)
                {
                    for (int col = 1; col <= dgv.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 1, col].Value = dgv.Rows[row - 1].Cells[col - 1].Value;
                    }
                }

                excelPackage.Save();
            }
            MoExcelFile(filePath+fileName);
        }
        //Xuất file Excel từ DGV không mở file
        public static void XuatExceltuDGBKhongMo(DataGridView dgv, string TenThuMuc)
        {
            string filePath = Application.StartupPath + "/" + TenThuMuc + "/";
            string fileName = "LH" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            FileInfo newFile = new FileInfo(filePath + fileName);
            using (ExcelPackage excelPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                // Đổ tiêu đề cột vào Excel
                for (int col = 1; col <= dgv.Columns.Count; col++)
                {
                    worksheet.Cells[1, col].Value = dgv.Columns[col - 1].HeaderText;
                }

                // Đổ dữ liệu từ DataGridView vào Excel
                for (int row = 1; row <= dgv.Rows.Count; row++)
                {
                    for (int col = 1; col <= dgv.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 1, col].Value = dgv.Rows[row - 1].Cells[col - 1].Value;
                    }
                }

                excelPackage.Save();
            }
            
        }
        //Xuất file exce từ sàn tiktok
        /*
        public static void XuatExcelTiktok(DataGridView dgv, string TenThuMuc)
        {

            string filePath = Application.StartupPath + "/" + TenThuMuc + "/";
            string fileName = "TT" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            FileInfo newFile = new FileInfo(filePath + fileName);
            using (ExcelPackage excelPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                // Đổ tiêu đề cột vào Excel
                for (int col = 1; col <= dgv.Columns.Count; col++)
                {
                    worksheet.Cells[1, col].Value = dgv.Columns[col - 1].HeaderText;
                }

                // Đổ dữ liệu từ DataGridView vào Excel
                for (int row = 1; row <= dgv.Rows.Count; row++)
                {
                    for (int col = 1; col <= dgv.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 1, col].Value = dgv.Rows[row - 1].Cells[col - 1].Value;
                    }
                }

                excelPackage.Save();
            }
            MoExcelFile(filePath + fileName);
        }
        */
        public static void XuatExcelTiktok(DataGridView dgv, string TenThuMuc)
        {
            string filePath = Application.StartupPath + "/" + TenThuMuc + "/";
            string fileName = "TT" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            FileInfo newFile = new FileInfo(filePath + fileName);
            using (ExcelPackage excelPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // Ánh xạ giữa tên cột trên DataGridView và tên cột bạn muốn xuất ra Excel
                Dictionary<string, string> columnMapping = new Dictionary<string, string>
        {
            { "Seller SKU", "MaHang" },
            { "Product Name", "TenHang" },
            { "Variation", "BienThe" },
            { "Quantity", "soluong" },
            { "Tracking ID", "VanDon" }
        };

                // Đổ tiêu đề cột vào Excel
                for (int col = 1; col <= dgv.Columns.Count; col++)
                {
                    string dgvColumnName = dgv.Columns[col - 1].HeaderText;
                    if (columnMapping.ContainsKey(dgvColumnName))
                    {
                        string excelColumnName = columnMapping[dgvColumnName];
                        worksheet.Cells[1, col].Value = excelColumnName;
                    }
                }

                // Đổ dữ liệu từ DataGridView vào Excel
                for (int row = 1; row <= dgv.Rows.Count; row++)
                {
                    for (int col = 1; col <= dgv.Columns.Count; col++)
                    {
                        string dgvColumnName = dgv.Columns[col - 1].HeaderText;
                        if (columnMapping.ContainsKey(dgvColumnName))
                        {
                            string excelColumnName = columnMapping[dgvColumnName];
                            worksheet.Cells[row + 1, col].Value = dgv.Rows[row - 1].Cells[col - 1].Value;
                        }
                    }
                }

                excelPackage.Save();
            }
            MoExcelFile(filePath + fileName);
        }
        //Mở file excel vừa tạo
        public static void MoExcelFile(string filePath)
        {
            try
            {
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể mở file Excel: " + ex.Message);
            }
        }
    }
}
