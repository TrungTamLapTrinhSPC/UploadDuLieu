using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
namespace UploadBaoCao
{
    public partial class FormMain : DevExpress.XtraEditors.XtraForm
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            this.CenterToParent();
            cmb_LoaiBaoCao.SelectedIndex = 0;
            loadSetting();

            cmb_NgayTao.DateTimeOffset = DateTimeOffset.Now;
            txt_Post.Text = Bien.ftpPost;
            txt_Host.Text = Bien.ftpHost;
            txt_User.Text = Bien.ftpUser;
            txt_Pass.Text = Bien.ftpPassword;
        }
        private void getDataByNameFile()
        {
            int nam = cmb_NgayTao.DateTimeOffset.Year;
            int thang = cmb_NgayTao.DateTimeOffset.Month;
            int ngay = cmb_NgayTao.DateTimeOffset.Day;
            string date = nam.ToString() + thang.ToString() + ngay.ToString();
            string foldersave = "";
            string filename = "";
            if (cmb_LoaiBaoCao.SelectedIndex==0)
            {
                filename= Bien.MaDonVi + "_revenue_daily_kqi_" + date + ".txt";
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang + "\\" + ngay +"\\" + filename;
            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 1)
            {
                filename = Bien.MaDonVi + "_revenue_monthly_kqi_" + date + ".txt";
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang + "\\" + filename;
            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 2)
            {
                filename = Bien.MaDonVi + "_revenue_monthly_kh_kqi_" + date + ".txt";
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang + "\\" + filename;
            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 3)
            {
                filename = Bien.MaDonVi + "_quarterly_kh_kqi_" + date + ".txt";
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang + "\\" + filename;
            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 4)
            {
                filename = Bien.MaDonVi + "_revenue_yearly_kh_kqi_" + date + ".txt";
                foldersave = Bien.pathProject + "\\" + nam + "\\" + filename;
            }

            if(File.Exists(foldersave))
            {
                string[] noidung = File.ReadAllLines(foldersave);
                if (noidung.Count() > 0)
                {
                    CellRange usedrange = Bien.worksheet.GetUsedRange();
                    int rowcount = usedrange.RowCount;
                    int dongbatdau = 3;
                    int ind_Ma = 0;
                    int ind_Value = 5;
                    int ind_Value_Month = 6;
                    for (int i = 0; i < rowcount;i++)
                    {
                        string ma_ex = Bien.worksheet.Cells[dongbatdau + i, ind_Ma].Value.ToString();
                        foreach (string str in noidung)
                        {

                            string[] lst = str.Split('|');
                            if (filename.Contains("daily"))
                            {
                                string ma = lst[0];
                                string value = lst[1];
                                string value_month = lst[2];
                                if (ma == ma_ex)
                                {
                                    Bien.worksheet.Cells[dongbatdau + i, ind_Value].Value = value;
                                    Bien.worksheet.Cells[dongbatdau + i, ind_Value_Month].Value = value_month;
                                }
                            }
                            else
                            {
                                string ma = lst[0];
                                string value = lst[1];
                                if (ma == ma_ex)
                                {
                                    Bien.worksheet.Cells[dongbatdau + i, ind_Value].Value = value;
                                }
                            }
                        }
                    }
                   
                }
            }
        }
        private void loadSetting()
        {
            DataSet set = new DataSet();
            set.ReadXml(Bien.pathSetting);
            DataTable tb = set.Tables[0];

            Bien.pathProject = tb.Rows[0]["DuongDanProject"].ToString();
            Bien.ftpPost = tb.Rows[0]["SFPTPOST"].ToString();
            Bien.ftpHost = tb.Rows[0]["SFPTHOST"].ToString();
            Bien.ftpUser = tb.Rows[0]["SFPTUSER"].ToString();
            Bien.ftpPassword = tb.Rows[0]["SFTPPASS"].ToString();
            Bien.MaDonVi = tb.Rows[0]["MADONVI"].ToString();
        }
        public DataTable CreateTable(string tableName,  string[] colName)
        {
            DataTable table = new DataTable();
            try
            {
                table.TableName = tableName;
                for (int i = 0; i <= colName.Count() - 1; i++)
                {
                    DataColumn c = new DataColumn();
                    c.ColumnName = colName[i];
                    c.DataType = typeof(string);
                    c.DefaultValue = "";
                    table.Columns.Add(c);
                }
            }
            catch (Exception ex)
            {
            }
            return table;
        }
        private void saveSetting()
        {
            DataSet set = new DataSet();
            DataTable tb = CreateTable("tb_Setting", new string[] { "DuongDanProject", "SFPTHOST", "SFPTPOST", "SFPTUSER", "SFTPPASS", "MADONVI" });
            DataRow r = tb.NewRow();
            r["DuongDanProject"] = Bien.pathProject;
            r["SFPTPOST"] = Bien.ftpPost;
            r["SFPTHOST"] = Bien.ftpHost;
            r["SFPTUSER"] = Bien.ftpUser;
            r["SFTPPASS"] = Bien.ftpPassword;
            r["MADONVI"] = Bien.MaDonVi;
            tb.Rows.Add(r);
            set.Tables.Add(tb.Copy());

            set.WriteXml(Bien.pathSetting);
        }
        private void btn_KetNoi_Click(object sender, EventArgs e)
        {
            pn_FTP.Visible = false;
            Bien.ftpHost = txt_Host.Text;
            Bien.ftpPassword = txt_Pass.Text;
            Bien.ftpPost = txt_Post.Text;
            Bien.ftpUser = txt_User.Text;
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            pn_FTP.Visible = true;
        }

        private void cmb_LoaiBaoCao_SelectedIndexChanged(object sender, EventArgs e)
        {
            Bien.workbook = spreadsheetControl1.Document;
            cmb_Folder.SelectedIndex = cmb_LoaiBaoCao.SelectedIndex;
            if (cmb_LoaiBaoCao.SelectedIndex == 0)
            {
                Bien.workbook.LoadDocument(Bien.pathTemplate_Ngay);
                Bien.worksheet = spreadsheetControl1.Document.Worksheets[0];

            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 1)
            {
                Bien.workbook.LoadDocument(Bien.pathTemplate_ChotThang);
                Bien.worksheet = spreadsheetControl1.Document.Worksheets[0];

            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 2)
            {
                Bien.workbook.LoadDocument(Bien.pathTemplate_Thang);
                Bien.worksheet = spreadsheetControl1.Document.Worksheets[0];

            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 3)
            {
                Bien.workbook.LoadDocument(Bien.pathTemplate_Quy);
                Bien.worksheet = spreadsheetControl1.Document.Worksheets[0];

            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 4)
            {
                Bien.workbook.LoadDocument(Bien.pathTemplate_Nam);
                Bien.worksheet = spreadsheetControl1.Document.Worksheets[0];

            }
            getDataByNameFile();
        }

        private void btn_ChuyenDoi_Click(object sender, EventArgs e)
        {
            int nam = cmb_NgayTao.DateTimeOffset.Year;
            int thang = cmb_NgayTao.DateTimeOffset.Month;
            int ngay = cmb_NgayTao.DateTimeOffset.Day;
            if (cmb_LoaiBaoCao.SelectedIndex == 0)
            {
                //1.	Đối với dữ liệu ngày
                string date = nam.ToString() + thang.ToString() + ngay.ToString();
                ConvertThucHienNgay(date);
            }
            else if (cmb_LoaiBaoCao.SelectedIndex ==1)
            {
                //1.Đối với dữ liệu chốt tháng
                string date = nam.ToString() + thang.ToString() + ngay.ToString();
                ConvertThucHienChotThang(date);
            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 2)
            {
                //1.	Đối với dữ liệu Kế hoạch tháng
                string date = nam.ToString() + thang.ToString() + ngay.ToString();
                ConvertThucHienThang(date);
            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 3)
            {
                //1.	Đối với dữ liệu Kế hoạch Quý 
                string date = nam.ToString() + thang.ToString() + ngay.ToString();
                ConvertThucHienQuy(date);
            }
            else if (cmb_LoaiBaoCao.SelectedIndex == 4)
            {
                //1.	Đối với dữ liệu Kế hoạch Năm 
                string date = nam.ToString() + thang.ToString() + ngay.ToString();
                ConvertThucHienNam(date);
            }

        }
        private void ConvertThucHienNgay(string date)
        {
            //Fomat tên file:  <Mã_đơn_vị>_revenue_daily_kqi_<PRD_ID>.txt
            //Ví dụ: Nội dung file vtt_revenue_daily_kqi_20211202.txt
            //Nội dung file: 
            string temp = "SERVICE_PK|F_VALUE|F_VALUE_MONTH|PRD_ID|DEP_ID|UNIT_ID";
            string noidung = "";
            int ind_Ma = 0;
            int ind_Value = 5;
            int ind_Value_Month = 6;
            int ind_DonViTinh = 7;
           
            CellRange usedrange = Bien.worksheet.GetUsedRange();
            int rowcount = usedrange.RowCount;
            int dongbatdau = 3;
            for (int i =0; i< rowcount;i++)
            {
                string value = Bien.worksheet.Cells[dongbatdau + i, ind_Value].Value.ToString();
                if (value !="" && value != "0" && char.IsNumber(value,0))
                {
                    string ma = Bien.worksheet.Cells[dongbatdau + i, ind_Ma].Value.ToString();
                    string valuemonth = Bien.worksheet.Cells[dongbatdau + i, ind_Value_Month].Value.ToString();
                    string unit = Bien.worksheet.Cells[dongbatdau + i, ind_DonViTinh].Value.ToString();
                    string noidung_id = temp.Replace("SERVICE_PK",ma)
                        .Replace("F_VALUE_MONTH", valuemonth)
                        .Replace("F_VALUE", value)
                        .Replace("PRD_ID", date)
                        .Replace("DEP_ID", Bien.MaDonVi)
                        .Replace("UNIT_ID", unit);

                        noidung += noidung_id + "\n";
              
                }
            }

            if (noidung != "")
            {
                txt_NoiDung.Text = noidung;
                txt_TenFile.Text = Bien.MaDonVi + "_revenue_daily_kqi_" + date + ".txt";
            }
        }
        private void ConvertThucHienThang(string date)
        {
            //Fomat tên file:  <Mã_đơn_vị>_revenue_daily_kqi_<PRD_ID>.txt
            //Ví dụ: Nội dung file vtt_revenue_daily_kqi_20211202.txt
            //Nội dung file: 
            string temp = "SERVICE_PK|F_VALUE_MONTH|PRD_ID|DEP_ID|UNIT_ID";
            string noidung = "";
            int ind_Ma = 0;
            int ind_Value = 5;
            int ind_DonViTinh = 6;

            CellRange usedrange = Bien.worksheet.GetUsedRange();
            int rowcount = usedrange.RowCount;
            int dongbatdau = 3;
            for (int i = 0; i < rowcount; i++)
            {
                string value = Bien.worksheet.Cells[dongbatdau + i, ind_Value].Value.ToString();
                if (value != "" && value != "0" && char.IsNumber(value, 0))
                {
                    string ma = Bien.worksheet.Cells[dongbatdau + i, ind_Ma].Value.ToString();
                    string unit = Bien.worksheet.Cells[dongbatdau + i, ind_DonViTinh].Value.ToString();
                    string noidung_id = temp.Replace("SERVICE_PK", ma)
                        .Replace("F_VALUE_MONTH", value)
                        .Replace("PRD_ID", date)
                        .Replace("DEP_ID", Bien.MaDonVi)
                        .Replace("UNIT_ID", unit);
                    noidung += noidung_id + "\n";

                }
            }

            if (noidung != "")
            {
                txt_NoiDung.Text = noidung;
                txt_TenFile.Text = Bien.MaDonVi + "_revenue_monthly_kh_kqi_" + date + ".txt";
            }
        }
        private void ConvertThucHienChotThang(string date)
        {
            //Fomat tên file:  <Mã_đơn_vị>_revenue_daily_kqi_<PRD_ID>.txt
            //Ví dụ: Nội dung file vtt_revenue_daily_kqi_20211202.txt
            //Nội dung file: 
            string temp = "SERVICE_PK|F_VALUE_MONTH|PRD_ID|DEP_ID|UNIT_ID";
            string noidung = "";
            int ind_Ma = 0;
            int ind_Value = 5;
            int ind_DonViTinh = 6;

            CellRange usedrange = Bien.worksheet.GetUsedRange();
            int rowcount = usedrange.RowCount;
            int dongbatdau = 3;
            for (int i = 0; i < rowcount; i++)
            {
                string value = Bien.worksheet.Cells[dongbatdau + i, ind_Value].Value.ToString();
                if (value != "" && value != "0" && char.IsNumber(value, 0))
                {
                    string ma = Bien.worksheet.Cells[dongbatdau + i, ind_Ma].Value.ToString();
                    string unit = Bien.worksheet.Cells[dongbatdau + i, ind_DonViTinh].Value.ToString();
                    string noidung_id = temp.Replace("SERVICE_PK", ma)
                        .Replace("F_VALUE_MONTH", value)
                        .Replace("PRD_ID", date)
                        .Replace("DEP_ID", Bien.MaDonVi)
                        .Replace("UNIT_ID", unit);

                    noidung += noidung_id + "\n";

                }
            }

            if (noidung != "")
            {
                txt_NoiDung.Text = noidung;
                txt_TenFile.Text = Bien.MaDonVi + "_revenue_monthly_kqi_" + date + ".txt";
            }
        }
        private void ConvertThucHienNam(string date)
        {
            //Fomat tên file:  <Mã_đơn_vị>_revenue_daily_kqi_<PRD_ID>.txt
            //Ví dụ: Nội dung file vtt_revenue_daily_kqi_20211202.txt
            //Nội dung file: 
            string temp = "SERVICE_PK|F_VALUE_YEAR|PRD_ID|DEP_ID|UNIT_ID";
            string noidung = "";
            int ind_Ma = 0;
            int ind_Value = 5;
            int ind_DonViTinh = 6;

            CellRange usedrange = Bien.worksheet.GetUsedRange();
            int rowcount = usedrange.RowCount;
            int dongbatdau = 3;
            for (int i = 0; i < rowcount; i++)
            {
                string value = Bien.worksheet.Cells[dongbatdau + i, ind_Value].Value.ToString();
                if (value != "" && value != "0" && char.IsNumber(value, 0))
                {
                    string ma = Bien.worksheet.Cells[dongbatdau + i, ind_Ma].Value.ToString();
                    string unit = Bien.worksheet.Cells[dongbatdau + i, ind_DonViTinh].Value.ToString();
                    string noidung_id = temp.Replace("SERVICE_PK", ma)
                        .Replace("F_VALUE_YEAR", value)
                        .Replace("PRD_ID", date)
                        .Replace("DEP_ID", Bien.MaDonVi)
                        .Replace("UNIT_ID", unit);
                    noidung += noidung_id + "\n";

                }
            }

            if (noidung != "")
            {
                txt_NoiDung.Text = noidung;
                txt_TenFile.Text = Bien.MaDonVi + "_revenue_yearly_kh_kqi_" + date + ".txt";
            }
        }
        private void ConvertThucHienQuy(string date)
        {
            //Fomat tên file:  <Mã_đơn_vị>_revenue_daily_kqi_<PRD_ID>.txt
            //Ví dụ: Nội dung file vtt_revenue_daily_kqi_20211202.txt
            //Nội dung file: 
            string temp = "SERVICE_PK|F_VALUE_YEAR|PRD_ID|DEP_ID|UNIT_ID";
            string noidung = "";
            int ind_Ma = 0;
            int ind_Value = 5;
            int ind_DonViTinh = 6;

            CellRange usedrange = Bien.worksheet.GetUsedRange();
            int rowcount = usedrange.RowCount;
            int dongbatdau = 3;
            for (int i = 0; i < rowcount; i++)
            {
                string value = Bien.worksheet.Cells[dongbatdau + i, ind_Value].Value.ToString();
                if (value != "" && value != "0" && char.IsNumber(value, 0))
                {
                    string ma = Bien.worksheet.Cells[dongbatdau + i, ind_Ma].Value.ToString();
                    string unit = Bien.worksheet.Cells[dongbatdau + i, ind_DonViTinh].Value.ToString();
                    string noidung_id = temp.Replace("SERVICE_PK", ma)
                        .Replace("F_VALUE_YEAR", value)
                        .Replace("PRD_ID", date)
                        .Replace("DEP_ID", Bien.MaDonVi)
                        .Replace("UNIT_ID", unit);
                    noidung += noidung_id + "\n";
                }
            }
            if (noidung != "")
            {
                txt_NoiDung.Text = noidung;
                txt_TenFile.Text = Bien.MaDonVi + "_quarterly_kh_kqi_" + date + ".txt";
            }
        }
        private void uploadFile(string filename)
        {
           if( clsDongBo.UpLoadFile(cmb_Folder.Text, filename,Bien.ftpUser,Bien.ftpPassword))
                MessageBox.Show("Upload dữ liệu thành công!", "UPload", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
           else
                MessageBox.Show("Upload dữ liệu không thành công!", "UPload", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
        private void savefile(string filename)
        {
            //Lấy loại file, ngày tháng năm'
            string[] getdate = filename.Split('_');
            string date = getdate[getdate.Count() -1].Replace(".txt", "");
            string tenexcel = "";
            string loaifile = filename.Replace(getdate + ".txt","");
            string nam = date[0].ToString() + date[1].ToString() + date[2].ToString() + date[3].ToString();
            string thang = date[4].ToString() + date[5].ToString();
            string ngay = date[6].ToString() + date[7].ToString();
            //
            string foldersave = Bien.pathProject + "\\" + nam + "\\" + thang + "\\" + ngay;

            if (loaifile.Contains("daily"))
            {
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang + "\\" + ngay;
                tenexcel = "TVTK_THUC HIEN NGAY.xlsx";
            }
            else if (loaifile.Contains("revenue_monthly"))
            {
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang;
                tenexcel = "TVTK_THUC HIEN CHOT THANG.xlsx";
            }
            else if (loaifile.Contains("monthly_kh"))
            {
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang;
                tenexcel = "TVTK_THUC HIEN THANG.xlsx";
            }
            else if (loaifile.Contains("quarterly_kh_kqi"))
            {
                foldersave = Bien.pathProject + "\\" + nam + "\\" + thang ;
                tenexcel = "TVTK_THUC HIEN QUY_"+ thang + ".xlsx";
            }
            else if (loaifile.Contains("yearly_kh_kqi"))
            {
                foldersave = Bien.pathProject + "\\" + nam;
                tenexcel = "TVTK_THUC HIEN NAM.xlsx";
            }

            string foldersavefile = foldersave + "\\" + filename;
            string foldersaveexcel = foldersave + "\\" + tenexcel;

            if (!Directory.Exists(foldersave))
            {
                Directory.CreateDirectory(foldersave);
            }
            if (Directory.Exists(foldersave))
            {
                string noidung = txt_NoiDung.Text;
                File.WriteAllText(foldersavefile, noidung);
 
               using (var stream = new FileStream(foldersaveexcel, FileMode.Create, FileAccess.ReadWrite))
                {
                    Bien.workbook.SaveDocument(stream, DocumentFormat.Xlsx);
                }
                if (File.Exists(foldersavefile))
                {
                    uploadFile(foldersavefile);
                    MessageBox.Show("Đã lưu dữ liệu thành công!", "Save file", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            
              
        }
        private void btn_Upload_Click(object sender, EventArgs e)
        {
            savefile(txt_TenFile.Text);
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            saveSetting();
        }

        private void btn_ChonFolder_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Form_SelectedFolder frm = new Form_SelectedFolder();
            frm.ShowDialog();
        }

        private void cmb_NgayTao_EditValueChanged(object sender, EventArgs e)
        {
            getDataByNameFile();
        }
    }
}
