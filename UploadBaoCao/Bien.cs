using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UploadBaoCao
{
    class Bien
    {
        public static string linkfpt = "ftp://ftp_gbtd@LocalHost/2022";
        public static string pathProject = @"D:\0.ITsoft\code\PhanMemUpload\Tài liệu\2022";
        public static string pathSetting = Environment.CurrentDirectory +"\\Setting.xml";
        public static string pathTemplate = Environment.CurrentDirectory + "\\Temp\\";
        public static string ftpUser = "ftp_gbtd";
        public static string ftpPassword = "891997";
        public static string ftpHost = "LocalHost";
        public static string ftpPost = "21";
        public static string MaDonVi = "TVTK";
        public static IWorkbook workbook;
        public static Worksheet worksheet;
        public static string pathTemplate_Ngay = pathTemplate + "TVTK_THUC HIEN NGAY.xltx";
        public static string pathTemplate_Thang = pathTemplate + "TVTK_KE HOACH THANG.xltx";
        public static string pathTemplate_ChotThang = pathTemplate + "TVTK_THUC HIEN THANG CHOT.xltx";
        public static string pathTemplate_Quy = pathTemplate + "TVTK_THUC HIEN QUY.xltx";
        public static string pathTemplate_Nam = pathTemplate + "TVTK_THUC HIEN NAM.xltx";

    }
}
