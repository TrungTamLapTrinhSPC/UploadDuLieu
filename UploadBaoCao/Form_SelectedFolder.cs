using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace UploadBaoCao
{
    public partial class Form_SelectedFolder : DevExpress.XtraEditors.XtraForm
    {
        public Form_SelectedFolder()
        {
            InitializeComponent();
        }

        private void btn_Select_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txt_Folder.Text = dlg.SelectedPath;

            }
        }

        private void Form_SelectedFolder_Load(object sender, EventArgs e)
        {
            this.CenterToParent();
            txt_Folder.Text= Bien.pathProject;
        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            Bien.pathProject = txt_Folder.Text;
            this.Close();
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();

        }
    }
}