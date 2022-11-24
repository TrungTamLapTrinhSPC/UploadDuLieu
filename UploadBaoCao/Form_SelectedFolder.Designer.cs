namespace UploadBaoCao
{
    partial class Form_SelectedFolder
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txt_Folder = new DevExpress.XtraEditors.TextEdit();
            this.btn_Select = new DevExpress.XtraEditors.SimpleButton();
            this.btn_OK = new DevExpress.XtraEditors.SimpleButton();
            this.btn_Cancel = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.txt_Folder.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Thư mục lưu trữ dữ liệu:";
            // 
            // txt_Folder
            // 
            this.txt_Folder.Location = new System.Drawing.Point(15, 25);
            this.txt_Folder.Name = "txt_Folder";
            this.txt_Folder.Size = new System.Drawing.Size(418, 20);
            this.txt_Folder.TabIndex = 1;
            // 
            // btn_Select
            // 
            this.btn_Select.Location = new System.Drawing.Point(439, 23);
            this.btn_Select.Name = "btn_Select";
            this.btn_Select.Size = new System.Drawing.Size(75, 23);
            this.btn_Select.TabIndex = 2;
            this.btn_Select.Text = "Browse";
            this.btn_Select.Click += new System.EventHandler(this.btn_Select_Click);
            // 
            // btn_OK
            // 
            this.btn_OK.Location = new System.Drawing.Point(148, 51);
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.Size = new System.Drawing.Size(75, 23);
            this.btn_OK.TabIndex = 3;
            this.btn_OK.Text = "OK";
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(229, 51);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_Cancel.TabIndex = 4;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // Form_SelectedFolder
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 81);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_OK);
            this.Controls.Add(this.btn_Select);
            this.Controls.Add(this.txt_Folder);
            this.Controls.Add(this.label1);
            this.IconOptions.Image = global::UploadBaoCao.Properties.Resources.folders;
            this.Name = "Form_SelectedFolder";
            this.Text = "Chọn thư mục";
            this.Load += new System.EventHandler(this.Form_SelectedFolder_Load);
            ((System.ComponentModel.ISupportInitialize)(this.txt_Folder.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private DevExpress.XtraEditors.TextEdit txt_Folder;
        private DevExpress.XtraEditors.SimpleButton btn_Select;
        private DevExpress.XtraEditors.SimpleButton btn_OK;
        private DevExpress.XtraEditors.SimpleButton btn_Cancel;
    }
}