namespace testOutlookAddIn
{
    partial class ParametersForm
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
            this.txtSendUrl = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtBUListUrl = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtCategoryListUrl = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.cmbCategoryName = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtSendUrl
            // 
            this.txtSendUrl.Location = new System.Drawing.Point(208, 12);
            this.txtSendUrl.Name = "txtSendUrl";
            this.txtSendUrl.Size = new System.Drawing.Size(502, 20);
            this.txtSendUrl.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "URL для отправки заявки";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "URL получения списка БЕ";
            // 
            // txtBUListUrl
            // 
            this.txtBUListUrl.Location = new System.Drawing.Point(208, 38);
            this.txtBUListUrl.Name = "txtBUListUrl";
            this.txtBUListUrl.Size = new System.Drawing.Size(502, 20);
            this.txtBUListUrl.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(179, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "URL получения списка Категорий";
            // 
            // txtCategoryListUrl
            // 
            this.txtCategoryListUrl.Location = new System.Drawing.Point(208, 64);
            this.txtCategoryListUrl.Name = "txtCategoryListUrl";
            this.txtCategoryListUrl.Size = new System.Drawing.Size(502, 20);
            this.txtCategoryListUrl.TabIndex = 4;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(337, 128);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // cmbCategoryName
            // 
            this.cmbCategoryName.FormattingEnabled = true;
            this.cmbCategoryName.Location = new System.Drawing.Point(208, 91);
            this.cmbCategoryName.Name = "cmbCategoryName";
            this.cmbCategoryName.Size = new System.Drawing.Size(502, 21);
            this.cmbCategoryName.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(23, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(168, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Категория для пометки письма";
            // 
            // ParametersForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 163);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbCategoryName);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtCategoryListUrl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtBUListUrl);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtSendUrl);
            this.Name = "ParametersForm";
            this.Text = "Параметры надстройки";
            this.Load += new System.EventHandler(this.ParametersForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSendUrl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtBUListUrl;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtCategoryListUrl;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.ComboBox cmbCategoryName;
        private System.Windows.Forms.Label label4;
    }
}