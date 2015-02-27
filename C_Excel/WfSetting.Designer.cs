namespace C_Excel
{
    partial class WfSetting
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
            this.label2 = new System.Windows.Forms.Label();
            this.T_ShowUp = new System.Windows.Forms.TextBox();
            this.T_Dissmis = new System.Windows.Forms.TextBox();
            this.B_Valide = new System.Windows.Forms.Button();
            this.B_Cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "最晚打卡时间：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "最早下班时间：";
            // 
            // T_ShowUp
            // 
            this.T_ShowUp.Location = new System.Drawing.Point(105, 21);
            this.T_ShowUp.Name = "T_ShowUp";
            this.T_ShowUp.Size = new System.Drawing.Size(100, 21);
            this.T_ShowUp.TabIndex = 2;
            // 
            // T_Dissmis
            // 
            this.T_Dissmis.Location = new System.Drawing.Point(105, 69);
            this.T_Dissmis.Name = "T_Dissmis";
            this.T_Dissmis.Size = new System.Drawing.Size(100, 21);
            this.T_Dissmis.TabIndex = 3;
            // 
            // B_Valide
            // 
            this.B_Valide.Location = new System.Drawing.Point(24, 113);
            this.B_Valide.Name = "B_Valide";
            this.B_Valide.Size = new System.Drawing.Size(75, 23);
            this.B_Valide.TabIndex = 4;
            this.B_Valide.Text = "应用";
            this.B_Valide.UseVisualStyleBackColor = true;
            this.B_Valide.Click += new System.EventHandler(this.B_Valide_Click);
            // 
            // B_Cancel
            // 
            this.B_Cancel.Location = new System.Drawing.Point(130, 113);
            this.B_Cancel.Name = "B_Cancel";
            this.B_Cancel.Size = new System.Drawing.Size(75, 23);
            this.B_Cancel.TabIndex = 5;
            this.B_Cancel.Text = "取消";
            this.B_Cancel.UseVisualStyleBackColor = true;
            this.B_Cancel.Click += new System.EventHandler(this.B_Cancel_Click);
            // 
            // WfSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(229, 154);
            this.Controls.Add(this.B_Cancel);
            this.Controls.Add(this.B_Valide);
            this.Controls.Add(this.T_Dissmis);
            this.Controls.Add(this.T_ShowUp);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "WfSetting";
            this.Text = "设置";
            this.Load += new System.EventHandler(this.WfSetting_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox T_ShowUp;
        private System.Windows.Forms.TextBox T_Dissmis;
        private System.Windows.Forms.Button B_Valide;
        private System.Windows.Forms.Button B_Cancel;
    }
}