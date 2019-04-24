namespace myStatisticalAnalysis
{
    partial class Login
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
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.btLogin = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // tbPassword
            // 
            this.tbPassword.Location = new System.Drawing.Point(59, 92);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.Size = new System.Drawing.Size(154, 20);
            this.tbPassword.TabIndex = 0;
            // 
            // btLogin
            // 
            this.btLogin.Location = new System.Drawing.Point(24, 172);
            this.btLogin.Name = "btLogin";
            this.btLogin.Size = new System.Drawing.Size(84, 42);
            this.btLogin.TabIndex = 1;
            this.btLogin.Text = "登录";
            this.btLogin.UseVisualStyleBackColor = true;
            this.btLogin.Click += new System.EventHandler(this.btLogin_Click);
            // 
            // btCancel
            // 
            this.btCancel.Location = new System.Drawing.Point(169, 172);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(81, 42);
            this.btCancel.TabIndex = 2;
            this.btCancel.Text = "取消";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(57, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "请输入管理密码：";
            // 
            // Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 288);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btLogin);
            this.Controls.Add(this.tbPassword);
            this.Name = "Login";
            this.Text = "Login";
            this.Shown += new System.EventHandler(this.LoginShown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbPassword;
        private System.Windows.Forms.Button btLogin;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Label label1;
    }
}