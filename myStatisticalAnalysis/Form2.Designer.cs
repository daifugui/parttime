namespace myStatisticalAnalysis
{
    partial class Form2
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
            this.btSave = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.btLoad = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btClear = new System.Windows.Forms.Button();
            this.btDelete = new System.Windows.Forms.Button();
            this.btAdd = new System.Windows.Forms.Button();
            this.btTest = new System.Windows.Forms.Button();
            this.btInsertImage = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // btSave
            // 
            this.btSave.Location = new System.Drawing.Point(21, 362);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(106, 46);
            this.btSave.TabIndex = 0;
            this.btSave.Text = "Save";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(0, 0);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 23;
            this.dataGridView2.Size = new System.Drawing.Size(831, 301);
            this.dataGridView2.TabIndex = 1;
            // 
            // btLoad
            // 
            this.btLoad.Location = new System.Drawing.Point(0, 428);
            this.btLoad.Name = "btLoad";
            this.btLoad.Size = new System.Drawing.Size(89, 46);
            this.btLoad.TabIndex = 2;
            this.btLoad.Text = "load";
            this.btLoad.UseVisualStyleBackColor = true;
            this.btLoad.Visible = false;
            this.btLoad.Click += new System.EventHandler(this.btLoad_Click);
            // 
            // btCancel
            // 
            this.btCancel.Location = new System.Drawing.Point(682, 363);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(108, 46);
            this.btCancel.TabIndex = 3;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // btClear
            // 
            this.btClear.Location = new System.Drawing.Point(548, 362);
            this.btClear.Name = "btClear";
            this.btClear.Size = new System.Drawing.Size(102, 45);
            this.btClear.TabIndex = 4;
            this.btClear.Text = "Clear";
            this.btClear.UseVisualStyleBackColor = true;
            this.btClear.Click += new System.EventHandler(this.btClear_Click);
            // 
            // btDelete
            // 
            this.btDelete.Location = new System.Drawing.Point(416, 362);
            this.btDelete.Name = "btDelete";
            this.btDelete.Size = new System.Drawing.Size(104, 45);
            this.btDelete.TabIndex = 5;
            this.btDelete.Text = "Delete";
            this.btDelete.UseVisualStyleBackColor = true;
            this.btDelete.Click += new System.EventHandler(this.btDelete_Click);
            // 
            // btAdd
            // 
            this.btAdd.Location = new System.Drawing.Point(153, 363);
            this.btAdd.Name = "btAdd";
            this.btAdd.Size = new System.Drawing.Size(96, 45);
            this.btAdd.TabIndex = 6;
            this.btAdd.Text = "Add";
            this.btAdd.UseVisualStyleBackColor = true;
            this.btAdd.Click += new System.EventHandler(this.btAdd_Click);
            // 
            // btTest
            // 
            this.btTest.Location = new System.Drawing.Point(110, 428);
            this.btTest.Name = "btTest";
            this.btTest.Size = new System.Drawing.Size(92, 46);
            this.btTest.TabIndex = 7;
            this.btTest.Text = "Test";
            this.btTest.UseVisualStyleBackColor = true;
            this.btTest.Visible = false;
            this.btTest.Click += new System.EventHandler(this.btTest_Click);
            // 
            // btInsertImage
            // 
            this.btInsertImage.Location = new System.Drawing.Point(274, 362);
            this.btInsertImage.Name = "btInsertImage";
            this.btInsertImage.Size = new System.Drawing.Size(106, 45);
            this.btInsertImage.TabIndex = 8;
            this.btInsertImage.Text = "Insert Image";
            this.btInsertImage.UseVisualStyleBackColor = true;
            this.btInsertImage.Click += new System.EventHandler(this.btInsertImage_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form2
            // 
            this.ClientSize = new System.Drawing.Size(831, 472);
            this.Controls.Add(this.btInsertImage);
            this.Controls.Add(this.btTest);
            this.Controls.Add(this.btAdd);
            this.Controls.Add(this.btDelete);
            this.Controls.Add(this.btClear);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btLoad);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.btSave);
            this.Name = "Form2";
            this.Text = "后台管理";
            this.Shown += new System.EventHandler(this.form2shown);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btSave;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button btLoad;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btClear;
        private System.Windows.Forms.Button btDelete;
        private System.Windows.Forms.Button btAdd;
        private System.Windows.Forms.Button btTest;
        private System.Windows.Forms.Button btInsertImage;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}