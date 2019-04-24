namespace myStatisticalAnalysis
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ManageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sPCToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.monitorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.printPreviewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.printToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.printSetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openHelpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.zedGraphControl1 = new ZedGraph.ZedGraphControl();
            this.zedGraphControl2 = new ZedGraph.ZedGraphControl();
            this.label1 = new System.Windows.Forms.Label();
            this.zedGraphControl3 = new ZedGraph.ZedGraphControl();
            this.textUSL = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textLSL = new System.Windows.Forms.TextBox();
            this.btConnect = new System.Windows.Forms.Button();
            this.textRealTimeData = new System.Windows.Forms.TextBox();
            this.disConnect = new System.Windows.Forms.Button();
            this.btClearData = new System.Windows.Forms.Button();
            this.btSPC = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.logoPicBox = new System.Windows.Forms.PictureBox();
            this.productPicBox = new System.Windows.Forms.PictureBox();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBox4 = new System.Windows.Forms.ComboBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label9 = new System.Windows.Forms.Label();
            this.textSubGroupSize = new System.Windows.Forms.TextBox();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.btSaveData2DataBase = new System.Windows.Forms.Button();
            this.btLoadData = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.btDelete = new System.Windows.Forms.Button();
            this.btMonthTrend = new System.Windows.Forms.Button();
            this.comboBox5 = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.logoPicBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.productPicBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileFToolStripMenuItem,
            this.ManageToolStripMenuItem,
            this.sPCToolStripMenuItem,
            this.monitorToolStripMenuItem,
            this.testToolStripMenuItem,
            this.printToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1428, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileFToolStripMenuItem
            // 
            this.fileFToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openpToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileFToolStripMenuItem.Name = "fileFToolStripMenuItem";
            this.fileFToolStripMenuItem.Size = new System.Drawing.Size(53, 21);
            this.fileFToolStripMenuItem.Text = "File(&F)";
            this.fileFToolStripMenuItem.Click += new System.EventHandler(this.fileFToolStripMenuItem_Click);
            // 
            // openpToolStripMenuItem
            // 
            this.openpToolStripMenuItem.Name = "openpToolStripMenuItem";
            this.openpToolStripMenuItem.Size = new System.Drawing.Size(126, 22);
            this.openpToolStripMenuItem.Text = "Open(&O)";
            this.openpToolStripMenuItem.Click += new System.EventHandler(this.openpToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(126, 22);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(126, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // ManageToolStripMenuItem
            // 
            this.ManageToolStripMenuItem.Name = "ManageToolStripMenuItem";
            this.ManageToolStripMenuItem.Size = new System.Drawing.Size(68, 21);
            this.ManageToolStripMenuItem.Text = "后台管理";
            this.ManageToolStripMenuItem.Click += new System.EventHandler(this.ManageToolStripMenuItem_Click);
            // 
            // sPCToolStripMenuItem
            // 
            this.sPCToolStripMenuItem.Enabled = false;
            this.sPCToolStripMenuItem.Name = "sPCToolStripMenuItem";
            this.sPCToolStripMenuItem.Size = new System.Drawing.Size(42, 21);
            this.sPCToolStripMenuItem.Text = "SPC";
            this.sPCToolStripMenuItem.Visible = false;
            this.sPCToolStripMenuItem.Click += new System.EventHandler(this.sPCToolStripMenuItem_Click);
            // 
            // monitorToolStripMenuItem
            // 
            this.monitorToolStripMenuItem.Enabled = false;
            this.monitorToolStripMenuItem.Name = "monitorToolStripMenuItem";
            this.monitorToolStripMenuItem.Size = new System.Drawing.Size(67, 21);
            this.monitorToolStripMenuItem.Text = "Monitor";
            this.monitorToolStripMenuItem.Visible = false;
            this.monitorToolStripMenuItem.Click += new System.EventHandler(this.monitorToolStripMenuItem_Click);
            // 
            // testToolStripMenuItem
            // 
            this.testToolStripMenuItem.Name = "testToolStripMenuItem";
            this.testToolStripMenuItem.Size = new System.Drawing.Size(41, 21);
            this.testToolStripMenuItem.Text = "test";
            this.testToolStripMenuItem.Visible = false;
            this.testToolStripMenuItem.Click += new System.EventHandler(this.testToolStripMenuItem_Click);
            // 
            // printToolStripMenuItem
            // 
            this.printToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.printPreviewToolStripMenuItem,
            this.printToolStripMenuItem1,
            this.printSetToolStripMenuItem});
            this.printToolStripMenuItem.Name = "printToolStripMenuItem";
            this.printToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.printToolStripMenuItem.Text = "打印";
            // 
            // printPreviewToolStripMenuItem
            // 
            this.printPreviewToolStripMenuItem.Name = "printPreviewToolStripMenuItem";
            this.printPreviewToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.printPreviewToolStripMenuItem.Text = "打印预览";
            this.printPreviewToolStripMenuItem.Click += new System.EventHandler(this.printPreviewToolStripMenuItem_Click);
            // 
            // printToolStripMenuItem1
            // 
            this.printToolStripMenuItem1.Name = "printToolStripMenuItem1";
            this.printToolStripMenuItem1.Size = new System.Drawing.Size(124, 22);
            this.printToolStripMenuItem1.Text = "打印";
            this.printToolStripMenuItem1.Click += new System.EventHandler(this.printToolStripMenuItem1_Click);
            // 
            // printSetToolStripMenuItem
            // 
            this.printSetToolStripMenuItem.Name = "printSetToolStripMenuItem";
            this.printSetToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.printSetToolStripMenuItem.Text = "PrintSet";
            this.printSetToolStripMenuItem.Visible = false;
            this.printSetToolStripMenuItem.Click += new System.EventHandler(this.printSetToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openHelpToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.helpToolStripMenuItem.Text = "帮助";
            // 
            // openHelpToolStripMenuItem
            // 
            this.openHelpToolStripMenuItem.Name = "openHelpToolStripMenuItem";
            this.openHelpToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.openHelpToolStripMenuItem.Text = "打开帮助文档";
            this.openHelpToolStripMenuItem.Click += new System.EventHandler(this.openHelpToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.aboutToolStripMenuItem.Text = "关于";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(3, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(334, 120);
            this.dataGridView1.TabIndex = 1;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(0, 353);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(136, 26);
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // zedGraphControl1
            // 
            this.zedGraphControl1.BackColor = System.Drawing.SystemColors.Control;
            this.zedGraphControl1.IsShowPointValues = false;
            this.zedGraphControl1.Location = new System.Drawing.Point(0, 27);
            this.zedGraphControl1.Name = "zedGraphControl1";
            this.zedGraphControl1.PointValueFormat = "G";
            this.zedGraphControl1.Size = new System.Drawing.Size(761, 254);
            this.zedGraphControl1.TabIndex = 4;
            // 
            // zedGraphControl2
            // 
            this.zedGraphControl2.IsShowPointValues = false;
            this.zedGraphControl2.Location = new System.Drawing.Point(0, 286);
            this.zedGraphControl2.Name = "zedGraphControl2";
            this.zedGraphControl2.PointValueFormat = "G";
            this.zedGraphControl2.Size = new System.Drawing.Size(761, 297);
            this.zedGraphControl2.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(1276, 169);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "参数显示";
            // 
            // zedGraphControl3
            // 
            this.zedGraphControl3.IsShowPointValues = false;
            this.zedGraphControl3.Location = new System.Drawing.Point(793, 488);
            this.zedGraphControl3.Name = "zedGraphControl3";
            this.zedGraphControl3.PointValueFormat = "G";
            this.zedGraphControl3.Size = new System.Drawing.Size(627, 355);
            this.zedGraphControl3.TabIndex = 7;
            // 
            // textUSL
            // 
            this.textUSL.Enabled = false;
            this.textUSL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textUSL.Location = new System.Drawing.Point(842, 297);
            this.textUSL.Name = "textUSL";
            this.textUSL.Size = new System.Drawing.Size(100, 22);
            this.textUSL.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(799, 306);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 16);
            this.label2.TabIndex = 9;
            this.label2.Text = "USL";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label3.Location = new System.Drawing.Point(799, 337);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(33, 17);
            this.label3.TabIndex = 10;
            this.label3.Text = "LSL";
            // 
            // textLSL
            // 
            this.textLSL.Enabled = false;
            this.textLSL.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.textLSL.Location = new System.Drawing.Point(842, 334);
            this.textLSL.Name = "textLSL";
            this.textLSL.Size = new System.Drawing.Size(100, 23);
            this.textLSL.TabIndex = 11;
            // 
            // btConnect
            // 
            this.btConnect.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btConnect.Location = new System.Drawing.Point(801, 397);
            this.btConnect.Name = "btConnect";
            this.btConnect.Size = new System.Drawing.Size(97, 32);
            this.btConnect.TabIndex = 12;
            this.btConnect.Text = "连接";
            this.btConnect.UseVisualStyleBackColor = true;
            this.btConnect.Click += new System.EventHandler(this.btConnect_Click);
            // 
            // textRealTimeData
            // 
            this.textRealTimeData.Font = new System.Drawing.Font("SimSun", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textRealTimeData.Location = new System.Drawing.Point(898, 169);
            this.textRealTimeData.Name = "textRealTimeData";
            this.textRealTimeData.ReadOnly = true;
            this.textRealTimeData.Size = new System.Drawing.Size(236, 38);
            this.textRealTimeData.TabIndex = 14;
            // 
            // disConnect
            // 
            this.disConnect.Font = new System.Drawing.Font("SimSun", 14F, System.Drawing.FontStyle.Bold);
            this.disConnect.Location = new System.Drawing.Point(926, 397);
            this.disConnect.Name = "disConnect";
            this.disConnect.Size = new System.Drawing.Size(98, 34);
            this.disConnect.TabIndex = 15;
            this.disConnect.Text = "断开";
            this.disConnect.UseVisualStyleBackColor = true;
            this.disConnect.Click += new System.EventHandler(this.disConnect_Click);
            // 
            // btClearData
            // 
            this.btClearData.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold);
            this.btClearData.Location = new System.Drawing.Point(1057, 397);
            this.btClearData.Name = "btClearData";
            this.btClearData.Size = new System.Drawing.Size(93, 34);
            this.btClearData.TabIndex = 16;
            this.btClearData.Text = "数据清空";
            this.btClearData.UseVisualStyleBackColor = true;
            this.btClearData.Click += new System.EventHandler(this.btClearData_Click);
            // 
            // btSPC
            // 
            this.btSPC.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold);
            this.btSPC.Location = new System.Drawing.Point(1178, 397);
            this.btSPC.Name = "btSPC";
            this.btSPC.Size = new System.Drawing.Size(104, 34);
            this.btSPC.TabIndex = 17;
            this.btSPC.Text = "SPC";
            this.btSPC.UseVisualStyleBackColor = true;
            this.btSPC.Click += new System.EventHandler(this.btSPC_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("SimSun", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(797, 179);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 19);
            this.label4.TabIndex = 18;
            this.label4.Text = "实时数据:";
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(793, 238);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(99, 20);
            this.comboBox1.TabIndex = 19;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1Changed);
            // 
            // comboBox2
            // 
            this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox2.DropDownWidth = 130;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(913, 238);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(95, 20);
            this.comboBox2.TabIndex = 20;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBoxChanged);
            // 
            // comboBox3
            // 
            this.comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Items.AddRange(new object[] {
            "护管",
            "芯子"});
            this.comboBox3.Location = new System.Drawing.Point(1028, 238);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(86, 20);
            this.comboBox3.TabIndex = 21;
            this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBoxChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(796, 220);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(68, 16);
            this.label5.TabIndex = 22;
            this.label5.Text = "选择车型";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(910, 220);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 16);
            this.label6.TabIndex = 23;
            this.label6.Text = "产品编号";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(1025, 220);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(72, 16);
            this.label7.TabIndex = 24;
            this.label7.Text = "护管/芯子";
            // 
            // logoPicBox
            // 
            this.logoPicBox.Image = ((System.Drawing.Image)(resources.GetObject("logoPicBox.Image")));
            this.logoPicBox.Location = new System.Drawing.Point(1144, 27);
            this.logoPicBox.Name = "logoPicBox";
            this.logoPicBox.Size = new System.Drawing.Size(288, 130);
            this.logoPicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.logoPicBox.TabIndex = 25;
            this.logoPicBox.TabStop = false;
            // 
            // productPicBox
            // 
            this.productPicBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.productPicBox.Location = new System.Drawing.Point(1100, 306);
            this.productPicBox.Name = "productPicBox";
            this.productPicBox.Size = new System.Drawing.Size(161, 71);
            this.productPicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.productPicBox.TabIndex = 26;
            this.productPicBox.TabStop = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(1097, 292);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(68, 16);
            this.label8.TabIndex = 27;
            this.label8.Text = "产品图片";
            // 
            // comboBox4
            // 
            this.comboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox4.FormattingEnabled = true;
            this.comboBox4.Location = new System.Drawing.Point(801, 366);
            this.comboBox4.Name = "comboBox4";
            this.comboBox4.Size = new System.Drawing.Size(97, 20);
            this.comboBox4.TabIndex = 28;
            this.comboBox4.SelectedIndexChanged += new System.EventHandler(this.COMChanged);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(0, 585);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            this.dataGridView2.RowTemplate.Height = 23;
            this.dataGridView2.Size = new System.Drawing.Size(761, 267);
            this.dataGridView2.TabIndex = 29;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(793, 19);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(345, 144);
            this.tabControl1.TabIndex = 30;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(337, 118);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "原始数据";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(967, 322);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(71, 16);
            this.label9.TabIndex = 31;
            this.label9.Text = "子组大小:";
            // 
            // textSubGroupSize
            // 
            this.textSubGroupSize.Enabled = false;
            this.textSubGroupSize.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textSubGroupSize.Location = new System.Drawing.Point(968, 337);
            this.textSubGroupSize.Name = "textSubGroupSize";
            this.textSubGroupSize.Size = new System.Drawing.Size(100, 22);
            this.textSubGroupSize.TabIndex = 32;
            this.textSubGroupSize.Text = "5";
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            // 
            // printDialog1
            // 
            this.printDialog1.UseEXDialog = true;
            // 
            // btSaveData2DataBase
            // 
            this.btSaveData2DataBase.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold);
            this.btSaveData2DataBase.Location = new System.Drawing.Point(801, 439);
            this.btSaveData2DataBase.Name = "btSaveData2DataBase";
            this.btSaveData2DataBase.Size = new System.Drawing.Size(97, 36);
            this.btSaveData2DataBase.TabIndex = 33;
            this.btSaveData2DataBase.Text = "保存数据";
            this.btSaveData2DataBase.UseVisualStyleBackColor = true;
            this.btSaveData2DataBase.Click += new System.EventHandler(this.btSaveData2DataBase_Click);
            // 
            // btLoadData
            // 
            this.btLoadData.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold);
            this.btLoadData.Location = new System.Drawing.Point(1057, 439);
            this.btLoadData.Name = "btLoadData";
            this.btLoadData.Size = new System.Drawing.Size(93, 36);
            this.btLoadData.TabIndex = 34;
            this.btLoadData.Text = "打开数据";
            this.btLoadData.UseVisualStyleBackColor = true;
            this.btLoadData.Click += new System.EventHandler(this.btLoadData_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(926, 439);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(110, 21);
            this.dateTimePicker1.TabIndex = 35;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(927, 467);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(109, 21);
            this.dateTimePicker2.TabIndex = 36;
            // 
            // btDelete
            // 
            this.btDelete.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold);
            this.btDelete.Location = new System.Drawing.Point(1178, 439);
            this.btDelete.Name = "btDelete";
            this.btDelete.Size = new System.Drawing.Size(104, 36);
            this.btDelete.TabIndex = 37;
            this.btDelete.Text = "删除记录";
            this.btDelete.UseVisualStyleBackColor = true;
            this.btDelete.Click += new System.EventHandler(this.btDelete_Click);
            // 
            // btMonthTrend
            // 
            this.btMonthTrend.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btMonthTrend.Location = new System.Drawing.Point(1311, 400);
            this.btMonthTrend.Name = "btMonthTrend";
            this.btMonthTrend.Size = new System.Drawing.Size(105, 31);
            this.btMonthTrend.TabIndex = 38;
            this.btMonthTrend.Text = "月平均图";
            this.btMonthTrend.UseVisualStyleBackColor = true;
            this.btMonthTrend.Click += new System.EventHandler(this.btMonthTrend_Click);
            // 
            // comboBox5
            // 
            this.comboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox5.FormattingEnabled = true;
            this.comboBox5.Items.AddRange(new object[] {
            "护管",
            "芯子"});
            this.comboBox5.Location = new System.Drawing.Point(1133, 238);
            this.comboBox5.Name = "comboBox5";
            this.comboBox5.Size = new System.Drawing.Size(86, 20);
            this.comboBox5.TabIndex = 21;
            this.comboBox5.SelectedIndexChanged += new System.EventHandler(this.comboBoxChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(1130, 220);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(72, 16);
            this.label10.TabIndex = 24;
            this.label10.Text = "护管/芯子";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(790, 266);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(68, 16);
            this.label11.TabIndex = 40;
            this.label11.Text = "生产数量";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(860, 263);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(118, 22);
            this.textBox1.TabIndex = 39;
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(1065, 263);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(154, 22);
            this.textBox2.TabIndex = 39;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(1002, 266);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(57, 16);
            this.label12.TabIndex = 40;
            this.label12.Text = "LOT NO";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1428, 823);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btMonthTrend);
            this.Controls.Add(this.btDelete);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.btLoadData);
            this.Controls.Add(this.btSaveData2DataBase);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.textSubGroupSize);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.comboBox4);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.productPicBox);
            this.Controls.Add(this.logoPicBox);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBox5);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.comboBox3);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btSPC);
            this.Controls.Add(this.btClearData);
            this.Controls.Add(this.disConnect);
            this.Controls.Add(this.textRealTimeData);
            this.Controls.Add(this.btConnect);
            this.Controls.Add(this.textLSL);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textUSL);
            this.Controls.Add(this.zedGraphControl3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.zedGraphControl2);
            this.Controls.Add(this.zedGraphControl1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "SPC                    版权：精昌科技   022-29440088";
            this.Deactivate += new System.EventHandler(this.Form1_Deactivate);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.formclosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.logoPicBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.productPicBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ToolStripMenuItem sPCToolStripMenuItem;
        private System.Windows.Forms.PictureBox pictureBox1;
        private ZedGraph.ZedGraphControl zedGraphControl1;
        private ZedGraph.ZedGraphControl zedGraphControl2;
        private System.Windows.Forms.Label label1;
        private ZedGraph.ZedGraphControl zedGraphControl3;
        private System.Windows.Forms.ToolStripMenuItem monitorToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.TextBox textUSL;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textLSL;
        private System.Windows.Forms.Button btConnect;
        private System.Windows.Forms.TextBox textRealTimeData;
        private System.Windows.Forms.Button disConnect;
        private System.Windows.Forms.Button btClearData;
        private System.Windows.Forms.Button btSPC;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ToolStripMenuItem ManageToolStripMenuItem;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.PictureBox logoPicBox;
        private System.Windows.Forms.PictureBox productPicBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ToolStripMenuItem testToolStripMenuItem;
        private System.Windows.Forms.ComboBox comboBox4;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textSubGroupSize;
        private System.Windows.Forms.ToolStripMenuItem printToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem printPreviewToolStripMenuItem;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Windows.Forms.ToolStripMenuItem printToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem printSetToolStripMenuItem;
        private System.Windows.Forms.Button btSaveData2DataBase;
        private System.Windows.Forms.Button btLoadData;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Button btDelete;
        private System.Windows.Forms.Button btMonthTrend;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openHelpToolStripMenuItem;
        private System.Windows.Forms.ComboBox comboBox5;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label12;
    }
}

