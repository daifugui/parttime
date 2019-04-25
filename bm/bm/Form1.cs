using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace bm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Graphics g = pictureBox1.CreateGraphics();
            // g.DrawLine(new Pen(Color.Blue, 2), 10, 10, 50, 50);
            // g.DrawLine(new Pen(Color.Blue, 2), (float)0.1, (float)0.1, (float)0.5, (float)0.5);
            draw_dis_line(102, 32, 290, 120);

        }

        private void draw_dis_line(int x1, int y1, int x2, int y2)
        {
            Graphics g = pictureBox1.CreateGraphics();
            g.DrawString("直径1", new Font("Microsoft Sans Serif", 20), new SolidBrush(Color.Red), x1, y1);
            g.DrawLine(new Pen(Color.Blue, 2), x1, y1, x2, y2);
            double d = 5;
            if (y2 - y1 != 0)
            {
                double k = -(float)(x2 - x1) / (y2 - y1);
                double dx = Math.Sqrt(d * d / (1 + k * k));
                int x11 = x1 - (int)dx;
                int y11 = (int)(y1 - k * dx);
                int x12 = x1 + (int)dx;
                int y12 = (int)(y1 + k * dx);
                g.DrawLine(new Pen(Color.Blue, 2), x11, y11, x12, y12);

                int x21 = x2 - (int)dx;
                int y21 = (int)(y2 - k * dx);
                int x22 = x2 + (int)dx;
                int y22 = (int)(y2 + k * dx);
                g.DrawLine(new Pen(Color.Blue, 2), x21, y21, x22, y22);

            }
            else
            {
                g.DrawLine(new Pen(Color.Blue, 2), x1, (int)(y1 - d), x1, (int)(y1 + d));
                g.DrawLine(new Pen(Color.Blue, 2), x2, (int)(y2 - d), x2, (int)(y2 + d));
            }

        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
           // draw_dis_line(10, 20, 150, 170);
           // Graphics g = pictureBox1.CreateGraphics();
           // g.DrawLine(new Pen(Color.Blue, 2), 10, 10, 50, 50);

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        public void computeRangeMeanStd(double[] testdata, int subgroupsize, ArrayList R, ArrayList mean, ArrayList Std)
        {
            int kk = 0;
            double maxX = 0;
            double minX = 0;
            // int m=0;
            R.Clear();
            double sum = 0;
            double sum2 = 0;
            foreach (double x in testdata)
            {
                if (kk == 0)
                {
                    maxX = x;
                    minX = x;
                    sum = 0;
                    sum2 = 0;
                }
                if (x > maxX) maxX = x;
                else if (x < minX) minX = x;
                sum += x;
                sum2 += x * x;
                kk++;
                if (kk == subgroupsize)
                {
                    // R[m++] = maxX - minX;
                    R.Add(maxX - minX);
                    mean.Add(sum / subgroupsize);
                    Std.Add(Math.Sqrt((sum2 - sum * sum / subgroupsize) / (subgroupsize - 1)));
                    sum = 0;
                    kk = 0;
                }
            }
            //     int nn=testdata.Length;
            //   for(int i=0;i<nn;i++)
            // {

            // }

        }
        public void computeRangeMean(double[] testdata, int subgroupsize, ArrayList R, ArrayList mean)
        {
            int kk = 0;
            double maxX = 0;
            double minX = 0;
            // int m=0;
            R.Clear();
            double sum = 0;
            foreach (double x in testdata)
            {
                if (kk == 0)
                {
                    maxX = x;
                    minX = x;
                    sum = 0;
                }
                if (x > maxX) maxX = x;
                else if (x < minX) minX = x;
                sum += x;
                kk++;
                if (kk == subgroupsize)
                {
                    // R[m++] = maxX - minX;
                    R.Add(maxX - minX);
                    mean.Add(sum / subgroupsize);
                    sum = 0;
                    kk = 0;
                }
            }
            //     int nn=testdata.Length;
            //   for(int i=0;i<nn;i++)
            // {

            // }

        }
        /*
         25 0.600 0.153 0.606 0.9896 1.0105 0.565 1.435 0.559 1.420 3.931 0.2544 0.708 1.806 6.056 0.459 1.541
         24 0.612 0.157 0.619 0.9892 1.0109 0.555 1.445 0.549 1.429 3.895 0.2567 0.712 1.759 6.031 0.451 1.548
         23 0.626 0.162 0.633 0.9887 1.0114 0.545 1.455 0.539 1.438 3.858 0.2592 0.716 1.710 6.006 0.443 1.557
         22 0.640 0.167 0.647 0.9882 1.0119 0.534 1.466 0.528 1.448 3.819 0.2618 0.720 1.659 5.979 0.434 1.566
         21 0.655 0.173 0.663 0.9876 1.0126 0.523 1.477 0.516 1.459 3.778 0.2647 0.724 1.605 5.951 0.425 1.575
         20 0.671 0.180 0.680 0.9869 1.0133 0.510 1.490 0.504 1.470 3.735 0.2677 0.729 1.549 5.921 0.415 1.585
        19 0.688 0.187 0.698 0.9862 1.0140 0.497 1.503 0.490 1.483 3.689 0.2711 0.734 1.487 5.891 0.403 1.597
        18 0.707 0.194 0.718 0.9854 1.0148 0.482 1.518 0.475 1.496 3.640 0.2747 0.739 1.424 5.856 0.391 1.608
        17 0.728 0.203 0.739 0.9845 1.0157 0.466 1.534 0.458 1.511 3.588 0.2787 0.744 1.356 5.820 0.378 1.622
        16 0.750 0.212 0.763 0.9835 1.0168 0.448 1.552 0.440 1.526 3.532 0.2831 0.750 1.282 5.782 0.363 1.637
        15 0.775 0.223 0.789 0.9823 1.0180 0.428 1.572 0.421 1.544 3.472 0.2880 0.756 1.203 5.741 0.347 1.653
        14 0.802 0.235 0.817 0.9810 1.0194 0.406 1.594 0.399 1.563 3.407 0.2935 0.763 1.118 5.696 0.328 1.672
        13 0.832 0.249 0.850 0.9794 1.0210 0.382 1.618 0.374 1.585 3.336 0.2998 0.770 1.025 5.647 0.307 1.693
        12 0.866 0.266 0.886 0.9776 1.0229 0.354 1.646 0.346 1.610 3.258 0.3069 0.778 0.922 5.594 0.283 1.717
        11 0.905 0.285 0.927 0.9754 1.0252 0.321 1.679 0.313 1.637 3.173 0.3152 0.787 0.811 5.535 0.256 1.744
        10 0.949 0.308 0.975 0.9727 1.0281 0.284 1.716 0.276 1.669 3.078 0.3249 0.797 0.687 5.469 0.223 1.777
        9 1.000 0.337 1.032 0.9693 1.0317 0.239 1.761 0.232 1.707 2.970 0.3367 0.808 0.547 5.393 0.184 1.816
        8 1.061 0.373 1.099 0.9650 1.0363 0.185 1.815 0.179 1.751 2.847 0.3512 0.820 0.388 5.306 0.136 1.864
        7 1.134 0.419 1.182 0.9594 1.0423 0.118 1.882 0.113 1.806 2.704 0.3698 0.833 0.204 5.204 0.076 1.924
        6 1.225 0.483 1.287 0.9515 1.0510 0.030 1.970 0.029 1.874 2.534 0.3946 0.848 0 5.078 0 2.004
        5 1.342 0.577 1.427 0.9400 1.0638 0 2.089 0 1.964 2.326 0.4299 0.864 0 4.918 0 2.114
        4 1.500 0.729 1.628 0.9213 1.0854 0 2.266 0 2.088 2.059 0.4857 0.880 0 4.698 0 2.282
        3 1.732 1.023 1.954 0.8862 1.1284 0 2.568 0 2.276 1.693 0.5907 0.888 0 4.358 0 2.574
        2 2.121 1.880 2.659 0.7979 1.2533 0 3.267 0 2.606 1.128 0.8865 0.853 0 3.686 0 3.267
         
         * */

        double[] SPC_A = { 2.12100000000000, 1.73200000000000, 1.50000000000000, 1.34200000000000, 1.22500000000000, 1.13400000000000, 1.06100000000000, 1, 0.949000000000000, 0.905000000000000, 0.866000000000000, 0.832000000000000, 0.802000000000000, 0.775000000000000, 0.750000000000000, 0.728000000000000, 0.707000000000000, 0.688000000000000, 0.671000000000000, 0.655000000000000, 0.640000000000000, 0.626000000000000, 0.612000000000000, 0.600000000000000 };
        double[] SPC_A2 = { 1.8800, 1.0230, 0.7290, 0.5770, 0.4830, 0.4190, 0.3730, 0.3370, 0.3080, 0.2850, 0.2660, 0.2490, 0.2350, 0.2230, 0.2120, 0.2030, 0.1940, 0.1870, 0.1800, 0.1730, 0.1670, 0.1620, 0.1570, 0.1530 };
        double[] SPC_D3 = { 0, 0, 0, 0, 0, 0.0760, 0.1360, 0.1840, 0.2230, 0.2560, 0.2830, 0.3070, 0.3280, 0.3470, 0.3630, 0.3780, 0.3910, 0.4030, 0.4150, 0.4250, 0.4340, 0.4430, 0.4510, 0.4590 };
        double[] SPC_D4 = { 3.2670, 2.5740, 2.2820, 2.1140, 2.0040, 1.9240, 1.8640, 1.8160, 1.7770, 1.7440, 1.7170, 1.6930, 1.6720, 1.6530, 1.6370, 1.6220, 1.6080, 1.5970, 1.5850, 1.5750, 1.5660, 1.5570, 1.5480, 1.5410 };
        double[] SPC_C4 = { 0.7979, 0.8862, 0.9213, 0.9400, 0.9515, 0.9594, 0.9650, 0.9693, 0.9727, 0.9754, 0.9776, 0.9794, 0.9810, 0.9823, 0.9835, 0.9845, 0.9854, 0.9862, 0.9869, 0.9876, 0.9882, 0.9887, 0.9892, 0.9896 };
        double[] SPC_d2 = { 1.1280, 1.6930, 2.0590, 2.3260, 2.5340, 2.7040, 2.8470, 2.9700, 3.0780, 3.1730, 3.2580, 3.3360, 3.4070, 3.4720, 3.5320, 3.5880, 3.6400, 3.6890, 3.7350, 3.7780, 3.8190, 3.8580, 3.8950, 3.9310 };
        double[] SPC_1_C4 = { 1.2533, 1.1284, 1.0854, 1.0638, 1.0510, 1.0423, 1.0363, 1.0317, 1.0281, 1.0252, 1.0229, 1.0210, 1.0194, 1.0180, 1.0168, 1.0157, 1.0148, 1.0140, 1.0133, 1.0126, 1.0119, 1.0114, 1.0109, 1.0105 };
        double[] SPC_1_d2 = { 0.8865, 0.5907, 0.4857, 0.4299, 0.3946, 0.3698, 0.3512, 0.3367, 0.3249, 0.3152, 0.3069, 0.2998, 0.2935, 0.2880, 0.2831, 0.2787, 0.2747, 0.2711, 0.2677, 0.2647, 0.2618, 0.2592, 0.2567, 0.2544 };
        public bool computeLimit(double[] R, double[] mean, int subgroupsize, ref double Ave_R, ref double Ave_mean, ref double UCL, ref double LCL, ref double UCLR, ref double LCLR)
        {
            if (subgroupsize <= 1) return false;
            int m = R.Length;
            Ave_R = 0;
            Ave_mean = 0;
            for (int i = 0; i < m; i++)
            {
                Ave_R += R[i];
                Ave_mean += mean[i];
            }
            Ave_R /= m;
            // Ave_R = Math.Sqrt(Ave_R);
            Ave_mean /= m;
            UCL = Ave_mean + SPC_A2[subgroupsize - 2] * Ave_R;
            LCL = Ave_mean - SPC_A2[subgroupsize - 2] * Ave_R;

            UCLR = SPC_D4[subgroupsize - 2] * Ave_R;
            LCLR = SPC_D3[subgroupsize - 2] * Ave_R;
            return true;
        }
        double sigma = 0, Cp = 0, Cpu = 0, Cpl = 0, Cpk = 0;
        double UCL = 0;
        double LCL = 0;
        double UCLR = 0;
        double LCLR = 0;
        public bool computeCpk(double USL, double LSL, double Ave_mean, double Ave_R, int subgroupsize, ref double sigma, ref double Cp, ref double Cpu, ref double Cpl, ref double Cpk)
        {
            sigma = Ave_R * SPC_1_d2[subgroupsize - 2];
            Cp = (USL - LSL) / sigma / 6;
            Cpl = (Ave_mean - LSL) / sigma / 3;
            Cpu = (USL - Ave_mean) / sigma / 3;
            Cpk = Math.Min(Cpl, Cpu);
            return true;

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (this.comboBox1.SelectedIndex == 0)

                drawchart1(this.comboBox1.SelectedIndex);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.comboBox1.Items.Add("X_bar");
            this.comboBox1.Items.Add("R");
            this.comboBox2.Items.Add("直径1");
            this.comboBox2.Items.Add("直径2");
            this.comboBox2.Items.Add("直径3");
            this.comboBox2.Items.Add("圆度3");
            this.comboBox2.Items.Add("高度1");
            this.comboBox2.Items.Add("高度2");
            this.comboBox1.SelectedIndex = 0;
            this.comboBox2.SelectedIndex = 0;

            Random rd = new Random();
            for (int i = 0; i < 30; i++)
            {
                int index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString(); ;
                this.dataGridView1.Rows[index].Cells[1].Value = "D1";                
                this.dataGridView1.Rows[index].Cells[2].Value =(rd.Next()%10).ToString();

                index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString(); ;
                this.dataGridView1.Rows[index].Cells[1].Value = "D2";
                this.dataGridView1.Rows[index].Cells[2].Value = (rd.Next() % 10).ToString();
                index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString(); ;
                this.dataGridView1.Rows[index].Cells[1].Value = "D3";
                this.dataGridView1.Rows[index].Cells[2].Value = (rd.Next() % 10).ToString();

                index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString(); ;
                this.dataGridView1.Rows[index].Cells[1].Value = "H1";
                this.dataGridView1.Rows[index].Cells[2].Value = (rd.Next() % 10).ToString();
                index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = DateTime.Now.ToString(); ;
                this.dataGridView1.Rows[index].Cells[1].Value = "H2";
                this.dataGridView1.Rows[index].Cells[2].Value = (rd.Next() % 10).ToString();
            }
        
            for (int i = 0; i < 30; i++)
            {
                int index = this.dataGridView2.Rows.Add();
                //this.dataGridView2.Rows[index].Cells[0].Value = DateTime.Now.ToString(); ;
                //this.dataGridView2.Rows[index].Cells[10].Value = "D1";
                double d0 = (9+rd.Next() % 10);
                double d1 = (9+rd.Next() % 10);
                double d2 = (9+rd.Next() % 10);
                this.dataGridView2.Rows[index].Cells[0].Value = d0.ToString();
                this.dataGridView2.Rows[index].Cells[1].Value = d1.ToString();
                this.dataGridView2.Rows[index].Cells[2].Value = d2.ToString();

                this.dataGridView2.Rows[index].Cells[3].Value = ((Math.Max(Math.Max(d0,d1),d2)- Math.Min(Math.Min(d0, d1), d2))/2).ToString();
                
                this.dataGridView2.Rows[index].Cells[4].Value = (5+rd.Next() % 10).ToString();
                this.dataGridView2.Rows[index].Cells[5].Value = (5+rd.Next() % 10).ToString();
                this.dataGridView2.Rows[index].Cells[6].Value = DateTime.Now.ToString(); ;
                dataGridView2.Rows[index].HeaderCell.Value = (i + 1).ToString();
            }
           // dataGridView2.RowHeadersWidthSizeMode =;

            dataGridView2.RowHeadersWidth = 50;
            this.dataGridView2.Rows[1].Cells[3].Style.BackColor = Color.Red;

            this.dataGridView2.Rows[3].Cells[4].Style.BackColor = Color.Red;



        }
        Form2 form2 = new Form2();
        private void ConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.form2.ShowDialog(this);
        }

        private void PrintToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void HelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("SPC帮助.pdf");
        }

        public bool computeCpkC4(double USL, double LSL, double Ave_mean, double Ave_STD, int subgroupsize, ref double sigma, ref double Cp, ref double Cpu, ref double Cpl, ref double Cpk)
        {
            sigma = Ave_STD * SPC_1_C4[subgroupsize - 2];
            Cp = (USL - LSL) / sigma / 6;
            Cpl = (Ave_mean - LSL) / sigma / 3;
            Cpu = (USL - Ave_mean) / sigma / 3;
            Cpk = Math.Min(Cpl, Cpu);
            return true;

        }


        private void drawchart1(int charttype=0)
        {

            //this.comboBox1.Items.Add("R");
            //sigma = 0, Cp = 0, Cpu = 0, Cpl = 0, Cpk =

            int subgroupsize = 3;
            double Ave_R=0;
            double Ave_mean = 0;
            double UCL = 0;
            double LCL = 0;
            double UCLR = 0;
            double LCLR = 0;

           // double[] dmean = { 1, 4, 3, 6, 5, 2, 1, 4, 8, 4 };
           // double[] dR = { 2, 3, 1, 5, 5, 3, 1, 4, 2, 4 };

             double[] dmean=new double[10];
             double[] dR=new double[10];
            Random rd = new Random();
            for (int i = 0; i < 10; i++)
            {
                dmean[i] = 10+rd.Next()%10;
                dR[i] = 5+rd.Next() % 10;
            }
            this.computeLimit(dR, dmean, subgroupsize, ref Ave_R, ref Ave_mean, ref UCL, ref LCL, ref UCLR, ref LCLR);

            computeCpk(25, 3, Ave_mean, Ave_R, 3, ref sigma, ref Cp, ref Cpu, ref Cpl, ref Cpk);
            this.textBox4.Text = sigma.ToString();
            this.textBox5.Text = UCL.ToString();
            this.textBox6.Text = LCL.ToString();
            this.textBox7.Text = Cp.ToString();
            this.textBox8.Text = Cpk.ToString();

            this.chart1.Series.Clear();
            //this.chart1.ChartAreas[0].AxisX.LineDashStyle = ChartDashStyle.Dash;
            //this.chart1.ChartAreas[0].AxisY.LineDashStyle = ChartDashStyle.Dash;
            //this.chart1.ChartAreas[0].BorderDashStyle= ChartDashStyle.Dash;
         //   chart1.ChartAreas[0].AxisX.MajorGrid.LineWidth = 0;
          //  chart1.ChartAreas[0].AxisX.MinorGrid.LineWidth = 1;
          //  chart1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 1;
          //  chart1.ChartAreas[0].AxisY.MinorGrid.LineWidth = 1;
            chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.DarkGray;

            chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.DarkGray;

            //chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            // chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            //chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
            // chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false;

            // this.chart1.ChartAreas[0].Axes.

            if (charttype == 0)
            {
                Series S_d = new Series("data");
                S_d.ChartType = SeriesChartType.Line;
                S_d.Color = Color.Blue;
                S_d.BorderWidth = 2;

                for (int i=0;i<dmean.Length;i++)
                    S_d.Points.AddXY(i+1, dmean[i]);

                S_d.MarkerColor = Color.Red;
               // S_d.MarkerBorderColor = Color.Yellow;
                S_d.MarkerBorderWidth = 2;
                S_d.MarkerStyle = MarkerStyle.Circle; 
           
                this.chart1.Series.Add(S_d);
               




                Series S_USL= new Series("USL");
                S_USL.ChartType = SeriesChartType.Line;
                S_USL.Color = Color.Red;
                S_USL.Points.AddXY(1, 30);
                S_USL.Points.AddXY(11, 30);
                S_USL.BorderWidth = 3;
                S_USL.BorderDashStyle = ChartDashStyle.Dash;//设置图像边框线的样式(实线、虚线、点线)
                this.chart1.Series.Add(S_USL);

                Series S_LSL = new Series("LSL");
                S_LSL.ChartType = SeriesChartType.Line;
                S_LSL.Color = Color.Red;
                S_LSL.Points.AddXY(1, 1);
                S_LSL.Points.AddXY(11, 1);
                S_LSL.BorderWidth = 3;
                S_LSL.BorderDashStyle = ChartDashStyle.Dash;
                this.chart1.Series.Add(S_LSL);


                Series S_bar = new Series("X-bar");
                S_bar.ChartType = SeriesChartType.Line;
                S_bar.Color = Color.Green;
                S_bar.Points.AddXY(1, Ave_mean);
                S_bar.Points.AddXY(11, Ave_mean);
                S_bar.BorderWidth = 3;
                S_bar.BorderDashStyle = ChartDashStyle.Dash;
                this.chart1.Series.Add(S_bar);

                Series S_UCL = new Series("UCL");
                S_UCL.ChartType = SeriesChartType.Line;
                S_UCL.Color = Color.Yellow;
                S_UCL.Points.AddXY(1, UCL);
                S_UCL.Points.AddXY(11, UCL);
                S_UCL.BorderWidth = 3;
                S_UCL.BorderDashStyle = ChartDashStyle.Dash;

                this.chart1.Series.Add(S_UCL);

                Series S_LCL = new Series("LCL");
                S_LCL.ChartType = SeriesChartType.Line;
                S_LCL.Color = Color.Yellow;
                S_LCL.Points.AddXY(1, LCL);
                S_LCL.Points.AddXY(11, LCL);
                S_LCL.BorderWidth = 3;
                S_LCL.BorderDashStyle = ChartDashStyle.Dash;
                this.chart1.Series.Add(S_LCL);

            }
            else if (charttype == 1)
            {
                Series S_d = new Series("data");
                S_d.ChartType = SeriesChartType.Line;
                S_d.Color = Color.Blue;
                S_d.BorderWidth = 2;

                for (int i = 0; i < dR.Length; i++)
                    S_d.Points.AddXY(i + 1, dR[i]);

                S_d.MarkerColor = Color.Red;
                // S_d.MarkerBorderColor = Color.Yellow;
                S_d.MarkerBorderWidth = 2;
                S_d.MarkerStyle = MarkerStyle.Circle;
                this.chart1.Series.Add(S_d);


                Series s_aveR = new Series("aveR");
                s_aveR.ChartType = SeriesChartType.Line;
                s_aveR.Color = Color.Green;
                s_aveR.Points.AddXY(1, Ave_R);
                s_aveR.Points.AddXY(11, Ave_R);
                s_aveR.BorderWidth = 3;
                s_aveR.BorderDashStyle = ChartDashStyle.Dash;
                this.chart1.Series.Add(s_aveR);

                Series S_UCLR = new Series("UCLR");
                S_UCLR.ChartType = SeriesChartType.Line;
                S_UCLR.Color = Color.Yellow;
                S_UCLR.Points.AddXY(1, UCLR);
                S_UCLR.Points.AddXY(11, UCLR);
                S_UCLR.BorderWidth = 3;
                S_UCLR.BorderDashStyle = ChartDashStyle.Dash;
                this.chart1.Series.Add(S_UCLR);

                Series S_LCLR = new Series("LCLR");
                S_LCLR.ChartType = SeriesChartType.Line;
                S_LCLR.Color = Color.Yellow;
                S_LCLR.Points.AddXY(1, LCLR);
                S_LCLR.Points.AddXY(11, LCLR);
                S_LCLR.BorderWidth = 3;
                S_LCLR.BorderDashStyle = ChartDashStyle.Dash;
                this.chart1.Series.Add(S_LCLR);



            }


        }
        private void button2_Click(object sender, EventArgs e)
        {
            Series series = this.chart1.Series[0];
            int[] yv = {1, 3, 2, 4, 6, 3, 2 };
         /*   series.Points.AddY(1);
            series.Points.AddY(3);
            series.Points.AddY(2);
            series.Points.AddY(4);
            series.Points.AddY(4);
            series.Points.AddY(4);
            series.Points.AddY(4);
            */


            series.Points.AddXY(0,1);
            series.Points.AddXY(1, 5);
            series.Points.AddXY(2, 4);
            series.Points.AddXY(3, 3);
            series.Points.AddXY(4, -2);
            series.Points.AddXY(5, -5);
            series.Points.AddXY(6, 7);
            //series.Points.AddY()
            int nn = series.Points.Count();
            this.Text = nn.ToString();
            Series seriesH = new Series("LineH");
            seriesH.ChartType = SeriesChartType.Line;
            seriesH.Points.AddXY(0, 10);
            seriesH.Points.AddXY(nn, 10);
            this.chart1.Series.Add(seriesH);

            Series seriesL = new Series("LineL");
            seriesL.ChartType = SeriesChartType.Line;
            seriesL.Points.AddXY(0, -10);
            seriesL.Points.AddXY(nn, -10);
           
            this.chart1.Series.Add(seriesL);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            drawchart1(1);
        }
    }
}
