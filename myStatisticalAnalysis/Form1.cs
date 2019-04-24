using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
//using Microsoft.Reporting.WinForms;
using ZedGraph;
using System.Collections;
using System.IO.Ports;
using System.Threading;
using System.IO;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Drawing.Printing;
using System.Security.Cryptography;
namespace myStatisticalAnalysis
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public DataSet RawDataSet=new DataSet();
        public DataTable RawDataTable = new DataTable();
        public DataTable ManageDataTable = new DataTable();
        public DataSet readExcel(string fileName)
        {
             string connStr ;
             if (fileName.EndsWith(".xls"))
               connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
           else
                 connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\"";



        //    if (fileType == ".xls")
        //        connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
       //     else
          //      connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";

      //      string strCon = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + excelpath + ";Extended Properties=Excel 8.0;HDR=NO";
           OleDbConnection myConn = new OleDbConnection(connStr);
            myConn.Open();

            DataTable dtSheetName = myConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });

            string[] strTableNames = new string[dtSheetName.Rows.Count];
            for (int k = 0; k < dtSheetName.Rows.Count; k++)
            {
                strTableNames[k] = dtSheetName.Rows[k]["TABLE_NAME"].ToString();
            }

           // string strCom = " SELECT * FROM [Sheet1$] ";
              RawDataSet.Tables.Clear();

              for (int k = 0; k < strTableNames.Length; k++)
              {
                  string strCom = " SELECT * FROM [" + strTableNames[k] + "] ";

                  OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
                  myCommand.Fill(RawDataSet, "[" + strTableNames[k] + "]");
                 
                }
  
          //  myDataSet = new DataSet();
           
           // myCommand.Fill(RawDataSet, "[Sheet1$]");      
            
            //  myCommand.Fill(
           // RawDataSet.t
            myConn.Close();
            return RawDataSet;
        }
        private void openpToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  this.openFileDialog1.Filter = "*.xls";
            this.openFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx";
            this.openFileDialog1.Dispose();
            System.GC.Collect();
            if(this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = this.openFileDialog1.FileName;
                readExcel(FileName);
               // RawDataTable = RawDataSet.Tables["[demo1]"];
                RawDataTable = RawDataSet.Tables[0];
               // dataGridView1.DataMember = "[Sheet1$]";
               // this.dataGridView1.Rows.Clear();
               // DataTable dt = (DataTable)this.dataGridView1.DataSource;
              //  dt.Rows.Clear();
                //foreach (DataRow Dr in RawDataTable.Rows)
               // {
                    //MessageBox.Show(Dr[0].ToString());
                    //string ss = Dr[0].ToString();
              // }
              //  this.dataGridView1.Columns.Clear();
              //  this.dataGridView1.DataSource = RawDataTable;
                this.dataGridView1.Rows.Clear();
                this.datatabe2view(RawDataTable, dataGridView1);
              
               // dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
             /*   dataGridView1.RowHeadersWidth = 80;
                int rowNumber = 1;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.HeaderCell.Value = rowNumber.ToString();                   
                    rowNumber++;
                }
              */
            //    for (int i = 0; i < this.dataGridView1.Columns.Count; i++)
             //   {
               //     this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
              //  }

           //     dataGridV1_V2(SubGroupSize);
                /*
                if (dataGridView1.Columns.Count < 2)
                {
                    //dataGridV1_V2(SubGroupSize);
                    return;
                }
                int Coln = dataGridView1.Columns.Count;

                if (dataGridView1.Rows[0].Cells[Coln-2].Value.ToString() == "车型".ToString() && dataGridView1.Rows[0].Cells[Coln-1].Value != DBNull.Value)
                {
                    int i = 0;
                    for (i = 0; i < comboBox1.Items.Count; i++)
                    {
                        if (comboBox1.Items[i].ToString() == dataGridView1.Rows[0].Cells[Coln - 1].Value.ToString())
                        {
                            comboBox1.SelectedIndex = i;
                            break;
                        }

                    }
                   if (i == comboBox1.Items.Count)
                        comboBox1.SelectedItem = null;
                }

                if (dataGridView1.Rows[1].Cells[Coln - 2].Value.ToString() == "产品编号".ToString() && dataGridView1.Rows[1].Cells[Coln - 1].Value != DBNull.Value)
                {
                    int i = 0;
                    for (i = 0; i < comboBox2.Items.Count; i++)
                    {
                        if (comboBox2.Items[i].ToString() == dataGridView1.Rows[1].Cells[Coln - 1].Value.ToString())
                        {
                            comboBox2.SelectedIndex = i;
                            break;
                        }

                    }
                    if (i == comboBox2.Items.Count)
                        comboBox2.SelectedItem = null;
                }

                if (dataGridView1.Rows[2].Cells[Coln - 2].Value.ToString() == "护管/芯子".ToString() && dataGridView1.Rows[2].Cells[Coln - 1].Value != DBNull.Value)
                {
                    int i = 0;
                    for (i = 0; i < comboBox3.Items.Count; i++)
                    {
                        if (comboBox3.Items[i].ToString() == dataGridView1.Rows[2].Cells[Coln - 1].Value.ToString())
                        {
                            comboBox3.SelectedIndex = i;
                            break;
                        }

                    }
                    if (i == comboBox2.Items.Count)
                        comboBox2.SelectedItem = null;
                 
                }
               */

               //this.dataGridView1.Refresh();
                try
                {
                    SubGroupSize = System.Convert.ToInt32(textSubGroupSize.Text);
                }
                catch
                {
                    SubGroupSize = 5;
                }
              //  dataGridV1_V2(SubGroupSize);
                // 你的 处理文件路径代码 
            }
            this.openFileDialog1.Dispose();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        int i = 10;
        public Bitmap createChartImage(int Width, int Height)
        {
            System.Drawing.Bitmap image1 = new Bitmap(Width, Height);
            Graphics g = Graphics.FromImage(image1);
            g.Clear(Color.Black);
            g.DrawLine(Pens.Gold, 10 + i, 10, 110, 10 + i);
            Point[] pp ={ new Point(1, 2), new Point(2, 3), new Point(5, 7), new Point(9, 7), new Point(9, 1) };

            g.DrawLines(Pens.Gold, pp);
            i += 10;
            Font font = new Font("宋体", 30f); //字是什么样子的？

            Brush brush = Brushes.Red; //用红色涂上我的字吧；
            string str = "Baidu"; //写什么字？

//Font font =new  Font("宋体",30f); //字是什么样子的？

//Brush brush = Brushes.Red; //用红色涂上我的字吧；

PointF point = new PointF(10f,10f); //从什么地方开始写字捏？

 

//横着写还是竖着写呢？

System.Drawing.StringFormat sf = new System.Drawing.StringFormat();

//还是竖着写吧

sf.FormatFlags = StringFormatFlags.DirectionVertical;

 

//开始写咯

g.DrawString(str,font,brush,point,sf);

 
g.DrawString("hello", font, brush,20f, 20f);
            g.Dispose();
            return image1;
        }
        public Bitmap createChartImage1(int Width, int Height, double[] showDataY,double minY,double maxY)
        {
            System.Drawing.Bitmap image1 = new Bitmap(Width, Height);
            Graphics g = Graphics.FromImage(image1);
            g.Clear(Color.Black);
            double xd = (Width - 20) / (showDataY.Length - 1);
            for (int i = 0; i < showDataY.Length-1;i++)
            {
                double  py0= (double)Height * 0.9 - (double)(showDataY[i] - minY) * (Height * 0.9 - Height * 0.1) / (maxY - minY);
                double  py1= (double)Height * 0.9 - (double)(showDataY[i+1] - minY) * (Height * 0.9 - Height * 0.1) / (maxY - minY);
                g.DrawLine(Pens.Gold, 10 + (int)xd * i, (int)py0, 10 + (int)xd * (i + 1), (int)py1);
            }
        //    Point[] pp = { new Point(1, 2), new Point(2, 3), new Point(5, 7), new Point(9, 7), new Point(9, 1) };

        //    g.DrawLines(Pens.Gold, pp);
 
            Font font = new Font("宋体", 30f); //字是什么样子的？

            Brush brush = Brushes.Red; //用红色涂上我的字吧；
            string str = "Baidu"; //写什么字？


            PointF point = new PointF(10f, 10f); //从什么地方开始写字捏？



            //横着写还是竖着写呢？

            System.Drawing.StringFormat sf = new System.Drawing.StringFormat();

            //还是竖着写吧

            sf.FormatFlags = StringFormatFlags.DirectionVertical;



            //开始写咯

            g.DrawString(str, font, brush, point, sf);


            g.DrawString("hello", font, brush, 20f, 20f);
            g.Dispose();
            return image1;
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
                sum2 += x*x;
                kk++;
                if (kk == subgroupsize)
                {
                    // R[m++] = maxX - minX;
                    R.Add(maxX - minX);
                    mean.Add(sum / subgroupsize);
                    Std.Add(Math.Sqrt((sum2 - sum*sum / subgroupsize) / (subgroupsize - 1)));
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
            int kk=0;
            double maxX=0;
            double minX=0;
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
        double[] SPC_D3={ 0 , 0 ,0, 0 ,   0 , 0.0760, 0.1360,0.1840 , 0.2230,0.2560 ,0.2830,0.3070,0.3280,0.3470, 0.3630,0.3780,0.3910,0.4030 ,0.4150 ,0.4250 ,0.4340, 0.4430,0.4510,0.4590};
        double[] SPC_D4 = { 3.2670,2.5740,2.2820,2.1140 ,2.0040,1.9240 ,1.8640,1.8160,1.7770 ,1.7440,1.7170,1.6930,1.6720,1.6530,1.6370 ,1.6220,1.6080 ,1.5970,1.5850,1.5750,1.5660,1.5570,1.5480,1.5410};
        double[] SPC_C4 = { 0.7979, 0.8862, 0.9213, 0.9400, 0.9515, 0.9594, 0.9650, 0.9693, 0.9727, 0.9754, 0.9776, 0.9794, 0.9810, 0.9823, 0.9835, 0.9845, 0.9854, 0.9862, 0.9869, 0.9876, 0.9882, 0.9887, 0.9892, 0.9896 };
        double[] SPC_d2 = { 1.1280, 1.6930, 2.0590, 2.3260, 2.5340, 2.7040, 2.8470, 2.9700, 3.0780, 3.1730, 3.2580, 3.3360, 3.4070, 3.4720, 3.5320, 3.5880, 3.6400, 3.6890, 3.7350, 3.7780, 3.8190, 3.8580, 3.8950, 3.9310 };
        double[] SPC_1_C4 ={ 1.2533, 1.1284, 1.0854, 1.0638, 1.0510, 1.0423, 1.0363, 1.0317, 1.0281, 1.0252, 1.0229, 1.0210, 1.0194, 1.0180, 1.0168, 1.0157, 1.0148, 1.0140, 1.0133, 1.0126, 1.0119, 1.0114, 1.0109, 1.0105 };
        double[] SPC_1_d2 = { 0.8865, 0.5907, 0.4857, 0.4299, 0.3946, 0.3698, 0.3512, 0.3367, 0.3249, 0.3152, 0.3069, 0.2998, 0.2935, 0.2880, 0.2831, 0.2787, 0.2747, 0.2711, 0.2677, 0.2647, 0.2618, 0.2592, 0.2567, 0.2544 };
        public bool computeLimit(double[] R, double[] mean, int subgroupsize,  ref double  Ave_R, ref double  Ave_mean, ref double UCL, ref double LCL, ref double UCLR, ref double LCLR)
         {
             if (subgroupsize <= 1) return false;
             int m = R.Length;
            Ave_R=0;
            Ave_mean=0;
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
        public bool computeCpk(double USL, double LSL, double Ave_mean, double Ave_R, int subgroupsize, ref double sigma,ref double Cp,ref double Cpu, ref double Cpl, ref double Cpk)
        {
            sigma = Ave_R * SPC_1_d2[subgroupsize - 2];            
            Cp=(USL-LSL)/sigma/6;
            Cpl = (Ave_mean-LSL)/sigma/3;
            Cpu = (USL-Ave_mean)/sigma/3;
            Cpk=Math.Min(Cpl,Cpu);
            return true;

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
        public void ShowXbarRChart(DataTable RawDataTable, int subgroupsize, ZedGraph.ZedGraphControl ZedxbarChart, ZedGraph.ZedGraphControl ZedRchart)
        {
            ZedRchart.GraphPane.CurveList.Clear();
            ZedRchart.GraphPane.GraphItemList.Clear();
            ZedxbarChart.GraphPane.CurveList.Clear();
            ZedxbarChart.GraphPane.GraphItemList.Clear();      
            int nn = RawDataTable.Rows.Count;
            double[] dataY = new double[nn];
            //double[] dataX=new double[nn];
            for (int i = 0; i < nn; i++)
            {
                dataY[i] = (double)System.Convert.ToDouble(RawDataTable.Rows[i][0]);
            }

            //   double[] showdataY = {1,2,4,3,2,3,4,4,1,23,12,12,12,12,1,56,323};
            //     double[] showdataX = { 0,1, 2, 3, 4, 5, 6, 7, 8 };
            //    this.pictureBox1.Image = createChartImage1(pictureBox1.Width, pictureBox1.Height,showdataY,0,4000);
            // this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource(RawDataTable.TableName, RawDataTable));
            //  this.reportViewer1.RefreshReport();
            // subgroupsize = 5;
            ArrayList R = new ArrayList();
            ArrayList mean = new ArrayList();
            ArrayList std = new ArrayList();
            computeRangeMean(dataY, subgroupsize, R, mean);
          //  computeRangeMeanStd(dataY, subgroupsize, R, mean,std);
            double[] RR = (double[])R.ToArray(Type.GetType("System.Double"));
            double[] mm = (double[])mean.ToArray(Type.GetType("System.Double"));
            double Ave_R = 0;
            double Ave_mean = 0;
       /*    double Ave_Std=0;
            for(int i=0;i<std.Count;i++)
                Ave_Std+=(double)std[i];
            Ave_Std /= std.Count;
        */

            //defAxisBorderPenWidth=0.1;
            computeLimit(RR, mm, subgroupsize, ref Ave_R, ref Ave_mean, ref UCL, ref LCL, ref  UCLR, ref LCLR);
            computeCpk(USL, LSL, Ave_mean, Ave_R, subgroupsize, ref  sigma, ref  Cp, ref  Cpu, ref  Cpl, ref  Cpk);


            string printparameter_ss = string.Format("subgroupsize:{0}\nAve_R:{1:n6}\nAve_mean:{2:n6}\nUCL:{3:n6}\nLCL:{4:n6}\nUCLR:{5:n6}\nLCLR:{6:n6}\n", subgroupsize, Ave_R, Ave_mean, UCL, LCL, UCLR, LCLR);
         //   printparameter_ss
            string printparameter_ss1 = string.Format("Sigma:{0:n6}\nCp:{1:n6}\nCpu:{2:n6}\nCpl:{3:n6}\nCpk:{4:n6}", sigma, Cp, Cpu, Cpl, Cpk);
            this.label1.Text = printparameter_ss + printparameter_ss1;
            //     zedGraphControl1.=0.1;
            //      zedGraphControl1.GraphPane.AddCurve("line", showdataX, showdataY, Color.Red, SymbolType.None);
            //   zedGraphControl1.GraphPane.AddCurve();
            nn = R.Count;
            double[] showdataX = new double[nn];
            for (int i = 0; i < nn; i++)
                showdataX[i] = i;
            ZedRchart.GraphPane.AddCurve("Data", null, (double[])R.ToArray(Type.GetType("System.Double")), Color.Blue, SymbolType.Circle);

            PointPairList lineAveR = new PointPairList();
            lineAveR.Add(0, Ave_R);
            lineAveR.Add(nn, Ave_R);
            ZedRchart.GraphPane.AddCurve("Ave_R", lineAveR, Color.Green, SymbolType.None);

            PointPairList lineUCLR = new PointPairList();
            lineUCLR.Add(0, UCLR);
            lineUCLR.Add(nn, UCLR);
            //   zedGraphControl1.GraphPane.ScaledPenWidth(3, 1);

            ZedRchart.GraphPane.AddCurve("UCLR", lineUCLR, Color.Red, SymbolType.None);

            PointPairList lineLCLR = new PointPairList();
            lineLCLR.Add(0, LCLR);
            lineLCLR.Add(nn, LCLR);
            ZedRchart.GraphPane.AddCurve("LCLR", lineLCLR, Color.Red, SymbolType.None);
            ZedRchart.GraphPane.Title = "R chart";
            //  zedGraphControl1.GraphPane.
            ZedRchart.GraphPane.YAxis.IsZeroLine = false;

            //    zedGraphControl1.GraphPane.AxisBorder.IsVisible = false;
            // zedGraphControl1.GraphPane.PaneBorder.IsVisible = false;
            double minR = (double)R[0];
            double maxR = (double)R[0];
            foreach(double Ri in R)
            {
                if (Ri < minR) minR = Ri;
                else if (Ri > maxR) maxR = Ri;
            }
            double Ymin = minR < LCLR ? minR : LCLR - (maxR - minR) * 0.2;
            if (minR >= 0 && LCLR >= 0 && Ymin < 0)
                ZedRchart.GraphPane.YAxis.Min = Ymin;
            //ZedRchart.GraphPane.YAxis.Min = minR< LCLR?minR:LCLR - (maxR - minR) * 0.2;
            ZedRchart.AxisChange();
         //   double Ymin = ZedRchart.GraphPane.YAxis.Min;
          //  double Ymax = ZedRchart.GraphPane.YAxis.Max;
           
            //    zedGraphControl1.GraphPane.AddCurve("LCLR", lineLCLR, Color.Red, SymbolType.None);
            //   zedGraphControl1.Refresh();
            ZedRchart.Invalidate();
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            nn = mean.Count;
         //   double[] showdataX = new double[nn];
            for (int i = 0; i < nn; i++)
                showdataX[i] = i;
            ZedxbarChart.GraphPane.AddCurve("Data", null, (double[])mean.ToArray(Type.GetType("System.Double")), Color.Blue, SymbolType.Circle);

            PointPairList linemean = new PointPairList();
            linemean.Add(0, Ave_mean);
            linemean.Add(nn, Ave_mean);
            ZedxbarChart.GraphPane.AddCurve("X_bar", linemean, Color.Green, SymbolType.None);

            PointPairList lineUCL = new PointPairList();
            lineUCL.Add(0, UCL);
            lineUCL.Add(nn, UCL);
            //   zedGraphControl1.GraphPane.ScaledPenWidth(3, 1);

            ZedxbarChart.GraphPane.AddCurve("UCL", lineUCL, Color.Red, SymbolType.None);

            PointPairList lineLCL = new PointPairList();
            lineLCL.Add(0, LCL);
            lineLCL.Add(nn, LCL);
            ZedxbarChart.GraphPane.AddCurve("LCL", lineLCL, Color.Red, SymbolType.None);

            PointPairList lineUSL = new PointPairList();
            lineUSL.Add(0, USL);
            lineUSL.Add(nn, USL);
            //   zedGraphControl1.GraphPane.ScaledPenWidth(3, 1);

            ZedxbarChart.GraphPane.AddCurve("USL", lineUSL, Color.Purple, SymbolType.None);

            PointPairList lineLSL = new PointPairList();
            lineLSL.Add(0, LSL);
            lineLSL.Add(nn, LSL);
            ZedxbarChart.GraphPane.AddCurve("LSL", lineLSL, Color.Purple, SymbolType.None);

            ZedxbarChart.GraphPane.Title = "X_Bar chart";
            //  zedGraphControl1.GraphPane.
            ZedxbarChart.GraphPane.YAxis.IsZeroLine = false;

            //    zedGraphControl1.GraphPane.AxisBorder.IsVisible = false;
            // zedGraphControl1.GraphPane.PaneBorder.IsVisible = false;
           // double minM= (double)mean[0];
          //  double maxM = (double)mean[0];
           // foreach (double meani in mean)
           // {
              //  if (meani < minM) minM = meani;
              //  else if (meani > maxM) maxM = meani;
           // }
            //  Ymin=
            
            ZedxbarChart.GraphPane.YAxis.IsZeroLine = false;
            ZedxbarChart.AxisChange();
          //  ZedRchart.GraphPane.YAxis.MinAuto = false;
        //    ZedxbarChart.GraphPane.YAxis.Min = minM < LCL ? minM : LCL - (maxM - minM) * 0.2 ;
           // ZedxbarChart.AxisChange();
            //
            
            //   double Ymin = ZedRchart.GraphPane.YAxis.Min;
            //  double Ymax = ZedRchart.GraphPane.YAxis.Max;

            //    zedGraphControl1.GraphPane.AddCurve("LCLR", lineLCLR, Color.Red, SymbolType.None);
            //   zedGraphControl1.Refresh();
            
            ZedxbarChart.Invalidate();
             
        }
        public DataTable GetDgvToTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();

            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }

            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        private void sPCToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  ThreadRun = false;
        //    this.sPort.Close();
            DataTable dtSource = GetDgvToTable(this.dataGridView1);
            ShowXbarRChart(dtSource, 5, zedGraphControl1, zedGraphControl2);
            ShowDistributionChart(dtSource, zedGraphControl3);
        }

        public void ShowDistributionChart(DataTable RawDataTable, ZedGraph.ZedGraphControl ZedChart0)
        {
            ZedChart0.GraphPane.CurveList.Clear();
            ZedChart0.GraphPane.GraphItemList.Clear();
            ZedChart0.GraphPane.XAxis.Type = AxisType.Date;
            ZedChart0.GraphPane.XAxis.Title = "";
            ZedChart0.GraphPane.YAxis.Title = "";
          
            int nn = RawDataTable.Rows.Count;
            double[] dataY = new double[nn];
          //  double[] dataX=new double[nn];
            double maxY = System.Convert.ToDouble(RawDataTable.Rows[0][0]);
            double minY = maxY;
            double sumY = 0;
            double sumY2 = 0;
            for (int i = 0; i < nn; i++)
            {
                dataY[i] = System.Convert.ToDouble(RawDataTable.Rows[i][0]);
                if (dataY[i] > maxY) 
                    maxY = dataY[i];
                else if (dataY[i] < minY) 
                    minY = dataY[i];

                sumY += dataY[i];
                sumY2 += (dataY[i] * dataY[i]);
            }
            double meanY = sumY / nn;
            double VarY = (sumY2 - nn * meanY * meanY) / (nn - 1);
             double StdY=Math.Sqrt(VarY);
            double [] dataYn =new double[7]{0,0,0,0,0,0,0};
            double[] dataX0 =new double[7]{30,30.5,31,31.5,32,32.2,32.4};
            for (int i = 0; i < 7; i++)
            {
                dataX0[i] =minY+ i * (maxY - minY) / 7 + (maxY - minY)/14;
            }
                for (int i = 0; i < nn; i++)
                {
                    if (dataY[i] == maxY)
                        dataYn[6]++;
                    else
                        dataYn[(int)(7 * (dataY[i] - minY) / (maxY - minY))]++;

                }



               ZedChart0.GraphPane.AddBar("Data Distribution", dataX0, dataYn, Color.Blue);
               // Axis aa=ZedChart0.GraphPane.BarBaseAxis();
              //  ZedChart0.GraphPane.MinBarGap = (float)0.02;
                ZedChart0.GraphPane.ClusterScaleWidth = (float)(maxY-minY)/(7);

            PointPairList datanorm=new PointPairList();
            //    double[] datanorm = new double[100];
                for (int i = 0; i < 100; i++)
                {
                    double xi=0;
                    double yi=0;
                    xi=meanY - 4* StdY+ (8 * StdY) *i/ 100;
                    yi = Math.Exp(-(xi - meanY) * (xi - meanY) / VarY / 2) / Math.Sqrt(2 * Math.PI) / StdY ;
                    datanorm.Add(xi, yi * nn *StdY);
 
                }
          //      ZedChart0.GraphPane.XAxis.Min = meanY - 6 * StdY;
              //  ZedChart0.GraphPane.XAxis.Max=meanY + 6 * StdY;

                ZedChart0.GraphPane.AddCurve("norm", datanorm, Color.Green,SymbolType.None);
                  //  for (double x = meanY - 6 * StdY; x <= meanY + 6 * StdY; x += (12 * StdY) / 100)
                //    {
                  //  }

                    //   ZedChart0.GraphPane.MinClusterGap = (float)0.002;
                    //  ZedChart0.GraphPane.AddBar
                double[] LX ={ USL, USL };
                double[] LY ={ -1, 0.3989*nn };
                ZedChart0.GraphPane.AddCurve("USL",LX,LY, Color.Red, SymbolType.None);
                double[] LX1 ={ LSL, LSL };
                double[] LY1 ={ -1, 0.3989 * nn };
                ZedChart0.GraphPane.AddCurve("LSL", LX1, LY1, Color.Red, SymbolType.None);
                double[] LX2 ={ UCL, UCL };
                double[] LY2 ={ -1, 0.3989 * nn };
                ZedChart0.GraphPane.AddCurve("UCL", LX2, LY2, Color.Yellow, SymbolType.None);
                double[] LX3 ={ LCL, LCL };
                double[] LY3 ={ -1, 0.3989 * nn };
                ZedChart0.GraphPane.AddCurve("LCL", LX3, LY3, Color.Yellow, SymbolType.None);
                ZedChart0.GraphPane.Title = "CPK Chart";               
                ZedChart0.AxisChange();
                ZedChart0.Invalidate();
       
 
        }

        public void ShowDistributionPerMonthChart(DataTable RawDataTable, ZedGraph.ZedGraphControl ZedChart0)
        {
            ZedChart0.GraphPane.CurveList.Clear();
            ZedChart0.GraphPane.GraphItemList.Clear();
            ZedChart0.Controls.Clear();
            int nn = RawDataTable.Rows.Count;
           // double[] dataY = new double[nn];
            //  double[] dataX=new double[nn];
            //double data = System.Convert.ToDouble(RawDataTable.Rows[0][0]);
          //  double minY = maxY;
        //    double sumY = 0;
           // double sumY2 = 0;
            List <double> DataPerMonth=new List<double>();
           List  <string> DateMonth=new List<string>();
           double sumdata = 0;
           long  mDataCount = 0;
            for (int i = 0; i < nn; i++)
            {
              double tdata = System.Convert.ToDouble(RawDataTable.Rows[i][0]);
              System.DateTime  date = System.Convert.ToDateTime(RawDataTable.Rows[i][1].ToString());
              string tdate = date.ToString("yyyy.MM");
                int Mn=DateMonth.Count;
                if(Mn==0)
                    DateMonth.Add(tdate);
                else if (tdate != DateMonth[Mn - 1])
                {
                    DateMonth.Add(tdate);
                    DataPerMonth.Add(sumdata / mDataCount);
                    sumdata=0;
                    mDataCount=0; 
                }
                sumdata += tdata;
                mDataCount++;
                if (i == nn - 1 && mDataCount!=0)
                {                    
                    DataPerMonth.Add(sumdata / mDataCount);
                }

            
            }
   
            ZedChart0.GraphPane.AddBar("月均值", null, DataPerMonth.ToArray(), Color.Blue);


         // ZedChart0.GraphPane.AddBar("Data Distribution", null, dataYn, Color.Blue);
           
       //     ZedChart0.GraphPane.ClusterScaleWidth = (float)(maxY - minY) / (7);

           //ZedChart0.GraphPane.AddBar(

            //   ZedChart0.GraphPane.MinClusterGap = (float)0.002;
            //  ZedChart0.GraphPane.AddBar
       //     string[] TextLabels1 = new string []{ "nihao", "dajihao" };

            
            
            ZedChart0.GraphPane.XAxis.Type = AxisType.Text;
            ZedChart0.GraphPane.XAxis.TextLabels = DateMonth.ToArray();
            ZedChart0.GraphPane.Title = "月均值图";
            ZedChart0.GraphPane.XAxis.Title = "月份";
            ZedChart0.GraphPane.YAxis.Title = "平均值";

            ZedChart0.AxisChange();
            ZedChart0.Invalidate();


        }
        public string Encode(string data, string encryptKey)
        {
            string KEY_64 = encryptKey.Substring(0, 8);
            string IV_64 = encryptKey.Substring(0, 8);
            byte[] byKey = System.Text.ASCIIEncoding.ASCII.GetBytes(KEY_64);
            byte[] byIV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV_64);

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            int i = cryptoProvider.KeySize;
            MemoryStream ms = new MemoryStream();
            CryptoStream cst = new CryptoStream(ms, cryptoProvider.CreateEncryptor(byKey, byIV), CryptoStreamMode.Write);

            StreamWriter sw = new StreamWriter(cst);
            sw.Write(data);
            sw.Flush();
            cst.FlushFinalBlock();
            sw.Flush();
            return Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length);

        }

        public string Decode(string data, string encryptKey)
        {
            string KEY_64 = encryptKey.Substring(0, 8);
            string IV_64 = encryptKey.Substring(0, 8);
            byte[] byKey = System.Text.ASCIIEncoding.ASCII.GetBytes(KEY_64);
            byte[] byIV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV_64);

            byte[] byEnc;
            try
            {
                byEnc = Convert.FromBase64String(data);
            }
            catch
            {
                return null;
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream ms = new MemoryStream(byEnc);
            CryptoStream cst = new CryptoStream(ms, cryptoProvider.CreateDecryptor(byKey, byIV), CryptoStreamMode.Read);
            StreamReader sr = new StreamReader(cst);
            return sr.ReadToEnd();
        }
        bool read3ss(ref string ss0, ref string ss1, ref string ss2)
        {
            try
            {
                StreamReader sr = new StreamReader("License.dat");
                ss0 = sr.ReadLine();
                ss1 = sr.ReadLine();
                ss2 = sr.ReadLine();
                sr.Close();
                return true;
            }
            catch
            {
                return false;

            }

        }
        bool IsLicensed()
        {
            return true;
            string ss0 = "1";
            string ss1 = "1";
            string ss2 = "1";
            if (read3ss(ref ss0, ref ss1, ref ss2) == false)
                return false;
            Computer cm1 = new Computer();
            //bool ret = Encode(cm1.CpuID, "daifugui") == ss0 && Encode(cm1.DiskID, "daifugui") == ss1 && Encode(cm1.BiosSerialNumber, "daifugui") == ss2;
            bool ret = Encode(cm1.CpuID, "daifugui") == ss0 && Encode(cm1.DiskID, "daifugui") == ss1 && (Encode(cm1.BiosSerialNumber, "daifugui") == ss2 || Encode(cm1.MacAddress, "daifugui") == ss2);
            return ret;


        }

        Form2 form2 = new Form2();
        private void Form1_Load(object sender, EventArgs e)
        {
            sPort = new SerialPort();
            if (sPort.IsOpen)
            {
                sPort.Close();
            }
            Control.CheckForIllegalCrossThreadCalls = false;
            this.WindowState = FormWindowState.Maximized;
            
           // this.MinimizeBox = false;
            //this.MaximizeBox = false;
            string[] SerialPortNames = SerialPort.GetPortNames();
            this.comboBox4.Items.AddRange(SerialPortNames);
            if (comboBox4.Items.Count>0)
                comboBox4.SelectedIndex = 0;
          //  this.close
          
          //  sPort.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived); 
            DateTime a = System.DateTime.Today.Date;
            if (a.Year != 2018)
            {
                if (IsLicensed() == false)
                {
                    MessageBox.Show("试用期限已到，请购买软件！");
                    this.Close();
                }
            }

            

            DateTime b = new DateTime(2018,5,1);
            if (a.CompareTo(b) > 0)
            {
                if (IsLicensed() == false)
                {
                    MessageBox.Show("试用期限已到，请购买软件！");
                    this.Close();
                }
                //MessageBox.Show("服务期限已过，请购买软件！");
                //this.Close(); 
            }

         //   this.ManageDataTable = OpenCSV(Application.StartupPath+"\\config",8);
            try
            {
                ReadMDB ReadMdb1 = new ReadMDB();
                this.ManageDataTable = (DataTable)ReadMdb1.MDBreader("select * from Config", Application.StartupPath + "\\SPC_DATA.accdb");
            }
            catch
            {
                MessageBox.Show("Error in read database!!!");
                return ;
            }

            UpdateSysInfo(ManageDataTable);
         //   updatecomboBox(ManageDataTable);

            
        }
        SerialPort sPort;
        string Valuess=null;
        private void port_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            if (sPort.IsOpen)
            {
               // byte[] data = Convert.FromBase64String(sPort.ReadLine());
              //  this.label1.Text = Encoding.Unicode.GetString(data);
                
           //     int count = sPort.BytesToRead;
           //     byte[] data = new byte[count];
                // sPort.Read(data,0,count);
            // string ss=sPort.ReadTo("\n");
           // this.label1.Text = Encoding.Unicode.GetString(a);
            //this.label1.Text = ss;
           //   this.label1.Text = data.ToString();
               // this.label1.Text = Encoding.Unicode.GetString(data);
           //   this.label1.Text = Encoding.ASCII.GetString(data);

                try
                {
                    sPort.NewLine = "\r";

                    string ss = sPort.ReadLine();
                    if (ss == "S")
                    {
                        DataGridViewRow aa = new DataGridViewRow();
                        // this.dataGridView1.Rows.Add();
                        int index = this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[index].Cells[0].Value = Valuess;
                        this.dataGridView1.Refresh();
                    }
                    // this.label1.Text = Encoding.Unicode.GetString(a);
                    else
                        Valuess = ss;
                    this.label1.Text = ss;
                }
                catch
                {
                    ;
                }
            }
            ;
        }
        
        private void monitorToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.Columns.Clear();
            this.dataGridView1.Columns.Add("Data", "Data");
            this.dataGridView1.Columns.Add("Time", "Time");
            this.dataGridView1.Columns["Time"].Width = 120;
            dataGridView1.RowHeadersWidth = 80;
            if (ThreadRun)
                return;
            if (sPort.IsOpen) return;
     
           // dataGridView1.Controls.
         //   this.dataGridView1.Columns.
            sPort.PortName = "com1";//串口的portname 
            sPort.BaudRate = 9600;//串口的波特率
            sPort.DataBits = 8;
            //两个停止位
            sPort.StopBits = System.IO.Ports.StopBits.One;
            //无奇偶校验位
            sPort.Parity = System.IO.Ports.Parity.None;
            sPort.ReadTimeout = -1;
            sPort.WriteTimeout = -1;
            try
            {
                sPort.Open();
            }catch
                {
                    sPort.Open();
                }

                t = new Thread(SerialReadDataThread);
                ThreadRun = true;
                t.Start();     
      
        }
        [DllImport("kernel32.dll")]
        public static extern bool Beep(int freq, int duration);
        [DllImport("user32.dll")]
        public static extern bool MessageBeep(uint uType);
        Thread t;
        bool ThreadRun=false;
        private delegate void InvokeHandler();
        private void SerialReadDataThread()
        {
            while (ThreadRun)
            {
                
                if (sPort.IsOpen)
                {
                    
                    try
                    {
                        sPort.NewLine = "\r"; // test sscom \r\n  , \r really software 
                        string ss =sPort.ReadLine();
                     //   sPort.DataReceived
                     //   string ss=sPort.ReadExisting();
                        if (ss == "S")
                        {
                            // DataGridViewRow aa = new DataGridViewRow();
                            // this.dataGridView1.Rows.Add();
                            this.Invoke(new InvokeHandler(delegate()
                           {

                               int index = this.dataGridView1.Rows.Add();
                               this.dataGridView1.Rows[index].Cells[0].Value = System.Convert.ToDouble(Valuess);
                               this.dataGridView1.Rows[index].Cells[1].Value = DateTime.Now.ToString();
                                this.dataGridView1.Rows[index].HeaderCell.Value = (index+1).ToString();

                                this.dataGridView1.Rows[index].Cells[2].Value = this.comboBox1.SelectedItem.ToString();
                                this.dataGridView1.Rows[index].Cells[3].Value = this.comboBox2.SelectedItem.ToString();
                                this.dataGridView1.Rows[index].Cells[4].Value = this.comboBox3.SelectedItem.ToString();
                                this.dataGridView1.Rows[index].Cells[5].Value = this.comboBox5.SelectedItem.ToString();
      

                               if (dataGridView1.RowCount > 4)
                                   this.dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
                               double dataV = System.Convert.ToDouble(Valuess);
                               if (dataV > USL || dataV < LSL)
                               {
                                   MessageBeep(0x00000040);
                                   this.dataGridView1.Rows[index].Cells[0].Style.ForeColor = Color.Red;
                                   this.dataGridView1.Rows[index].Cells[1].Style.ForeColor = Color.Red;
                               }
                               dataGridView2AddData(dataV, SubGroupSize);
                               // this.dataGridView1.DisplayedRowCount();
                           }));

                            //  this.dataGridView1.Refresh();
                        }
                        // this.label1.Text = Encoding.Unicode.GetString(a);
                        else
                            Valuess = ss;
                      // this.label1.Text = ss;
                        this.textRealTimeData.Text = ss;
                    }
                    catch
                    {
                        ;
                    }
                    System.Threading.Thread.Sleep(100);
                }
            }
        }

        private void SerialReadDataThread1()
        {
            while (ThreadRun)
            {

                if (sPort.IsOpen)
                {

                    try
                    {
                        sPort.NewLine = "\r"; // test sscom \r\n  , \r really software 
                        string ss = sPort.ReadLine();
                        
                        //   sPort.DataReceived
                        //   string ss=sPort.ReadExisting();
                        string ss0 = ss.Substring(0, 3);
                        //Valuess = ss.Substring(
                        if (ss0 == "01A")
                        {
                            // DataGridViewRow aa = new DataGridViewRow();
                            // this.dataGridView1.Rows.Add();
                            this.Invoke(new InvokeHandler(delegate()
                            {
                                Valuess = ss.Substring(3);
                                int index = this.dataGridView1.Rows.Add();
                                this.dataGridView1.Rows[index].Cells[0].Value = System.Convert.ToDouble(Valuess);
                                this.dataGridView1.Rows[index].Cells[1].Value = DateTime.Now.ToString();
                                this.dataGridView1.Rows[index].HeaderCell.Value = (index + 1).ToString();

                                this.dataGridView1.Rows[index].Cells[2].Value = this.comboBox1.SelectedItem.ToString();
                                this.dataGridView1.Rows[index].Cells[3].Value = this.comboBox2.SelectedItem.ToString();
                                this.dataGridView1.Rows[index].Cells[4].Value = this.comboBox3.SelectedItem.ToString();
                                


                                if (dataGridView1.RowCount > 4)
                                    this.dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
                                double dataV = System.Convert.ToDouble(Valuess);
                                if (dataV > USL || dataV < LSL)
                                {
                                    MessageBeep(0x00000040);
                                    this.dataGridView1.Rows[index].Cells[0].Style.ForeColor = Color.Red;
                                    this.dataGridView1.Rows[index].Cells[1].Style.ForeColor = Color.Red;
                                }
                                dataGridView2AddData(dataV, SubGroupSize);
                                // this.dataGridView1.DisplayedRowCount();
                            }));

                            //  this.dataGridView1.Refresh();
                        }
                        // this.label1.Text = Encoding.Unicode.GetString(a);
                      //  else
                          //  Valuess = ss;
                        // this.label1.Text = ss;
                        this.textRealTimeData.Text = ss;
                    }
                    catch
                    {
                        ;
                    }
                    System.Threading.Thread.Sleep(100);
                }
            }
        }
        private void testGetdataToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  byte[] data = Convert.FromBase64String(sPort.ReadLine());
            sPort.NewLine ="\r";
            string ss = sPort.ReadLine();
           // this.label1.Text = Encoding.Unicode.GetString(a);
            this.label1.Text = ss;
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            //MessageBox.Show("close");
        }

        private void formclosing(object sender, FormClosingEventArgs e)
        {
           // MessageBox.Show("close");
            ThreadRun = false;
            sPort.Close();         
           
        }
            public void DataToExcel(DataGridView m_DataView)
        {
           SaveFileDialog kk = new SaveFileDialog(); 
            kk.Title = "保存EXECL文件"; 
            kk.Filter = "EXECL文件(*.xls) |*.xls |所有文件(*.*) |*.*"; 
          kk.FilterIndex = 1;
            if (kk.ShowDialog() == DialogResult.OK) 
            { 
                 string FileName = kk.FileName;
                 if (File.Exists(FileName))
                     File.Delete(FileName);
                 FileStream objFileStream; 
                StreamWriter objStreamWriter; 
                string strLine = ""; 
                objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write); 
                objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);
              //   for (int i = 0; i  < m_DataView.Columns.Count; i++) 
              //   { 
                //    if (m_DataView.Columns[i].Visible == true) 
                //     { 
                    //    strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9); 
                 //    } 
             //    } 
            //    objStreamWriter.WriteLine(strLine); 
                strLine = ""; 

             for (int i = 0; i  < m_DataView.Rows.Count; i++) 
               { 
                   if (m_DataView.Columns[0].Visible == true) 
                    { 
                       if (m_DataView.Rows[i].Cells[0].Value == null) 
                          strLine = strLine + " " + Convert.ToChar(9); 
                       else 
                           strLine = strLine + m_DataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9); 
                   } 
                    for (int j = 1; j  < m_DataView.Columns.Count; j++) 
                    { 
                        if (m_DataView.Columns[j].Visible == true) 
                        { 
                           if (m_DataView.Rows[i].Cells[j].Value == null) 
                                strLine = strLine + " " + Convert.ToChar(9); 
                            else 
                            { 
                                string rowstr = ""; 
                               rowstr = m_DataView.Rows[i].Cells[j].Value.ToString(); 
                                if (rowstr.IndexOf("\r\n") >  0) 
                                   rowstr = rowstr.Replace("\r\n", " "); 
                               if (rowstr.IndexOf("\t") >  0) 
                                    rowstr = rowstr.Replace("\t", " "); 
                                strLine = strLine + rowstr + Convert.ToChar(9); 
                            } 
                        } 
                    } 
                   objStreamWriter.WriteLine(strLine); 
                    strLine = ""; 
               } 
                objStreamWriter.Close(); 
                objFileStream.Close();
               MessageBox.Show(this,"保存EXCEL成功","提示",MessageBoxButtons.OK,MessageBoxIcon.Information); 
            }
        }
        void saveexcel(DataGridView m_DataView)
        {
            DataTable DT1=GetDgvToTable(m_DataView);
            DT1.TableName="data";
            String sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
     "Data Source=K:\\work111\\excel1.xls;" +
     "Extended Properties=Excel 8.0;";
            //实例化一个Oledbconnection类(实现了IDisposable,要using)
                using (OleDbConnection ole_conn = new OleDbConnection(sConnectionString))
                {
                    ole_conn.Open();
                 using (OleDbCommand ole_cmd = ole_conn.CreateCommand())
                   {
                  //     ole_cmd.CommandText = "CREATE TABLE data ([Data] VarChar,[Time] VarChar)";
                   //   ole_cmd.ExecuteNonQuery();
                     //   ole_cmd.CommandText = "insert into data values('DJ001','点击科技')";
                     //   ole_cmd.ExecuteNonQuery();
                       // ole_cmd.up
                       // MessageBox.Show("生成Excel文件成功并写入一条数据......");
                    }

                    string strCom = " SELECT * FROM data";

                  OleDbDataAdapter myDA = new OleDbDataAdapter(strCom, ole_conn); ;
                   // SqlDataAdapter myDA = new SqlDataAdapter(strCom, sConnectionString);
                   // SqlCommandBuilder cbUpdate = new SqlCommandBuilder(myDA);
                  OleDbCommandBuilder cb = new OleDbCommandBuilder(myDA);
                 // myDA.UpdateCommand = cb.GetUpdateCommand(); 
                    myDA.Update(DT1);
                }

        }
         public void ExportToExcel(DataGridView Tbl, string ExcelFilePath)
        {
            try
            {
                if (Tbl == null || Tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
              //  Microsoft.Office.Tools.Excel.Workbook excelapp ;//= new Microsoft.Office.Tools.Excel.Workbook();
            //    excelapp.NewSheet();
              //  Microsoft.Office.Tools.Excel.Worksheet workSheet = excelapp.Worksheets[0];
                //Microsoft.Office.Tools.Excel.
             Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
               excelApp.Workbooks.Add(true);
               
                // single worksheet
               Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)excelApp.ActiveSheet;

                // column headings
         //       for (int i = 0; i < Tbl.Columns.Count; i++)
           //     {
              //      workSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
           //     }

                // rows
                for (int i = 0; i < Tbl.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (int j = 0; j < Tbl.Columns.Count; j++)
                    {
                      //  workSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
                        workSheet.Cells[(i + 1), (j + 1)]=Tbl.Rows[i].Cells[j].Value;
                    }
                }
               

                // check fielpath
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        //workSheet.SaveAs(
                        workSheet.SaveAs(ExcelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlAddIn, null, null, null, null, null, null,null, null);
                        excelApp.Quit();
                        MessageBox.Show("Excel file saved!");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else    // no filepath is given
                {
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        public static bool AppendDataTableToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                app.Visible = false;

                //Microsoft.Office.Interop.Excel.Workbook wBook = app.Workbooks.Add(true);
                Microsoft.Office.Interop.Excel.Workbook wBook = app.Workbooks.Open(filePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet wSheet = wBook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

                //  int a = wSheet.Rows.Count; ;
                int usedRows = wSheet.UsedRange.Rows.Count;
                if (excelTable.Rows.Count > 0)
                {
                    //  int row = 0;
                    int row = excelTable.Rows.Count;
                    int col = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = excelTable.Rows[i][j].ToString();
                            wSheet.Cells[i + 1 + usedRows, j + 1] = str;
                        }
                    }
                }

                //   int size = excelTable.Columns.Count;
                //    for (int i = 0; i < size; i++)
                //    {
                //     wSheet.Cells[1, 1 + i] = excelTable.Columns[i].ColumnName;
                //   }
                //设置禁止弹出保存和覆盖的询问提示框 
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                //保存工作簿 
                wBook.Save();
                //wBook.sv
                //保存excel文件 
                //app.Save(;
              // app.Save(;
                //app.SaveWorkspace(filePath);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
            }

        }
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
           // DataToExcel(this.dataGridView1);

              SaveFileDialog kk = new SaveFileDialog(); 
            kk.Title = "保存EXECL文件"; 
          //  kk.Filter = "EXECL文件(*.xls) |*.xls |所有文件(*.*) |*.*";
           kk.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx";
          kk.FilterIndex = 1;
          kk.CheckFileExists = false;
          kk.OverwritePrompt = false;
          if (kk.ShowDialog() == DialogResult.OK)
          {
              string FileName = kk.FileName;
              if (File.Exists(FileName))
              {
                  AppendDataTableToExcel(GetDgvToTable(this.dataGridView1), FileName);
                  //File.Delete(FileName);
                  return;
              } 
              ExportToExcel(this.dataGridView1, FileName);
          }
        }

        private void fileFToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        public const int MB_ICONEXCLAMATION = 48;

      
        double USL = 0;
        double LSL = 0;
        int SubGroupSize = 5;
        private void btConnect_Click(object sender, EventArgs e)
        {
            //this.zedGraphControl1.BackColor = Color.Blue;
          /*  this.zedGraphControl1.GraphPane.AxisFill.Color= Color.Black;
            this.zedGraphControl1.GraphPane.AxisBorder.Color = Color.Blue;
            this.zedGraphControl1.GraphPane.Legend.Fill.Color= Color.Yellow;
            this.zedGraphControl1.GraphPane.PaneFill .Color=Color.Black;
            this.zedGraphControl1.GraphPane.XAxis.Color = Color.Blue;
            this.zedGraphControl1.GraphPane.XAxis.
            this.zedGraphControl1.Invalidate();
            this.zedGraphControl2.Invalidate();
           * */

            if (this.textLSL.Text == "" || this.textUSL.Text == "")
            {
                MessageBox.Show("请输入USL, LSL!!!");
                return;
            }
            try
            {
                USL = System.Convert.ToDouble(textUSL.Text);
                LSL = System.Convert.ToDouble(textLSL.Text);
            }
            catch
            {
                MessageBox.Show("请输入正确的USL, LSL值!!!");
                return;
            }
            try
            {
                // this.textSubGroupSize
                SubGroupSize = System.Convert.ToInt32(textSubGroupSize.Text);
                if (SubGroupSize < 2)
                {
                    MessageBox.Show("子组大小应大于1!!!");
                    this.textSubGroupSize.Focus();
                    return;
                }

            }
            catch
            {
                this.textSubGroupSize.Focus();
                MessageBox.Show("请输入正确的子组大小!!!");
                this.textSubGroupSize.Focus();
                return;
            }

            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
           // this.dataGridView1.Columns.Clear();
          //  this.dataGridView1.Columns.Add("Data", "Data");
           //// this.dataGridView1.Columns.Add("Time", "Time");
          //  this.dataGridView1.Columns["Time"].Width = 120;
            dataGridView1.RowHeadersWidth = 80;
            if (ThreadRun)
                return;
            if (sPort.IsOpen) return;

            // dataGridView1.Controls.
            //   this.dataGridView1.Columns.
            sPort.PortName = "COM1";
         /*   if(comboBox4.SelectedIndex==0)
                 sPort.PortName = "com1";//串口的portname 
             else if (comboBox1.SelectedIndex == 1)
                 sPort.PortName = "com2";//串口的portname */
            sPort.PortName = comboBox4.SelectedItem.ToString();
            sPort.BaudRate = 9600;//串口的波特率
            //sPort.BaudRate = 2400;//串口的波特率
            sPort.DataBits = 8;
            //两个停止位
            sPort.StopBits = System.IO.Ports.StopBits.One;
            //无奇偶校验位
            sPort.Parity = System.IO.Ports.Parity.None;
            sPort.ReadTimeout = -1;
            sPort.WriteTimeout = -1;
            sPort.DtrEnable = true;
            sPort.RtsEnable = true;
            sPort.ReceivedBytesThreshold = 10;
            try
            {
                sPort.Open();
            }
            catch
            {
                sPort.Open();
            }

            t = new Thread(SerialReadDataThread);
           ThreadRun = true;
           t.Start();

        }

        private void disConnect_Click(object sender, EventArgs e)
        {
           // String[] sPortNames = System.IO.Ports.SerialPort.GetPortNames();
            //this.Text = sPortNames.Length.ToString();         
            ThreadRun = false;
            sPort.Close();  

        }

        private void btClearData_Click(object sender, EventArgs e)
        {
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
          //  this.dataGridView1.Columns.Clear();
           // this.dataGridView1.Columns.Add("Data", "Data");
          //  this.dataGridView1.Columns.Add("Time", "Time");
         //   this.dataGridView1.Columns["Time"].Width = 120;
         //   dataGridView1.RowHeadersWidth = 80;
            this.dataGridView2.Rows.Clear();
        }

        private void btSPC_Click(object sender, EventArgs e)
        {
            //this.dataGridView1.Rows[1].Cells[0].Style.BackColor= Color.Red;
            //this.dataGridView1.Rows[1].HeaderCell.Style.BackColor = Color.Red;
            
          //  DateTime a = System.DateTime.Today.Date;
        //    if (a.Year != 2015 || a.Month!= 6)
         //  {
             //   MessageBox.Show("试用期限已到!!!");
             //   this.Close();
           // }

            if (this.textLSL.Text == "" || this.textUSL.Text == "")
            {
                MessageBox.Show("请输入USL, LSL!!!");
                if (this.textUSL.Text == "")
                    this.textUSL.Focus();
                else if (this.textLSL.Text == "")
                    this.textLSL.Focus();

                return;
            }
            try
            {
                USL = System.Convert.ToDouble(textUSL.Text);                
            }
            catch
            {
                MessageBox.Show("请输入正确的USL值!!!");
                this.textUSL.Focus();
                return;
            }
            try
            {                
                LSL = System.Convert.ToDouble(textLSL.Text);
            }
            catch
            {
                this.textLSL.Focus();
                MessageBox.Show("请输入正确的LSL值!!!");
                return;
            }
            try
            {
               // this.textSubGroupSize
                SubGroupSize = System.Convert.ToInt32(textSubGroupSize.Text);
                if (SubGroupSize < 2)
                {
                    MessageBox.Show("子组大小应大于1!!!");
                    this.textSubGroupSize.Focus();
                    return;
                }

            }
            catch
            {
                this.textSubGroupSize.Focus();
                MessageBox.Show("请输入正确的子组大小!!!");
                this.textSubGroupSize.Focus();
                return;
            }
            dataGridV1_V2(SubGroupSize);

            DataTable dtSource = GetDgvToTable(this.dataGridView1);
         //   DataTable dtSource = getSPCDataTable(dataGridView2,3);
            if (dtSource.Rows.Count < SubGroupSize)
            {
                MessageBox.Show("数据不足!!!");
                return;
            }

            ShowXbarRChart(dtSource, SubGroupSize, zedGraphControl1, zedGraphControl2);
           
            ShowDistributionChart(dtSource, zedGraphControl3);
        }

        public void UpdateComboBox(DataTable DT)
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            comboBox5.Items.Clear();
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                string ss0 = DT.Rows[i][0].ToString();
                string ss1 = DT.Rows[i][1].ToString();
                string ss2 = DT.Rows[i][2].ToString();
                string ss3 = DT.Rows[i][3].ToString();
                int j = 0;
                for (j = 0; j < comboBox1.Items.Count; j++)
                {
                    if (comboBox1.Items[j].ToString() == ss0)
                        break;
                }
                if (j == comboBox1.Items.Count)
                    comboBox1.Items.Add(ss0);

                for (j = 0; j < comboBox2.Items.Count; j++)
                {
                    if (comboBox2.Items[j].ToString() == ss1)
                        break;
                }
                if (j == comboBox2.Items.Count)
                    comboBox2.Items.Add(ss1);

                for (j = 0; j < comboBox3.Items.Count; j++)
                {
                    if (comboBox3.Items[j].ToString() == ss2)
                        break;
                }
                if (j == comboBox3.Items.Count)
                    comboBox3.Items.Add(ss2);


                for (j = 0; j < comboBox5.Items.Count; j++)
                {
                    if (comboBox5.Items[j].ToString() == ss3)
                        break;
                }
                if (j == comboBox5.Items.Count)
                    comboBox5.Items.Add(ss3);

            }

        }
        public void UpdateDataBase_Data(DataTable DT,int start,int end)
        {
            ReadMDB readmdb1 = new ReadMDB();
            if (readmdb1.IsTableExist("TestData"))
                return;
            else
            {
                string sqlstr = "create table TestData(数据 double, 日期时间 datetime";
                int nn=DT.Columns.Count;
                for (int i = start; i <= end; i++)
                    sqlstr += ("," + DT.Columns[i].ColumnName + " varchar(50)");
                sqlstr += ")";
             // sqlstr=@"create table TestData2( 数据 double, 日期时间 varchar(50),车型 varchar(50),编号 varchar(50),宽度or厚度 varchar(50))";
               // string sql = "create table MyTable ( 编号 autoincrement(1,1) , 姓名 varchar(50) , constraint pk_test_id primary key(编号))";
                readmdb1.MDBreader(sqlstr);
            }

      //   DataTable a=  (DataTable)readmdb1.MDBreader(@"SELECT Count (*) AS con FROM MSysObjects WHERE name='Data'");
        }
        private void UpdateDateGridView1()
        {
            ReadMDB readmdb1 = new ReadMDB();
            string[] ColNames = readmdb1.GetColumnsName("TestData");

         //   DataGridViewColumn Col00 = new DataGridViewColumn();
           //// Col00.Name = ColNames[0];
          //  Col00.ValueType = System.Type.GetType("System.Double");
         //   this.dataGridView1.Columns.Add(Col00);
            this.dataGridView1.Columns.Clear();
            for(int i=0;i<ColNames.Length;i++)
            {
              //  DataGridViewColumn Col0 = new DataGridViewColumn(DataGridViewCell.;
               // Col0.Name=ColNames[i];
               // Col0.HeaderText = ColNames[i];
                this.dataGridView1.Columns.Add(ColNames[i],ColNames[i]);
            
            }
            
           // DataGridViewColumn Colnames = new DataGridViewColumn();
           // Colnames.ValueType=

 
        }
        public void UpdateSysInfo(DataTable DT)
        {
         
            this.label5.Text = DT.Columns[0].ColumnName;
            this.label6.Text = DT.Columns[1].ColumnName;
            this.label7.Text = DT.Columns[2].ColumnName;
            this.label10.Text = DT.Columns[3].ColumnName;
            this.label11.Text = DT.Columns[4].ColumnName;
            this.label12.Text = DT.Columns[5].ColumnName;
            UpdateComboBox(DT);
            UpdateDataBase_Data(DT,0,3);
            UpdateDateGridView1();
           
            
        }
        /*
        public void updatecomboBox(DataTable DT)
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                string ss0=DT.Rows[i][0].ToString();
                string ss1 = DT.Rows[i][1].ToString();
                int j = 0;
                for (j = 0; j < comboBox1.Items.Count; j++)
                {
                    if (comboBox1.Items[j].ToString() == ss0)
                        break; 
                }
                if (j == comboBox1.Items.Count)
                    comboBox1.Items.Add(ss0);
           
                for (j = 0; j < comboBox2.Items.Count; j++)
                {
                    if (comboBox2.Items[j].ToString() == ss1)
                        break;
                }
                if (j == comboBox2.Items.Count)
                comboBox2.Items.Add(ss1);
            }
 
        }
         * */

        public void updateUSL_LSL(DataTable DT)
        {
            try
            {
                int i = 0;
                for ( i = 0; i < DT.Rows.Count; i++)
                {
                    if (DT.Rows[i][0].ToString() == comboBox1.SelectedItem.ToString() && DT.Rows[i][1].ToString() == comboBox2.SelectedItem.ToString() && DT.Rows[i][2].ToString() == comboBox3.SelectedItem.ToString() && DT.Rows[i][3].ToString() == comboBox5.SelectedItem.ToString())
                    {
                       // textBox1.Text = DT.Rows[i][4].ToString();
                       // textBox2.Text = DT.Rows[i][5].ToString();

                         textUSL.Text = DT.Rows[i][6].ToString();
                         textLSL.Text = DT.Rows[i][7].ToString();                      
                        textSubGroupSize.Text = DT.Rows[i][8].ToString();
                        productPicBox.ImageLocation = DT.Rows[i][9].ToString();                      
                        SubGroupSize = System.Convert.ToInt32(textSubGroupSize.Text);                            
                        break;
                    }
                }
                if (i == DT.Rows.Count)
                {
                   // textBox1.Text = "";
                    //textBox2.Text = "";
                    textUSL.Text = "";
                    textLSL.Text = "";
                    textSubGroupSize.Text = "";
                    productPicBox.Image = null;
                    SubGroupSize = 5;
                }
            }
            catch
            {
                //textBox1.Text = "";
                //textBox2.Text = "";
                textUSL.Text = "";
                textLSL.Text = "";
                textSubGroupSize.Text = "";
                productPicBox.Image = null;
                SubGroupSize = 5;
                ;
            }
        
        }
        public static DataTable OpenCSV(string filePath, int Coln)
        {
            // Encoding encoding = System.Data.Common.GetType(filePath); //Encoding.ASCII;//
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

            StreamReader sr = new StreamReader(fs, Encoding.UTF8);
            //StreamReader sr = new StreamReader();
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
           
            for (int i = 0; i < Coln; i++)
            {
                DataColumn dc = new DataColumn();
                dt.Columns.Add(dc);
            }
            while ((strLine = sr.ReadLine()) != null)
            {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    int columnCount = aryLine.Length;
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
               
            }
        

            sr.Close();
            fs.Close();
            return dt;
        }
        Login login = new Login();
        private void ManageToolStripMenuItem_Click(object sender, EventArgs e)
        {

            // DialogResult DR= form2.ShowDialog(this);
            if (login.ShowDialog(this) == DialogResult.OK)
            {
                if (form2.ShowDialog(this) == DialogResult.OK)
                {
                    ManageDataTable = form2.ManageDataTable1;
                   // updatecomboBox(ManageDataTable);
               //     this.UpdateSysInfo(ManageDataTable);
                    UpdateComboBox(ManageDataTable);
                }
                // GetDgvToTable(form2.da)

            }
        }

        private void UpdateComboBox1_2(DataTable DT)
        {
        //    string tmpss =null;
       //     if(comboBox2.SelectedItem!=null)
           //     tmpss = this.comboBox2.SelectedItem.ToString();
            this.comboBox2.Items.Clear();
            string ComboboxString1 = this.comboBox1.SelectedItem.ToString();
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                if (DT.Rows[i][0].ToString() == ComboboxString1)
                {
                     string ss1=DT.Rows[i][1].ToString() ;
                     int j = 0;
                     for (j = 0; j < comboBox2.Items.Count; j++)
                     {
                         if (comboBox2.Items[j].ToString() == ss1)
                             break;
                     }
                     if (j == comboBox2.Items.Count)
                         comboBox2.Items.Add(DT.Rows[i][1].ToString());
                }
            }
 
        }
        private void comboBox1Changed(object sender, EventArgs e)
        {
           // updateUSL_LSL(ManageDataTable);

            UpdateComboBox1_2(ManageDataTable);
            updateUSL_LSL(ManageDataTable);
        }
        private void comboBoxChanged(object sender, EventArgs e)
        {       
            updateUSL_LSL(ManageDataTable);
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x00A1 && m.WParam.ToInt32() == 2)
            {
                m.Msg = 0x0201;
                m.LParam = IntPtr.Zero;
            }
            base.WndProc(ref m);
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //sPort.NewLine = "\r";
           // string ss = sPort.ReadLine();
           // sPort
            int count = sPort.BytesToRead;
            while (sPort.BytesToRead == 0)
                ;
            string ss1=sPort.ReadExisting();
            // this.label1.Text = Encoding.Unicode.GetString(a);
            this.label1.Text = ss1;
        }

        private void COMChanged(object sender, EventArgs e)
        {
            ThreadRun = false;
            sPort.Close();  
        }

        bool dataGridV1_V2(int n)
        {
            DataGridView dgv1 = this.dataGridView1;            
            DataGridView dgv2 = this.dataGridView2;
             dgv2.Rows.Clear();
             if (dgv2.Columns.Count != (n+2))
            {
                //dgv2.Rows.Clear();
                dgv2.Columns.Clear();
                dgv2.Columns.Add("Measuring DETA", "Measuring DETA");
                for (int i = 1; i <= n; i++)
                {
                    dgv2.Columns.Add(i.ToString(), "BREADTHS");
                }
                dgv2.Columns.Add("平均值", "平均值");
                dgv2.RowHeadersWidth = 80;
             }
            double m1 = 0;
   
             for (int i = 0; i < dgv1.Rows.Count; i++)
             {
                 if (i % n == 0)
                 {
                     dgv2.Rows.Add();
                     dgv2.Rows[dgv2.Rows.Count - 1].HeaderCell.Value = (dgv2.Rows.Count).ToString();
                     if (dgv1.ColumnCount >= 1 && dgv1.Rows[i].Cells[1].Value.ToString().Length>10)
                     {
                         string ss=dgv1.Rows[i].Cells[1].Value.ToString();
                          dgv2.Rows[dgv2.Rows.Count-1].Cells[0].Value =ss.Substring(0,ss.IndexOf(' '));                     }
                     else
                           dgv2.Rows[dgv2.Rows.Count-1].Cells[0].Value ="----------";
                         


                  }

                    dgv2.Rows[dgv2.Rows.Count-1].Cells[i % n+1].Value = dgv1.Rows[i].Cells[0].Value;
                    m1 +=System.Convert.ToDouble(dgv2.Rows[dgv2.Rows.Count - 1].Cells[i % n+1].Value);

                    if ((i +1)% n == 0)
                    {                   
                        dgv2.Rows[dgv2.Rows.Count - 1].Cells[n+1].Value = m1 / n;
                        m1 = 0;
                    }

              }

            
            return true;
        }

        bool dataGridView2AddData(double d0,int n)
        {
            DataGridView dgv2 = this.dataGridView2;
            DataGridView dgv1= this.dataGridView1;
            if (dgv2.Columns.Count != n+2)
            {
                dgv2.Rows.Clear();
                dgv2.Columns.Clear();
                dgv2.Columns.Add("Measuring DETA", "Measuring DETA");
                for (int i = 1; i <= n; i++)
                {
                    dgv2.Columns.Add(i.ToString(), "BREADTHS");
                }
                dgv2.Columns.Add("平均值", "平均值");
                dgv2.RowHeadersWidth = 80;
            }
            int rn = this.dataGridView1.Rows.Count-1;
            if (rn % n == 0)
            {
                if (rn != 0)
                {
                    double m1 = 0;
                    for (int i = 0; i < n; i++)
                    {
                        m1 += (double)dgv2.Rows[dgv2.Rows.Count - 1].Cells[i+1].Value; 
                    }
                     dgv2.Rows[dgv2.Rows.Count - 1].Cells[n+1].Value = m1 / n;
                    
                }

                dgv2.Rows.Add();
                dgv2.Rows[dgv2.Rows.Count - 1].HeaderCell.Value = (dgv2.Rows.Count).ToString();
                 string ss=dgv1.Rows[dgv1.RowCount-1].Cells[1].Value.ToString();
                 dgv2.Rows[dgv2.Rows.Count-1].Cells[0].Value =ss.Substring(0,ss.IndexOf(' '));        
                 
                         
            }

            dgv2.Rows[dgv2.Rows.Count - 1].Cells[rn % n+1].Value =d0;
        

            return true;
 
        }

        DataTable getSPCDataTable(DataGridView dgv2, int cn)
        {
            DataTable DT=new DataTable();
            DT.Columns.Add();
     
            for (int i = 0; i < dgv2.Rows.Count; i++)
            {
                string tmp = Convert.ToString(dgv2.Rows[i].Cells[cn].Value);
                if (tmp == "") continue;
                DataRow dr = DT.NewRow();
                dr[0] = tmp;
                DT.Rows.Add(dr);
              }

            return DT;           

        }

        Bitmap tempBit = null;
        private void printPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
         //   printDocument1.Print
            tempBit = CopyScreenBitmap();
            printPreviewDialog1.Document = printDocument1;
             try
           {
               printPreviewDialog1.ShowDialog();
             }
            catch(Exception excep)
              {
               MessageBox.Show(excep.Message, "打印出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
               }
        }
        
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Graphics g = e.Graphics; //获得绘图对象
          //  g.DrawEllipse(new Pen(Color.Red), 10, 10, 300, 200);
              Font font = new Font("宋体", 10f); //字是什么样子的？
            Brush brush = Brushes.Black; //用红色涂上我的字吧；
            string s0,s1,s2;
            try
            {
                 s0 = this.comboBox1.SelectedItem.ToString();
            }
            catch
            {
                s0="  ";
            }

            try
            {
                s1 = this.comboBox2.SelectedItem.ToString();
            }
            catch
            {
                s1 = "  ";
            }
            try
            {
                s2 = this.comboBox3.SelectedItem.ToString();
            }
            catch
            {
                s2 = "  ";
            }
             string ss=string.Format("车型:{0}\t产品编号:{1}\t {2}",s0,s1,s2);
           g.DrawString(ss, font, brush, 10, 10);

           string ss1 = string.Format("USL:{0} \tLSL:{1}\t  子组大小:{2}",this.textUSL.Text, this.textLSL.Text, this.textSubGroupSize.Text);
           g.DrawString(ss1, font, brush, 10, 30);

           string ss2 = string.Format("UCL:{0:n2} \tLCL:{1:n2}\t     UCLR:{2:n2}\t LCLR:{3:n2}", this.UCL, this.LCL, this.UCLR, this.LCLR);
           g.DrawString(ss2, font, brush, 10, 50);
           string ss3 = string.Format("Sigma:{0:n2} \tCp:{1:n2}\t Cpu:{2:n2}\t Cpl:{3:n2}\tCpk:{4:n2}", this.sigma, this.Cp, this.Cpu, this.Cpl, this.Cpk);
           g.DrawString(ss3, font, brush, 10, 70);
            
            if (tempBit != null)
           {
              g.DrawImage(tempBit,10,100);
            }
         //   Bitmap aa=this.zedGraphControl1.Image;
            //g.DrawImage(aa ,10, 100);
         //   System.Windows.Forms.Clipboard.GetImage();
         
                g.DrawImage(this.logoPicBox.Image, 650, 0,200,100);          
            
        }

        private void printToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tempBit = CopyScreenBitmap();
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument1;
              
          if (printDialog.ShowDialog() == DialogResult.OK)
          {
              try
              {
                printDocument1.Print();
                  }
                   catch(Exception excep)
              {
                     MessageBox.Show(excep.Message, "打印出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                       printDocument1.PrintController.OnEndPrint(printDocument1,new PrintEventArgs());
               }
            }
              
        }
        Bitmap CopyScreenBitmap()
        {
            int width = 750;
            int hight = 810;
            Bitmap bit = new Bitmap(width, hight);
            Graphics g = Graphics.FromImage(bit);
            //g.CopyFromScreen(this.Location, new Point(0, 0), bit.Size);
            g.CopyFromScreen(this.Location.X, this.Location.Y+50, 0, 0, bit.Size);
         //   SaveFileDialog saveFileDialog = new SaveFileDialog();
         //   saveFileDialog.Filter = "jpg|*.jpg|gif|*.gif|bmp|*.bmp";
          //  if (saveFileDialog.ShowDialog() != DialogResult.Cancel)
          //  {
             //   bit.Save(saveFileDialog.FileName);
          //  }
            g.Dispose();
            return bit;
        }
        private void printSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  PrintDialog printDialog = new PrintDialog();
         //   printDialog.Document = printDocument1;
         //   printDialog.ShowDialog();
           
        }

        private void btSaveData2DataBase_Click(object sender, EventArgs e)
        {
            DataTable DT=GetDgvToTable(this.dataGridView1);
            ReadMDB readmdb1 = new ReadMDB();
            readmdb1.InsertmdbNoRepFast(DT, "TestData");
            this.btClearData.PerformClick();
        }

        private void datatabe2view(DataTable DT, DataGridView DGV)
        {
            DGV.RowHeadersWidth = 80;
            for (int count = 0; count < DT.Rows.Count; count++)
            {
                //DataGridViewRow dr = new DataGridViewRow();
                int index = DGV.Rows.Add();
                for (int countsub = 0; countsub < DT.Columns.Count  ; countsub++)
                {
                    
                    DGV.Rows[index].Cells[countsub].Value = Convert.ToString(DT.Rows[count][countsub]);
                   
                }
                DGV.Rows[index].HeaderCell.Value = Convert.ToString(index+1);
                
                

            }

        }

        private void btLoadData_Click(object sender, EventArgs e)
        {
            try                
            {
                dataGridView1.Rows.Clear();
                ReadMDB readmdb1 = new ReadMDB();
                string[] ColNames = readmdb1.GetColumnsName("config");
              //  List<string> ColNames = readmdb1.GetColumnsName1("config");
                string date0 = this.dateTimePicker1.Value.ToShortDateString();
                string date1 = this.dateTimePicker2.Value.AddDays(1).ToShortDateString();
                string sqlss = "select * from TestData where " + ColNames[0] + "= '" + comboBox1.SelectedItem.ToString() + "' and " + ColNames[1] + "= '" + comboBox2.SelectedItem.ToString() + "' and " + ColNames[2] + "= '" + comboBox3.SelectedItem.ToString() + "' and "
                    + ColNames[3] + "= '" + comboBox5.SelectedItem.ToString()                     
                    + "' and 日期时间 >=#" + date0 + "# and 日期时间 <=#" + date1 + "#" + " order by 日期时间";
                DataTable DT = (DataTable)readmdb1.MDBreader(sqlss);
                //  this.dataGridView1.Columns.Clear();
                 datatabe2view(DT, dataGridView1);
                //this.dataGridView1.DataSource = DT;

              //  ColNames.Clear();
               
                
            }
            catch
           {
           //   // this.Dispose();
                this.openFileDialog1.Dispose();
           }
           // this.Dispose();
            this.openFileDialog1.Dispose();
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            int ColN = dataGridView1.Columns.Count;
           // if (ColN <= 0 || this.dataGridView1.SelectedRows.Count <= 0)
             //   return;
            if (DialogResult.OK == MessageBox.Show("确认删除数据！删除后无法恢复", "警告", MessageBoxButtons.OKCancel))
            {
                ReadMDB readmdb1 = new ReadMDB();
                foreach (DataGridViewRow DGVR in this.dataGridView1.SelectedRows)
                {
                    string sqlss = "delete * from testdata where ";
                    sqlss += " " + dataGridView1.Columns[0].Name + " = " + DGVR.Cells[0].Value;
                    sqlss += " and " + dataGridView1.Columns[1].Name + " = #" + DGVR.Cells[1].Value + "#";
                    for (int i = 2; i < ColN; i++)
                    {
                        sqlss += " and " + dataGridView1.Columns[i].Name + " = '" + DGVR.Cells[i].Value + "'";
                        //DGVR.Cells[i].ValueType.Name.ToString() == "String"
                     //   if (i>0)
                      //     sqlss += " " + dataGridView1.Columns[i].Name + " = '" + DGVR.Cells[i].Value + "'";
                       // else
                           // sqlss += " " + dataGridView1.Columns[i].Name + " = " + DGVR.Cells[i].Value;
                      //  if (i != ColN - 1)
                         //   sqlss += " and ";
                    }

                    readmdb1.MDBreader(sqlss);

                    this.dataGridView1.Rows.Remove(DGVR);

                }
            }
        }

        private void btMonthTrend_Click(object sender, EventArgs e)
        {
            DataTable dtSource = GetDgvToTable(this.dataGridView1);
            this.ShowDistributionPerMonthChart(dtSource, this.zedGraphControl3);
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            copyright CPDlg = new copyright();
            CPDlg.ShowDialog();

        }

        private void openHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("SPC帮助.pdf");
        }

       

    }
}