using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace myStatisticalAnalysis
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            //    ManageDataTable1 = OpenCSV(System.Environment.CurrentDirectory +"\\config");
            //  dataGridView2.Rows.Clear();
            //  datatabe2view(ManageDataTable1, dataGridView2);
        }
        void saveconfig(DataGridView m_DataView, string FileName)
        {
            if (File.Exists(FileName))
                File.Delete(FileName);
            FileStream objFileStream;
            StreamWriter objStreamWriter;
            string strLine = "";
            objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write);
            objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);
            /*  for (int i = 0; i  < m_DataView.Columns.Count; i++) 
              { 
               if (m_DataView.Columns[i].Visible == true) 
                 { 
                   strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9); 
                   } 
               } 
              objStreamWriter.WriteLine(strLine); 
             */

            strLine = "";

            for (int i = 0; i < m_DataView.Rows.Count; i++)
            {
                if (m_DataView.Columns[0].Visible == true)
                {
                    if (m_DataView.Rows[i].Cells[0].Value == null)
                        strLine = strLine + " " + Convert.ToChar(9);
                    else
                        strLine = strLine + m_DataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9);
                }
                for (int j = 1; j < m_DataView.Columns.Count; j++)
                {
                    if (m_DataView.Columns[j].Visible == true)
                    {
                        if (m_DataView.Rows[i].Cells[j].Value == null)
                            strLine = strLine + " " + Convert.ToChar(9);
                        else
                        {
                            string rowstr = "";
                            rowstr = m_DataView.Rows[i].Cells[j].Value.ToString();
                            if (rowstr.IndexOf("\r\n") > 0)
                                rowstr = rowstr.Replace("\r\n", " ");
                            if (rowstr.IndexOf("\t") > 0)
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
            MessageBox.Show(this, "保存成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        void loadconfig(string FileName, DataGridView m_DataView)
        {
            // if (File.Exists(FileName))
            //    File.Delete(FileName);
            FileStream objFileStream;
            StreamReader objStreamReader;
            string strLine = "";
            objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Read);
            objStreamReader = new StreamReader(objFileStream, System.Text.Encoding.Unicode);
            /*  for (int i = 0; i  < m_DataView.Columns.Count; i++) 
              { 
               if (m_DataView.Columns[i].Visible == true) 
                 { 
                   strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9); 
                   } 
               } 
              objStreamWriter.WriteLine(strLine); 
             */
            int i = 0;
            //  objStreamReader.ReadLine(
            while (true)
            {
                strLine = objStreamReader.ReadLine();
                if (strLine == null) break;
                string[] strLines = strLine.Split(new Char[] { ',' });
                m_DataView.Rows.Add();
                for (int j = 0; j < strLines.Length; j++)
                    m_DataView.Rows[i].Cells[j].Value = strLines[j];
                i++;
            }

            objStreamReader.Close();
            objFileStream.Close();
        }
        public static void SaveCSV(DataTable dt, string fullPath)
        {
            FileInfo fi = new FileInfo(fullPath);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            FileStream fs = new FileStream(fullPath, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            //StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            string data = "";
            //写出列名称
            /*   for (int i = 0; i < dt.Columns.Count; i++)
               {
                   data += dt.Columns[i].ColumnName.ToString();
                   if (i < dt.Columns.Count - 1)
                   {
                       data += ",";
                   }
               }
               sw.WriteLine(data);
             * 
             * */

            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //if(j==dt.Columns.Count-1)
                    //dt.Rows[i][j].tip
                    string str = dt.Rows[i][j].ToString();
                    //  str = str.Replace("\"", "\"\"");
                    //   if (str.Contains(',') || str.Contains('"')|| str.Contains('\r') || str.Contains('\n'))
                    //   {
                    //     str = string.Format("\"{0}\"", str);
                    //     }

                    data += str;
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
            // MessageBox.Show("CSV文件保存成功！");


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
            string[] tableHead = null;
            //标示列数
            // int columnCount = 0;
            //标示是否是读取的第一行
            // bool IsFirst = true;
            //逐行读取CSV中的数据
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
                if (columnCount == Coln - 1)
                {
                    for (int j = 0; j < columnCount-1; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dr[columnCount - 1] = 5; 
                    dr[columnCount] = aryLine[columnCount - 1];
                }
                else
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                }
                dt.Rows.Add(dr);

            }


            sr.Close();
            fs.Close();
            return dt;
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
                        workSheet.Cells[(i + 1), (j + 1)] = Tbl.Rows[i].Cells[j].Value;
                    }
                }

                // check fielpath
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        //workSheet.SaveAs(
                        workSheet.SaveAs(ExcelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlAddIn, null, null, null, null, null, null, null, null);
                        excelApp.Quit();
                        //MessageBox.Show("!");
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
                for (int countsub = 0; countsub < dgv.Columns.Count - 1; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                    //dr[countsub] = dgv.Rows[count].Cells[countsub].Value;
                }
                if (dgv.Rows[count].Cells[dgv.Columns.Count - 1].Tag == null)
                    dr[dgv.Columns.Count - 1] = "null";
                else
                    dr[dgv.Columns.Count - 1] = Convert.ToString(dgv.Rows[count].Cells[dgv.Columns.Count - 1].Tag);

                dt.Rows.Add(dr);
            }
            return dt;
        }
        public DataTable ManageDataTable1 = new DataTable();
        private void btSave_Click(object sender, EventArgs e)
        {
            //   loadconfig( "config",dataGridView2);
            //   (Form1)this.Parent.
            // Environment.CurrentDirectory();
            ManageDataTable1 = GetDgvToTable(dataGridView2);
         //   SaveCSV(ManageDataTable1, Application.StartupPath + "\\config");
            ReadMDB readmdb1 = new ReadMDB();
            readmdb1.MDBreader("DELETE from Config");
            readmdb1.SaveDataTable("Config", ManageDataTable1);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        private void datatabe2view(DataTable DT, DataGridView DGV)
        {
            for (int count = 0; count < DT.Rows.Count; count++)
            {
                //DataGridViewRow dr = new DataGridViewRow();
                int index = DGV.Rows.Add();
                for (int countsub = 0; countsub < DT.Columns.Count - 1; countsub++)
                {
                    // dr.Cells[index].Value = Convert.ToString(DT.Rows[count][countsub]);
                    DGV.Rows[index].Cells[countsub].Value = Convert.ToString(DT.Rows[count][countsub]);
                    //dr.Cells[
                    //dr.Cells[]
                }
                if (DT.Rows[count][DT.Columns.Count - 1] != DBNull.Value && DT.Rows[count][DT.Columns.Count - 1].ToString() != "".ToString() && DT.Rows[count][DT.Columns.Count - 1].ToString() != "null".ToString())
                {
                    DGV.Rows[index].Cells[DT.Columns.Count - 1].Tag = DT.Rows[count][DT.Columns.Count - 1].ToString();
                    if(File.Exists(DGV.Rows[index].Cells[DT.Columns.Count - 1].Tag.ToString()))
                         DGV.Rows[index].Cells[DT.Columns.Count - 1].Value = new Bitmap(DT.Rows[count][DT.Columns.Count - 1].ToString());

                }

            }

        }
        private void btLoad_Click(object sender, EventArgs e)
        {
            ManageDataTable1 = OpenCSV(Application.StartupPath + "\\config", dataGridView2.ColumnCount);
            dataGridView2.Rows.Clear();
            datatabe2view(ManageDataTable1, dataGridView2);
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();

        }
        private void FillDataGridViewWithDataSource(DataGridView dataGridView, DataTable dTable)
        {
            //1.清空旧数据
            dataGridView.Rows.Clear();
            //2.填充新数据
            if (dTable != null && dTable.Rows.Count > 0)
            {
                //设置DataGridView列数据
                dataGridView.Columns["ITEM_NO"].DataPropertyName = "ITEM_NO";
                dataGridView.Columns["ITEM_NAME"].DataPropertyName = "ITEM_NAME";
                dataGridView.Columns["INPUT_CODE"].DataPropertyName = "INPUT_CODE";

                //设置数据源，部分显示数据
                dataGridView.DataSource = dTable;
                dataGridView.AutoGenerateColumns = false;
            }
        }
        private void FillDataGridViewWithForeach(DataGridView dataGridView, DataTable dTable)
        {
            //1.清空旧数据
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
            foreach(DataColumn DC in dTable.Columns)
            {
                if (DC.ColumnName == "图片")
                {
                    DataGridViewImageColumn column = new DataGridViewImageColumn();
                    column.HeaderText = "图片";
                    column.ImageLayout = DataGridViewImageCellLayout.Zoom;
                    int index=dataGridView.Columns.Add(column);
                    dataGridView.Columns[index].Name = "图片";
                    
                }
                else
                    dataGridView.Columns.Add(DC.ColumnName, DC.ColumnName);
             //   dataGridView.Columns[2].CellType = DataGridView.
            }
            //2.赋值新数据
            foreach (DataRow row in dTable.Rows)
            {
                int index = dataGridView.Rows.Add();

             
               /*
                foreach (DataColumn DC in dTable.Columns)
                {
                    try
                    {

                        dataGridView.Rows[index].Cells[DC.ColumnName].Value = row[DC.ColumnName];
                    }
                    catch
                    {
                        dataGridView.Rows[index].Cells[DC.ColumnName].Value = null;
                        dataGridView.Rows[index].Cells[DC.ColumnName].Tag = null;
                        ;
                    }
               
                }
             */
            //   int i=0;
                for (int i = 0; i < dTable.Columns.Count - 1; i++)
                {
                    dataGridView.Rows[index].Cells[i].Value = row[i]; 
                }
                int nn=dTable.Columns.Count;

                if (row[nn - 1] == null || row[nn - 1].ToString() == "null" || row[nn - 1].ToString() == "")
                {
                   dataGridView.Rows[index].Cells[nn - 1].Value = null;
                    dataGridView.Rows[index].Cells[nn - 1].Tag = null;
                }
                else
                {
                    dataGridView.Rows[index].Cells[nn - 1].Value = new Bitmap(row[nn - 1].ToString());
                    dataGridView.Rows[index].Cells[nn - 1].Tag = row[nn - 1].ToString();
 
                }
               
            }
        }

        private void form2shown(object sender, EventArgs e)
        {
           // dataGridView2.Rows.Clear();
            //ManageDataTable1 = OpenCSV(Application.StartupPath + "\\config", dataGridView2.ColumnCount);
            ReadMDB readmdb1 = new ReadMDB();
            ManageDataTable1 =(DataTable) readmdb1.MDBreader("select * from config");
           // dataGridView2.DataSource = ManageDataTable1;
           FillDataGridViewWithForeach(dataGridView2, ManageDataTable1);

          //  dataGridView2.Rows.Clear();
          //  datatabe2view(ManageDataTable1, dataGridView2);
        }

        private void btClear_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();

        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            //for (int i = 0; i < dataGridView2.SelectedRows.Count; i++)
            //dataGridView2.Rows.RemoveAt();
            foreach (DataGridViewRow DGVR in dataGridView2.SelectedRows)
            {
                dataGridView2.Rows.Remove(DGVR);

            }


        }

        private void btAdd_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Add();

        }

        private void btTest_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell DGVC in dataGridView2.SelectedCells)
            {
                if (DGVC.ColumnIndex == 7)
                {
                    string imgpath = Application.StartupPath + "\\1.jpg";
                    Bitmap img1 = new Bitmap(imgpath);

                    DGVC.Value = img1;
                    DGVC.Tag = imgpath;
                }
            }
        }

        private void btInsertImage_Click(object sender, EventArgs e)
        {
             this.openFileDialog1.Filter = "(*.jpg)|*.jpg|(*.bmp)|*.bmp";
             if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
             {
                 string FileName = this.openFileDialog1.FileName;
                 foreach (DataGridViewCell DGVC in dataGridView2.SelectedCells)
                 {
                     if (DGVC.ColumnIndex == dataGridView2.ColumnCount-1)
                     {
                       //  string imgpath = Application.StartupPath + "\\1.jpg";
                     //    FileName
                         Bitmap img1 = new Bitmap(FileName);

                         DGVC.Value = img1;
                         DGVC.Tag = FileName;
                     }
                 }
             }

        }




    }
}