using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace bm
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            int index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = "直径1";
            this.dataGridView1.Rows[index].Cells[1].Value = "3";
            this.dataGridView1.Rows[index].Cells[2].Value = "15";
            this.dataGridView1.Rows[index].Cells[3].Value = "5";



            index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = "直径2";
            this.dataGridView1.Rows[index].Cells[1].Value = "3";
            this.dataGridView1.Rows[index].Cells[2].Value = "15";
            this.dataGridView1.Rows[index].Cells[3].Value = "5";
            index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = "直径3";
            this.dataGridView1.Rows[index].Cells[1].Value = "3";
            this.dataGridView1.Rows[index].Cells[2].Value = "15";
            this.dataGridView1.Rows[index].Cells[3].Value = "5";

            index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = "圆度";
            this.dataGridView1.Rows[index].Cells[1].Value = "3";
            this.dataGridView1.Rows[index].Cells[2].Value = "0.03";
            this.dataGridView1.Rows[index].Cells[3].Value = "0";


            index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = "高度1";
            this.dataGridView1.Rows[index].Cells[1].Value = "3";
            this.dataGridView1.Rows[index].Cells[2].Value = "15";
            this.dataGridView1.Rows[index].Cells[3].Value = "10";
            index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = "高度2";
            this.dataGridView1.Rows[index].Cells[1].Value = "3";
            this.dataGridView1.Rows[index].Cells[2].Value = "15";
            this.dataGridView1.Rows[index].Cells[3].Value = "10";

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow DGVR in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(DGVR);

            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
