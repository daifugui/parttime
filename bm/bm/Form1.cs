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
            draw_dis_line(40, 170, 250, 170);

        }

        private void draw_dis_line(int x1, int y1, int x2, int y2)
        {
            Graphics g = pictureBox1.CreateGraphics();
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
    }
}
