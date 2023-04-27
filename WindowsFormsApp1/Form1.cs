using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
            Login test = (LoginLoad.Verif(textBox1.Text, textBox2.Text));
            if (test.Rol == null)
            {
                MessageBox.Show("Introdu Username-ul si Parola corecta!");
            }
            else
            {
                this.Hide();
                Form2 m = new Form2();
                m.Show();
            }
        }
    }
    }

