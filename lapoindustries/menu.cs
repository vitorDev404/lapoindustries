using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace lapoindustries
{
    public partial class menu : Form
    {
        public menu()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            usuario usu;
            usu = new usuario();
            usu.Show();
            Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cliente cli;
            cli = new cliente();
            cli.Show();
            Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fornecedor forn;
            forn = new fornecedor();
            forn.Show();
            Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            funcionario func;
            func = new funcionario();
            func.Show();
            Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            produtos pro;
            pro = new produtos();
            pro.Show();
            Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            login log;
            log = new login();
            log.Show();
            Hide();
        }
    }
}
