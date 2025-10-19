using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
namespace lapoindustries
{
    public partial class login : Form
    {
        string stringdeconexao=
            "server = localhost;uid=root;pwd=123456;database=dados2025";
        MySql.Data.MySqlClient.MySqlConnection conn;
        string strdemensagem;
        public login()
        {
            try
            {
                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = stringdeconexao;
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                strdemensagem = ex.Message;
                MessageBox.Show(strdemensagem);
            }
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || txtSenha.Text == "")
            {
                MessageBox.Show("Campos em Branco");
            }
            else 
            {
                int login = int.Parse(textBox1.Text.Trim());
                string senha = txtSenha.Text.Trim();
                string strsql = "SELECT * FROM login WHERE log_log =" + login;
                try
                {
                    MySqlCommand cmd = new MySqlCommand(strsql, conn);
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt.Rows.Count>0) 
                    {
                        string senhabd = Convert.ToString(dt.Rows[0][1]);
                        if (senha == senhabd)
                        {
                            menu me;
                            me = new menu();
                            me.Show();
                            Hide();
                        }
                        else
                        {
                            MessageBox.Show("Senha Incorreta");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Login Incorreto");
                    }
                }
                catch (MySql.Data.MySqlClient.MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
            System.Windows.Forms.Application.Exit();
        }
 
        private void button3_Click(object sender, EventArgs e)
        {
            
        }
        private bool senhaVisivel = false;
        private void txtSenha_TextChanged(object sender, EventArgs e)
        {             
            txtSenha.UseSystemPasswordChar = true;
            pbMostrarSenha.Image = Properties.Resources.fechado;
        }
        private void pbMostrarSenha_Click(object sender, EventArgs e)
        {
            senhaVisivel = !senhaVisivel;

            if (senhaVisivel)
            {
                txtSenha.UseSystemPasswordChar = false;
                pbMostrarSenha.Image = Properties.Resources.aberto;
            }
            else
            {
                txtSenha.UseSystemPasswordChar = true;
                pbMostrarSenha.Image = Properties.Resources.fechado;
            }
        }
    }
}