using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace lapoindustries
{
    public partial class funcionario : Form
    {
        string stringdeconexao =
           "server = localhost;uid=root;pwd=123456;database=dados2025";
        MySql.Data.MySqlClient.MySqlConnection conn;
        string strdemensagem;
        public funcionario()
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

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void funcionario_Load(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            menu me;
            me = new menu();
            me.Show();
            Hide();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            login log;
            log = new login();
            log.Show();
            Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || textBox9.Text == "" || textBox10.Text == "" || textBox11.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                string strsql = "INSERT INTO funcionario(cod_fun,nom_fun,cpf_fun,rg_fun,ema_fun,tel_fun,end_fun,set_fun,est_fun,ida_fun,car_fun) " +
                                              "VALUES (" +
                int.Parse(textBox1.Text.Trim()) + "," +
                 "'" + textBox2.Text.Trim() + "'" + "," +
                 "'" + textBox3.Text.Trim() + "'" + "," +
                 "'" + textBox4.Text.Trim() + "'" + "," +
                 "'" + textBox5.Text.Trim() + "'" + "," +
                 "'" + textBox6.Text.Trim() + "'" + "," +
                 "'" + textBox7.Text.Trim() + "'" + "," +
                 "'" + textBox8.Text.Trim() + "'" + "," +
                 "'" + textBox9.Text.Trim() + "'" + "," +
                 int.Parse(textBox10.Text.Trim()) + "," +
                  "'" + textBox11.Text.Trim() + "');";
                try
                {
                    MySqlCommand cmd = new MySqlCommand(strsql, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Cadastro realizado com sucesso.");
                }
                catch (MySql.Data.MySqlClient.MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                int codigo = int.Parse(textBox1.Text.Trim());
                string strsql = "SELECT * FROM funcionario WHERE cod_fun = " + codigo;
                try
                {
                    MySqlCommand cmd = new MySqlCommand(strsql, conn);
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        textBox1.Text = Convert.ToString(dt.Rows[0][0]);
                        textBox2.Text = Convert.ToString(dt.Rows[0][1]);
                        textBox3.Text = Convert.ToString(dt.Rows[0][2]);
                        textBox4.Text = Convert.ToString(dt.Rows[0][3]);
                        textBox5.Text = Convert.ToString(dt.Rows[0][4]);
                        textBox6.Text = Convert.ToString(dt.Rows[0][5]);
                        textBox7.Text = Convert.ToString(dt.Rows[0][6]);
                        textBox8.Text = Convert.ToString(dt.Rows[0][7]);
                        textBox9.Text = Convert.ToString(dt.Rows[0][8]);
                        textBox10.Text = Convert.ToString(dt.Rows[0][9]);
                        textBox11.Text = Convert.ToString(dt.Rows[0][10]);

                    }
                    else
                    {
                        MessageBox.Show("Não encontrado");
                    }
                }
                catch (MySql.Data.MySqlClient.MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || textBox9.Text == "" || textBox10.Text == "" || textBox11.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                int codigo = int.Parse(textBox1.Text.Trim());

                string strsql = "SELECT * FROM funcionario WHERE cod_fun = " + codigo;
                try
                {
                    MySqlCommand cmd = new MySqlCommand(strsql, conn);
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt.Rows.Count <= 0)
                    {
                        MessageBox.Show("Não encontrado");
                    }
                    else
                    {
                        string strsql2 = "UPDATE funcionario SET " +
                            "nom_fun = " + "'" + textBox2.Text.Trim() + "'" + "," +
                            "cpf_fun = " + "'" + textBox3.Text.Trim() + "'" + "," +
                            "rg_fun = " + "'" + textBox4.Text.Trim() + "'" + "," +
                            "ema_fun = " + "'" + textBox5.Text.Trim() + "'" + "," +
                            "tel_fun = " + "'" + textBox6.Text.Trim() + "'" + "," +
                            "end_fun = " + "'" + textBox7.Text.Trim() + "'" + "," +
                            "set_fun = " + "'" + textBox8.Text.Trim() + "'" + "," +
                            "est_fun = " + "'" + textBox9.Text.Trim() + "'" + "," +
                            "ida_fun = " + int.Parse(textBox10.Text.Trim()) + "," +
                            "car_fun = " + "'" + textBox11.Text.Trim() + "'" +
                                          " WHERE cod_fun = " + codigo;
                        if (MessageBox.Show("Deseja Alterar", "Alterar", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            try
                            {
                                MySqlCommand cmd2 = new MySqlCommand(strsql2, conn);
                                cmd2.ExecuteNonQuery();
                                MessageBox.Show("Funcionario modificado com sucesso");
                            }
                            catch (MySql.Data.MySqlClient.MySqlException ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }
                    }
                }
                catch (MySql.Data.MySqlClient.MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                int codigo = int.Parse(textBox1.Text.Trim());

                string strsql = "SELECT * FROM funcionario WHERE cod_fun = " + codigo;
                try
                {
                    MySqlCommand cmd = new MySqlCommand(strsql, conn);
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt.Rows.Count <= 0)
                    {
                        MessageBox.Show("Não encontrado");
                    }
                    else
                    {
                        string strsql2 = "DELETE FROM funcionario WHERE cod_fun = " + codigo;
                        if (MessageBox.Show("Deseja Excluir", "Excluir", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            try
                            {
                                MySqlCommand cmd2 = new MySqlCommand(strsql2, conn);
                                cmd2.ExecuteNonQuery();
                                MessageBox.Show("Excluído com sucesso.");
                            }
                            catch (MySql.Data.MySqlClient.MySqlException ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }
                    }
                }
                catch (MySql.Data.MySqlClient.MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
           
                string conexaoString = "server=localhost;user=root;password=123456;database=dados2025;";
                string nomeTabela = "funcionario"; 

                using (MySqlConnection conexao = new MySqlConnection(conexaoString))
                {
                    conexao.Open();

                    string query = $"SELECT * FROM funcionario";
                    MySqlCommand cmd = new MySqlCommand(query, conexao);

                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Nenhum dado encontrado na tabela.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    
                    using (var workbook = new XLWorkbook())
                    {
                        var planilha = workbook.Worksheets.Add("Relatório");
                        planilha.Cell(1, 1).InsertTable(dt); 

                        
                        string caminhoArquivo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                            + $"\\Relatorio_{nomeTabela}_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";

                        workbook.SaveAs(caminhoArquivo);

                        MessageBox.Show($"Relatório gerado com sucesso!\n\n{caminhoArquivo}",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                        {
                            FileName = caminhoArquivo,
                            UseShellExecute = true
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao gerar relatório:\n" + ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}



