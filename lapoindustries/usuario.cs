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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ClosedXML.Excel;

namespace lapoindustries
{
    public partial class usuario : Form
    {
        string stringdeconexao =
            "server = localhost;uid=root;pwd=123456;database=dados2025";
        MySql.Data.MySqlClient.MySqlConnection conn;
        string strdemensagem;
        public usuario()
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

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                string strsql = "INSERT INTO login(log_log,sen_log,nda_log) " +
                                              "VALUES (" +
                int.Parse(textBox1.Text.Trim()) + "," +
                                              "'" + textBox2.Text.Trim() + "'" + "," +
                int.Parse(textBox3.Text.Trim()) + ");";
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
                string strsql = "SELECT * FROM login WHERE log_log = " + codigo;
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

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                int codigo = int.Parse(textBox1.Text.Trim());

                string strsql = "SELECT * FROM login WHERE log_log = " + codigo;
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
                        string strsql2 = "DELETE FROM login WHERE log_log = " + codigo;
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

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                int codigo = int.Parse(textBox1.Text.Trim());

                string strsql = "SELECT * FROM login WHERE log_log = " + codigo;
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
                        string strsql2 = "UPDATE login SET " +
                                         "sen_log = " + int.Parse(textBox2.Text.Trim()) + "," +
                                          "nda_log = " + int.Parse(textBox3.Text.Trim()) +
                                          " WHERE log_log = " + codigo;
                        if (MessageBox.Show("Deseja Alterar", "Alterar", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            try
                            {
                                MySqlCommand cmd2 = new MySqlCommand(strsql2, conn);
                                cmd2.ExecuteNonQuery();
                                MessageBox.Show("Uuario modificado com sucesso");
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

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {

                string conexaoString = "server=localhost;user=root;password=123456;database=dados2025;";
                string nomeTabela = "login";

                using (MySqlConnection conexao = new MySqlConnection(conexaoString))
                {
                    conexao.Open();

                    string query = $"SELECT * FROM login";
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
        private void button9_Click(object sender, EventArgs e)
        {
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void usuario_Load(object sender, EventArgs e)
        {

        }
    }
}
                


