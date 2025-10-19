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
    public partial class fornecedor : Form
    {
        string stringdeconexao =
            "server = localhost;uid=root;pwd=123456;database=dados2025";
        MySql.Data.MySqlClient.MySqlConnection conn;
        string strdemensagem;
        public fornecedor()
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
            if (textBox1.Text == "" || textBox2.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || textBox9.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                string strsql = "INSERT INTO fornecedor(cod_for,cli_for,nom_for,ema_for,cnp_for,est_for,cid_for,end_for,tel_for) " +
                                              "VALUES (" +
                int.Parse(textBox1.Text.Trim()) + "," +
                 "'" + textBox2.Text.Trim() + "'" + "," +
                 "'" + textBox3.Text.Trim() + "'" + "," +
                 "'" + textBox4.Text.Trim() + "'" + "," +
                 "'" + textBox5.Text.Trim() + "'" + "," +
                 "'" + textBox6.Text.Trim() + "'" + "," +
                 "'" + textBox7.Text.Trim() + "'" + "," +
                 "'" + textBox8.Text.Trim() + "'" + "," +
                 "'" + textBox9.Text.Trim() + "');";
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
                string strsql = "SELECT * FROM fornecedor WHERE cod_for = " + codigo;
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
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || textBox9.Text == "")
            {
                MessageBox.Show("Campo em branco");
            }
            else
            {
                int codigo = int.Parse(textBox1.Text.Trim());

                string strsql = "SELECT * FROM fornecedor WHERE cod_for = " + codigo;
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
                        string strsql2 = "UPDATE fornecedor SET " +
                            "cli_for = " + "'" + textBox2.Text.Trim() + "'" + "," +
                            "nom_for = " + "'" + textBox3.Text.Trim() + "'" + "," +
                            "ema_for = " + "'" + textBox4.Text.Trim() + "'" + "," +
                            "cnp_for = " + long.Parse(textBox5.Text.Trim()) + "," +
                            "est_for = " + "'" + textBox6.Text.Trim() + "'" + "," +
                            "cid_for = " + "'" + textBox7.Text.Trim() + "'" + "," +
                            "end_for = " + "'" + textBox8.Text.Trim() + "'" + "," +
                            "tel_for = " + "'" + textBox9.Text.Trim() + "'"  +              
                                          " WHERE cod_for = " + codigo;
                        if (MessageBox.Show("Deseja Alterar", "Alterar", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            try
                            {
                                MySqlCommand cmd2 = new MySqlCommand(strsql2, conn);
                                cmd2.ExecuteNonQuery();
                                MessageBox.Show("Fornecedor modificado com sucesso");
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

                string strsql = "SELECT * FROM fornecedor WHERE cod_for = " + codigo;
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
                        string strsql2 = "DELETE FROM fornecedor WHERE cod_for = " + codigo;
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
                // ======== CONFIGURAÇÃO DO BANCO DE DADOS ========
                string conexaoString = "server=localhost;user=root;password=123456;database=dados2025;";
                string nomeTabela = "fornecedor"; // 🔹 Você pode trocar depois para cliente, fornecedor, etc.

                using (MySqlConnection conexao = new MySqlConnection(conexaoString))
                {
                    conexao.Open();

                    string query = $"SELECT * FROM fornecedor";
                    MySqlCommand cmd = new MySqlCommand(query, conexao);

                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Nenhum dado encontrado na tabela.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // ======== CRIAR ARQUIVO EXCEL ========
                    using (var workbook = new XLWorkbook())
                    {
                        var planilha = workbook.Worksheets.Add("Relatório");
                        planilha.Cell(1, 1).InsertTable(dt); // insere o DataTable direto na planilha

                        // Nome do arquivo
                        string caminhoArquivo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                            + $"\\Relatorio_{nomeTabela}_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";

                        workbook.SaveAs(caminhoArquivo);

                        MessageBox.Show($"Relatório gerado com sucesso!\n\n{caminhoArquivo}",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Abre automaticamente no Excel
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




