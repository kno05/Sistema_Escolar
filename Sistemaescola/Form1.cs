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
using System.IO;
using ClosedXML.Excel;

namespace Sistemaescola
{
    public partial class button_JSON : Form
    {
        MySqlCommand mySqlCommand;
        MySqlDataReader reader;

        public button_JSON()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CargaDatos();
        }

        private void CargaDatos()
        {
            try
            {
                MySqlConnection mySqlConnection = new MySqlConnection
                    ("host=localhost;user=root;password=1234;database=escolar");
                mySqlConnection.Open();

                MySqlDataAdapter mySqlDataAdapter = new MySqlDataAdapter
                    ("SELECT matricula as 'Expediente', ap1 as 'Apellido Paterno'," +
                    " ap2 as 'Apellido de la mama', nombre as 'Nombre', fecha_nacimiento as 'Fecha De Nacimiento'," +
                    "Correo, telefono as 'Telefono' FROM alumnos ORDER By matricula", mySqlConnection);



                DataTable dataTable = new DataTable();

                mySqlDataAdapter.Fill(dataTable);

                dataGridView1.DataSource = dataTable;
            }
            catch (Exception error)
            {

                MessageBox.Show(error.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String query = "INSERT INTO alumnos (matricula,ap1,ap2,nombre,fecha_nacimiento,correo,telefono) VALUES (+" +
                textBox1.Text + ",'" +
                textBox2.Text + "','" +
                textBox3.Text + "','" +
                textBox4.Text + "','" +
                dateTimePicker1.Value.Year + "-" +
                dateTimePicker1.Value.Month + "-" +
                dateTimePicker1.Value.Day + "','" +
                textBox5.Text + "'," +
                textBox6.Text + ")";
            MessageBox.Show(query);

            try
            {
                MySqlConnection mySqlConnection = new MySqlConnection
                    ("host=localhost;user=root;password=1234;database=escolar");
                mySqlConnection.Open();
                mySqlCommand = new MySqlCommand(query, mySqlConnection);

                mySqlCommand.ExecuteNonQuery();

                MessageBox.Show("Agregado");

            }
            catch (Exception error)
            {

                MessageBox.Show(error.ToString(), "titulo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            CargaDatos();

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String query = "DELETE FROM alumnos WHERE matricula=" +
                textBox1.Text;


            try
            {
                MySqlConnection mySqlConnection = new MySqlConnection
                    ("host=localhost;user=root;password=1234;database=escolar");
                mySqlConnection.Open();
                mySqlCommand = new MySqlCommand(query, mySqlConnection);

                mySqlCommand.ExecuteNonQuery();

                MessageBox.Show("ELIMINADO");

                CargaDatos();

            }
            catch (Exception error)
            {

                MessageBox.Show(error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            CargaDatos();

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
        }

        private void button_sql_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "Guardar SQL";
            saveFileDialog1.FileName = "*.sql";
            saveFileDialog1.InitialDirectory = @"C:\Users\kno_0\Google Drive\UNIKINO\Base de Datos 2\06-04-19";
            saveFileDialog1.Filter = "archivo sql |*.sql";

            saveFileDialog1.ShowDialog();

            String archivo;

            archivo = saveFileDialog1.FileName;

            MessageBox.Show(archivo);
            StreamWriter writer = new StreamWriter(archivo);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                writer.WriteLine("INSERT INTO alumnos (matricula,ap1,ap2,nombre,fecha_nacimiento,correo,telefono) VALUES (" +
                dataGridView1[0, i].Value.ToString() + ",'" +
                dataGridView1[1, i].Value.ToString() + "','" +
                dataGridView1[2, i].Value.ToString() + "','" +
                dataGridView1[3, i].Value.ToString() + "','" +
                /* dateTimePicker1.Value.Year + "-" +
                 dateTimePicker1.Value.Month + "-" +
                 dateTimePicker1.Value.Day + "','" +*/
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Year + "-" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Month + "-" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Day + "-" +
                dataGridView1[5, i].Value.ToString() + "'," +
                dataGridView1[6, i].Value.ToString() + ");");
            }
            writer.Close();

        }

        private void button_CCV_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "Guardar csv";
            saveFileDialog1.FileName = ".csv";
            saveFileDialog1.InitialDirectory = @"C:\Users\kno_0\Google Drive\UNIKINO\Base de Datos 2\06-04-19";
            saveFileDialog1.Filter = "archivo csv |*.csv";

            saveFileDialog1.ShowDialog();

            String archivo;

            archivo = saveFileDialog1.FileName;

            MessageBox.Show(archivo);
            StreamWriter writer = new StreamWriter(archivo);
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                if (i == dataGridView1.ColumnCount)
                {
                    writer.Write((dataGridView1.Columns[i].HeaderText));
                }
                else
                {
                    writer.Write((dataGridView1.Columns[i].HeaderText) + ";");
                }
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                writer.WriteLine(
                dataGridView1[0, i].Value.ToString() + ";" +
                dataGridView1[1, i].Value.ToString() + ";" +
                dataGridView1[2, i].Value.ToString() + ";" +
                dataGridView1[3, i].Value.ToString() + ";" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Year + ";" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Month + ";" +
                Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Day + ";" +
                dataGridView1[5, i].Value.ToString() + ";" +
                dataGridView1[6, i].Value.ToString() + ")");
            }
            writer.Close();
        }

        private void button_PDFF_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "Guardar JSON";
            saveFileDialog1.FileName = "*.json";
            saveFileDialog1.InitialDirectory = @"C:\Users\kno_0\Google Drive\UNIKINO\Base de Datos 2\06-04-19";
            saveFileDialog1.Filter = "archivo json |*.json";

            saveFileDialog1.ShowDialog();

            String archivo;

            archivo = saveFileDialog1.FileName;

            MessageBox.Show(archivo);
            StreamWriter writer = new StreamWriter(archivo);
            writer.WriteLine("{\"sistema_escolar\" :");
            writer.WriteLine("\t{");
            writer.WriteLine("\t\t \"alumnos\" :[");
            writer.WriteLine("\t\t {");
            writer.WriteLine("\t{");
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                if (i == dataGridView1.ColumnCount)
                {
                    writer.WriteLine("\t\t{");
                    writer.WriteLine(("\t\t\"Matricula\": ") + dataGridView1[0, i].Value.ToString() + ",");
                    writer.WriteLine(("\t\t\"Apellido Paterno\": ") + dataGridView1[1, i].Value.ToString() + ",");
                    writer.WriteLine(("\t\t\"Apellido Materno\": ") + dataGridView1[2, i].Value.ToString() + ",");
                    writer.WriteLine(("\t\t\"Nombre\": ") + dataGridView1[3, i].Value.ToString() + ",");
                    writer.WriteLine(("\t\t\"Fecha\": ") +
                    Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Year + "-" +
                    Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Month + "-" +
                    Convert.ToDateTime(dataGridView1[4, i].Value.ToString()).Day + ",");
                    writer.WriteLine(("\t\t\"Correo\": ") + dataGridView1[5, i].Value.ToString() + ",");
                    writer.WriteLine(("\t\t\"Telefono\": ") + dataGridView1[6, i].Value.ToString() + ",");
                    writer.WriteLine("\t\t},");
                }
            }
        }

        private void button_exel_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "Guardar xlsx";
            saveFileDialog1.FileName = ".xlsx";
            saveFileDialog1.InitialDirectory = @"C:\Users\kno_0\Google Drive\UNIKINO\Base de Datos 2\06-04-19";
            saveFileDialog1.Filter = "archivo xlsx |*.xlsx";

            saveFileDialog1.ShowDialog();

            String archivo;

            archivo = saveFileDialog1.FileName;

            MessageBox.Show(archivo);
            var workbook = new XLWorkbook();
            var hoja = workbook.Worksheets.Add("Hoja1");
            hoja.Cell(1, 1).Value = "Matricula";
            hoja.Cell(1, 1).Style.Font.Bold=true;
            hoja.Cell(1, 2).Value = "Apellido P";
            hoja.Cell(1, 2).Style.Font.Bold = true;
            hoja.Cell(1, 3).Value = "Apellido M";
            hoja.Cell(1, 3).Style.Font.Bold = true;
            hoja.Cell(1, 4).Value = "Nombre";
            hoja.Cell(1, 4).Style.Font.Bold = true;
            hoja.Cell(1, 5).Value = "Fecha de Nacimiento";
            hoja.Cell(1, 5).Style.Font.Bold = true;
            hoja.Cell(1, 6).Value = "Correo";
            hoja.Cell(1, 6).Style.Font.Bold = true;
            hoja.Cell(1, 7).Value = "Telefono";
            hoja.Cell(1, 7).Style.Font.Bold = true;
            for (int i = 0; i < dataGridView1; i++)
            {
                for (int k = 0; k < dataGridView1.ColumnCount; k++)
                {
                    hoja.Cell((i + 2), (k + 1)).Value = dataGridView1.Rows[i].Cells[k].Value.ToString();
                }
            }
            workbook.SaveAs(archivo);

           
        }
    }
    }
            
          
        
    

