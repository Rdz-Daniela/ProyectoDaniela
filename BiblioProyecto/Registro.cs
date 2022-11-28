using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

namespace BibliotecaFime
{
    public partial class Registro : Form
    {
        SqlConnection conn = new SqlConnection(@"Data Source=LAPTOP-NRJGQ1AE\SQLEXPRESS;Initial Catalog=Biblioteca;Integrated Security=True;");
        public Registro()
        {
            InitializeComponent();
           
        }

        private void Registro_Load(object sender, EventArgs e)
        {
            Alumnos();
       }
        private void Alumnos()
        {
            DsConexionTableAdapters.alumnosTableAdapter ta = new DsConexionTableAdapters.alumnosTableAdapter();
            DsConexion.alumnosDataTable dt = ta.GetData();
            dataGridView1.DataSource = dt;
        }

        private void btnRegresar_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.Show();
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
          
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.alumnosTableAdapter ta = new DsConexionTableAdapters.alumnosTableAdapter();
            ta.UpdateAlumnos(txtNombre.Text, cmbCarrera.Text, int.Parse(txtSemestre.Text), cmbTurno.Text, txtTelefono.Text, txtMatricula.Text, txtMatricula.Text);
            Alumnos();
            Limpiar();
        }

        private void btnRegistro_Click(object sender, EventArgs e)
        {
           
        }

        private void btnbusca_Click(object sender, EventArgs e)
        {
            
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            conn.Open();
            try { 
            string consulta = "Select * From alumnos where " + cmbBuscar.Text + " like '%" + txtBusqueda.Text + "%'";
            SqlDataAdapter adap = new SqlDataAdapter(consulta, conn);
            DataTable dt = new DataTable();
            adap.Fill(dt);
            dataGridView1.DataSource = dt;
            SqlCommand cmd = new SqlCommand(consulta, conn);
            SqlDataReader lector;
            lector = cmd.ExecuteReader();
            }
            catch
            {
                MessageBox.Show("Favor de Completar los Datos!!");
            }
            conn.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMatricula.Text = dataGridView1.SelectedCells[0].Value.ToString();
            txtNombre.Text= dataGridView1.SelectedCells[1].Value.ToString();
            cmbCarrera.Text= dataGridView1.SelectedCells[2].Value.ToString();
            txtSemestre.Text= dataGridView1.SelectedCells[3].Value.ToString();
            cmbTurno.Text= dataGridView1.SelectedCells[4].Value.ToString();
            txtTelefono.Text= dataGridView1.SelectedCells[5].Value.ToString();
            this.btnModificar.Visible = true;
            this.btnAgregar.Visible = false;
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            Limpiar();
        }
        private void Limpiar()
        {
            txtMatricula.Text = "";
            txtNombre.Text = "";
            txtSemestre.Text = "";
            txtTelefono.Text = "";
            cmbCarrera.Text = "";
            cmbTurno.Text = "";
            this.btnAgregar.Visible = true;
            this.btnModificar.Visible = false;
        }

        private void btnAgregar_Click_1(object sender, EventArgs e)
        {
          
          
               try { 
            DsConexionTableAdapters.alumnosTableAdapter ta = new DsConexionTableAdapters.alumnosTableAdapter();
            ta.agregaralumnos(txtMatricula.Text, txtNombre.Text, cmbCarrera.Text, int.Parse(txtSemestre.Text), cmbTurno.Text, txtTelefono.Text);
            MessageBox.Show("Datos Agregados Correctamente!");
            Alumnos();
            Limpiar();

            }
            catch
            {
                MessageBox.Show("Favor de Completar los Datos!");
            }
        }
      

        private void btnTodos_Click(object sender, EventArgs e)
        {
            txtBusqueda.Text = "";
            cmbBuscar.Text = "";
            Alumnos();
        }

        private void txtNombre_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtNombre_Enter(object sender, EventArgs e)
        {
            MessageBox.Show("Escribir sin ACENTOS!!");
        }

        private void btnBorrar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.alumnosTableAdapter ta = new DsConexionTableAdapters.alumnosTableAdapter();
            ta.deletealumnos(txtMatricula.Text);
            Alumnos();
            Limpiar();
        }

        private void btnPdf_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF (*.pdf)|*.pdf";
                save.FileName = "Alumnos.pdf";
                bool ErrorMessage = false;
                if (save.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(save.FileName))
                    {
                        try
                        {
                            File.Delete(save.FileName);
                        }
                        catch (Exception ex)
                        {
                            ErrorMessage = true;
                            MessageBox.Show("No se puede escribir la lista" + ex.Message);
                        }
                    }
                    if (!ErrorMessage)
                    {
                        try
                        {
                            PdfPTable pTable = new PdfPTable(dataGridView1.Columns.Count);
                            pTable.DefaultCell.Padding = 2;
                            pTable.WidthPercentage = 100;
                            pTable.HorizontalAlignment = Element.ALIGN_LEFT;

                            foreach (DataGridViewColumn col in dataGridView1.Columns)
                            {
                                PdfPCell pCell = new PdfPCell(new Phrase(col.HeaderText));
                                pTable.AddCell(pCell);
                            }
                            foreach (DataGridViewRow viewRow in dataGridView1.Rows)
                            {
                                foreach (DataGridViewCell dcell in viewRow.Cells)
                                {
                                    pTable.AddCell(dcell.Value.ToString());
                                }
                            }

                            using (FileStream fileStream = new FileStream(save.FileName, FileMode.Create))
                            {
                                Document document = new Document(PageSize.A4, 8f, 16f, 16f, 8f);
                                document.Open();
                                document.Add(pTable);
                                document.Close();
                                fileStream.Close();
                            }
                            MessageBox.Show("Exportacion Exitosa!", "info");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error en la Exportacion" + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No se Exporto", "info");
            }
        }

        private void btnXlsx_Click(object sender, EventArgs e)
        {
            ExportarDatos(dataGridView1);
        }
        public void ExportarDatos(DataGridView datalistado)
        {
            Microsoft.Office.Interop.Excel.Application exportarexcel = new Microsoft.Office.Interop.Excel.Application();

            exportarexcel.Application.Workbooks.Add(true);

            int indicecolumn = 0;
            foreach (DataGridViewColumn columna in dataGridView1.Columns)
            {
                indicecolumn++;

                exportarexcel.Cells[1, indicecolumn] = columna.Name;
            }
            int indicefila = 0;
            foreach (DataGridViewRow fila in dataGridView1.Rows)
            {
                indicefila++;
                indicecolumn = 0;

                foreach (DataGridViewColumn columna in dataGridView1.Columns)
                {
                    indicecolumn++;
                    exportarexcel.Cells[indicefila + 1, indicecolumn] = fila.Cells[columna.Name].Value;
                }
            }
            exportarexcel.Visible = true;
        }

        private void btnBloc_Click(object sender, EventArgs e)
        {
            GuardarNotas();
        }
        private void GuardarNotas()
        {
            StreamWriter bloc = new StreamWriter("Alumnos.txt", true);
            bloc.WriteLine(txtMatricula.Text);
            bloc.WriteLine(txtNombre.Text);
            bloc.WriteLine(cmbCarrera.Text);
            bloc.WriteLine(txtSemestre.Text);
            bloc.WriteLine(cmbTurno.Text);
            bloc.WriteLine(txtTelefono.Text);
            bloc.Close();
        }
    }
    
}
