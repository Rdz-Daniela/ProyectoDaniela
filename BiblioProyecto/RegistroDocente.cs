using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

namespace BibliotecaFime
{
    public partial class RegistroDocente : Form
    {
        SqlConnection conn = new SqlConnection(@"Data Source=LAPTOP-NRJGQ1AE\SQLEXPRESS;Initial Catalog=Biblioteca;Integrated Security=True;");
        public RegistroDocente()
        {
            InitializeComponent();
        }

        private void btnRegresar_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.Show();
        }
        private void docentes()
        {
            DsConexionTableAdapters.docentesTableAdapter ta = new DsConexionTableAdapters.docentesTableAdapter();
            DsConexion.docentesDataTable dt = ta.GetData();
            dataGridView1.DataSource = dt;
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.docentesTableAdapter ta = new DsConexionTableAdapters.docentesTableAdapter();
            ta.updatedocente(txtNombre.Text, txtTelefono.Text, txtId.Text, txtId.Text);
            docentes();
            Limpiar();
        }

        private void RegistroDocente_Load(object sender, EventArgs e)
        {
            docentes();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtId.Text =  dataGridView1.SelectedCells[0].Value.ToString();
            txtNombre.Text = dataGridView1.SelectedCells[1].Value.ToString();
            txtTelefono.Text = dataGridView1.SelectedCells[2].Value.ToString();
            this.btnModificar.Visible = true;
            this.btnAgregar.Visible = false;
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            conn.Open();
            try
            {            
            string consulta = "Select * From docentes where " + cmbBuscar.Text + " like '%" + txtBusqueda.Text + "%'"; 
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
                MessageBox.Show("Favor de Completar los Datos!");
            }
            conn.Close();
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            Limpiar();
        }
        private void Limpiar()
        {
            txtId.Text = "";
            txtNombre.Text = "";
            txtTelefono.Text = "";
            this.btnAgregar.Visible = true;
            this.btnModificar.Visible = false;
        }

        private void btnAgregar_Click_1(object sender, EventArgs e)
        {
           
                try
                {
                    DsConexionTableAdapters.docentesTableAdapter ta = new DsConexionTableAdapters.docentesTableAdapter();
                    ta.insertdocente(txtId.Text, txtNombre.Text, txtTelefono.Text);
                    MessageBox.Show("Datos Agregados Correctamente!");
                    docentes();
                    Limpiar();
                }
                catch
                {
                    MessageBox.Show("Favor de Completar los Datos!!");
                }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtBusqueda.Text = "";
            cmbBuscar.Text = "";
            docentes(); 
        }

        private void txtNombre_Enter(object sender, EventArgs e)
        {
            MessageBox.Show("Escribir sin ACENTOS!!");
        }

        private void btnBorrar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.docentesTableAdapter ta = new DsConexionTableAdapters.docentesTableAdapter();
            ta.deletedocentes(txtId.Text);
            docentes();
            Limpiar();
        }

        private void btnPdf_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF (*.pdf)|*.pdf";
                save.FileName = "Docentes.pdf";
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
                            MessageBox.Show("Exportacion Exitosa!", "informacion");
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
                MessageBox.Show("No se Exporto", "informacion");
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
            StreamWriter bloc = new StreamWriter("Docentes.txt", true);
            bloc.WriteLine(txtId.Text);
            bloc.WriteLine(txtNombre.Text);
            bloc.WriteLine(txtTelefono.Text);
            bloc.Close();
        }
    }
}
