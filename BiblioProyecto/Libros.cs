using System;
using System.Data;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

namespace BibliotecaFime
{
    public partial class Libros : Form
    {
        SqlConnection conn = new SqlConnection(@"Data Source=LAPTOP-NRJGQ1AE\SQLEXPRESS;Initial Catalog=Biblioteca;Integrated Security=True;");
        public Libros()
        {
            InitializeComponent();
            
        }

        private void Libros_Load(object sender, EventArgs e)
        {

            tablalibros();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }
        private void tablalibros()
        {
            DsConexionTableAdapters.librosTableAdapter ta = new DsConexionTableAdapters.librosTableAdapter();
            DsConexion.librosDataTable dt = ta.GetData();
            dataGridView1.DataSource = dt;
        }
       
        private void btnModificar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.librosTableAdapter ta = new DsConexionTableAdapters.librosTableAdapter();
            ta.updatelibros(txtClave.Text, txtTitulo.Text, txtAutor.Text, int.Parse(txtCantidad.Text), cmbClasificacion.Text, txtId.Text, txtId.Text);
            tablalibros();
            limpiar();
         
        }
        private void btnAgregar_Click(object sender, EventArgs e)
        {
            
            try { 
            DsConexionTableAdapters.librosTableAdapter ta = new DsConexionTableAdapters.librosTableAdapter();
            ta.AgregarLibros(txtId.Text, txtClave.Text, txtTitulo.Text, txtAutor.Text, int.Parse(txtCantidad.Text), cmbClasificacion.Text);
                    MessageBox.Show("Datos Agregados Correctamente!");
                    limpiar();
                    tablalibros();
            } catch
            {
                    MessageBox.Show("Favor de Completar los Datos!");
            }
            

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

       

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
           
                txtId.Text = dataGridView1.SelectedCells[0].Value.ToString();
                txtClave.Text = dataGridView1.SelectedCells[1].Value.ToString();
                txtTitulo.Text = dataGridView1.SelectedCells[2].Value.ToString();
                txtAutor.Text= dataGridView1.SelectedCells[3].Value.ToString();
                txtCantidad.Text = dataGridView1.SelectedCells[4].Value.ToString();
                cmbClasificacion.Text = dataGridView1.SelectedCells[5].Value.ToString();
                this.btnModificar.Visible = true;
                this.btnAgregar.Visible = false;
           
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            conn.Open();
            try { 
            string consulta = "Select * From libros where "+cmbBuscar.Text+" like '%" + txtBusca.Text + "%'";
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

        private void btnRegistro_Click(object sender, EventArgs e)
        {
           
        }

        private void btnbusca_Click(object sender, EventArgs e)
        {
           
        }
        private void limpiar()
        {
            txtId.Text = "";
            txtClave.Text = "";
            txtAutor.Text = "";
            txtCantidad.Text = "";
            txtTitulo.Text = "";
            cmbClasificacion.Text = "";
            this.btnAgregar.Visible = true;
            this.btnModificar.Visible = false;
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            limpiar();
        }

        private void txtBusca_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtBusca.Text = "";
            cmbBuscar.Text = "";
            tablalibros();
        }

        private void txtTitulo_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void txtAutor_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void txtTitulo_Enter(object sender, EventArgs e)
        {
            MessageBox.Show("Escribir sin ACENTOS!!");
        }

        private void txtAutor_Enter(object sender, EventArgs e)
        {
            MessageBox.Show("Escribir sin ACENTOS!!");
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void btnBorrar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.librosTableAdapter ta = new DsConexionTableAdapters.librosTableAdapter();
            ta.deletelibros(txtId.Text);
            tablalibros();
            limpiar();
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
            StreamWriter bloc = new StreamWriter("libros.txt", true);
            bloc.WriteLine(txtId.Text);
            bloc.WriteLine(txtClave.Text);
            bloc.WriteLine(txtTitulo.Text);
            bloc.WriteLine(txtAutor.Text);
            bloc.WriteLine(cmbClasificacion.Text);
            bloc.WriteLine(txtCantidad.Text);
            bloc.Close();
        }

        private void btnPdf_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF (*.pdf)|*.pdf";
                save.FileName = "Libros.pdf";
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        
    }
}
