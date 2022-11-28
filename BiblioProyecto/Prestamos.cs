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
    public partial class Prestamos : Form
    {
        SqlConnection conn = new SqlConnection(@"Data Source=LAPTOP-NRJGQ1AE\SQLEXPRESS;Initial Catalog=Biblioteca;Integrated Security=True;");
        public Prestamos()
        {
            InitializeComponent();
        }
 
        private void tablaPrestamos()
        {
            DsConexionTableAdapters.PrestamosTableAdapter ta = new DsConexionTableAdapters.PrestamosTableAdapter();
            DsConexion.PrestamosDataTable dt = ta.GetData();
            dataGridView1.DataSource = dt;
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {           
                try { 
             DsConexionTableAdapters.PrestamosTableAdapter ta = new DsConexionTableAdapters.PrestamosTableAdapter();
            ta.aPrestamos(txtId.Text, txtMatricula.Text, txtFolio.Text, dateSalida.Value, dateEntrega.Value);
                    MessageBox.Show("Datos Agregados Correctamente!");
                    Limpiar();
                    tablaPrestamos();
            }
                catch                 
                {
                   
                    MessageBox.Show("Favor de Completar los Datos!!");
                
            }
               
            
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.PrestamosTableAdapter ta = new DsConexionTableAdapters.PrestamosTableAdapter();
            ta.moPrestamos(txtId.Text, txtMatricula.Text, txtFolio.Text, dateSalida.Value, dateEntrega.Value,int.Parse(txtPrestamo.Text));
            tablaPrestamos();
            Limpiar();
        }

        private void Prestamos_Load(object sender, EventArgs e)
        {
            tablaPrestamos();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtPrestamo.Text = dataGridView1.SelectedCells[0].Value.ToString();
            txtId.Text=dataGridView1.SelectedCells[1].Value.ToString();
            txtMatricula.Text= dataGridView1.SelectedCells[2].Value.ToString();
            txtFolio.Text= dataGridView1.SelectedCells[3].Value.ToString();
            this.btnAgregar.Visible = false;
            this.btnModificar.Visible = true;
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            conn.Open();
            try
            {
                string consulta = "Select * From Prestamos where " + cmbBuscar.Text + " like '%" + txtBusqueda.Text + "%'";
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
        private void Limpiar()
        {
            txtFolio.Text = "";
            txtMatricula.Text = "";
            txtId.Text = "";
            txtPrestamo.Text = "";
            this.btnAgregar.Visible = true;
            this.btnModificar.Visible = false;
            this.label2.Visible = false;
            this.txtId.Visible = false;
            this.label3.Visible = false;
            this.txtMatricula.Visible = false;
            this.radioButton1.Checked = false;
            this.radioButton2.Checked = false;
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            Limpiar();
        }

        private void btnTodos_Click(object sender, EventArgs e)
        {
            txtBusqueda.Text = "";
            cmbBuscar.Text = "";
            tablaPrestamos();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                this.label2.Visible = true;
                this.txtId.Visible = true;
                this.label3.Visible =false;
                this.txtMatricula.Visible = false;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                this.label3.Visible = true;
                this.txtMatricula.Visible = true;
                this.label2.Visible = false;
                this.txtId.Visible = false;
            }
        }

        private void btnAdocente_Click(object sender, EventArgs e)
        {
            

        }

        private void btnAalumno_Click(object sender, EventArgs e)
        {
            
        }

        private void txtMatricula_Leave(object sender, EventArgs e)
        {
            
        }

        private void btnBorrar_Click(object sender, EventArgs e)
        {
            DsConexionTableAdapters.PrestamosTableAdapter ta = new DsConexionTableAdapters.PrestamosTableAdapter();
            ta.deleteprestamos(int.Parse(txtPrestamo.Text));
            tablaPrestamos();
            Limpiar();
        }

        private void btnPdf_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF (*.pdf)|*.pdf";
                save.FileName = "Prestamos.pdf";
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
            StreamWriter bloc = new StreamWriter("Prestamos.txt", true);
            bloc.WriteLine(txtPrestamo.Text);
            bloc.WriteLine(txtId.Text);
            bloc.WriteLine(txtMatricula.Text);
            bloc.WriteLine(txtFolio.Text);
            bloc.WriteLine(dateSalida.Text);
            bloc.WriteLine(dateEntrega.Text);
            bloc.Close();
        }
    }
}
