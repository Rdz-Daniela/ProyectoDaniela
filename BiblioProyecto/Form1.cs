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


namespace BibliotecaFime
{
    public partial class Form1 : Form
    {
        SqlConnection conn = new SqlConnection(@"Data Source=LAPTOP-NRJGQ1AE\SQLEXPRESS;Initial Catalog=Biblioteca;Integrated Security=True;");
        SqlCommand cmd = new SqlCommand();
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void txtUsuario_TextChanged(object sender, EventArgs e)
        {
        }
       

        private void txtContraseña_TextChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlDataAdapter data = new SqlDataAdapter("Select * from Usuario", conn);

            try
            {
                conn.Open();
                cmd = new SqlCommand(("select usuario , contraseña from usuario where usuario = @pusuario and contraseña = @pcontraseña"), conn);
                cmd.Parameters.AddWithValue("@pusuario", txtUsuario.Text);
                cmd.Parameters.AddWithValue("@pcontraseña", txtContraseña.Text);
                SqlDataReader readcomand = cmd.ExecuteReader();
                if (readcomand.Read())
                {
                    this.Hide();
                    Menu menu = new Menu();
                    menu.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error", ex.Message);
                conn.Close();
            }
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void txtUsuario_Leave_1(object sender, EventArgs e)
        {
            if (txtUsuario.Text == "")
            {
                txtUsuario.Text = "Usuario";
                txtUsuario.ForeColor = Color.DimGray;
            }
        }

        private void txtContraseña_Leave_1(object sender, EventArgs e)
        {
            if (txtContraseña.Text == "")
            {
                txtContraseña.Text = "Contraseña";
                txtContraseña.ForeColor = Color.DimGray;
                txtContraseña.UseSystemPasswordChar = false;
            }
        }

        private void txtUsuario_Enter(object sender, EventArgs e)
        {
            if (txtUsuario.Text == "Usuario")
            {
                txtUsuario.Text = "";
                txtUsuario.ForeColor = Color.Black;
            }
        }

        private void txtContraseña_Enter(object sender, EventArgs e)
        {
            if (txtContraseña.Text == "Contraseña")
            {
                txtContraseña.Text = "";
                txtContraseña.ForeColor = Color.Black;
                txtContraseña.UseSystemPasswordChar = true;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
    }
    

