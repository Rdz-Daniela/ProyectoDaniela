using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BibliotecaFime
{
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
        }

        private void btnLibros_Click(object sender, EventArgs e)
        {
            formhijo(new Libros());
        }

        private void btnRegistro_Click(object sender, EventArgs e)
        {
            formhijo(new Registro());
            
        }

        private void btnDocentes_Click(object sender, EventArgs e)
        {
            formhijo(new RegistroDocente());
        }

        private void btnPrestamo_Click(object sender, EventArgs e)
        {
            formhijo(new Prestamos());
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void formhijo(object form)
        {
            if (this.panelsubmenu.Controls.Count > 0)
                this.panelsubmenu.Controls.RemoveAt(0);
            Form fh = form as Form;
            fh.TopLevel = false;
            fh.Dock= DockStyle.Fill;
            fh.Show();
            this.panelsubmenu.Controls.Add(fh);
            this.panelsubmenu.Tag = fh;
            
        }

        private void btnSesion_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 form = new Form1();
            form.Show();
        }
    }
}
