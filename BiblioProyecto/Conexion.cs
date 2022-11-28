using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace BibliotecaFime
{
    internal class Conexion
    {
        static SqlConnection cnx;
        static string cadena = @"Server=LAPTOP-NRJGQ1AE\SQLEXPRESS;Database=Biblioteca;Trusted_Connection=True;";

        public static void Conectar()
        {
            cnx = new SqlConnection(cadena);
            cnx.Open();
        }
        public static void Desconectar()
        {
            cnx.Close();
            cnx = null;
        }

        internal static DataTable EjecutaSeleccion(string consulta)
        {
            throw new NotImplementedException();
        }

        public  int EjecutaConsulta(string consulta)
        {
            int filasAfectadas = 0;
            Conectar();
            SqlCommand cmd = new SqlCommand(consulta, cnx);
            filasAfectadas = cmd.ExecuteNonQuery();
            Desconectar();
            return filasAfectadas;
        }
        public static DataTable ConsultarLibros() {
            string query = "Select * from libros";
            SqlCommand cmd = new SqlCommand(query, cnx);
            SqlDataAdapter data = new SqlDataAdapter(cmd);
            DataTable tabla = new DataTable();
            data.Fill(tabla);

            return tabla;
        }
    }
}
