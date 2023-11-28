using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creacion_PDF_HelloLetter
{
    public class ConexionDB
    {
        public static string Connection = "Data Source=192.168.50.48;Initial Catalog=DespachoMc;Persist Security Info=True;User ID=sa;Password=Binabiq2018_;MultipleActiveResultSets=True";
        public static SqlConnection Connect = new SqlConnection(Connection);

        public static void AbrirConexion()
        {
            if (Connect.State == ConnectionState.Closed) { Connect.Open(); }
        }

        public static void CerrarConexion()
        {
            if (Connect.State == ConnectionState.Open) { Connect.Close(); }
        }

        public SqlConnection ObtenerConexion() { return Connect; }

        public void Insert(String query)
        {
            SqlCommand command = new SqlCommand(query, Connect);
            command.ExecuteNonQuery();
        }

        public DataTable Select(String query)
        {
            SqlCommand command = new SqlCommand(query, Connect);
            SqlDataAdapter adapter = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dt.Rows.InsertAt(dt.NewRow(), 0);
            return dt;
        }

        public void Update(String query)
        {
            SqlCommand command = new SqlCommand(query, Connect);
            command.ExecuteNonQuery();
        }

        public void Delete(String query)
        {
            SqlCommand command = new SqlCommand(query, Connect);
            command.ExecuteNonQuery();
        }
    }
}
