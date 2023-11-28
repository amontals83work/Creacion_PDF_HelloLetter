using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Creacion_PDF_HelloLetter
{
    public class MCCommand
    {
        public SqlCommand command = new SqlCommand();
        public SqlConnection connection = new SqlConnection();

        public string CommandText
        {
            get { return command.CommandText; }
            set { command.CommandText = value; }
        }

        private IDbConnection Connection
        {
            get { return command.Connection; }
            set { connection = command.Connection; }
        }

        public int ExecuteNonQuery()
        {
            return command.ExecuteNonQuery();
        }

        public object ExecuteScalar()
        {
            if (command.Connection == null)
            {
                throw new InvalidOperationException("La conexión no se ha inicializado.");
            }
            return command.ExecuteScalar();
        }

        public IDataReader ExecuteReader()
        {
            return command.ExecuteReader();
        }

        public IDataReader ExecuteReader(CommandBehavior behavior)
        {
            return command.ExecuteReader(behavior);
        }

        public void Cancel()
        {
            command.Cancel();
        }

        public int CommandTimeout
        {
            get { return command.CommandTimeout; }
            set { command.CommandTimeout = value; }
        }

        public CommandType CommandType
        {
            get { return command.CommandType; }
            set { command.CommandType = value; }
        }

        public IDbDataParameter CreateParameter()
        {
            return command.CreateParameter();
        }

        public IDataParameterCollection Parameters
        {
            get { return command.Parameters; }
        }

        public void Prepare()
        {
            command.Prepare();
        }

        public IDbTransaction Transaction
        {
            get { return command.Transaction; }
            set { command.Transaction = (SqlTransaction)value; }
        }

        public UpdateRowSource UpdatedRowSource
        {
            get { return command.UpdatedRowSource; }
            set { command.UpdatedRowSource = value; }
        }

    }
}
