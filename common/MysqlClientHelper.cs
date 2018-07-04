using System;
using MySql.Data.MySqlClient;
using System.Data;

namespace MyDBQuery.common
{
    //MySql connect helper
    //H.Z.XIN 2017-09-07
    public class MySqlClientHelper
    {
        public static int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params MySqlParameter[] commandParameters)
        {
            var cmd = new MySqlCommand();

            //Create a connection
            using (var connection = new MySqlConnection(connectionString))
            {
                //Prepare the command
                PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
                //Execute the command
                int val = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                return val;
            }
        }

        public static int ExecuteNonQuery(MySqlTransaction trans, CommandType cmdType, string cmdText, params MySqlParameter[] commandParameters)
        {
            var cmd = new MySqlCommand();
            PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        public static int ExecuteNonQuery(MySqlConnection connection, CommandType cmdType, string cmdText, params MySqlParameter[] commandParameters)
        {
            var cmd = new MySqlCommand();

            PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        public static MySqlDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params MySqlParameter[] commandParameters)
        {
            var cmd = new MySqlCommand();
            var conn = new MySqlConnection(connectionString);

            try
            {
                //Prepare the command to execute
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);

                //Execute the query, stating that the connection should close when the resulting datareader has been read
                var rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return rdr;

            }
            catch (Exception ex)
            {

                //If an error occurs close the connection as the reader will not be used and we expect it to close the connection
                conn.Close();
                throw;
            }
        }

        public static MySqlDataReader ExecuteReader(MySqlConnection conn, CommandType cmdType, string cmdText, params MySqlParameter[] commandParameters)
        {
            //Create the command and connection
            var cmd = new MySqlCommand();
            //Prepare the command to execute
            PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);

            //Execute the query, stating that the connection should close when the resulting datareader has been read
            var rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            cmd.Parameters.Clear();
            return rdr;
        }

        public static object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params MySqlParameter[] commandParameters)
        {
            var cmd = new MySqlCommand();

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                object val = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return val;
            }
        }

        public static object ExecuteScalar(MySqlTransaction transaction, CommandType commandType, string commandText, params MySqlParameter[] commandParameters)
        {
            if (transaction == null)
                throw new ArgumentNullException("transaction");
            if (transaction != null && transaction.Connection == null)
                throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");

            // Create a command and prepare it for execution
            var cmd = new MySqlCommand();
            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters);

            // Execute the command & return the results
            object retval = cmd.ExecuteScalar();
            // Detach the SqlParameters from the command object, so they can be used again
            cmd.Parameters.Clear();
            return retval;
        }

        public static object ExecuteScalar(MySqlConnection connectionString, CommandType cmdType, string cmdText, params MySqlParameter[] commandParameters)
        {
            var cmd = new MySqlCommand();
            PrepareCommand(cmd, connectionString, null, cmdType, cmdText, commandParameters);
            object val = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return val;
        }

        #region 内部
        private static void PrepareCommand(MySqlCommand cmd, MySqlConnection conn, MySqlTransaction trans, CommandType cmdType, string cmdText, MySqlParameter[] commandParameters)
        {
            //Open the connection if required
            if (conn.State != ConnectionState.Open)
                conn.Open();

            //Set up the command
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            cmd.CommandType = cmdType;

            //Bind it to the transaction if it exists
            if (trans != null)
                cmd.Transaction = trans;

            // Bind the parameters passed in
            if (commandParameters != null)
            {
                foreach (MySqlParameter parm in commandParameters)
                {
                    cmd.Parameters.Add(parm);
                }
            }
        }
        #endregion
    }
}
