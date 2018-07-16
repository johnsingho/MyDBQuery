using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;

namespace MyDBQuery.common
{
    //MySql connect helper
    //H.Z.XIN 2018-07-16
    public class SqliteHelper
    {
        public static int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params SQLiteParameter[] commandParameters)
        {
            var cmd = new SQLiteCommand();

            //Create a connection
            using (var connection = new SQLiteConnection(connectionString))
            {
                //Prepare the command
                PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
                //Execute the command
                int val = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                return val;
            }
        }

        public static int ExecuteNonQuery(SQLiteTransaction trans, CommandType cmdType, string cmdText, params SQLiteParameter[] commandParameters)
        {
            var cmd = new SQLiteCommand();
            PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        public static int ExecuteNonQuery(SQLiteConnection connection, CommandType cmdType, string cmdText, params SQLiteParameter[] commandParameters)
        {
            var cmd = new SQLiteCommand();

            PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        public static SQLiteDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params SQLiteParameter[] commandParameters)
        {
            var cmd = new SQLiteCommand();
            var conn = new SQLiteConnection(connectionString);

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

        public static SQLiteDataReader ExecuteReader(SQLiteConnection conn, CommandType cmdType, string cmdText, params SQLiteParameter[] commandParameters)
        {
            //Create the command and connection
            var cmd = new SQLiteCommand();
            //Prepare the command to execute
            PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);

            //Execute the query, stating that the connection should close when the resulting datareader has been read
            var rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            cmd.Parameters.Clear();
            return rdr;
        }

        public static object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params SQLiteParameter[] commandParameters)
        {
            var cmd = new SQLiteCommand();

            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                object val = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return val;
            }
        }

        public static object ExecuteScalar(SQLiteTransaction transaction, CommandType commandType, string commandText, params SQLiteParameter[] commandParameters)
        {
            if (transaction == null)
                throw new ArgumentNullException("transaction");
            if (transaction != null && transaction.Connection == null)
                throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");

            // Create a command and prepare it for execution
            var cmd = new SQLiteCommand();
            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters);

            // Execute the command & return the results
            object retval = cmd.ExecuteScalar();
            // Detach the SqlParameters from the command object, so they can be used again
            cmd.Parameters.Clear();
            return retval;
        }

        public static object ExecuteScalar(SQLiteConnection connectionString, CommandType cmdType, string cmdText, params SQLiteParameter[] commandParameters)
        {
            var cmd = new SQLiteCommand();
            PrepareCommand(cmd, connectionString, null, cmdType, cmdText, commandParameters);
            object val = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return val;
        }

        #region 内部
        private static void PrepareCommand(SQLiteCommand cmd, SQLiteConnection conn, SQLiteTransaction trans, CommandType cmdType, string cmdText, SQLiteParameter[] commandParameters)
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
                foreach (SQLiteParameter parm in commandParameters)
                {
                    cmd.Parameters.Add(parm);
                }
            }
        }


        public static ulong BulkToDB(string connStr, DataTable dt, string sTarTable, out string sErr)
        {
            sErr = string.Empty;
            ulong nRec = 0;

            var colNames = DataTableHelper.GetDataTableColNames(dt);
            if (0 == colNames.Count)
            {
                return 0;
            }
            var sCols = string.Join(",", colNames);
            var arrAtCols = new List<string>();
            colNames.ForEach((s) =>
            {
                var sColN = string.Format("@{0}", s);
                arrAtCols.Add(sColN);
            });
            var sAtCols = string.Join(",", arrAtCols);

            string sql = string.Format("@INSERT INTO [{0}] ({1}) VALUES({1});", sCols, sTarTable, sAtCols);
            try
            {
                using (var cn = new SQLiteConnection(connStr))
                {
                    cn.Open();
                    using (var transaction = cn.BeginTransaction())
                    {
                        using (var cmd = cn.CreateCommand())
                        {
                            cmd.CommandText = sql;
                            arrAtCols.ForEach((s) =>
                            {
                                cmd.Parameters.AddWithValue(s, null);
                            });

                            foreach (DataRow dr in dt.Rows)
                            {
                                arrAtCols.ForEach((s) =>
                                {
                                    cmd.Parameters[s].Value = dr[s];
                                });
                            }

                            cmd.ExecuteNonQuery();
                            nRec++;
                        }
                        transaction.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
                nRec = 0;
            }
            return nRec;
        }
        #endregion
    }
}
