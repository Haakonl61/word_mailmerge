using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

using word_mailmerge;

namespace Exports
{
    //https://docs.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqlcommand?redirectedfrom=MSDN&view=dotnet-plat-ext-3.1
    static class SqlHelper
    {
        // Set the connection, command, and then execute the command with non query.  
        public static SqlConnection GetConnection(String connectionString)
        {
            var conn = new SqlConnection(connectionString);
            conn.Open();
            return conn;

        }

        public static Int32 ExecuteNonQuery(SqlConnection conn, String commandText, CommandType commandType, params SqlParameter[] parameters)
        {
            using (SqlCommand cmd = new SqlCommand(commandText, conn))
            {
                // There're three command types: StoredProcedure, Text, TableDirect. The TableDirect   
                // type is only for OLE DB.    
                cmd.CommandType = commandType;

                if (parameters != null)
                {
                    cmd.Parameters.AddRange(parameters);
                }
                var res = cmd.ExecuteNonQuery();
                return res;
            }
        }


        public static Int32 ExecuteNonQuery(String connectionString, String commandText,
            CommandType commandType, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(commandText, conn))
                {
                    // There're three command types: StoredProcedure, Text, TableDirect. The TableDirect   
                    // type is only for OLE DB.    
                    cmd.CommandType = commandType;

                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }
                    conn.Open();
                    var res = cmd.ExecuteNonQuery();
                    conn.Close();
                    return res;

                }
            }
        }

        // Set the connection, command, and then execute the command and only return one value.  
        public static Object ExecuteScalar(String connectionString, String commandText,
            CommandType commandType, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(commandText, conn))
                {
                    cmd.CommandType = commandType;
                    cmd.Parameters.AddRange(parameters);

                    conn.Open();
                    return cmd.ExecuteScalar();
                }
            }
        }

        public static SqlDataReader ExecuteReader(SqlConnection conn, String commandText, CommandType commandType, params SqlParameter[] parameters)
        {
            using (SqlCommand cmd = new SqlCommand(commandText, conn))
            {
                cmd.CommandType = commandType;
                cmd.Parameters.AddRange(parameters);

                // When using CommandBehavior.CloseConnection, the connection will be closed when the   
                // IDataReader is closed.  
                SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.Default);

                return reader;
            }
        }

        // Set the connection, command, and then execute the command with query and return the reader.  
        public static SqlDataReader ExecuteReader(String connectionString, String commandText,
            CommandType commandType, params SqlParameter[] parameters)
        {
            SqlConnection conn = new SqlConnection(connectionString);

            using (SqlCommand cmd = new SqlCommand(commandText, conn))
            {
                cmd.CommandType = commandType;
                cmd.Parameters.AddRange(parameters);

                conn.Open();
                // When using CommandBehavior.CloseConnection, the connection will be closed when the   
                // IDataReader is closed.  
                SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                return reader;
            }
        }
    }

    public class SqlExport
    {
        string connectionString = "Data Source=abg-db10-osl;Initial Catalog=ABGSC_DW;Integrated Security=True";
        SqlConnection conn;

        public SqlExport()
        {
            conn = SqlHelper.GetConnection(connectionString);
        }

        public void Write(List<SmtpMailDetails> mails)
        {
            foreach(var m in mails)
            {
                if (Exists(m) == false)
                {
                    Insert(m);
                }
            }
            //BulkAddRates(connectionString, rates);
        }

        private void BulkAddRates(List<SmtpMailDetails> mails)
        {
            throw (new NotImplementedException());
            string values = string.Empty;
            foreach (var mail in mails)
            {
                //string value = $"('{r.Currency}','{r.Date}',{r.Rate})";
                //if (string.Empty != values) { values += ","; }
                //values += value;
            }

            string commandText = "INSERT INTO [Staging].[OpenExchangeRates]([currency],[valuation_date],[rate]) VALUES";
            commandText += values;

            //var parameterCredits = new SqlParameter("@Credits", 0);

            var rows = SqlHelper.ExecuteNonQuery(connectionString, commandText, CommandType.Text, null);
        }

        public int Insert(SmtpMailDetails mail)
        {
            var x = mail.mime_mail_to_list;
            string commandText = $"INSERT INTO [dbo].[smtp_mail_details] " +
                                $" ([smtp_mail_batch_id], [mime_mail_to_list], [mime_mail_to_name_list], [mime_attachment_list]) " +
                                $" VALUES ({mail.smtp_mail_batch_id}, '{mail.mime_mail_to_list.First().Item1 }', '{mail.mime_mail_to_list.First().Item2}', '{mail.mime_attachment_list}')";
            return SqlHelper.ExecuteNonQuery(conn, commandText, CommandType.Text, null);
        }

        public bool Exists(SmtpMailDetails mail)
        {
            string commandText = $"SELECT COUNT(*) as numRows FROM dbo.smtp_mail_details WHERE [smtp_mail_batch_id]={mail.smtp_mail_batch_id} "
                                + $" AND [mime_mail_to_list] = '{mail.mime_mail_to_list.First().Item1}' "
                                + $" AND [mime_mail_to_name_list] = '{mail.mime_mail_to_list.First().Item2}' "
                                + $" AND [mime_attachment_list] = '{mail.mime_attachment_list.First()}' ";
            var reader = SqlHelper.ExecuteReader(conn, commandText, CommandType.Text);
            
            int numRows = 0;
            if (reader.Read() == true) 
            { 
                numRows = Convert.ToInt32(reader["numRows"]); 
            }
            reader.Close();
            return numRows != 0;
        }
    }
}
