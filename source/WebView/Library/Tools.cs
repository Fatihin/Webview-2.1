using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.EntityClient;
using System.Data.EntityModel;
using Oracle.DataAccess.Client;
using System.Data.Objects;
using System.Data.Common;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;

namespace WebView.Library
{
    public class Tools
    {
        public static bool IsValidEmail(string strIn)
        {
            // Return true if strIn is in valid e-mail format.
            return Regex.IsMatch(strIn,
                   @"^(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +
                   @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$");
        }

        public bool ExecuteSql(ObjectContext c, string sql)
        {
            var entityConnection = (System.Data.EntityClient.EntityConnection)c.Connection;
            DbConnection conn = entityConnection.StoreConnection;

            ConnectionState initialState = conn.State;
            try
            {
                if (initialState != ConnectionState.Open)
                    conn.Open();  // open connection if not already open


                using (DbCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();                    
                }

                return true;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.ToString());
                return false;
            }
            finally
            {
                if (initialState != ConnectionState.Open)
                    conn.Close(); // only close connection if not initially open
            }
        }


        public void Execute(string conn, string sql)
        {
            // Specify the parameter value.
            int paramValue = 5;

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OracleConnection connection =
                new OracleConnection(conn))
            {
                // Create the Command and Parameter objects.
                OracleCommand command = new OracleCommand(sql, connection);
                //command.Parameters.AddWithValue("@pricePoint", paramValue);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                //System.Diagnostics.Debug.WriteLine("try");
                try
                {
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        System.Diagnostics.Debug.WriteLine("!");
                        System.Diagnostics.Debug.WriteLine(reader[0].ToString() + " : " + reader[1].ToString() +
                            reader[2].ToString());
                        //Console.WriteLine("\t{0}\t{1}}",
                        //    reader[0], reader[1]);*/
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Console.ReadLine();
            }
        }

        public string ExecuteStr(string conn, string sql)
        {
            // Specify the parameter value.
            int paramValue = 0;
            string str = "";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OracleConnection connection =
                new OracleConnection(conn))
            {
                // Create the Command and Parameter objects.
                OracleCommand command = new OracleCommand(sql, connection);
                //command.Parameters.AddWithValue("@pricePoint", paramValue);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                //System.Diagnostics.Debug.WriteLine("try");
                string data = "";
                try
                {
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        paramValue++;
                        System.Diagnostics.Debug.WriteLine("!");
                        System.Diagnostics.Debug.WriteLine(reader[0].ToString() + " : " + reader[1].ToString() + " : " +
                            reader[2].ToString());
                        //Console.WriteLine("\t{0}\t{1}}",
                        //    reader[0], reader[1]);*/
                        data = data + reader[0].ToString() + ":" + reader[1].ToString() + ":" + reader[2].ToString() + "!";

                    }
                    reader.Close();

                    str = data; //paramValue.ToString();
                }
                catch (Exception ex)
                {
                    str = ex.Message;
                    Console.WriteLine(ex.Message);
                }
                Console.ReadLine();
            }

            return str;
        }

        // region Mubin - CR74-20180330
        public string ExecuteStored(string conn, string sql, CommandType cType, OracleParameter[] oraPrm, bool hasReturnValue)
        {
            OracleParameter p_ReturnValue;

            using (OracleConnection connection =
                new OracleConnection(conn))
            {
                // Create the Command and Parameter objects.
                OracleCommand command = new OracleCommand(sql, connection);

                try
                {
                    connection.Open();
                    command.CommandText = sql;
                    command.CommandType = cType;

                    if (oraPrm.Count() > 0)
                    {
                        for(int i = 0; i < oraPrm.Count(); i++)
                            command.Parameters.Add(oraPrm[i]);
                    }
 
                    p_ReturnValue = new OracleParameter("p_retval", OracleDbType.Varchar2, 100);
                    p_ReturnValue.Direction = ParameterDirection.Output;

                    if (hasReturnValue)
                    {
                        command.Parameters.Add(p_ReturnValue);
                    }
                    //OracleParameter retval = new OracleParameter("retval", OracleDbType.Varchar2);
                    //retval.Direction = ParameterDirection.ReturnValue;
                    //command.Parameters.Add(retval);

                    command.ExecuteNonQuery();

                    if (hasReturnValue) System.Diagnostics.Debug.WriteLine("affected: " + p_ReturnValue.Value.ToString());
                    //System.Diagnostics.Debug.WriteLine("affected: " + retval.Value); 
                    command.Parameters.Clear();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Message);
                    command.Parameters.Clear();
                    connection.Close();
                    return "fail";
                }
            }

            if (hasReturnValue)
                return p_ReturnValue.Value.ToString();
            else
                return "ok";
        }
        // endRegion
        


        public string[] GetFileList(string location)
        {
            string[] filePaths = new string[0];
            string[] tmpSplit;

            try
            {
                filePaths = Directory.GetFiles(@location);

                if (filePaths.Count() > 0)
                {
                    for (int i = 0; i < filePaths.Count(); i++)
                    {
                        tmpSplit = filePaths[i].Split('\\');
                        filePaths[i] = tmpSplit[tmpSplit.Count() -1];
                    }
                }

            }
            catch(Exception e)
            {
                System.Diagnostics.Debug.Write("########");
                System.Diagnostics.Debug.WriteLine(e.ToString());
            }

            return filePaths;
        }

        public void Log(string msg)
        {
            System.Diagnostics.Debug.Write("#> ");
            System.Diagnostics.Debug.WriteLine(msg);
        }
    }
}