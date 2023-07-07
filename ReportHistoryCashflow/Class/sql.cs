using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ReportHistoryCashflow.Class
{
        class sql
        {
            string _strConn = ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;

            /**
            * Execute query (insert, update, delete) menggunakan sql query
            */
            public int ExecuteNonQuery(string strSqlQuery)
            {
                SqlConnection sqlConn = null;
                SqlCommand sqlCmd = null;
                int iRowsAffected = 0;

                try
                {
                    sqlConn = new SqlConnection();
                    sqlConn.ConnectionString = _strConn;
                    if (sqlConn.State == ConnectionState.Closed)
                    {
                        sqlConn.Open();
                    }
                    sqlCmd = new SqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandText = strSqlQuery;
                    iRowsAffected = sqlCmd.ExecuteNonQuery();
                }
                catch (Exception)
                {
                    throw ;
                }
                finally
                {
                    if (sqlConn.State == ConnectionState.Open)
                    {
                        sqlCmd.Dispose();
                        sqlCmd = null;
                        sqlConn.Close();
                        sqlConn.Dispose();
                        sqlConn = null;
                    }
                }

                return (iRowsAffected);
            }

            /**
             * Mengambil data dari table menggunakan sql query
             */

            public DataTable GetDataTable(string strSqlQuery)
            {
                SqlConnection sqlConn = null;
                SqlDataAdapter sqlDa = null;
                SqlCommand sqlCmd;
                DataTable dt = null;

                try
                {
                    sqlConn = new SqlConnection();
                    sqlConn.ConnectionString = _strConn;
                    if (sqlConn.State == ConnectionState.Closed)
                    {
                        sqlConn.Open();
                    }
                    sqlCmd = new SqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandText = strSqlQuery;
                    sqlDa = new SqlDataAdapter(sqlCmd);
                    dt = new DataTable();
                    sqlDa.Fill(dt);
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    if (sqlConn.State == ConnectionState.Open)
                    {
                        sqlDa.Dispose();
                        sqlDa = null;
                        sqlConn.Close();
                        sqlConn.Dispose();
                        sqlConn = null;
                    }
                }

                return (dt);
            }


            public DataRow GetDataRow(string strSqlQuery)
            {
                SqlConnection sqlConn = null;
                SqlDataAdapter sqlDa = null;
                SqlCommand sqlCmd;
                DataTable dt = null;

                try
                {
                    sqlConn = new SqlConnection();
                    sqlConn.ConnectionString = _strConn;
                    if (sqlConn.State == ConnectionState.Closed)
                    {
                        sqlConn.Open();
                    }
                    sqlCmd = new SqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandText = strSqlQuery;
                    sqlDa = new SqlDataAdapter(sqlCmd);
                    dt = new DataTable();
                    sqlDa.Fill(dt);
                }
                catch (Exception)
                {
                    dt = null;
                }
                finally
                {
                    if (sqlConn.State == ConnectionState.Open)
                    {
                        sqlDa.Dispose();
                        sqlDa = null;
                        sqlConn.Close();
                        sqlConn.Dispose();
                        sqlConn = null;
                    }
                }

                if ((dt.Rows.Count == 0) || (dt == null))
                {
                    return null;
                }
                else
                {
                    return dt.Rows[0];
                }
            }




            //via store procedure
            /**S
          * Execute query (insert, update, delete) menggunakan stored procedure
          */
            public int ExecuteNonQuery(string strStoredProcedureName, List<SqlParameter> ListSqlParams)
            {
                SqlConnection sqlConn = null;
                SqlCommand sqlCmd = null;
                SqlParameter sqlParam;
                int iRowsAffected = 0;

                try
                {
                    sqlConn = new SqlConnection();
                    sqlConn.ConnectionString = _strConn;
                    if (sqlConn.State == ConnectionState.Closed)
                    {
                        sqlConn.Open();
                    }
                    sqlCmd = new SqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandText = strStoredProcedureName;
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    //param
                    if (ListSqlParams != null)
                    {
                        for (int iCount = 0; iCount < ListSqlParams.Count; iCount++)
                        {
                            sqlParam = new SqlParameter();
                            sqlParam = ListSqlParams[iCount];
                            sqlCmd.Parameters.Add(sqlParam);
                        }
                    }

                    iRowsAffected = sqlCmd.ExecuteNonQuery();
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    if (sqlConn.State == ConnectionState.Open)
                    {
                        sqlCmd.Dispose();
                        sqlCmd = null;
                        sqlConn.Close();
                        sqlConn.Dispose();
                        sqlConn = null;
                    }
                }

                return (iRowsAffected);
            }

            /**
             * Mengambil data dari table menggunakan stored procedure
             */

            public DataTable GetDataTable(string strStoredProcedureName, List<SqlParameter> ListSqlParams)
            {
                SqlConnection sqlConn = null;
                SqlDataAdapter sqlDa = null;
                SqlCommand sqlCmd;
                SqlParameter sqlParam;
                DataTable dt = null;

                try
                {
                    sqlConn = new SqlConnection();
                    sqlConn.ConnectionString = _strConn;
                    if (sqlConn.State == ConnectionState.Closed)
                    {
                        sqlConn.Open();
                    }
                    sqlCmd = new SqlCommand();
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.CommandText = strStoredProcedureName;
                    //param
                    if (ListSqlParams != null)
                    {
                        for (int iCount = 0; iCount < ListSqlParams.Count; iCount++)
                        {
                            sqlParam = new SqlParameter();
                            sqlParam = ListSqlParams[iCount];
                            sqlCmd.Parameters.Add(sqlParam);
                        }
                    }

                    sqlDa = new SqlDataAdapter(sqlCmd);
                    dt = new DataTable();
                    sqlDa.Fill(dt);
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    if (sqlConn.State == ConnectionState.Open)
                    {
                        sqlDa.Dispose();
                        sqlDa = null;
                        sqlConn.Close();
                        sqlConn.Dispose();
                        sqlConn = null;
                    }
                }

                return (dt);
            }

        }
    }

