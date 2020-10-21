using System;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Xml;
using System.Data.Common;

namespace PSH_BOne_AddOn.Database
{
    public class PSH_DatabaseHelpClass
    {
        private static readonly string connectString = "Data Source=191.1.1.220; Initial Catalog = PSHDB; Persist Security Info=True; User ID = sa; Password=password1!";
        private SqlConnection cn = null;

        /// <summary>
        /// 생성자(기본)
        /// </summary>
        public PSH_DatabaseHelpClass()
        {
            try
            {
                if (cn == null)
                    cn = CreateConnection(connectString);

                if (cn.State != ConnectionState.Open)
                    cn.Open();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 생성자(서버 정보를 동적 인자로 전달 받음)
        /// </summary>
        /// <param name="pSourceIPAddress">서버 IP</param>
        /// <param name="pDatabaseName">서버 DB 이름</param>
        /// <param name="pUserID">DB UserID</param>
        /// <param name="pPassword">DB Password</param>
        public PSH_DatabaseHelpClass(string pSourceIPAddress, string pDatabaseName, string pUserID, string pPassword)
        {
            string contString = "Data Source = " + pSourceIPAddress + "; Initial Catalog = " + pDatabaseName + "; Persist Security Info = True; User ID = " + pUserID + "; Password = " + pPassword;
            
            try
            {
                if (cn == null)
                {
                    cn = CreateConnection(contString);
                }

                if (cn.State != ConnectionState.Open)
                {
                    cn.Open();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region 연결(SqlConnection), 명령(SqlCommand)관련 메서드
        /// <summary>
        /// CreateConnection
        /// </summary>
        /// <param name="connectString"></param>
        /// <returns></returns>
        public SqlConnection CreateConnection(string connectString)
        {
            cn = new SqlConnection(connectString);

            return cn;
        }

        /// <summary>
        /// CreateInParam
        /// </summary>
        /// <param name="paramName"></param>
        /// <param name="type"></param>
        /// <param name="size"></param>
        /// <param name="paramValue"></param>
        /// <returns></returns>
        public SqlParameter CreateInParam(string paramName, SqlDbType type, int size, object paramValue)
        {
            SqlParameter param = new SqlParameter(paramName, type, size);
            param.Value = paramValue;

            return param;
        }

        /// <summary>
        /// CreateOutParam
        /// </summary>
        /// <param name="paramName"></param>
        /// <param name="type"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public SqlParameter CreateOutParam(string paramName, SqlDbType type, int size)
        {
            SqlParameter param = new SqlParameter(paramName, type, size);
            param.Direction = ParameterDirection.Output;

            return param;
        }

        /// <summary>
        /// CreateReturnParam
        /// </summary>
        /// <param name="paramName"></param>
        /// <returns></returns>
        public SqlParameter CreateReturnParam(string paramName)
        {
            SqlParameter param = new SqlParameter(paramName, SqlDbType.Int, 4);
            param.Direction = ParameterDirection.ReturnValue;

            return param;
        }

        /// <summary>
        /// CreateCommand
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public SqlCommand CreateCommand(SqlConnection connection, string command, CommandType type)
        {
            SqlCommand cmd = new SqlCommand(command, connection);
            cmd.CommandType = type;
            cmd.CommandTimeout = 0;

            return cmd;
        }

        /// <summary>
        /// CreateCommand
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public SqlCommand CreateCommand(SqlConnection connection, string command, CommandType type, SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand(command, connection);
            cmd.CommandType = type;
            cmd.CommandTimeout = 0;

            if (parameters != null)
            {
                foreach (SqlParameter p in parameters)
                {
                    cmd.Parameters.Add(p);
                }   
            }

            return cmd;
        }

        /// <summary>
        /// 연결 종료
        /// </summary>
        public void Close()
        {
            if (cn != null)
            {
                cn.Close();
            }   
        }
        #endregion

        #region 데이터베이스 조회(SELECT)와 관련 메서드
        /// <summary>
        /// ExecuteReader
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public SqlDataReader ExecuteReader(string command, CommandType type)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type);
            SqlDataReader dr = null;

            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                return dr;
            }
            catch (Exception ex)
            {
                if (dr != null)
                {
                    dr.Close();
                    dr = null;
                }
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        /// <summary>
        /// ExecuteReader
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public SqlDataReader ExecuteReader(string command, CommandType type, SqlParameter[] parameters)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type, parameters);
            SqlDataReader dr = null;

            try
            {
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                return dr;
            }
            catch (Exception ex)
            {
                if (dr != null)
                {
                    dr.Close();
                    dr = null;
                }
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        /// <summary>
        /// ExecuteDataTable
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public DataTable ExecuteDataTable(string command, CommandType type)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type);
            SqlDataAdapter adp = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();

            try
            {
                adp.Fill(ds);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    return ds.Tables[0];
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                adp.Dispose();
                cmd.Dispose();
                ds.Dispose();
            }
        }

        /// <summary>
        /// ExecuteDataTable
        /// </summary>
        /// <param name="command">SQL 구문</param>
        /// <param name="type">SQL TYPE(Procedure, String)</param>
        /// <param name="parameters">SQL 매개변수</param>
        /// <returns></returns>
        public DataTable ExecuteDataTable(string command, CommandType type, SqlParameter[] parameters)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type, parameters);
            SqlDataAdapter adp = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();

            try
            {
                adp.Fill(ds);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    return ds.Tables[0];
                }
                else
                {
                    return null;
                }   
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                adp.Dispose();
                cmd.Dispose();
                ds.Dispose();
            }
        }

        /// <summary>
        /// ExecuteXmlReader
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public XmlReader ExecuteXmlReader(string command, CommandType type)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type);
            XmlReader reader = null;

            try
            {
                reader = cmd.ExecuteXmlReader();
                return reader;
            }
            catch (Exception ex)
            {
                if (reader != null)
                {
                    reader.Close();
                }
                    
                reader = null;
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        /// <summary>
        /// ExecuteXmlReader
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public XmlReader ExecuteXmlReader(string command, CommandType type, SqlParameter[] parameters)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type, parameters);
            XmlReader reader = null;

            try
            {
                reader = cmd.ExecuteXmlReader();
                return reader;
            }
            catch (Exception ex)
            {
                if (reader != null)
                {
                    reader.Close();
                }

                reader = null;
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        /// <summary>
        /// ExecuteScalar
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public object ExecuteScalar(string command, CommandType type)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type);

            try
            {
                return cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        /// <summary>
        /// ExecuteScalar
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public object ExecuteScalar(string command, CommandType type, SqlParameter[] parameters)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type, parameters);

            try
            {
                return cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        /// <summary>
        /// ExecuteDataSet
        /// </summary>
        /// <param name="strQuery"></param>
        /// <param name="strAlias"></param>
        /// <param name="dsDataSet"></param>
        /// <param name="paramArray"></param>
        /// <param name="cmdType"></param>
        /// <param name="dataTableMappings"></param>
        /// <param name="TimeOut">커멘드 타임아웃</param>
        /// <returns></returns>
        public DataSet ExecuteDataSet(string strQuery, string strAlias, DataSet dsDataSet, SqlParameter[] paramArray, CommandType cmdType, DataTableMapping[] dataTableMappings, int TimeOut)
        {
            try
            {
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(strQuery, this.cn);
                sqlAdapter.SelectCommand.CommandType = cmdType;
                sqlAdapter.SelectCommand.CommandTimeout = TimeOut;

                if (dsDataSet == null)
                {
                    dsDataSet = new DataSet();
                }

                if (paramArray != null)
                {
                    foreach (SqlParameter param in paramArray)
                    {
                        if (strAlias != null)
                        {
                            sqlAdapter.SelectCommand.Parameters.Add(param);
                        }
                    }
                }

                if (strAlias != null)
                {
                    sqlAdapter.Fill(dsDataSet, strAlias);
                }
                else
                {
                    if (dataTableMappings != null)
                    {
                        sqlAdapter.TableMappings.AddRange(dataTableMappings);
                    }
                    sqlAdapter.Fill(dsDataSet);
                }

                sqlAdapter.SelectCommand.Parameters.Clear();
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            finally
            {

            }

            return dsDataSet;
        }

        /// <summary>
        /// ExecuteDataSet
        /// </summary>
        /// <param name="commandText"></param>
        /// <param name="strAlias"></param>
        /// <param name="dsDataSet"></param>
        /// <param name="paramArray"></param>
        /// <param name="cmdType"></param>
        /// <returns></returns>
        public DataSet ExecuteDataSet(string commandText, string strAlias, DataSet dsDataSet, SqlParameter[] paramArray, CommandType cmdType)
        {
            return ExecuteDataSet(commandText, strAlias, dsDataSet, paramArray, cmdType, null, 3600);
        }

        /// <summary>
        /// ExecuteDataSet
        /// </summary>
        /// <param name="commandText"></param>
        /// <param name="strAlias"></param>
        /// <param name="dsDataSet"></param>
        /// <param name="paramArray"></param>
        /// <returns></returns>
        public DataSet ExecuteDataSet(string commandText, string strAlias, DataSet dsDataSet, SqlParameter[] paramArray)
        {
            return ExecuteDataSet(commandText, strAlias, dsDataSet, paramArray, CommandType.Text, null, 3600);
        }

        /// <summary>
        /// ExecuteDataSet
        /// </summary>
        /// <param name="commandText"></param>
        /// <param name="strAlias"></param>
        /// <param name="dsDataSet"></param>
        /// <param name="paramArray"></param>
        /// <param name="TimeOut">커멘드 타임아웃</param>
        /// <returns></returns>
        public DataSet ExecuteDataSet(string commandText, string strAlias, DataSet dsDataSet, SqlParameter[] paramArray, int TimeOut)
        {
            return ExecuteDataSet(commandText, strAlias, dsDataSet, paramArray, CommandType.Text, null, TimeOut);
        }

        /// <summary>
        /// Fill
        /// </summary>
        /// <param name="commandText"></param>
        /// <param name="strAlias"></param>
        /// <param name="dsDataSet"></param>
        /// <param name="paramArray"></param>
        /// <returns></returns>
        public DataSet Fill(string commandText, string strAlias, DataSet dsDataSet, SqlParameter[] paramArray)
        {
            return ExecuteDataSet(commandText, strAlias, dsDataSet, paramArray, CommandType.Text, null, 3600);
        }

        /// <summary>
        /// Fill
        /// </summary>
        /// <param name="commandText"></param>
        /// <param name="strAlias"></param>
        /// <param name="dsDataSet"></param>
        /// <param name="paramArray"></param>
        /// <param name="TimeOut">커맨드타임아웃</param>
        /// <returns></returns>
        public DataSet Fill(string commandText, string strAlias, DataSet dsDataSet, SqlParameter[] paramArray, int TimeOut)
        {
            return ExecuteDataSet(commandText, strAlias, dsDataSet, paramArray, CommandType.Text, null, TimeOut);
        }

        /// <summary>
        /// ExecuteDataSet
        /// </summary>
        /// <param name="commandText"></param>
        /// <param name="strAlias"></param>
        /// <param name="dsDataSet"></param>
        /// <returns></returns>
        public DataSet ExecuteDataSet(string commandText, string strAlias, DataSet dsDataSet)
        {
            return ExecuteDataSet(commandText, strAlias, dsDataSet, null, CommandType.Text, null, 3600);
        }
        #endregion

        #region 데이터베이스의 테이터 변경(UPDATE, INSERT, DELETE)과 관련 메서드
        /// <summary>
        /// ExecuteNonQuery
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public int ExecuteNonQuery(string command, CommandType type)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type);

            try
            {
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        /// <summary>
        /// ExecuteNonQuery
        /// </summary>
        /// <param name="command"></param>
        /// <param name="type"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public int ExecuteNonQuery(string command, CommandType type, SqlParameter[] parameters)
        {
            SqlCommand cmd = CreateCommand(this.cn, command, type, parameters);

            try
            {
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Dispose();
            }
        }
        #endregion
    }
}
