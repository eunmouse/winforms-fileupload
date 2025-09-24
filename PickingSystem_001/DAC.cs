using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace PickingSystem_001
{
    // DB 연결/쿼리 실행/데이터 반환 전담 클래스 
    public class Dac
    {
        private string dbAddr = ConfigurationManager.ConnectionStrings["DBCON"].ConnectionString;
        private SqlConnection connection; // DB 연결
        private SqlCommand command; // DB 에 쿼리문 전달
        private SqlDataAdapter adapter; // DB 에서 반환된 데이터를 DataSet 클래스로 변환

        public void MssqlOpen()
        {
            using (connection = new SqlConnection(dbAddr))
            {
                try
                {
                    // 가장 먼저 DB 연결 문제없는지 확인
                    connection.Open();

                    if (connection.State != ConnectionState.Open)
                    {
                        Console.WriteLine("연결 실패");
                    }
                }
                catch (ConfigurationException cex) // 구성 오류 
                {
                    Console.WriteLine(cex.Message);
                }
                catch (SqlException sqlex) // 접속 오류 
                {
                    Console.WriteLine(sqlex.Message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        public List<Dictionary<string, object>> ExcuteQuery(string sql, string date)
        {
            // 쿼리 실행 (SELECT 문) 
            using (connection = new SqlConnection(dbAddr))
            {
                connection.Open();
                adapter = new SqlDataAdapter();
                DataSet ds = new DataSet();
                List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
                
                // 파라미터 타입 및 값 입력
                using (command = new SqlCommand(sql, connection))
                {
                    command.Parameters.AddWithValue("@date", date);
                    adapter.SelectCommand = command;
                    adapter.Fill(ds);
                 
                    if (ds.Tables.Count > 0)
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            Dictionary<string, object> dictionary = new Dictionary<string, object>();
                            dictionary["picking_date"] = dr["PICKING_DATE"].ToString();
                            dictionary["cust_code"] = dr["CUST_CODE"].ToString();
                            dictionary["cust_name"] = dr["CUST_NAME"].ToString();
                            dictionary["picking_code"] = dr["PICKING_CODE"].ToString();
                            dictionary["item_code"] = dr["ITEM_CODE"].ToString();
                            dictionary["item_name"] = dr["ITEM_NAME"].ToString();
                            dictionary["qty"] = dr["QTY"].ToString();
                            list.Add(dictionary);
                        }
                        return list;
                    }
                    else
                    {
                        return list;
                    }
                }
            }
        }
        public void ExcuteNonQuery(string sql, DataRowCollection rows)
        {
            // 쿼리 실행 (INSERT / UPDATE / DELETE 문)
            using (connection = new SqlConnection(dbAddr))
            {
                connection.Open();
                //int result = 0;

                using (command = new SqlCommand(sql, connection))
                {
                    command.Parameters.Add("@pickingDate", SqlDbType.NVarChar);
                    command.Parameters.Add("@custCode", SqlDbType.NVarChar);
                    command.Parameters.Add("@custName", SqlDbType.NVarChar);
                    command.Parameters.Add("@pickingCode", SqlDbType.NVarChar);
                    command.Parameters.Add("@itemCode", SqlDbType.NVarChar);
                    command.Parameters.Add("@itemName", SqlDbType.NVarChar);
                    command.Parameters.Add("@qty", SqlDbType.Int);

                    foreach (DataRow dr in rows)
                    {
                        // SQL 파라미터 바인딩
                        command.Parameters["@pickingDate"].Value = dr["pickingDate"]; 
                        command.Parameters["@custCode"].Value = dr["custCode"];
                        command.Parameters["@custName"].Value = dr["custName"];
                        command.Parameters["@pickingCode"].Value = dr["pickingCode"];
                        command.Parameters["@itemCode"].Value = dr["itemCode"];
                        command.Parameters["@itemName"].Value = dr["itemName"];
                        command.Parameters["@qty"].Value = dr["qty"];
                        command.ExecuteNonQuery();
                    }
                    //return result;
                }
            }
        }

        public void MssqlClose()
        {
            // DB 연결 해제 
            try
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Dispose();
                    connection = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
