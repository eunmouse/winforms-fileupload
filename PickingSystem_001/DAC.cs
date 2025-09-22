using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace PickingSystem_001
{
    // DB 연결/쿼리 실행/데이터 반환 전담 클래스 
    public class DAC
    {
        private string dbAddr = ConfigurationManager.ConnectionStrings["DBCON"].ConnectionString;
        private SqlConnection connection; // DB 연결
        private SqlCommand command; // DB 에 쿼리문 전달
        private SqlDataAdapter adapter; // DB 에서 반환된 데이터를 DataSet 클래스로 변환

        public void MssqlOpen()
        {
            using (connection = new SqlConnection(dbAddr))
            {
                // DB 연결되는지 확인
                try
                {

                    connection.Open();

                    if (connection.State != ConnectionState.Open)
                    {
                        Console.WriteLine("연결 실패");
                    }
                }
                catch (ConfigurationException cex)
                {
                    Console.WriteLine(cex.Message);
                }
                catch (SqlException sqlex)
                {
                    Console.WriteLine(sqlex.Message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        public void ExcuteNonQuery(string sql, DataRowCollection rows)
        {
            // 쿼리 실행 (INSERT / UPDATE / DELETE 문)
            using (connection = new SqlConnection(dbAddr))
            {
                connection.Open();

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
