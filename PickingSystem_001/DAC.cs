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

        public void ExcuteNonQuery(string sql, Dictionary<string, object> rawData)
        {
            // 쿼리 실행 (INSERT / UPDATE / DELETE 문)
            using (connection = new SqlConnection(dbAddr))
            {
                connection.Open();
           
                using (command = new SqlCommand(sql, connection))
                {
                    foreach (var param in rawData)
                    {
                        command.Parameters.AddWithValue(param.Key, param.Value ?? DBNull.Value);
                    }
                    command.ExecuteNonQuery();
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
