using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace PickingSystem_001
{
    // DB 연결/쿼리 실행/데이터 반환 전담 클래스 
    public class DAC
    {
        // private SqlCommand command;

        // app.config 에 있는 연결 문자열 읽어오기 
        public static SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBCON"].ConnectionString);

        public void MssqlOpen()
        {
            try
            {
                // SqlConnection 연결 
                connection.Open();

                if (connection.State != ConnectionState.Open)
                {
                    Console.WriteLine("연결 실패");
                }

                //string sql = "SELECT * FROM tb_rawdata";
                //
                //command = new SqlCommand(sql, connection);
                //
                //connection.Close();
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
}
