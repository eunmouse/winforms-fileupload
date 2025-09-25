using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using DataTable = System.Data.DataTable;

namespace PickingSystem_001
{
    // DB 연결/쿼리 실행/데이터 반환 전담 클래스 
    public class Dac
    {
        private string dbAddr = ConfigurationManager.ConnectionStrings["DBCON"].ConnectionString;
        private SqlConnection connection; // DB 연결
        private SqlCommand command; // DB 에 쿼리문 전달

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
        public DataTable ExcuteQuery(string sql, string date)
        {
            // 쿼리 실행 (SELECT 문) 
            using (connection = new SqlConnection(dbAddr))
            {
                connection.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(); // DB 에서 반환된 데이터를 DataSet 클래스로 변환 
                DataSet ds = new DataSet();
                
                // 파라미터 타입 및 값 입력
                using (command = new SqlCommand(sql, connection))
                {
                    command.Parameters.AddWithValue("@date", date);
                    adapter.SelectCommand = command;
                    adapter.Fill(ds);
                 
                    if (ds.Tables.Count > 0)
                    {
                        foreach (DataTable dt in ds.Tables)
                        {
                            return dt;
                        }
                    }
                    return null;
                }
            }
        }
        public void BulkInsert(DataTable dt)
        {
            using (connection = new SqlConnection(dbAddr))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction())
                {
                    // SqlBulkCopy 생성자 (SqlConnection, SqlBulkCopyOptions, SqlTransaction)
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.Default, transaction)) // .Default INSERT 할 때 DB 가 알아서 자동값 부여 
                    {
                        bulkCopy.DestinationTableName = "dbo.TB_RAWDATA";

                        // DataTable 컬럼명<>DB 컬럼명 매핑
                        bulkCopy.ColumnMappings.Add("피킹일자", "PICKING_DATE");
                        bulkCopy.ColumnMappings.Add("피킹번호", "PICKING_CODE");
                        bulkCopy.ColumnMappings.Add("거래처코드", "CUST_CODE");
                        bulkCopy.ColumnMappings.Add("거래처명", "CUST_NAME");
                        bulkCopy.ColumnMappings.Add("제품코드", "ITEM_CODE");
                        bulkCopy.ColumnMappings.Add("제품명", "ITEM_NAME");
                        bulkCopy.ColumnMappings.Add("피킹수량", "QTY");

                        try
                        {
                            bulkCopy.BatchSize = 10000; // 한 번에 전송할 행 개수 
                            bulkCopy.WriteToServer(dt);
                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("BulkInsert 중 터짐ㅠㅠ 사유 : " + ex.Message);
                            transaction.Rollback();
                        }
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
