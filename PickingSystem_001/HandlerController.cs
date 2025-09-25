using System;
using System.Data;
using System.Globalization;
using System.IO;
using ExcelDataReader;
using DataTable = System.Data.DataTable;

namespace PickingSystem_001
{
    class HandlerController
    {
        private frmSystem _form;
        private Dac dac = new Dac();

        // 생성자로 frmSystem 객체를 받아옴
        public HandlerController(frmSystem form)
        {
            _form = form;
        }

        public DataTable ReadStoredData(string date)
        {
            string query = "SELECT * FROM tb_rawdata WHERE picking_date = @date " +
                "ORDER BY PICKING_DATE ASC";
            _form.writeRtbNotice("tb_rawdata 테이블 조회 시작..");

            try
            {
                // DB 조회 
                DataTable dt = dac.ExcuteQuery(query, date);
                return dt;
            }
            catch (Exception ex)
            {
                _form.writeRtbNotice("데이터 불러오는 도중 오류 발생...");
                Console.WriteLine("ReadStoredData 에서 오류 발생 : " + ex.Message);
            }
            return null;
        }
        
        public void ExcelToDataTable(string fileName)
        {
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // 읽은 Excel 파일을 DataSet 으로 변환 
                        // ExcelDataSetConfiguration 을 통해 설정 지정 
                        _form.writeRtbNotice("파일 읽기 시작...");

                        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            // 람다식 / 매개변수(IExcelDataReader)를 받아서 ExcelDataTableConfiguration 객체를 돌려준다. 
                            // 람다식에서 매개변수를 반드시 써야 하는데, 실제로 그 값을 안쓸 경우 _ 같은 식으로 표시 
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true, // 첫번째 행을 DataTable 컬럼명으로 사용 (true)
                                FilterColumn = (rowReader, columnIndex) =>
                                {
                                    return columnIndex == 1 || columnIndex == 3 || columnIndex == 4 || columnIndex == 5 || columnIndex == 6 || columnIndex == 7 || columnIndex == 8; // 몇번째 컬럼을 포함할지 결정하는 필터 
                                }
                            }
                        });
                        DataTable excelDt = dataSet.Tables["RAW출고"]; // 엑셀 원본 

                        DataTable dt = new DataTable(); // 내가 쓸 DataTable (DB 스키마에 맞게 타입 지정) 
                        dt.Columns.Add("피킹일자", typeof(DateTime)); // Date 컬럼
                        dt.Columns.Add("피킹번호", typeof(string));
                        dt.Columns.Add("거래처코드", typeof(string));
                        dt.Columns.Add("거래처명", typeof(string));
                        dt.Columns.Add("제품코드", typeof(string));
                        dt.Columns.Add("제품명", typeof(string));
                        dt.Columns.Add("피킹수량", typeof(int));

                        // 변환 루프 
                        foreach (DataRow row in excelDt.Rows)
                        {
                            DataRow newRow = dt.NewRow();

                            // 날짜 변환 
                            var pickingDate = DateTime.ParseExact(row["피킹\n일자"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                            newRow["피킹일자"] = pickingDate;

                            // 문자열 
                            newRow["피킹번호"] = row["피킹\n번호"].ToString();
                            newRow["거래처코드"] = row["거래처\n코드"].ToString();
                            newRow["거래처명"] = row["거래처명"].ToString();
                            newRow["제품코드"] = row["제품\n코드"].ToString();
                            newRow["제품명"] = row["제품명"].ToString();

                            // 정수 변환 
                            if (int.TryParse(row["피킹\n수량"].ToString(), out int qty))
                                newRow["피킹수량"] = qty;
                            else
                                newRow["피킹수량"] = 0;

                            dt.Rows.Add(newRow);
                            
                        }
                        dac.BulkInsert(dt);

                        _form.writeRtbNotice("데이터 입력성공 축하...");
                    }
                } catch (Exception ex)
                {
                    Console.WriteLine("ExcelToDataTable 에서 오류 터짐 ㅠㅠ : " + ex.Message);
                }
                
            }
        }
    }
}