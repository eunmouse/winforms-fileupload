using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
// using DataTable = Microsoft.Office.Interop.Excel.DataTable;
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

        public void ReadStoredData(string date)
        {
            string query = "SELECT * FROM tb_rawdata WHERE picking_date = @date " +
                "ORDER BY PICKING_DATE ASC";
            _form.writeRtbNotice("tb_rawdata 테이블 조회 시작..");

            try
            {
                // DB 조회 
                List<Dictionary<string, object>> dataList = dac.ExcuteQuery(query, date);
                foreach (Dictionary<string, object> dictionary in dataList)
                {
                    string pickingDate = dictionary["picking_date"].ToString();
                    string custCode = dictionary["cust_code"].ToString();
                    string custName = dictionary["cust_name"].ToString();
                    string pickingCode = dictionary["picking_code"].ToString();
                    string itemCode = dictionary["item_code"].ToString();
                    string itemName = dictionary["item_name"].ToString();
                    string qty = dictionary["qty"].ToString();
                    _form.writeRtbNotice(pickingDate + " " + custCode + " " + custName + " " + pickingCode + " " + itemCode + " " + itemName + " " + qty);
                }
            }
            catch (Exception ex)
            {
                _form.writeRtbNotice("데이터 불러오는 도중 오류 발생...");
                Console.WriteLine("ReadStoredData 에서 오류 발생 : " + ex.Message);
            }
            
        }

        public void UploadExcelSheet(string filePath)
        {
            Application application = new Application();
            Workbook workbook = application.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Worksheets[1]; // 1번째 Worksheet 에 대한 객체 가져오기 

            // A 열부터 I 열까지 가져오기
            Range range = worksheet.Range["A:I"]; // Range used = worksheet.UsedRange;
            _form.writeRtbNotice("파일 데이터 읽고 데이터 DB에 바로 넣을거임... 시간 꽤 소요됨...");

            // DataTable 생성 및 컬럼 설정
            DataTable dt = new DataTable();
            dt.Columns.Add("pickingDate", typeof(string));
            dt.Columns.Add("custCode", typeof(string));
            dt.Columns.Add("custName", typeof(string));
            dt.Columns.Add("pickingCode", typeof(string));
            dt.Columns.Add("itemCode", typeof(string));
            dt.Columns.Add("itemName", typeof(string));
            dt.Columns.Add("qty", typeof(int));

            try
            {
                int row = 2; // 두번째 행 부터 시작 
                do
                {
                    // 엑셀 A열 셀이 공백/Null 이면 반복문에서 break; 로 빠지도록 수정 
                    string temp = Convert.ToString(((Range)range.Cells[row, 1]).Value2); // A열 출고구분
                    if (string.IsNullOrWhiteSpace(temp))
                    {
                        break;
                    }
                    DataRow dr = dt.NewRow();
                    // Object 타입 반환 (실제로는 COM 객체 -> Excel.Range) 
                    dr["pickingDate"] = Convert.ToString(((Range)range.Cells[row, 2]).Value2); // B열 피킹일자
                    dr["custCode"] = Convert.ToString(((Range)range.Cells[row, 5]).Value2); // E열 거래처코드
                    dr["custName"] = Convert.ToString(((Range)range.Cells[row, 6]).Value2); // F열 거래처명
                    dr["pickingCode"] = Convert.ToString(((Range)range.Cells[row, 4]).Value2); // D열 피킹번호
                    dr["itemCode"] = Convert.ToString(((Range)range.Cells[row, 7]).Value2); // G열 제품코드
                    dr["itemName"] = Convert.ToString(((Range)range.Cells[row, 8]).Value2); // H열 제품명 
                    dr["qty"] = Convert.ToInt32(((Range)range.Cells[row, 9]).Value2); // I열 피킹수량

                    // 행 추가 
                    dt.Rows.Add(dr);   
                    row++;
                }
                while (true);
                // 한줄씩 
                string query = "INSERT INTO tb_rawdata (PICKING_DATE, CUST_CODE, CUST_NAME, PICKING_CODE, ITEM_CODE, ITEM_NAME, QTY)" +
                    "VALUES (@pickingDate, @custCode, @custName, @pickingCode, @itemCode, @itemName, @qty)";
                dac.ExcuteNonQuery(query, dt.Rows);
                //int result = dac.ExcuteNonQuery(query, dt.Rows);

                _form.writeRtbNotice("데이터 입력성공 축하...");
                //_form.writeRtbNotice("총 " + result + "건의 데이터가 추가되었습니다.");

                DisposeObject(range);
                DisposeObject(worksheet);
                DisposeObject(workbook);
                DisposeObject(application);
            }
            catch (Exception ex)
            {
                _form.writeRtbNotice("파일 읽는 도중 오류 발생...");
                Console.WriteLine("ReadExcelSheet 에서 오류 발생 : " + ex.Message);

                DisposeObject(range);
                DisposeObject(worksheet);
                DisposeObject(workbook);
                DisposeObject(application);
            }
        }
        
        #region 엑셀 객체 해제 메서드 
        public void DisposeObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    // 액셀 객체 해제 
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("DisposeObject 에서 오류 발생 : " + ex.Message);
                throw; // 예외는 위로(호출자 방향으로) 전달 됨, 로그 찍은 후 예외 그대로 다시 던짐 
            }
        }
        #endregion
    }
}