using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace PickingSystem_001
{
    class HandlerController
    {
        private frmSystem _form;
        private DAC dac = new DAC();

        // 생성자로 frmSystem 객체를 받아옴
        public HandlerController(frmSystem form)
        {
            _form = form;
        }

        public void ReadExcelSheet(string filePath)
        {
            Application application = new Application();
            Workbook workbook = application.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Worksheets[1]; // 1번째 Worksheet 에 대한 객체 가져오기 

            // A 열부터 I 열까지 가져오기
            Range range = worksheet.Range["A:I"];

            try
            { 
                _form.writeRtbNotice("파일 데이터 읽는중... 시간 꽤 소요됨...");
                List<Dictionary<string, object>> paramList = new List<Dictionary<string, object>>();
                _form.writeRtbNotice("데이터 DB에 넣기 시작할거임...");
                int row = 2; // 두번째 행 부터 시작 
                do
                {    
                    Dictionary<string, object> parameters = new Dictionary<string, object>();

                    for (int column = 1; column <= range.Columns.Count; column++)
                    {
                        // Object 타입 반환 (실제로는 COM 객체 -> Excel.Range) 
                        parameters["@pickingDate"] = Convert.ToString(((Range)range.Cells[row, 2]).Value2); // B열 피킹일자
                        parameters["@custCode"] = Convert.ToString(((Range)range.Cells[row, 5]).Value2); // E열 거래처코드
                        parameters["@custName"] = Convert.ToString(((Range)range.Cells[row, 6]).Value2); // F열 거래처명
                        parameters["@pickingCode"] = Convert.ToString(((Range)range.Cells[row, 4]).Value2); // D열 피킹번호
                        parameters["@itemCode"] = Convert.ToString(((Range)range.Cells[row, 7]).Value2); // G열 제품코드
                        parameters["@itemName"] = Convert.ToString(((Range)range.Cells[row, 8]).Value2); // H열 제품명 
                        parameters["@qty"] = Convert.ToInt32(((Range)range.Cells[row, 9]).Value2); // I열 피킹수량
                    }
                    // paramList.Add(parameters);

                    // 한줄씩 
                    string query = "INSERT INTO tb_rawdata (PICKING_DATE, CUST_CODE, CUST_NAME, PICKING_CODE, ITEM_CODE, ITEM_NAME, QTY)" +
                        "VALUES (@pickingDate, @custCode, @custName, @pickingCode, @itemCode, @itemName, @qty)";
                    dac.ExcuteNonQuery(query, parameters);
                    row++;
                }
                while (row > range.Rows.Count);
                
                _form.writeRtbNotice("데이터 입력성공 축하...");

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