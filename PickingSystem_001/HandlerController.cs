using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace PickingSystem_001
{
    class HandlerController
    {
        public void ReadExcelSheet(string filePath)
        {
            try
            { 
                Application application = new Application();
                Workbook workbook = application.Workbooks.Open(filePath);
                //Worksheet worksheet = workbook.Worksheets.Item[1]; // 1번째 Worksheet 에 대한 객체 가져오기 
                Worksheet worksheet = workbook.Worksheets[0];

                // 사용중인 셀 범위 가져오기
                // Range range = worksheet.Cells["A1:G2"];
                Range range = worksheet.UsedRange;

                for (int row = 1; row <= range.Rows.Count; row++) // 가져온 행만큼 반복
                {
                    for (int column = 1; column <= range.Columns.Count; column++) // 가져온 열만큼 반복
                    {
                        Object value = range.Cells[row, column].Value; // 무조건 Object 타입 반환 

                        string str = value?.ToString(); // value 가 null 이면 str 도 null, value 가 null 이 아니면 ToString() 호출하여 문자열 반환
                    }
                }

                DisposeObject(worksheet);
                DisposeObject(workbook);
                DisposeObject(application);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ReadExcelSheet 에서 오류 발생 : " + ex.Message);
            }
        }

        // 엑셀 객체 해제 메서드 
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
    }
}
