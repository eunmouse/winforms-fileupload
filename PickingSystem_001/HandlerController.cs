using System;
using System.Runtime.InteropServices;
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
                Worksheet worksheet = workbook.Worksheets[1]; // 1번째 Worksheet 에 대한 객체 가져오기 

                // 실제로 사용된 셀 범위 가져오기
                Range range = worksheet.UsedRange;

                // 가져온 행, 열만큼 반복 
                for (int row = 1; row <= range.Rows.Count; row++) 
                {
                    for (int column = 1; column <= range.Columns.Count; column++) 
                    {
                        // Object 타입 반환 (실제로는 COM 객체 -> Excel.Range) 
                        Object value = ((Range)range.Cells[row, column]).Value2;
                        string str = Convert.ToString(value);
                        
                        Console.Write(value + " ");
                    }
                    Console.WriteLine();
                }
                DisposeObject(range);
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
