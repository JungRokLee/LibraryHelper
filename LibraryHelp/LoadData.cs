using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using agi = HtmlAgilityPack;
using HtmlAgilityPack;

namespace LibraryHelp
{
    public class LoadData  //엑셀 파일을 불러오는 클래스
    {
        Excel.Application excelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        /* <tag> String </tag> ==> String 만 추출  */
        /*public static string GetMiddleString(string str, string begin, string end)
        {
            if (string.IsNullOrEmpty(str))
            {
                return null;
            }

            string result = null;
            if (str.IndexOf(begin) > -1)
            {
                str = str.Substring(str.IndexOf(begin) + begin.Length);
                if (str.IndexOf(end) > -1) result = str.Substring(0, str.IndexOf(end));
                else result = str;
            }
            return result;
        }
        */

        /* Excel 불러오기 */
        public void ReadExcelData(string path, int SheetNum, int ColNum)
        {
            /* path는 Excel파일의 전체 경로
             * SheetNum은 작업을 시작 할 시트 번호
             * ColNum은 작업을 시작 할 열 번호 
             * */

            try
            {
                excelApp = new Excel.Application();   // 엑셀 객체 생성
                excelApp.Visible = true; // 엑셀 시트 화면 On
                wb = excelApp.Workbooks.Open(path); // 작업 경로 

                ws = wb.Worksheets.get_Item(SheetNum) as Excel.Worksheet; //원하는 시트
                Excel.Range range = ws.UsedRange;   // 시트 설정
                object[,] data = range.Value; // 현재 Worksheet에서 사용된 셀 전체를 선택

                int book_col = 3; //데이터에서 bokname의 열위치
                int url_col = 9;  //데이터에서 URL의 열위치

                // 열들에 들어있는 Data를 배열 (One-based array)로 받아온다
                for (int r = ColNum; r <= data.GetLength(0); r++)
                {
                    string ExcelBookName = (string)data[r, book_col]; //원본 데이터의 책 제목
                    string URL = (string)data[r, url_col];            //원본 데이터의 URL 필드의 값 

                    //////////////////////////////////////////////// Html Agility Pack 사용하여 URL파싱 //////////////////////////////////////////////////////////////

                    WebClient wc = new WebClient(); // WebClient 객체 생성
                    wc.Encoding = Encoding.UTF8; // 해당 객체를 UTF-8 형식으로 인코딩
                    string html = wc.DownloadString(URL); // URL 필드의 값을 요청하여 리소스를 string으로 저장
                    agi.HtmlDocument doc = new agi.HtmlDocument(); // HtmlDocument 인스턴스 생성
                    doc.LoadHtml(html); // 요청한 리소스 로드
                    HtmlNode divContainer = doc.DocumentNode.SelectSingleNode("//div[@class='header']"); // 로드한 결과에서 header 클래스 내부의 string만 추출
                    string webBookName = divContainer.InnerText; // 최종적으로 divContainer 값을 string 형식으로 저장

                    string convertedWebBookName = WebUtility.HtmlDecode(webBookName); // 특수문자 일반화 ex) &amp -> &
                    String trimedWebBookName = convertedWebBookName.Trim(); // convertedWebBookName 앞 뒤의 공백제거

                    /*  KMP알고리즘으로 WEBBOOKNAME에서 등장한 EXCELBOOKNAME을 찾는다  */
                    KMP stringCompare = new KMP();
                    List<int> list = stringCompare.kmpAlgorithm(trimedWebBookName, ExcelBookName);  //KMP 함수의 반환값(List) 저장, 즉, trimedWebBookName에 등장한 ExcelBookName 패턴 개수 찾기
                    int equalSize = list.Count;                       // 찾은 ExcelBookName 패턴이 몇개인지 저장할 변수, 1개라면 찾은 것임

                    if (equalSize == 0) // 패턴이 등장하지 않았으면 --> Red 색칠
                    {
                        ws.Cells[r, 11] = "1";  // 11번째 attribute에 '1'출력 (사용자의 요구)
                        ws.Cells[r, 12] = trimedWebBookName;  // 제일 오른쪽에 웹페이지의 원본 책제목 출력 (오류 확인용)
                        ws.Cells[r, book_col].interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else // 같다면 해당 row -> Yellow
                    {
                        ws.Cells[r, book_col].interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    }
                }
            }
            catch (Exception ex)
            {

                System.Windows.Forms.MessageBox.Show("오류발생");
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
                //throw ex;
            }
            finally
            {
                System.Windows.Forms.MessageBox.Show("작업완료!");
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }
        }
        public static void ReleaseExcelObject(object obj) //프로세스 종료시키는 메소드
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
            }
        }

    }
}
