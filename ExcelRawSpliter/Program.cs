using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using LinqToExcel;
using LinqToExcel.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRawSpliter {
    class Program {
        static ConsoleColor FColor = Console.ForegroundColor;
        static ConsoleColor BColor = Console.BackgroundColor;        
        static void Main(string[] args) {
            args = new string[] { "TEST.xls" };

            Console.WriteLine("Excel列切割工具 - 徐培堯 - v1.0.0");
            if (args?.Length == 0) {
                Error("請將Excel檔案拖曳至本程式上，程式即可成功執行");
                Console.WriteLine("請按任意鍵關閉本程式...");
                Console.ReadKey();
                return;
            }
            ExcelQueryFactory excel = null;
            Process($"讀取Excel檔案({args[0]})", () => {
                excel = new ExcelQueryFactory(args[0]);
            });
            string[] worksheetNames = null;
            Process("讀取工作表列表", () => {
                worksheetNames = excel.GetWorksheetNames().ToArray();
            });
            Console.WriteLine($"您所選定的文件內有以下工作表，請選定執行工作表");
            Console.WriteLine($"[代號]\t名稱");
            for (int i = 0; i < worksheetNames.Length; i++) {
                Console.WriteLine($"[{i}]\t{worksheetNames[i]}");
            }
            
            int selected = int.Parse(Input("請輸入上列清單中的代號",5));
            Alert($"您已經選擇工作表 {worksheetNames[selected]}");

            IQueryable<RowNoHeader> worksheet = null;
            int HeaderEnd = int.Parse(Input("該工作表表頭終止列(由0起始至此，如未有表頭則輸入0)", 5));
            Process("初始化工作表", () => {
                worksheet = excel.WorksheetNoHeader(selected).Skip(HeaderEnd);
            });

            int splitCount = int.Parse(Input("請輸入每幾筆進行切割", 20));
            bool addHeader = false;
            if(HeaderEnd > 0) {
                addHeader = Input("是否在每個分割檔案加入表頭?(Y/N)", 5).ToUpper() == "Y";
            }


            //excel.Worksheet()
        }
        
        static void Error(string Message) {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(Message);
            Console.ForegroundColor = FColor;
        }
        static void Alert(string Message) {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(Message);
            Console.ForegroundColor = FColor;
        }


        static string Input(string Message,int InputSize = 20) {
            Console.Write($"{Message}: ");
            Console.BackgroundColor = FColor;
            Console.ForegroundColor = BColor;
            Console.Write(new string(' ', InputSize + 2));
            Console.CursorLeft -= InputSize + 1;
            string result = Console.ReadLine();
            Console.BackgroundColor = BColor;
            Console.ForegroundColor = FColor;
            return result;
        }

        delegate void ProcessCallback();
        static void Process(string Message, ProcessCallback Process) {
            Console.Write(Message + "...");
            Process();
            Console.WriteLine("OK");
        }
    }
}
