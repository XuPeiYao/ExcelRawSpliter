using ClosedXML.Excel;
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
            //args = new string[] { @"C:\Users\XuPeiYao\Documents\GitHub\ExcelRawSpliter\ExcelRawSpliter\bin\Debug\TEST.xlsx" };

            Header();
            if (args?.Length == 0) {
                Error("請將Excel檔案拖曳至本程式上，程式即可成功執行");
                Console.WriteLine("請按任意鍵關閉本程式...");
                Console.ReadKey();
                return;
            }
            XLWorkbook excel =null;
            Process($"讀取Excel檔案({args[0]})", () => {
                excel = new XLWorkbook(args[0]);
            });
            string[] worksheetNames = null;
            Process("讀取工作表列表", () => {
                worksheetNames = excel.Worksheets.Select(x=>x.Name).ToArray();
            });
            Console.WriteLine($"您所選定的文件內有以下工作表，請選定執行工作表");
            Console.WriteLine($"[代號]\t名稱");
            for (int i = 0; i < worksheetNames.Length; i++) {
                Console.WriteLine($"[{i}]\t{worksheetNames[i]}");
            }

            int selected = -1;
            do {
                selected = InputInt("請輸入上列清單中的代號", 5);
            } while (selected < 0 || selected >= worksheetNames.Length);
            Alert($"您已經選擇工作表 {worksheetNames[selected]}");

            IEnumerable<object[]> worksheet = null;
            IEnumerable<object[]> header = null;
            int HeaderEnd = InputInt("該工作表表頭終止列(由0起始至此，如未有表頭則輸入0)", 5);
            Process("初始化工作表", () => {
                worksheet = excel.Worksheet(worksheetNames[selected]).Rows().Select(x => {
                    return x.Cells().Select(y=>y.Value).ToArray();
                });
                header = worksheet.Take(HeaderEnd);
                worksheet = worksheet.Skip(HeaderEnd);
            });

            int splitCount = InputInt("請輸入每幾筆進行切割", 20);
            bool addHeader = false;
            if(HeaderEnd > 0) {
                addHeader = Input("是否在每個分割檔案加入表頭?(Y/N)", 5).ToUpper() == "Y";
            }

            Alert("開始進行分割，請等候完成...");

            string fileName = args[0].Split('\\').Last();
            fileName = fileName.Substring(0, fileName.LastIndexOf('.'));
            string path =args[0].Substring(0, args[0].LastIndexOf('\\')) +'\\';

            for(int i= 0; i < worksheet.Count(); i+= splitCount) {
                var tempWorkbook = new XLWorkbook();
                var tempSheet = tempWorkbook.Worksheets.Add("Sheet1");

                //tempSheet.Rows().AdjustToContents();

                IXLCell tempCell = tempSheet.Cell(1,1);
                if (addHeader) {
                    tempCell.InsertData(header);
                    tempCell = tempSheet.Cell(HeaderEnd + 1, 1);
                }
                tempCell.InsertData(worksheet.Skip(i).Take(splitCount));

                double p = i / (double)worksheet.Count();
                string NewPath = path + fileName + $"-{Math.Floor(i / (double)splitCount) + 1}.xlsx";
                Console.WriteLine($"{Math.Round(p * 100)}%\t儲存{NewPath}");
                tempWorkbook.SaveAs(path + fileName + $"-{Math.Floor(i/(double)splitCount) + 1}.xlsx");
            }

            Success("100%\t分割動作已經完成!");

            Console.WriteLine("請按任意鍵關閉本程式...");
            Console.ReadKey();
        }

        static void Header() {
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Blue;
            string title = "Excel 2007+ Rows Spliter - 20160531 - Pei-Yao,Xu";
            Console.WriteLine(
                new string(' ', (Console.WindowWidth - title.Length) / 2)+
                title+
                new string(' ', (Console.WindowWidth - title.Length) / 2)
            );
            Console.BackgroundColor = BColor;
            Console.ForegroundColor = FColor;
        }

        static void Success(string Message) {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(Message);
            Console.ForegroundColor = FColor;
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

        static int InputInt(string Message, int InputSize = 20) {
            Console.Write($"{Message}: ");
            Console.BackgroundColor = FColor;
            Console.ForegroundColor = BColor;
            Console.Write(new string(' ', InputSize + 2));
            Console.CursorLeft -= InputSize + 1;
            int result = 0;
            if(!int.TryParse(Console.ReadLine(),out result)) {
                Console.BackgroundColor = BColor;
                Console.ForegroundColor = FColor;
                Error("不正確的數字");
                result = InputInt(Message, InputSize);
            }
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
