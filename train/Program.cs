using System;
using System.IO; // 提供檔案與資料夾操作的功能
using OfficeOpenXml; // EPPlus 函式庫，用於處理 Excel 檔案

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // 提示使用者輸入資料夾路徑
            Console.WriteLine("請輸入資料夾路徑:");
            string inputFolder = Console.ReadLine(); // 接收使用者輸入的資料夾路徑

            // 檢查輸入的資料夾路徑是否為空或不存在
            if (string.IsNullOrWhiteSpace(inputFolder) || !Directory.Exists(inputFolder))
            {
                Console.WriteLine("輸入的資料夾路徑無效或不存在。");
                return; // 結束程式執行
            }

            // 呼叫處理資料夾的核心方法
            ProcessFolder(inputFolder);
        }
        catch (Exception ex)
        {
            // 捕捉未處理的例外並顯示錯誤訊息
            Console.WriteLine($"程式執行過程中發生未處理的錯誤: {ex.Message}");
        }
    }

    public static void ProcessFolder(string inputFolder)
    {
        // 定義輸出資料夾的路徑
        string outputFolder = Path.Combine(inputFolder, "output");

        try
        {
            // 嘗試創建輸出資料夾，如果不存在則建立
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // 尋找輸入資料夾中的 Excel 檔案
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");
            if (excelFiles.Length == 0)
            {
                Console.WriteLine("未找到任何 Excel 文件。");
                return; // 如果沒有 Excel 檔案則退出
            }

            // 假設處理第一個找到的 Excel 檔案
            string excelFile = excelFiles[0];
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 設定 EPPlus 函式庫的非商業許可

            using (var package = new ExcelPackage(new FileInfo(excelFile)))
            {
                // 讀取 Excel 文件的第一個工作表
                var worksheet = package.Workbook.Worksheets[0];
                // 定義輸出文字檔案的路徑
                string txtFilePath = Path.Combine(outputFolder, "output.txt");

                // 開啟文字檔案寫入器
                using (var writer = new StreamWriter(txtFilePath))
                {
                    // 從第二行開始迭代 Excel 的資料列（跳過標題行）
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        try
                        {
                            // 取得 Excel 中的檔案名稱與文本數據
                            string wavFileName = worksheet.Cells[row, 1].Text; // 第一列為 WAV 檔案名稱
                            string textData = worksheet.Cells[row, 2].Text; // 第二列為文本數據
                            string newFileName = $"{row - 1:D4}.wav"; // 將行號轉為新檔案名稱，例如 0001.wav

                            // 定義來源檔案與目標檔案的路徑
                            string sourcePath = Path.Combine(inputFolder, wavFileName);
                            string destinationPath = Path.Combine(outputFolder, newFileName);

                            if (File.Exists(sourcePath))
                            {
                                // 如果來源檔案存在，複製並重命名
                                File.Copy(sourcePath, destinationPath, true);
                                Console.WriteLine($"複製並重命名: {sourcePath} -> {destinationPath}");
                            }
                            else
                            {
                                // 如果來源檔案不存在，記錄錯誤
                                Console.WriteLine($"找不到對應的 WAV 檔案: {sourcePath}");
                            }

                            // 將文本數據寫入輸出的文字檔案
                            writer.WriteLine($"{textData}");
                        }
                        catch (Exception fileEx)
                        {
                            // 捕捉處理單行資料時的錯誤
                            Console.WriteLine($"行 {row} 處理失敗: {fileEx.Message}");
                        }
                    }
                }
            }

            Console.WriteLine("處理完成。"); // 當處理完成時提示使用者
        }
        catch (Exception ex)
        {
            // 捕捉處理資料夾過程中的錯誤
            Console.WriteLine($"資料夾處理過程中發生錯誤: {ex.Message}");
        }
    }

}

