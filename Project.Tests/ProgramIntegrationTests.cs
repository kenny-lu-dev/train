using System;
using System.IO;
using Xunit; // xUnit 是用於單元測試與整合測試的框架
using OfficeOpenXml; // EPPlus 函式庫，用於處理 Excel 檔案

// 定義測試類別
public class ProgramIntegrationTests
{
    // [Fact] 是 xUnit 中的屬性，用於標註這是一個獨立的測試案例
    [Fact]
    public void FullIntegrationTest_ValidScenario()
    {
        // -------------------
        // Arrange: 設置測試環境
        // -------------------

        // 設定 EPPlus 非商業許可，避免建立 ExcelPackage 時拋出授權例外
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // 定義測試用的資料夾路徑
        string testFolder = Path.Combine(Directory.GetCurrentDirectory(), "TestInput");
        // 創建測試輸入資料夾
        Directory.CreateDirectory(testFolder);

        // 定義輸出資料夾路徑與測試用的 Excel 與 WAV 檔案路徑
        string outputFolder = Path.Combine(testFolder, "output");
        string excelFilePath = Path.Combine(testFolder, "test.xlsx");
        string wavFilePath = Path.Combine(testFolder, "test.wav");

        // 創建測試用的 Excel 檔案
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            // 建立一個工作表並設置測試數據
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells[1, 1].Value = "FileName"; // 第一列標題
            worksheet.Cells[1, 2].Value = "TextData"; // 第二列標題
            worksheet.Cells[2, 1].Value = "test.wav"; // 第一行的檔案名稱
            worksheet.Cells[2, 2].Value = "Test Content"; // 第一行的文本數據
            // 儲存 Excel 檔案
            package.Save();
        }

        // 創建一個測試用的 WAV 檔案
        File.WriteAllText(wavFilePath, "Dummy WAV content"); // 寫入一些虛擬內容

        // -------------------
        // Act: 呼叫被測試的方法
        // -------------------

        // 呼叫主專案中的 ProcessFolder 方法，傳入測試資料夾路徑
        Program.ProcessFolder(testFolder);

        // -------------------
        // Assert: 驗證輸出結果
        // -------------------

        // 驗證輸出資料夾是否被正確創建
        Assert.True(Directory.Exists(outputFolder));

        // 驗證 output.txt 檔案是否被正確生成
        Assert.True(File.Exists(Path.Combine(outputFolder, "output.txt")));

        // 驗證 WAV 檔案是否被複製並重新命名為 0001.wav
        Assert.True(File.Exists(Path.Combine(outputFolder, "0001.wav")));

        // -------------------
        // 清理測試環境
        // -------------------

        // 刪除測試用的資料夾與檔案，防止污染後續測試
        Directory.Delete(testFolder, true);
    }
}
