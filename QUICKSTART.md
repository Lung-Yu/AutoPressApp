# 快速使用指南

## 立即執行

1. **開啟 PowerShell 或 命令提示字元**
2. **導航到專案目錄**:
   ```powershell
   cd d:\projects\auto_press\AutoPressApp
   ```
3. **執行程式**:
   ```powershell
   dotnet run
   ```
   或直接執行：
   ```powershell
   .\run.bat
   ```

## 程式操作

1. 點「開始記錄」→ 按你的按鍵序列 → 再按一次「開始記錄」結束
2. （可選）按「回放記錄」測試單次
3. 勾「循環」+ 按「開始」= 無限回放
4. 速度倍率：下方選擇 0.5x~2.0x（影響延遲）
5. 匯出 / 匯入：JSON 檔保存或載入按鍵序列
6. 全域開始：Ctrl+Shift+S（背景也可）
7. 緊急停止：ESC 或 Ctrl+Shift+X
8. 清除記錄：按「清除記錄」

### 間隔時間說明
- 最小間隔：0.1 秒（每秒按 10 次）
- 最大間隔：60.0 秒（每分鐘按 1 次）
- 調整精度：0.1 秒

## 建置獨立執行檔

如果要建立不需要安裝 .NET 的獨立執行檔：

```powershell
dotnet publish -c Release -r win-x64 --self-contained
```

執行檔會產生在：
`bin\Release\net8.0-windows\win-x64\publish\AutoPressApp.exe`

## 注意事項

- 程式會模擬真實按鍵輸入
- 錄製時所有按鍵都會記錄（含功能鍵）
- 目標視窗可於開始前下拉選擇
- ESC / Ctrl+Shift+X 立即停止（安全）
- Alt+Tab 不會自動終止回放，請用停止熱鍵
