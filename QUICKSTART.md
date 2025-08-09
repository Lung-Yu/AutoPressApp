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

1. **選擇按鍵**: 從下拉選單選擇要自動按的按鍵（預設為 F1）
2. **設定間隔**: 調整按鍵間隔時間，範圍 0.1-60.0 秒（預設為 1.0 秒）
3. **開始/停止**: 點擊按鈕控制自動按鍵
4. **查看狀態**: 底部狀態列顯示目前執行狀況

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
`bin\Release\net5.0-windows\win-x64\publish\AutoPressApp.exe`

## 注意事項

- 程式會模擬真實按鍵輸入
- 請確保目標視窗具有焦點
- 不建議在重要工作中使用
- 按 Alt+Tab 切換到其他視窗可以停止對當前視窗的影響
