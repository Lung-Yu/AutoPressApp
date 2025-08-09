# AutoPressApp - 自動按鍵程式

一個使用 .NET Windows Forms 開發的自動按鍵程式，可以定時發送指定的按鍵輸入。

## 功能特色

- 🎯 支援多種按鍵選擇（F1-F5、Space、Enter、A-E等）
- ⏱️ 可調整按鍵間隔時間（0.1-60秒，精度0.1秒）
- 🔄 一鍵啟動/停止功能
- 📊 即時狀態顯示
- 🖥️ 簡潔的中文使用者介面

## 專案結構

```
AutoPressApp/
├── AutoPressApp.csproj          # 專案檔案
├── Program.cs                   # 程式進入點
├── Form1.cs                     # 主要邏輯程式碼
├── Form1.Designer.cs            # UI設計程式碼
├── README.md                    # 專案說明文件
└── obj/                         # 編譯暫存資料夾
```

## 系統需求

- Windows 作業系統
- .NET 6.0 或更高版本
- Visual Studio 2022 或 VS Code with C# extension

## 建立專案指令

```powershell
# 1. 建立專案資料夾
mkdir auto_press
cd auto_press

# 2. 建立 Windows Forms 專案
dotnet new winforms -n AutoPressApp

# 3. 進入專案目錄
cd AutoPressApp

# 4. 建置專案
dotnet build

# 5. 執行程式
dotnet run
```

## 使用方法

### 1. 啟動程式
```powershell
dotnet run
```

### 2. 設定參數
- **選擇按鍵**: 從下拉選單選擇要自動按的按鍵
- **設定間隔**: 調整按鍵間隔時間（0.1-60.0秒，每次調整0.1秒）

### 3. 控制程式
- 點擊「開始」按鍵開始自動按鍵
- 點擊「停止」按鍵停止自動按鍵
- 狀態列會顯示目前執行狀態

## 技術實作

### 核心技術
- **Windows Forms**: 使用者介面框架
- **System.Timers**: 定時器控制
- **Windows API**: `user32.dll` 的 `keybd_event` 函式進行按鍵模擬

### 關鍵程式碼

```csharp
// Windows API 宣告
[DllImport("user32.dll")]
static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

// 發送按鍵
private void SendKeyPress(Keys key)
{
    byte vkCode = (byte)key;
    keybd_event(vkCode, 0, 0, UIntPtr.Zero);           // 按下
    keybd_event(vkCode, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // 放開
}
```

## 支援的按鍵

| 按鍵類型 | 支援按鍵 |
|---------|---------|
| 功能鍵   | F1, F2, F3, F4, F5 |
| 特殊鍵   | Space, Enter |
| 字母鍵   | A, B, C, D, E |

## 編譯和發佈

### 開發模式執行
```powershell
dotnet run
```

### 發佈為可執行檔
```powershell
# 發佈為自包含執行檔
dotnet publish -c Release -r win-x64 --self-contained

# 發佈為框架相依執行檔
dotnet publish -c Release -r win-x64 --no-self-contained
```

## 注意事項

⚠️ **重要提醒**:
- 本程式會模擬按鍵輸入，請在合適的環境下使用
- 某些應用程式可能會偵測到自動化操作
- 建議在測試環境中先試用
- 使用時請確保不會干擾其他重要工作

## 疑難排解

### 常見問題

**Q: 程式無法執行？**
A: 請確認已安裝 .NET 6.0 或更高版本

**Q: 按鍵沒有作用？**
A: 請確認目標應用程式是否為焦點視窗

**Q: 編譯失敗？**
A: 檢查專案檔案是否完整，或重新執行 `dotnet restore`

## 授權

此專案僅供學習和個人使用。

## 更新歷史

- v1.0.0 - 初始版本
  - 基本按鍵模擬功能
  - 中文使用者介面
  - 可調整間隔時間
