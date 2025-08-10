# AutoPressApp - 自動按鍵程式

一個使用 .NET Windows Forms 開發的自動按鍵程式，可以定時發送指定的按鍵輸入。

## 功能特色

- 🎯 支援多種按鍵選擇（F1-F5、Space、Enter、A-E等）
- 🎬 按鍵記錄功能（含延遲）
- � 序列回放 + 循環播放
- � ESC / Ctrl+Shift+X 全域緊急停止
- ▶️ Ctrl+Shift+S 全域開始回放
- 🪟 視窗目標選擇
- � 狀態與操作說明即時顯示
- ⚙️ (內部) 改用 SendInput 提升穩定性

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
- .NET 5.0 (TargetFramework net5.0-windows)
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

### 2. 記錄與回放工作流程
1. 點「開始記錄」→ 在任何視窗輸入按鍵 → 再按一次「開始記錄」結束
2. 可選擇「回放記錄」做單次測試
3. 勾選「循環」+ 按「開始」→ 無限循環回放
4. 按 ESC / Ctrl+Shift+X 隨時停止
5. Ctrl+Shift+S 亦可全域啟動（程式需在執行中）

### 3. 速度倍率
- 下拉選單 0.5x / 1.0x / 1.5x / 2.0x（倍率越大播放越快）

### 4. 控制程式
- 「開始」→ 回放序列（有循環則無限）
- 「停止」/ ESC / Ctrl+Shift+X → 停止
- 「回放記錄」→ 單次回放不影響主開始鍵

## 技術實作

### 核心技術
- **Windows Forms**: 使用者介面框架
- **System.Timers**: 定時器控制
- **Windows API**: `user32.dll` 的 `keybd_event` 函式進行按鍵模擬

### 關鍵程式碼
### 3. 按鍵記錄與回放區塊
```csharp
// Windows API 宣告
[DllImport("user32.dll")]
static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

// 發送按鍵
private void SendKeyPress(Keys key)
{
    byte vkCode = (byte)key;
### 4. 開始/停止按鈕（序列回放模式）
- 「開始」：若已有記錄 → 依原延遲回放（勾循環則無限）
- 「停止」：停止回放（與循環）
- 若無記錄按鈕提示需先錄製

### 5. 熱鍵 / 系統控制
| 功能 | 操作 |
|------|------|
| 緊急停止 | ESC / Ctrl+Shift+X |
| 全域開始 | Ctrl+Shift+S |
| 匯出序列 | 匯出 (JSON) |
| 匯入序列 | 匯入 (JSON) |
| 開始錄製 | 點一次「開始記錄」 |
| 停止錄製 | 錄製中再點一次「開始記錄」 |
| 單次回放 | 點「回放記錄」 |
| 循環回放 | 勾選「循環」後點「開始」 |
| 清除記錄 | 點「清除記錄」 |

### 6. 狀態列訊息
- 顯示：記錄中 / 回放第 N 鍵 / 循環重新開始 / 停止 / 目標視窗等
- 錄製第一鍵延遲顯示為 0ms

### 7. 目標視窗選擇
- 下拉選單可選特定程式視窗
- 「重新整理」可更新視窗清單
- 未選擇或選第一項 = 使用目前焦點視窗

### 8. 內建輔助說明
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
