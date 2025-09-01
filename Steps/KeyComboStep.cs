using System;
using System.Linq;
using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    public class KeyComboStep : Step
    {
        // Example: ["Ctrl","Shift","S"]
        public string[] Keys { get; set; } = Array.Empty<string>();
        public int AfterDelayMs { get; set; } = 150;
    // 新增：送出前的預延遲 (讓切換視窗後有時間準備)
    public int PreDelayMs { get; set; } = 40;
    // 新增：主鍵按住時間 (某些遊戲需要 >0ms 才辨識，預設 50ms)
    public int HoldMs { get; set; } = 50;

        public override async System.Threading.Tasks.Task ExecuteAsync(StepContext ctx)
        {
            string display = string.Join("+", Keys);
            ctx.Log.Info($"[KeyCombo] {display} (pre={PreDelayMs}ms hold={HoldMs}ms)");
            await System.Threading.Tasks.Task.Run(() => ctx.Input.SendKeyCombo(Keys, HoldMs, PreDelayMs, ctx.TargetWindowHandle, ctx.DispatchMode), ctx.CancellationToken);
            if (AfterDelayMs > 0)
                await System.Threading.Tasks.Task.Delay(AfterDelayMs, ctx.CancellationToken);
        }
    }
}
