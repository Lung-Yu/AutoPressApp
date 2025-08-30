using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    // 精細鍵盤事件序列，可表達同時鍵 / 長按 (以 Down/Up 事件與相對延遲組成)
    /// <summary>
    /// 精細鍵盤事件序列，可表達同時鍵 / 長按 (以 Down/Up 事件與相對延遲組成)
    /// 範例：
    /// Events = [
    ///   { Key="Left", Down=true, DelayMsBefore=0 },
    ///   { Key="Z", Down=true, DelayMsBefore=20 },
    ///   { Key="Z", Down=false, DelayMsBefore=60 },
    ///   { Key="Left", Down=false, DelayMsBefore=100 }
    /// ]
    /// 表示：先按住 Left，20ms 後按 Z，60ms 後放開 Z，100ms 後放開 Left。
    /// </summary>
    public class KeySequenceStep : Step
    {
        public List<KeyEventItem> Events { get; set; } = new();
        // 方便預覽：顯示前幾個事件摘要
        public override async Task ExecuteAsync(StepContext ctx)
        {
            ctx.Log.Info($"[KeySequence] events={Events.Count}");
            int i = 0;
            foreach (var ev in Events)
            {
                if (ctx.CancellationToken.IsCancellationRequested) break;
                if (ev.DelayMsBefore > 0)
                    await Task.Delay(ev.DelayMsBefore, ctx.CancellationToken);
                ctx.Input.SendKeyEvent(ev.Key, ev.Down);
                if (++i % 20 == 0) // 週期性快照資訊
                    ctx.Log.Info($"[KeySequence] progressed {i}/{Events.Count}");
            }
        }
    }

    public class KeyEventItem
    {
        public string Key { get; set; } = string.Empty; // Display 名稱 (與 InputService.MapKeyStringToVk 相容)
        public bool Down { get; set; } // true=keydown false=keyup
        public int DelayMsBefore { get; set; } // 與前一事件的間隔
    }
}
