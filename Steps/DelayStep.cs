using System;
using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    public class DelayStep : Step
    {
        public int Ms { get; set; }
        public override Task ExecuteAsync(StepContext ctx)
        {
            double multiplier = ctx.DelayMultiplier <= 0 ? 1.0 : ctx.DelayMultiplier;
            int effective = Ms;
            if (multiplier != 1.0 && Ms > 0)
            {
                // Faster when multiplier > 1.0 => shorter delay; slower when <1.0 => longer delay
                effective = (int)Math.Max(1, Math.Round(Ms / multiplier));
            }
            if (effective != Ms)
                ctx.Log.Info($"[Delay] {Ms}ms -> {effective}ms (x{multiplier:0.###})");
            else
                ctx.Log.Info($"[Delay] {Ms}ms");
            return Task.Delay(effective, ctx.CancellationToken);
        }
    }
}
