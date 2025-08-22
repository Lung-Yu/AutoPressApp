using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    public class FocusWindowStep : Step
    {
        public string? TitleContains { get; set; }
        public string? ProcessName { get; set; }
        public int TimeoutMs { get; set; } = 5000;

        public override async Task ExecuteAsync(StepContext ctx)
        {
            ctx.Log.Info($"[FocusWindow] title~='{TitleContains}', process='{ProcessName}'");
            await Task.Run(() => ctx.Window.FocusWindow(TitleContains, ProcessName, TimeoutMs), ctx.CancellationToken);
        }
    }
}
