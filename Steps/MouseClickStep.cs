using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    public enum CoordMode { Screen, Window }
    public class MouseClickStep : Step
    {
        public int X { get; set; }
        public int Y { get; set; }
        public CoordMode Mode { get; set; } = CoordMode.Screen;
        public string Button { get; set; } = "left"; // left/right
        public int AfterDelayMs { get; set; } = 200;

        public override async Task ExecuteAsync(StepContext ctx)
        {
            ctx.Log.Info($"[MouseClick] {Button} @ ({X},{Y}) mode={Mode}");
            await Task.Run(() => ctx.Input.Click(X, Y, Button, Mode), ctx.CancellationToken);
            if (AfterDelayMs > 0) await Task.Delay(AfterDelayMs, ctx.CancellationToken);
        }
    }
}
