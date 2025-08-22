using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    public class DelayStep : Step
    {
        public int Ms { get; set; }
        public override Task ExecuteAsync(StepContext ctx)
        {
            ctx.Log.Info($"[Delay] {Ms}ms");
            return Task.Delay(Ms, ctx.CancellationToken);
        }
    }
}
