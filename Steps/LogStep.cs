using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    public class LogStep : Step
    {
        public string Message { get; set; } = string.Empty;
        public override Task ExecuteAsync(StepContext ctx)
        {
            ctx.Log.Info($"[Log] {Message}");
            return Task.CompletedTask;
        }
    }
}
