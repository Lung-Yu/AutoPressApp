using System.Threading.Tasks;

namespace AutoPressApp.Steps
{
    // Base class for steps; polymorphic JSON handled by custom StepJsonConverter
    public abstract class Step
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public bool Enabled { get; set; } = true;
        public abstract Task ExecuteAsync(StepContext ctx);
    }

    public class StepContext
    {
        public Services.WindowService Window { get; } = new Services.WindowService();
        public Services.InputService Input { get; } = new Services.InputService();
        public Services.LogService Log { get; } = new Services.LogService();
        public System.Threading.CancellationToken CancellationToken { get; set; }
    }
}
