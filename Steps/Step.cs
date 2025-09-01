using System; // Needed for IntPtr
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
    // Playback speed multiplier (1.0 = normal). >1.0 means faster (delays shortened), <1.0 slower.
    public double DelayMultiplier { get; set; } = 1.0;
    // 輸入派送模式：前景模擬 (SendInput) 或 背景視窗訊息 (PostMessage)
    public InputDispatchMode DispatchMode { get; set; } = InputDispatchMode.ForegroundSendInput;
    // 目標視窗 (背景模式時使用)，Foreground 模式若為 IntPtr.Zero 則使用目前前景
    public IntPtr TargetWindowHandle { get; set; } = IntPtr.Zero;
    }

    public enum InputDispatchMode
    {
        ForegroundSendInput = 0,
        BackgroundPostMessage = 1
    }
}
