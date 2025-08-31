using System;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using AutoPressApp.Models;
using AutoPressApp.Steps;

namespace AutoPressApp.Services
{
    public class WorkflowRunner
    {
        public event Action<int, Step>? OnStepExecuting;
    public event Action<int,int,bool>? OnLoopStarting; // currentLoop, totalLoops (or int.MaxValue), isInfinite
    public event Action<int>? OnIntervalTick; // remaining milliseconds before next loop starts
        private readonly LogService _log;
        public double PlaybackSpeed { get; set; } = 1.0; // 1.0=normal, 2.0=2x speed (half delays)

        public WorkflowRunner(LogService log, double playbackSpeed = 1.0)
        {
            _log = log;
            if (playbackSpeed <= 0) playbackSpeed = 1.0;
            PlaybackSpeed = playbackSpeed;
        }

        public async Task RunAsync(Workflow wf, CancellationToken ct)
        {
            var ctx = new StepContext { CancellationToken = ct, DelayMultiplier = PlaybackSpeed };
            ctx.Log.OnLog += m => _log.Info(m);

            bool infinite = wf.LoopEnabled && wf.LoopCount == null;
            int loops = wf.LoopEnabled ? (wf.LoopCount ?? int.MaxValue) : 1;
            for (int i = 0; i < loops && !ct.IsCancellationRequested; i++)
            {
                OnLoopStarting?.Invoke(i + 1, loops, infinite);
                _log.Info($"[Workflow] Run '{wf.Name}' loop {i + 1}/{(infinite ? -1 : loops)}");
                for (int stepIdx = 0; stepIdx < wf.Steps.Count; stepIdx++)
                {
                    var step = wf.Steps[stepIdx];
                    if (!step.Enabled) continue;
                    ct.ThrowIfCancellationRequested();
                    OnStepExecuting?.Invoke(stepIdx, step);
                    await step.ExecuteAsync(ctx);
                }
                // Interval countdown
                if (wf.LoopEnabled && wf.LoopIntervalMs > 0 && (infinite || i + 1 < loops))
                {
                    int remaining = wf.LoopIntervalMs;
                    const int slice = 200; // update every 200ms
                    while (remaining > 0)
                    {
                        ct.ThrowIfCancellationRequested();
                        OnIntervalTick?.Invoke(remaining);
                        int delay = remaining < slice ? remaining : slice;
                        await Task.Delay(delay, ct);
                        remaining -= delay;
                    }
                    OnIntervalTick?.Invoke(0);
                }
            }
        }

        public static Workflow LoadFromJson(string json)
        {
            var opts = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNameCaseInsensitive = true
            };
            opts.Converters.Add(new StepJsonConverter());
            return JsonSerializer.Deserialize<Workflow>(json, opts) ?? new Workflow();
        }

        public static string SaveToJson(Workflow wf)
        {
            var opts = new JsonSerializerOptions { WriteIndented = true };
            opts.Converters.Add(new StepJsonConverter());
            return JsonSerializer.Serialize(wf, opts);
        }
    }
}
