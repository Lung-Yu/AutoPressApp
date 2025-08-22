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
        private readonly LogService _log;
        public WorkflowRunner(LogService log) => _log = log;

        public async Task RunAsync(Workflow wf, CancellationToken ct)
        {
            var ctx = new StepContext { CancellationToken = ct };
            ctx.Log.OnLog += m => _log.Info(m);

            int loops = wf.LoopEnabled ? (wf.LoopCount ?? int.MaxValue) : 1;
            for (int i = 0; i < loops && !ct.IsCancellationRequested; i++)
            {
                _log.Info($"[Workflow] Run '{wf.Name}' loop {i + 1}/{loops}");
                foreach (var step in wf.Steps)
                {
                    if (!step.Enabled) continue;
                    ct.ThrowIfCancellationRequested();
                    await step.ExecuteAsync(ctx);
                }
                if (wf.LoopEnabled && wf.LoopIntervalMs > 0)
                    await Task.Delay(wf.LoopIntervalMs, ct);
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
