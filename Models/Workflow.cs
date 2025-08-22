using System;
using System.Collections.Generic;
using AutoPressApp.Steps;

namespace AutoPressApp.Models
{
    // Workflow definition (serializable)
    public class Workflow
    {
        public string SchemaVersion { get; set; } = "1.0";
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = "Untitled Workflow";
        public List<Step> Steps { get; set; } = new();

        // Loop options
        public bool LoopEnabled { get; set; }
        public int? LoopCount { get; set; }
        public int LoopIntervalMs { get; set; } = 1000;
    }
}
