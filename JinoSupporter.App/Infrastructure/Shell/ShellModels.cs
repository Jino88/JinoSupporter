using System;
using System.Collections.Generic;

namespace JinoSupporter.App.Infrastructure.Shell;

public sealed record ShellActionDefinition(string Id, string Label, bool Primary = false);

public sealed record ShellSummaryDefinition(string Label, string Value, string Body, bool Accent = false);

public sealed record ShellSectionDefinition(
    string Id,
    string Label,
    string Eyebrow,
    string Title,
    string Description,
    IReadOnlyList<ShellSummaryDefinition> Summaries,
    IReadOnlyList<string> HeroTargets,
    IReadOnlyList<string> ModuleTargets);

public sealed record ShellModuleDefinition(
    string Target,
    string SectionId,
    string Title,
    string Summary,
    string Tag,
    string Detail,
    IReadOnlyList<string> Notes,
    Func<System.Windows.FrameworkElement> CreateContent,
    IWorkbenchModuleBackend? Backend = null);

public sealed record ShellMetric(string Label, string Value, string Body = "", bool Accent = false);
