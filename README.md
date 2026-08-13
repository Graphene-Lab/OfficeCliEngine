# OfficeCliEngine

In-process Office document manipulation engine (DOCX / XLSX / PPTX), vendored from
the [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) project (Apache-2.0).

It exposes the OfficeCLI **engine** (`OfficeCli.Core` + `OfficeCli.Handlers`) as a
library: deterministic JSON output, path-based addressing (`/slide[1]/shape[2]`),
schema-driven property validation, template merge, dump/batch round-trip and HTML
rendering — without the CLI shell.

This package is the engine behind the `OfficeTool` agent tool
(`AIOrchestrator\API\OfficeTool.cs`). See `VENDOR.md` for the upstream sync
procedure and `NOTICE.md` for attribution.

## Quick start

```csharp
using OfficeCli.Handlers;               // DocumentHandlerFactory
using OfficeCli.Core;                   // IDocumentHandler, CliException

using var handler = DocumentHandlerFactory.Open("/tmp/report.docx", editable: true);
var outline = handler.ViewOutline();
var json    = handler.Get("/body/p[1]", depth: 1);
handler.Set("/body/p[1]", new Dictionary<string, object?> { ["bold"] = "true" });
handler.Save();
```

Errors are surfaced as `CliException` with structured `Code`, `Suggestion`,
`ValidValues` and `Help` fields — the same self-healing contract the CLI uses.

## Vendored content

- `src/officecli/` — upstream project tree, byte-identical, minus `Program.cs`
- `skills/` — agent skills (embedded `skills/…` resources)
- `schemas/help/` — help schemas (embedded `schemas/help/…` resources)

Do **not** edit files under `src/officecli/`: they are upstream-copied. See
`VENDOR.md`.

## Updating from upstream

The vendor tracks the **stable release** (the "Source code (zip)" asset of
`https://github.com/iOfficeAI/OfficeCLI/releases`), never the repository branch.

```
.\update-officecli.ps1                 # latest release: sync + gap analysis + builds + tests
.\update-officecli.ps1 -Tag v1.0.144   # pin a specific release
```

The updater vendors the release byte-identical, reports any CLI command that has no
`OfficeTool` method yet (`sync-gap-report.md`), verifies embedded-resource parity,
builds engine + AIOrchestrator and runs the `OfficeTool.Tests` harness. Nothing is
committed or pushed — review, fix reported gaps, then commit (CI publishes NuGet).
Full procedure: `VENDOR.md`.
