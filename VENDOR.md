# VENDOR — OfficeCliEngine

This project vendors the engine of **OfficeCLI** (upstream repository
`iOfficeAI/OfficeCLI`), so that the in-process document manipulation layer
(Core + Handlers) can be referenced from AIOrchestrator without the CLI shell.

## Current sync

| Field | Value |
|---|---|
| Upstream repository | https://github.com/iOfficeAI/OfficeCLI |
| Upstream version | v1.0.143 |
| Upstream commit | `(fill on next sync — git rev-parse HEAD in OfficeCLI checkout)` |
| Sync date | 2026-08-13 |
| Upstream license | Apache-2.0 (see NOTICE.md) |

## Vendored layout (mirrors upstream `OfficeCLI` repo root)

| Here | Upstream | Content |
|---|---|---|
| `src/officecli/` | `src/officecli/` | entire project tree, byte-identical, **minus `Program.cs`** (the CLI entry point; the only allowed DELETE) |
| `skills/` | `skills/` | agent skills (embedded as `skills/…` resources) |
| `schemas/help/` | `schemas/help/` | help schemas (embedded as `schemas/help/…` resources) |

The root `OfficeCliEngine.csproj` is **ours** (packaging + auto-version): it compiles
`src/officecli/**/*.cs` and embeds the resources with the same logical names upstream
uses, so the vendored code resolves them unchanged.

## Fidelity rules

1. **Zero modifications to vendored files.** No reformat, rename, "improvement" or
   inline fix: any divergence breaks the diff against upstream.
   - A fix/feature we need that upstream lacks → propose it **upstream first**, then
     bring it here with the next sync.
   - An unavoidable local workaround → isolate it in `patches/`, applied by the sync
     script, never mixed into vendored files.
2. **The only allowed operation on vendored files is DELETE** (`Program.cs`, and
   whatever `sync-exclude.txt` lists). A deleted file shows in `git diff`; a modified
   one does not.
3. **Version traceability**: this file records upstream version + commit. The value
   of a sync is that the next upstream release shows as a diff between two recorded
   versions.

## Sync procedure

Run `sync-from-upstream.ps1` from this directory. It:

1. copies `src/officecli` (minus `bin`, `obj`, and `sync-exclude.txt` entries) from
   the local upstream checkout (or the `-UpstreamPath` argument);
2. copies `skills/` and `schemas/help/`;
3. prints `git diff --stat` as the change report.

Then update the version/commit rows above, build, run the test suite
(AIOrchestrator\OfficeTool.Tests) and re-check `--output-schema-crc` parity with the
upstream CLI binary (no help-schema drift).
