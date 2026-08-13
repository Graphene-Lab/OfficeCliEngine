# VENDOR — OfficeCliEngine

This project vendors the engine of **OfficeCLI** (upstream repository
`iOfficeAI/OfficeCLI`), so that the in-process document manipulation layer
(Core + Handlers) can be referenced from AIOrchestrator without the CLI shell.

## Current sync

| Field | Value |
|---|---|
| Upstream repository | https://github.com/iOfficeAI/OfficeCLI |
| Upstream version | v1.0.144 |
| Upstream commit | 1ced45e900782c5083ed550ddf328ee974e425e7 |
| Sync date | 2026-08-13 |
| Upstream license | Apache-2.0 (see NOTICE.md) |

> The sync pipeline updates the *version*, *commit* and *date* rows above
> automatically — never edit them by hand.

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
3. **Version traceability**: the *Current sync* table records upstream version + commit.
   The value of a sync is that the next upstream release shows as a diff between two
   recorded versions.

## Sync source: the stable release, not the repository

The sync downloads the **"Source code (zip)" of the upstream GitHub release**
(`https://github.com/iOfficeAI/OfficeCLI/releases`), never the repository default
branch. This keeps the vendor on the **stable release** — no in-progress work, no
unreleased commits. The commit recorded in the table is the release's
`target_commitish` (the exact commit the tag points to).

## Sync procedure

### One-shot (mechanical)

```
.\sync-from-upstream.ps1                 # latest release
.\sync-from-upstream.ps1 -Tag v1.0.144   # pin a specific release
```

It downloads the release zip, vendors `src/officecli` / `skills` / `schemas` (pruning
files gone from the release), verifies **byte-identical parity** (SHA-256 of every
vendored file), updates the version/commit/date rows, and prints `git diff --stat`.

### Full update (recommended)

```
.\update-officecli.ps1                   # sync + gap analysis + builds + tests
.\update-officecli.ps1 -Tag v1.0.144     # pin a specific release
```

`update-officecli.ps1` is the semi-automatic updater: after the sync it

1. **gap analysis** — parses the vendored CLI commands and view modes and compares
   them against the `OfficeTool` methods: every command the shipped documentation can
   mention must exist as a method (coherence principle, OfficePorting.md §3); new
   commands land in `sync-gap-report.md` with the method to add;
2. **embedded-resource parity** — compares the upstream csproj embedded resources
   against `OfficeCliEngine.csproj` (a new embedded file missing here breaks the
   engine at runtime);
3. **builds** the engine (Release) and AIOrchestrator (surfaces engine API drift and
   OfficeTool breakage);
4. **runs** `OfficeTool.Tests` (the deterministic harness, docx/xlsx/pptx);
5. **packs** (`SkipNuGetPush`) to verify the NuGet pipeline.

Nothing is committed or pushed. After a green run:

1. Read `sync-gap-report.md` and the `git diff` (sync already verified byte-identity).
2. Add OfficeTool methods (+ harness tests) for any new command the report lists.
3. Bump the version in `NOTICE.md` to the new release tag.
4. Update OfficePorting.md §11 status rows if the surface changed.
5. Commit and push to `Graphene-Lab/OfficeCliEngine` — the CI workflow
   (`.github/workflows/publish.yml`) packs and publishes the NuGet package, which
   AIOrchestrator picks up via `PackageReference 1.*` where the project is absent.

### Offline / troubleshooting

- `.\sync-from-upstream.ps1 -UpstreamPath <dir>` syncs from a local source tree
  (release zip extract or git checkout) without downloading.
- The parity check must report *"OK — N files byte-identical"*. A `DIFFERS`/`MISSING`
  line means a vendored file was hand-edited or the sync was interrupted — investigate
  before committing.

## Verification after a sync

- `dotnet build OfficeCliEngine.csproj -c Release` — 0 errors.
- `dotnet run --project ..\AIOrchestrator\OfficeTool.Tests` — ALL TESTS PASSED.
- The parity line in the sync output — byte-identical.
