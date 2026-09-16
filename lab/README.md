# lab

Scripts that build and verify the test corpus in a **test tenant**. They write.
The analyzer at the repository root never does.

They are kept in the language they were written in. Translating them without
re-running them against a tenant would mean shipping untested changes to the
only thing that establishes the analyzer is correct.

## Corpus

Nine objects, covering the six row states plus the `NotResolved` counter.

| Case | What it establishes |
|------|---------------------|
| A | Declared and not granted. `UnderCoverage`. |
| B | Granted after the declaration was removed. Microsoft documents that removing a permission does not revoke the grant. |
| C | A principal that holds grants and has no local registration. Only the manifest can carry its declared state. |
| D | Declared and granted. `CorrectCoverage`. |
| E | Negative control. Nothing declared, nothing granted, no rows. Its assertion is on `conclusion.result`, not on a row state. |
| F | A negative expectation. `requiredResourceAccess` cannot express one, so only the manifest can. `CorrectExclusion`. |
| G | 25 assignments on one principal, to exercise the documented 20-item `$expand` limit. Without an individual re-read it produces 20 correct rows and 5 silent absences. |

## Four passes

The corpus is evaluated four times against three manifests. What proves
something is not the conformity of any one pass but the differences between
them, and each difference isolates a single variable.

1. **No manifest.** Grants report as `Observed`, completeness `Absent`.
2. **Complete manifest.** `B`, `C` and `G` become `OverCoverage`, `F` appears.
3. **Partial manifest.** `G` is removed from the manifest, `B`'s entry is set to
   `complete: false`. Both become `NotInManifest` while `C` stays `OverCoverage`:
   exhaustiveness is read per principal, not only at manifest level.
4. **Unresolved manifest.** The complete manifest plus one entry naming a
   principal that does not exist. `C` falls to `NotInManifest` although its own
   entry did not change: an unresolved entry degrades the whole tenant, where
   `complete: false` degrades one principal.

## Scripts

```
agd-lab-fabrique.ps1     builds cases A, B, D, E        writes
agd-lab-cas-c-v2.ps1     builds case C                  writes
agd-lab-corpus.ps1       builds case F, the three manifests and the registry
agd-lab-analyse.ps1      the analyzer plus the registry assertion  read-only
agd-lab-sonde-expand.ps1 measures the cost and completeness of $expand  read-only
```

`agd-lab-analyse.ps1` is the root analyzer with the registry confrontation kept
in. That confrontation is the harness, which is why it is not in the shipped
tool. The registry's expected states are written by hand and never derived from
a run, otherwise the corpus would confirm whatever the analyzer happens to do.
