---
name: merge-listed-and-clinical-pools
description: Validate, align, compare, and consolidate China-listed and China-clinical Excel molecule or asset pools for one therapeutic area. Use when Codex must freeze a run-specific full-pool schema and Clean schema, generate high-recall cross-library candidates, execute approved keep/delete/exclude decisions, create an authoritative full baseline, optionally coordinate marketed-information supplements for clinical-source marketed rows, and derive a Working or Final Clean view without altering the source workbooks.
---

# Merge Listed And Clinical Pools

Build one auditable pool from cleaned listed and clinical sources. Treat the normal business grain as `asset or fixed component set × dosage form × mapped Disease`. Do not deduplicate by molecule name, trial number, company, target, or marketed stage alone.

Before acting, read:

- [references/workflow.md](references/workflow.md) for the end-to-end gates and decision matrix.
- [references/decision-contract.md](references/decision-contract.md) before creating or importing decisions.
- [references/output-and-supplement-contract.md](references/output-and-supplement-contract.md) before supplementing marketed information or deriving Clean.
- [references/baseline-schema.json](references/baseline-schema.json) when preparing the run config.

## Core model

- Treat the **full pool** as the single fact baseline after cross-library decisions close.
- Derive every Clean version from the latest full pool; never rebuild Clean independently from the two raw sources.
- Permit a **Working Clean** before marketed-information supplements or downstream modules close.
- Require a **Final Clean** to disclose and close supplement/module states, then bind it to the current full-pool fingerprint.
- Return every late supplement to the full pool first. Regenerate Clean afterward and recheck only modules affected by changed fields.
- Keep BI ownership, MOA overlap, VBP derivation, Giant MNC/TA-leading-company judgment, and commercial prioritization outside this Skill. Carry their approved outputs only.

## Required workflow

Follow the gates in order. A later gate may run in parallel only where this contract explicitly allows it.

1. **Collect inputs.** Obtain the TA; listed and clinical workbook/sheet paths; upstream Mapping version; run scope dates; stable source IDs or approved run-frozen fallback; output directory; full and Clean field requirements; supplement trigger and fields; downstream module list; and missing-display rule.
2. **Create a run config.** Copy `references/baseline-schema.json` outside the Skill. Replace every placeholder. Keep `schema_approved` false.
3. **Inspect read-only.** Run `inspect`. Report dimensions, headers, source-ID quality, required-field gaps, Disease blanks/multivalue cells, within-source duplicate signatures, stage/status values, formulas/cached-value risk, input hashes, and source-sheet fingerprints.
4. **Confirm the Run Schema.** Freeze `wide_schema`, `clean_schema`, optional `clean_auxiliary_schema`, logical `wide_fields`, sheet names, conditional field rules, and supplement/module settings. Do not treat 37, 40, 21, or any TA-specific count as a constant. Wait for explicit approval; then set `schema_approved` true.
5. **Build the draft.** Run `build-draft --confirmed-schema`. Verify `listed rows + clinical rows = draft rows`, the approved full-field count, static mapped values, source provenance, raw-sheet cell/formula fingerprints, input hashes, config hash, and run fingerprint.
6. **Generate candidates.** Run `build-candidates`. Treat all results as review material, never automatic deletions. Report unique clinical rows, candidate rows, one-to-many groups, initial relation types, no-candidate marketed rows, and the candidate fingerprint.
7. **Resolve decisions.** Apply `workflow.md`. Use one action per unique clinical source row. Import only a JSON decision set bound to the emitted candidate fingerprint.
8. **Preview the full baseline.** Run `finalize-full --mode preview`. Require closed nonconflicting actions, valid listed carriers, zero `HOLD`/Pending groups, matching run/candidate fingerprints, and valid row arithmetic. Pause for explicit execution approval.
9. **Generate the full baseline.** Run `finalize-full --mode generate --confirmed`. Reopen and verify the full schema, sequence, source counts, relation mappings, raw-sheet fingerprints, static values, and hidden run manifest. Do not calculate downstream module fields here.
10. **Proceed flexibly after the full baseline.** The following may run in parallel:
    - generate a Working Clean;
    - export/collect/apply marketed-information supplements;
    - run downstream ownership, MOA, and company-classification modules.
11. **Apply late information safely.** Generate a supplement template with `build-supplement-template`. Apply returned information with `apply-supplements`; match only by the locked full-pool ID, write the full pool first, never overwrite a valid value without explicit permission, retain closed statuses, and mark any existing Clean stale.
12. **Derive Clean.** Run `derive-clean --stage working` at any time after the full baseline. Run `derive-clean --stage final` only after required supplement and module states are closed. Replace any earlier Clean only in a new output workbook and bind it to the current full fingerprint.
13. **Report completion.** Provide clickable paths, input/action/full/Clean counts, full/Clean ID and order agreement, supplement/module states, stale-state resolution, raw-sheet verification, fingerprints, exceptions, and QC.

## Commands

```powershell
python .\scripts\pool_workflow.py inspect --config "C:\path\run-config.json"

python .\scripts\pool_workflow.py build-draft `
  --config "C:\path\run-config.json" `
  --output "C:\path\TA_pool_draft.xlsx" `
  --confirmed-schema

python .\scripts\pool_workflow.py build-candidates `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_draft.xlsx" `
  --output "C:\path\TA_pool_candidates.xlsx"

python .\scripts\pool_workflow.py import-decisions `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_candidates.xlsx" `
  --decisions "C:\path\decisions.json" `
  --output "C:\path\TA_pool_decided.xlsx"

python .\scripts\pool_workflow.py finalize-full `
  --mode preview `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_decided.xlsx"

python .\scripts\pool_workflow.py finalize-full `
  --mode generate --confirmed `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_decided.xlsx" `
  --output "C:\path\TA_pool_full.xlsx"

python .\scripts\pool_workflow.py build-supplement-template `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_full.xlsx" `
  --output "C:\path\TA_marketed_supplement.xlsx"

python .\scripts\pool_workflow.py apply-supplements `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_full_or_working_clean.xlsx" `
  --supplement "C:\path\TA_marketed_supplement_returned.xlsx" `
  --output "C:\path\TA_pool_supplemented.xlsx"

python .\scripts\pool_workflow.py derive-clean `
  --stage working `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_full.xlsx" `
  --output "C:\path\TA_pool_working_clean.xlsx"

python .\scripts\pool_workflow.py derive-clean `
  --stage final --confirmed-current-full `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_latest.xlsx" `
  --output "C:\path\TA_pool_final.xlsx"
```

## Hard rules

- Never overwrite an input, intermediate, or delivery.
- Never build before schema approval or use a decision set from another candidate fingerprint.
- Never use marketed stage, registration-number difference, company similarity, title similarity, target equality, or missing launch information as a standalone keep/delete rule.
- Never count candidate rows as unique clinical rows.
- Never renumber the full baseline after it is locked; use its ID for every supplement, module, and Clean join.
- Never maintain Full and Clean independently. Full wins every conflict.
- Never make marketed-information completion a prerequisite for Working Clean.
- Never label a Clean as final while required supplement/module states remain open or it is bound to an older full fingerprint.
- Never update Clean first when late information arrives.
- Never recalculate VBP or execute BI/MOA/company-classification modules here.
- Never silently overwrite nonblank full-pool values with supplement blanks or conflicting values.
- Never use a cardiovascular, kidney, or other TA fixture's row/column counts as another TA's target.

## Stop conditions

Stop when input versions are unclear; required fields or stable source IDs are unusable; source Mapping states are unresolved; Disease cells contain unresolved multiple results; source schemas or TA lists conflict; a candidate requires unresolved medical/chemical/formulation judgment; one-to-many decisions conflict; a decision fingerprint drifts; Pending/HOLD remains before full generation; supplement IDs are missing/duplicated/unknown; a conflicting supplement overwrite lacks approval; a Final Clean gate is open; Full/Clean IDs or order diverge; formulas lack cached values needed for static output; source/raw fingerprints drift; or row arithmetic/QC fails.

## Completion report

Report the input paths and fingerprints; approved full/Clean field counts; listed, clinical, draft, decision, and full counts; action counts; full sequence; supplement candidate and status counts; downstream module states; Working/Final Clean state; full/Clean ID and order check; raw-sheet/formula checks; exceptions; and output paths.
