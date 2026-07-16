---
name: merge-listed-and-clinical-pools
description: Validate, align, compare, and finalize China-listed and China-clinical Excel molecule or asset pools for one therapeutic area. Use when Codex must confirm a run-specific output schema, create an auditable combined draft, generate high-recall cross-library candidates, concentrate medical or business exceptions for approval, and apply closed keep/delete/exclude decisions without altering the source sheets.
---

# Merge Listed And Clinical Pools

Create one traceable final molecule pool from a cleaned China-listed pool and a cleaned China-clinical pool. Treat the business unit as `asset or fixed combination × dosage form × mapped disease`; do not deduplicate by name or registration number alone.

Before acting, read:

- [references/workflow.md](references/workflow.md) for the full workflow and decision matrix.
- [references/decision-contract.md](references/decision-contract.md) before generating, importing, or executing decisions.
- [references/baseline-schema.json](references/baseline-schema.json) when proposing the run-specific field schema.

## Required workflow

Follow the gates in order. Never jump directly from input receipt or candidate generation to finalization.

1. **Collect inputs.** Request the TA name, listed and clinical workbook paths and sheet names, the effective TA Disease mapping field/list, inclusion dates/scope, output directory, and field additions or removals requested for this run.
2. **Create a run config.** Copy `references/baseline-schema.json` outside the Skill folder, fill the input paths and sheet names, and adjust the logical source fields. Keep `schema_approved` false.
3. **Inspect read-only.** Run `inspect`. Report sheet dimensions, headers, stable IDs, required-field gaps, Disease blanks/multivalue cells, within-source duplicate signatures, status values, formulas, and external-link risk.
4. **Propose and confirm Run Schema.** Start from the common 21-field baseline, but allow fields to be kept, removed, renamed, added, or reordered when a TA has an approved exception. Confirm that listed `持证商` and clinical `申办方` both map into `License Holder`, while their source-specific meanings remain visible on the raw sheets. Confirm that `是否VBP` is an upstream value to map, not a value calculated here. Report the final field count `C`, both source mappings, missing-value behavior, and audit-only fields. Wait for explicit approval, then freeze the approved field list in the run config and set `schema_approved` true.
5. **Build the combined draft.** Run `build-draft` only with `--confirmed-schema`. Verify `listed rows + clinical rows = draft rows`, the approved `C` fields, source counts tracked in hidden workflow metadata, sequential temporary IDs, and preserved raw sheets. Do not add a visible `数据来源` column unless the approved Run Schema explicitly requests one.
6. **Generate candidates.** Run `build-candidates`. Treat its result as high-recall review material, never as an automatic deletion list. Report unique clinical rows, candidate rows, one-to-many cases, initial relation types, and no-candidate marketed records.
7. **Resolve decisions.** Apply the rules in `workflow.md`. Fill the structured decision fields, group unresolved items for one concentrated confirmation, and perform targeted professional-database checks only when a small missing fact can close a decision. Use `import-decisions` when decisions are stored in JSON.
8. **Preview finalization.** Run `finalize --mode preview`. Require one nonconflicting action per unique clinical source row, valid carrier IDs for listed-covered deletions, closed decision statuses, and zero `HOLD`/Pending items. Show the expected row arithmetic and pause for explicit execution approval.
9. **Generate and verify.** Run `finalize --mode generate --confirmed`. Reopen the result and verify the final sequence, source counts, action counts, relation mappings, raw-sheet fingerprints, exact Run Schema, no trailing formatted blank rows, and no external formulas in the final pool.
10. **Report completion.** Provide a clickable output path, input/action/final counts, Pending count, raw-sheet preservation, and QC result.

## Run the script

The script requires Python and `openpyxl`. It refuses to overwrite an existing output.

```powershell
python .\scripts\pool_workflow.py inspect `
  --config "C:\path\run-config.json"

python .\scripts\pool_workflow.py build-draft `
  --config "C:\path\run-config.json" `
  --output "C:\path\TA_pool_draft.xlsx" `
  --confirmed-schema

python .\scripts\pool_workflow.py build-candidates `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_draft.xlsx" `
  --output "C:\path\TA_pool_candidates.xlsx"

python .\scripts\pool_workflow.py import-decisions `
  --workbook "C:\path\TA_pool_candidates.xlsx" `
  --decisions "C:\path\decisions.json" `
  --output "C:\path\TA_pool_decided.xlsx"

python .\scripts\pool_workflow.py finalize `
  --mode preview `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_decided.xlsx"

python .\scripts\pool_workflow.py finalize `
  --mode generate `
  --confirmed `
  --config "C:\path\run-config.json" `
  --workbook "C:\path\TA_pool_decided.xlsx" `
  --output "C:\path\TA_pool_final.xlsx"
```

## Hard rules

- Never overwrite an input, intermediate, or approved delivery.
- Never build a draft before Run Schema approval.
- Never delete because a name is similar or a China stage says `已上市`.
- Never treat candidate rows as unique clinical rows; one clinical row may have several listed candidates.
- Never split `License Holder` back into separate holder and sponsor columns in the common 21-field output; preserve those source distinctions on the raw sheets.
- Never recalculate VBP in this Skill. Map the already-prepared upstream `是否VBP` value and stop if the required run mapping is unclear.
- Never merge a different Disease without checking the indication evidence.
- Keep new Disease, new dosage form, new research molecule, fixed combination, and materially distinct chemical-form candidates unless a confirmed rule says otherwise.
- Exclude diagnostic/non-therapeutic assets and non-fixed regimens only after the asset nature is established.
- Do not force clinical sponsor or trial details into a listed row. Keep the listed row and preserve clinical evidence in the relation and raw sheets.
- Preserve the complete main experimental-drug name when a new molecule is combined with background therapy.
- Do not finalize with blank actions, conflicting actions, open statuses, `HOLD`, or an invalid listed carrier.
- Keep the copied raw sheets unchanged. Build the final pool from static values and physically omit removed rows.
- Treat the cardiovascular regression case as a test fixture only. Never use its counts, rates, or decisions as another TA's target or default.

## Stop conditions

Stop and request direction when input versions are unclear; required fields or stable IDs are missing; a Disease cell contains unresolved multiple results; source libraries still contain unapproved within-library duplicates; the TA Disease lists differ; Run Schema is not approved; a candidate requires medical, chemical-form, fixed-combination, or marketed-gap judgment; decisions conflict across one-to-many candidates; Pending remains; or row arithmetic/raw-sheet verification fails.

## Completion report

Report:

- listed, clinical, and draft input rows;
- approved Run Schema field count;
- unique clinical rows, candidate rows, and one-to-many count;
- listed-covered deletions and each exclusion action count;
- retained clinical rows and final source composition;
- final sequence range and Pending count;
- raw-sheet and row-arithmetic verification;
- exceptions and the output path.
