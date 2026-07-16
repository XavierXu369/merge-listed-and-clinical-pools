# Listed and Clinical Pool Workflow

## Contents

1. Purpose and boundaries
2. Input and schema gates
3. Draft and candidate generation
4. Decision order and matrix
5. Concentrated confirmation
6. Finalization and QC

## 1. Purpose and boundaries

Use this workflow after both source libraries have completed their own upstream cleaning and Disease mapping. The desired business unit is normally:

```text
asset or fixed component set × dosage form × mapped Disease
```

The workflow aligns two unlike schemas, creates a combined draft, generates cross-library comparison candidates, closes exceptions, and executes structured decisions with an audit trail.

It does not perform BI-product matching, MOA deduplication, VBP reduction, MNC tagging, commercial prioritization, or a wholesale remapping of the TA Disease taxonomy.

Required upstream state:

- The listed pool is already split/mapped and consolidated within the listed source.
- The clinical pool is already cleaned to an approved source-row grain.
- Each Disease cell contains one resolved result or an explicit review state.
- Both pools use the same TA Disease list/version.
- Every source row has a stable ID; generate a technical `Sheet!R<row>` ID only when no reliable clinical ID exists.

## 2. Input and schema gates

### Read-only preflight

Check before writing:

- files, sheets, headers, hidden/trailing cells, formulas, and external links;
- blank/duplicate stable IDs;
- required asset, indication, Disease, and dosage-form values;
- unresolved multivalue Disease cells;
- project/product stage and trial phase/status value variants;
- obvious within-source `asset + dosage form + Disease` duplicates;
- diagnostic, imaging, testing, combination, and marketed-no-candidate signals.

Stop when a missing value or version ambiguity prevents stable matching or auditing.

### Run Schema confirmation

The 21 fields in `baseline-schema.json` are the current cross-TA common starting point. They remain a run schema rather than a permanent technical constant; a TA-specific change still requires explicit confirmation. For every run:

1. Mark each baseline field keep/remove/rename/reorder.
2. Define every added field's business purpose, source mapping on both sides, missing rule, and whether it affects matching.
3. Preserve the minimum logical fields: source ID in audit metadata, asset name, related indication, mapped Disease, dosage form, and the state fields needed for the run. A visible `数据来源` field is optional because source provenance is also retained in hidden workflow metadata and the raw sheets.
4. Freeze the approved ordered field list as the Run Schema.
5. Reopen this gate after any later field change.

The common 21-field contract includes two special mappings:

- listed `持证商` and clinical `申办方` both map to `License Holder`; the raw sheets preserve their original headers and definitions;
- both sources map one already-prepared `是否VBP` field; this workflow does not derive, split, or recompute VBP status.

Do not use column positions as meaning. Always map by confirmed header name.

## 3. Draft and candidate generation

### Combined draft

- Place listed rows first and clinical rows second.
- Assign temporary sequential molecule IDs across both sections.
- Track a source label in hidden workflow metadata. Add a visible source field only when it belongs to the approved Run Schema.
- Fill unavailable source-specific fields with the configured export placeholder; never replace real `0` or `False`.
- Keep copied raw listed and clinical sheets.
- Use static values in the combined main sheet.

Required arithmetic:

```text
draft data rows = listed input rows + clinical input rows
```

### Candidate generation

Generate candidates with high recall. Matching tiers are:

1. normalized clinical asset equals a listed product or generic name;
2. normalized base names agree after non-destructive dosage-form handling;
3. component sets overlap;
4. names are similar enough to warrant review;
5. no direct candidate.

Normalization may harmonize whitespace, punctuation, full-width characters, and separators. Do not silently remove salt, stereochemistry, controlled/extended-release, or component information. Those differences are review signals.

One clinical row may create several candidate rows. All final counts and actions must be grouped by the unique clinical source row, not the candidate-row count.

## 4. Decision order and matrix

Evaluate in this order:

```text
in scope
→ therapeutic asset
→ independent asset/fixed combination versus regimen only
→ same asset/components
→ same dosage form
→ same Disease
→ indication evidence for a real new Disease
→ material chemical/release-form difference
→ marketed-pool coverage gap
```

| Scenario | Required check | Default result |
|---|---|---|
| Same asset, form, Disease | Indications are the same class/subset/population | Delete clinical row; listed row carries it |
| Same asset/form, different Disease | Mapping error versus true new indication | Correct and delete if same; retain if truly new Disease |
| Same asset/Disease, different form | Dosage-form difference is real | Retain as potential formulation innovation |
| Same asset, different form and Disease | Check Disease first, then form | Retain when either difference is real |
| Partial component overlap | New research molecule or fixed product exists | Retain the complete clinical asset name |
| Only already-listed drugs used together, no fixed product | Regimen rather than asset | Exclude |
| Salt/stereoisomer/free acid/ester/release difference | Difference changes asset identity | Retain or confirm; never auto-delete |
| Diagnostic, imaging, or testing-only asset | No treatment purpose | Exclude |
| Clinical-only therapeutic asset | Belongs to target TA | Retain |
| Marketed stage but no listed-pool candidate | First launch, form, and approved indication | Follow marketed-gap rule below |

### Disease and indication

- Treat Disease as the final classification and related indication as its main checking evidence.
- Different Disease labels do not prove a new opportunity; compare the indications.
- When indications are equivalent and only mapping differs, correct Disease and reclassify.
- When indication evidence shows a true new Disease, retain the clinical row even for the same asset/form.
- Keep an upstream Disease for outcome/risk descriptions unless a clear contradiction supports a targeted correction.

### Asset combinations

- Retain a stable fixed combination as an independent asset.
- Retain any row containing a new research molecule, including the full source main-drug name with background therapy.
- Exclude a regimen made only from already-listed drugs when it is not a fixed product and contains no new molecule.

### Marketed stage with no listed candidate

Do not auto-delete. Check:

1. therapeutic versus non-therapeutic purpose;
2. China first-launch date against the run scope;
3. marketed dosage form;
4. approved indications versus the clinical Disease;
5. new molecule/fixed combination/formulation evidence.

Exclude only when the asset was marketed before the run window and the clinical row adds no new Disease, dosage form, molecule, or fixed-product value. Retain a true new Disease/form/asset. Use `HOLD` only as an intermediate state when evidence is missing.

Do not copy trial registration, sponsor, or status data into the listed carrier row after a duplicate decision. Preserve that evidence in the relation and raw sheets.

## 5. Concentrated confirmation

Resolve clear rows first. Batch only material uncertainties:

- A: keep or change an ambiguous upstream Disease;
- B: independent chemical form or naming variant;
- C: fixed combination or regimen only;
- D: outcome/risk description and upstream Disease;
- E: new molecule plus background therapy naming;
- F: marketed-pool coverage gap.

For each item report the clinical source ID, candidates, form/Disease/indication comparison, known facts, open question, recommendation, and effect.

Research only when a small missing fact can close the item. Prefer professional or regulatory sources for launch and approved-indication facts. Record source, date, query, supported conclusion, and affected source row.

## 6. Finalization and QC

Require one action per unique clinical row and zero Pending/HOLD. Keep all listed rows unless an upstream defect is returned to the source process. Apply only structured correction values, not free-text advice.

Final arithmetic:

```text
final rows
= listed input rows
+ clinical input rows
- unique clinical rows covered by listed rows
- unique excluded clinical rows
```

Final workbook requirements:

- `最终分子池` uses the approved Run Schema and static values.
- `跨库关系判断` retains candidates, decisions, evidence, actions, carriers, and final mappings.
- raw listed and clinical sheets retain their original cell values/formulas and order.
- final sequence is exactly `1..K`.
- every kept clinical row has a final sequence; deleted/excluded rows do not.
- every listed-covered deletion points to one or more valid listed carriers.
- one-to-many candidates have one consistent clinical-row action.
- no trailing formatted blank rows remain in the final pool.

Treat any cardiovascular regression fixture as a script test only. Do not load its counts or decisions as rules for another TA.
