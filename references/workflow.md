# Listed and Clinical Pool Workflow

## Contents

1. Purpose and boundaries
2. Inputs and Run Schema
3. Draft and candidate generation
4. Decision matrix
5. Full baseline
6. Parallel post-baseline work
7. Clean derivation and final QC

## 1. Purpose and boundaries

Use this workflow only after both source libraries complete upstream cleaning, indication Mapping, and approved within-source consolidation. The normal business grain is:

```text
asset or fixed component set × dosage form × mapped Disease
```

This workflow aligns unlike schemas, generates cross-library candidates, closes decisions, locks the full-pool fact baseline, coordinates optional marketed-information supplements, and derives Clean views.

It does not execute:

- company-owned-product matching;
- MOA overlap matching;
- VBP derivation;
- Giant MNC or TA-leading-company classification;
- commercial prioritization;
- wholesale TA Disease remapping.

Required upstream state:

- Listed data are already split, mapped, and consolidated to the approved listed grain.
- Clinical data are already cleaned to the approved clinical source-row grain.
- Every included row has one resolved Disease.
- Both sources use the same approved TA Disease list/version.
- A reliable source ID exists, or a run-frozen fallback ID is explicitly approved and bound to the source fingerprint.

## 2. Inputs and Run Schema

### Read-only preflight

Inspect:

- input file hashes, selected sheets, effective bounds, headers, hidden/trailing cells, formulas, cached values, and external links;
- blank/duplicate source IDs and any fallback-ID risk;
- required asset, indication, Disease, dosage-form, stage, and audit fields;
- unresolved multivalue Disease cells;
- project/product stage and trial phase/status variants;
- within-source `asset + dosage form + Disease` duplicate signatures;
- diagnostic, testing, non-fixed-regimen, new-combination, and marketed-no-candidate signals.

Do not construct output when a gap prevents stable matching or auditing.

### Run Schema contract

Freeze four independent structures for every run:

1. `wide_schema`: the authoritative full-pool fields and listed/clinical mappings.
2. `wide_fields`: logical fields used by matching, decisions, provenance, IDs, and stages.
3. `clean_schema`: the ordered core Clean fields derived only from the full pool.
4. `clean_auxiliary_schema`: optional run-specific fields appended to a Working Clean for manual downstream judgments.

The field counts are run outputs, not constants. A historical 37-, 40-, or 21-field workbook is a fixture only.

Every mapping rule must state its missing behavior. Use header names, never column letters. Preserve internal nulls until export; apply one approved display placeholder only at the output boundary.

### Supported mapping patterns

Full-pool source rules may use:

- `sequence`: temporary draft sequence;
- `constant`;
- `column`;
- `column_or_missing`;
- `coalesce_columns`;
- `missing`.

Clean rules may additionally use:

- `conditional`: ordered cases with a default rule.

Use conditional rules for source/stage-dependent values such as License Holder. Do not encode such logic as a manual afterthought.

## 3. Draft and candidate generation

### Combined draft

- Place listed rows first and clinical rows second.
- Use static cached values for the mapped draft; retain formula expressions only on raw audit sheets.
- Assign temporary sequential IDs for candidate review.
- Track source label, source row ID, original row number, input hash, and run fingerprint in hidden metadata.
- Include visible `数据来源` when the approved full schema requests it.
- Fill unavailable source-specific values with the configured export placeholder without replacing real `0` or `False`.
- Verify copied raw-sheet value/formula fingerprints against the original input sheets.

Required arithmetic:

```text
draft rows = listed input rows + clinical input rows
```

### Candidate generation

Generate high-recall candidates in this order:

1. normalized clinical asset equals a listed product or generic name;
2. base names agree after conservative dosage-form handling;
3. component sets overlap;
4. names are sufficiently similar for review;
5. no direct candidate.

Normalization may harmonize whitespace, punctuation, width, and separators. Preserve salt, stereochemistry, release form, fixed-combination, and component differences as review signals.

One clinical row may yield several candidate rows. Bind the candidate workbook and any decision JSON to one candidate fingerprint calculated from immutable candidate columns.

## 4. Decision matrix

Evaluate in this order:

```text
in scope
→ therapeutic asset
→ independent asset/fixed combination versus regimen only
→ same asset/components
→ same dosage form
→ same Disease
→ indication evidence for a true new Disease
→ material chemical/release-form difference
→ marketed-pool coverage gap
```

| Scenario | Required check | Default result |
|---|---|---|
| Same asset, form, Disease | Indications are equivalent, subset, or population expression | Delete clinical row; listed row carries it |
| Same asset/form, different Disease | Mapping difference versus true new indication | Correct/delete if equivalent; retain if a real new Disease |
| Same asset/Disease, different form | Form difference is real | Retain as a formulation opportunity |
| Same asset, different form and Disease | Check Disease then form | Retain when either difference is real |
| Partial component overlap | A new research molecule or fixed product exists | Retain the complete clinical asset name |
| Only already-listed drugs used together | No fixed product and no new molecule | Exclude as a regimen |
| Salt/stereoisomer/free acid/ester/release difference | Difference changes asset identity | Retain or confirm; never auto-delete |
| Diagnostic/testing/imaging asset | No therapeutic purpose | Exclude |
| Clinical-only therapeutic asset | Belongs to the target TA | Retain |
| Marketed stage but no listed candidate | Launch/form/approved-indication context | Apply marketed-gap review; do not auto-delete |

### Evidence that is not decisive alone

Do not keep/delete solely because:

- registration numbers differ or match;
- sponsors, holders, or groups differ or match;
- trial titles differ or look similar;
- target order differs;
- project stage says `已上市`;
- launch information is absent;
- asset names are similar.

Target order may be normalized as identity evidence, but target equality does not replace asset/form/Disease analysis. Distinct molecules remain distinct even when companies, targets, or Diseases overlap.

### Disease and indication

- Treat Disease as the classification and indication text as the primary checking evidence.
- Different Disease labels do not prove a new opportunity.
- Correct a targeted Mapping error only through the approved correction field.
- Retain a true new Disease even for the same asset and form.

### Combinations and background therapy

- Retain a stable fixed combination as an independent asset.
- Retain any row containing a new research molecule and preserve its complete main-drug expression.
- Exclude a regimen composed only of already-listed drugs when it is neither a fixed product nor a new molecule.

### Marketed stage with no listed candidate

Check therapeutic purpose, China first launch against scope, marketed form, approved indications, and new molecule/fixed-combination/formulation evidence. Exclude only when the asset predates scope and adds no new Disease, form, asset, or fixed-product value. Missing marketed information is a data-completion issue, not a merge decision.

## 5. Full baseline

Require one closed action per unique clinical source row and zero Pending/HOLD groups. Keep all valid listed rows. Apply only structured correction values.

```text
full rows
= listed input rows
+ clinical input rows
- unique clinical rows covered by listed rows
- unique excluded clinical rows
```

The generated full pool must:

- use exactly the approved `wide_schema`;
- contain static values;
- assign final sequence `1..K` once and never renumber it later;
- preserve source provenance in visible fields or hidden metadata;
- map each kept clinical row to one final sequence;
- map deleted/excluded clinical rows to no final sequence;
- preserve relation decisions and raw source audit sheets;
- store input, config, run, candidate, and full fingerprints in the hidden manifest.

The full pool is now the sole fact baseline. Any later manual or imported update must target it first by locked final ID.

## 6. Parallel post-baseline work

After the full baseline, allow these branches to proceed independently:

```text
                 ┌─ Working Clean
Full baseline ───┼─ marketed-information supplement
                 └─ downstream ownership/MOA/company modules
```

The supplement branch is not a hard prerequisite for Working Clean. It is a closure requirement only when the run defines it as required for Final Clean.

When supplement data return:

1. match by locked full ID only;
2. update the full pool first;
3. reject blanks as overwrite values;
4. reject a conflict with an existing valid value unless explicit overwrite approval is supplied;
5. store one closed status per candidate;
6. mark any existing Clean stale;
7. regenerate Clean from the updated full pool.

Review downstream modules selectively:

- holder/group changes may affect company ownership and TA-leading-company results;
- target/MOA changes may affect MOA overlap;
- name, launch date, or packaging changes usually require only Full/Clean synchronization;
- unchanged module-driving fields do not justify a full rerun.

## 7. Clean derivation and final QC

### Working Clean

Permit generation whenever the full baseline exists and passes schema/ID checks. Disclose open supplement/module states in the command report and manifest. Append run-specific auxiliary columns only when configured.

### Final Clean

Require:

- current full-pool fingerprint explicitly confirmed;
- every required supplement candidate closed as `已补充`, `未查到`, `不适用`, or `经批准跳过`;
- every required downstream module closed as `completed`, `not_applicable`, or `approved_skip`;
- no stale Clean binding;
- exact ID set and order agreement with the full pool;
- exact approved Clean headers;
- static values and no trailing formatted blank rows.

License Holder must follow the approved conditional rule. The common pattern is:

- listed source → holder;
- active clinical source → sponsor;
- clinical source already marketed with valid supplemented holder → holder;
- marketed information not found → retain sponsor and retain `未查到` in the supplement audit state.

Treat cardiovascular, kidney, and all other completed TA workbooks as regression fixtures only. Never reuse their counts, company lists, or decisions as defaults.
