# Full, Supplement, and Clean Contract

## Contents

1. Output roles
2. Run states
3. Marketed-information supplement
4. Downstream module coordination
5. Clean derivation
6. Conditional License Holder

## 1. Output roles

### Full pool

Treat `最终分子池_完整版` as the authoritative fact table. It contains the approved wide schema, final stable sequence, source provenance, and all later supplements/module outputs.

### Working Clean

Treat a Working Clean as a usable projection of the current full pool. It may exist while supplement or module tasks are still open. Record those open states; do not describe it as final.

### Final Clean

Treat a Final Clean as a closed projection bound to the latest full-pool fingerprint. Require identical IDs, row count, and order with the full pool. Its column count is run-specific.

### Audit sheets

Retain the relation sheet and raw listed/clinical sheets. Keep the run manifest and supplement audit hidden. Do not force extra visible explanation sheets when the user chooses a lightweight human-fill workflow.

## 2. Run states

Use these core states:

```text
DRAFT_BUILT
RELATIONS_BUILT
DECISIONS_PENDING
DECISIONS_CLOSED
FULL_BASELINE_LOCKED
SUPPLEMENT_PENDING
SUPPLEMENT_CLOSED
SUPPLEMENT_NOT_APPLICABLE
CLEAN_WORKING
CLEAN_STALE
CLEAN_FINAL
```

Store the current state, full fingerprint, Clean-bound full fingerprint, supplement counts, and module statuses in hidden metadata.

## 3. Marketed-information supplement

### Candidate gate

Define the trigger in the run config. The common pattern is:

```text
数据来源 = 中国临床库
AND 项目中国阶段 ∈ approved marketed-stage values
```

Do not hardcode the number of candidates. Export the locked full ID plus context fields and only the approved target fields.

### Status contract

Use:

- `待补充`: open;
- `已补充`: closed with usable returned information;
- `未查到`: closed after a documented unsuccessful search;
- `不适用`: closed because the candidate does not require supplementation;
- `经批准跳过`: closed through explicit approval.

Working Clean may coexist with `待补充`. Final Clean may not when supplementation is required.

### Apply contract

- Join only by the locked full ID.
- Reject blank, duplicate, and unknown IDs.
- Treat name, company, and registration number as verification context, never join keys.
- Ignore blank returned target cells.
- Fill an empty full-pool value by default.
- Reject a different nonblank replacement unless `--allow-overwrite` is explicitly supplied after approval.
- Update the full pool before Clean.
- Retain returned statuses and notes in the hidden supplement audit.
- Mark an existing Clean stale after any applied full-pool change.

## 4. Downstream module coordination

Configure every module with:

- a stable module name;
- whether it is required for Final Clean;
- its current status;
- its full-pool output fields;
- the upstream fields that can invalidate it.

Use statuses:

```text
pending
completed
not_applicable
approved_skip
```

This Skill does not calculate module outputs. It verifies their closure and carries approved values through Clean rules or auxiliary columns.

After a supplement, compare changed fields with each module's `affected_by` list. Report affected modules. Recheck only those modules.

## 5. Clean derivation

Derive rows in current full-pool order. Use the full sequence as the join/order contract. Support:

- direct column selection;
- coalesced columns;
- constants;
- missing placeholders;
- ordered conditional cases;
- optional auxiliary fields.

Never read raw listed or clinical sheets to populate Clean after the full baseline exists.

For Final Clean:

1. require `--confirmed-current-full`;
2. verify full headers and unique sequential IDs;
3. verify supplement/module gates;
4. replace any existing Clean only in a newly created output workbook;
5. verify identical IDs, row count, and order;
6. store the bound full fingerprint;
7. verify static values and no trailing blank rows.

## 6. Conditional License Holder

Represent License Holder as a Clean conditional rule rather than a fixed clinical mapping. A reusable pattern is:

```text
IF 数据来源 = 中国已上市库
  THEN 持证商
ELSE IF 数据来源 = 中国临床库
        AND 项目中国阶段 is marketed
        AND 持证商 is nonblank
  THEN 持证商
ELSE
  申办方
```

When supplementation closes as `未查到`, keep the sponsor rather than inventing a holder. Preserve the status in the supplement audit.
