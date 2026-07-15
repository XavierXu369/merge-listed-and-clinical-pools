# Decision Contract

## Relation sheet

`build-candidates` creates one or more rows for every unique clinical source row. The fixed relation columns are:

```text
判断编号
临床来源行ID
临床分子序号
候选已上市来源行ID
候选已上市分子序号
临床登记号
项目中国阶段
品种中国阶段
试验分期
试验状态
临床药品资产名称
候选已上市药品资产名称
候选已上市通用名
临床剂型
已上市剂型
临床对应疾病
已上市对应疾病
临床相关适应症
已上市相关适应症
临床集团
已上市集团
名称匹配层级
初始关系类型
初始判断依据
适应症核对结论
Disease修正建议
Disease修正值
资产名称修正建议
资产名称修正值
修正后关系类型
最终动作代码
最终承接上市分子序号
判断建议依据
证据来源
集中确认项
判断状态
主表同步状态
阶段执行结果
最终分子序号
```

Candidate rows are audit records, not unique action records. Group by `临床来源行ID` or `临床分子序号`. Every row in one group must receive the same final action, correction values, and carrier.

## Final action codes

| Code | Meaning | Carrier required | Enters final pool |
|---|---|---:|---:|
| `KEEP` | Retain the clinical row | No | Yes |
| `DELETE_LISTED_COVERED` | One or more listed rows cover the same asset/form/Disease value | Yes | No |
| `EXCLUDE_NON_THERAPEUTIC` | Diagnostic/testing/imaging rather than treatment | No | No |
| `EXCLUDE_NON_FIXED_REGIMEN` | Already-listed drugs used together without a fixed product/new molecule | No | No |
| `EXCLUDE_PRE_SCOPE_MARKETED` | Marketed before the run window with no new Disease/form/asset value | No | No |
| `HOLD` | Missing decision evidence | No | Never allowed in finalization |

Do not invent additional action text in the machine field. Add a new action code to both this contract and the script before using it.

## Decision status

Closed examples:

```text
规则明确
用户已确认
专业数据库已确认
轻调研已确认
```

Any blank status, `待判断`, `待用户确认`, or `待专业数据库核验` is open. Finalization must fail while any unique clinical row is open.

## Correction values

- Put the exact new Disease only in `Disease修正值`.
- Put the exact new asset display name only in `资产名称修正值`.
- Keep reasoning in the corresponding advice/rationale columns.
- Leave correction-value fields empty when no change is required.
- Do not put `无需修改`, `—`, or prose into a correction-value field.

The finalizer applies correction values only to clinical rows. It does not rewrite the raw clinical sheet.

## Decision JSON

`import-decisions` accepts UTF-8 JSON:

```json
{
  "decisions": [
    {
      "clinical_source_id": "中国临床_原始!R2",
      "clinical_sequence": 68,
      "indication_check": "Clinical and listed indications are equivalent.",
      "disease_correction": "",
      "disease_advice": "No change.",
      "asset_name_correction": "",
      "asset_name_advice": "No change.",
      "revised_relation_type": "完全重复（同资产×同剂型×同Disease）",
      "final_action": "DELETE_LISTED_COVERED",
    "carrier_listed_sequence": "12；13",
      "rationale": "Same asset, dosage form, Disease, and indication value.",
      "evidence_source": "Source workbooks",
      "confirmation_group": "",
      "decision_status": "规则明确"
    }
  ]
}
```

Provide at least one of `clinical_source_id` or `clinical_sequence`. If both are present, both must identify the same relation group.

## Preview requirements

Before generation, `finalize --mode preview` must confirm:

- every unique clinical row has exactly one allowed action;
- all one-to-many candidate rows agree;
- no action is blank or `HOLD`;
- all statuses are closed;
- every listed-covered deletion has at least one valid listed carrier sequence; multiple carriers may be separated by `；`, `/`, or `,` and every value must exist;
- correction values agree within each clinical group;
- expected final rows satisfy the row formula.

Generation additionally verifies final sequences, source composition, relation mappings, raw-sheet fingerprints, and absence of trailing blank records.
