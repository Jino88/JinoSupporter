"""Quick payload builder + committer.

Reads a compact YAML/JSON-like spec from stdin and expands into the full
AI_EXCEL_PROC payload before calling commit_dataset. Designed so the agent
only has to write the unique fields per dataset, not the boilerplate scaffold.

Spec format (JSON in stdin):
{
  "name": "<dataset>",
  "doc": {"title":..., "model":..., "report_date":..., "department":..., "marker":..., "line":..., "report_type":...,
          "primary_defect":"...", "aliases":[...], "related_defects":[...], "parts":[...], "processes":[...],
          "purpose":"...", "content":[...], "source_sheet":"...", "source_cells":{...}},
  "conditions": [ {short fields} ],   # condition_id auto cnd0..
  "results":    [ {short fields} ],   # result_id auto rid0..
  "conclusions":[ {short fields} ],   # auto concl0..
  "hints":      [ {short fields} ],   # auto hint0..
  "log":        {"confidence":, "assumptions":[], "warnings":[], "decision_rationale":""},
  "tr": {
    "ko": {"title","purpose","content","conclusions":[{"topic","st","ni"}],"hints":[{"item","why"}],"log":{"assumptions":[],"warnings":[],"dr":""}},
    "en": {...same shape...},
    "vi": {...same shape...}
  }
}
"""
import sys, json, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _ai_batch_helper as H


def build(spec):
    name = spec['name']
    doc = spec.get('doc', {})
    conds_in = spec.get('conditions', [])
    res_in = spec.get('results', [])
    conc_in = spec.get('conclusions', [])
    hints_in = spec.get('hints', [])
    log_in = spec.get('log', {})

    primary = doc.get('primary_defect', '')
    aliases = doc.get('aliases', [])

    full_doc = {
        'document_id': '', 'source_file': name,
        'source_sheet': doc.get('source_sheet', ''),
        'title': doc.get('title', ''),
        'model': doc.get('model', ''),
        'report_date': doc.get('report_date', ''),
        'department': doc.get('department', ''),
        'marker': doc.get('marker', ''),
        'line': doc.get('line', ''),
        'report_type': doc.get('report_type', 'ng_without_baseline'),
        'primary_defect': {'canonical_name': primary, 'aliases_in_document': aliases},
        'related_defects': doc.get('related_defects', []),
        'parts': doc.get('parts', []),
        'processes': doc.get('processes', []),
        'purpose': doc.get('purpose', ''),
        'content': doc.get('content', []),
        'source_cells': doc.get('source_cells', {}),
    }

    conditions = []
    for i, c in enumerate(conds_in):
        cid = f'cnd{i}'
        conditions.append({
            'condition_id': cid,
            'condition_group': c.get('group', ''),
            'line': c.get('line', ''),
            'process': c.get('process', ''),
            'changed_factor': c.get('factor', ''),
            'before_value': c.get('before'),
            'after_value': c.get('after'),
            'unit': c.get('unit'),
            'machine': c.get('machine'),
            'jig': c.get('jig'),
            'material_lot': c.get('material_lot'),
            'supplier': c.get('supplier'),
            'dry_time_sec': c.get('dry_time_sec'),
            'temperature': c.get('temperature'),
            'pressure': c.get('pressure'),
            'bond_amount': c.get('bond_amount'),
            'uv_energy': c.get('uv_energy'),
            'source_file': name,
            'sheet_name': c.get('sheet', ''),
            'source_cells': c.get('cells', []),
        })

    results = []
    for i, r in enumerate(res_in):
        cid_ref = r.get('cond')
        cid = f'cnd{cid_ref}' if isinstance(cid_ref, int) else (cid_ref or '')
        rate_pct = r.get('rate_pct')
        rate_dec = r.get('rate')
        if rate_dec is None and rate_pct is not None:
            rate_dec = rate_pct / 100.0
        if rate_pct is None and rate_dec is not None:
            rate_pct = rate_dec * 100.0
        results.append({
            'result_id': f'rid{i}',
            'condition_id': cid,
            'measurement_type': r.get('mtype', ''),
            'condition_group': r.get('group', ''),
            'date': r.get('date', ''),
            'line': r.get('line', ''),
            'input_count': r.get('input'),
            'ok_count': r.get('ok'),
            'ng_count': r.get('ng'),
            'ng_rate_decimal': rate_dec,
            'ng_rate_percent': rate_pct,
            'metric_name': r.get('metric', ''),
            'metric_value': r.get('mvalue'),
            'unit': r.get('unit'),
            'judgement': r.get('judge'),
            'ng_breakdown': r.get('breakdown') or {},
            'source_file': name,
            'sheet_name': r.get('sheet', ''),
            'source_cells': r.get('cells', []),
        })

    conclusions = []
    for i, c in enumerate(conc_in):
        conclusions.append({
            'conclusion_id': f'concl{i}',
            'topic': c.get('topic', ''),
            'statement_from_report': c.get('st', ''),
            'normalized_interpretation': c.get('ni', ''),
            'source_file': name,
            'sheet_name': c.get('sheet', ''),
            'source_cells': c.get('cells', []),
        })

    hints = []
    for i, h in enumerate(hints_in):
        hints.append({
            'hint_id': f'hint{i}',
            'check_item': h.get('item', ''),
            'reason': h.get('why', ''),
            'evidence_strength': h.get('strength', ''),
            'related_process': h.get('process', ''),
            'related_part': h.get('part', ''),
            'source_file': name,
            'sheet_name': h.get('sheet', ''),
            'source_cells': h.get('cells', []),
        })

    troubleshooting = {
        'defect_name': primary,
        'when_user_asks': spec.get('when_asks', []),
        'suggested_checks': hints,
        'limitations': spec.get('limitations', []),
    }

    full = {
        'schema_version': '0.1',
        'document': full_doc,
        'test_conditions': conditions,
        'results': results,
        'conclusions': conclusions,
        'troubleshooting_index': troubleshooting,
        'ai_extraction_log': {
            'confidence': log_in.get('confidence', 0.5),
            'assumptions': log_in.get('assumptions', []),
            'warnings': log_in.get('warnings', []),
            'decision_rationale': log_in.get('decision_rationale', ''),
        },
    }

    def expand_tr(t):
        if not t:
            return {}
        td = {
            'document': {
                'title': t.get('title', ''),
                'purpose': t.get('purpose', ''),
                'content': t.get('content', []),
            },
            'conclusions': {},
            'hints': {},
            'log': {
                'assumptions': (t.get('log') or {}).get('assumptions', []),
                'warnings': (t.get('log') or {}).get('warnings', []),
                'decision_rationale': (t.get('log') or {}).get('dr', ''),
            },
        }
        for i, tc in enumerate(t.get('conclusions', []) or []):
            td['conclusions'][f'concl{i}'] = {
                'topic': tc.get('topic', ''),
                'statement_from_report': tc.get('st', ''),
                'normalized_interpretation': tc.get('ni', ''),
            }
        for i, th in enumerate(t.get('hints', []) or []):
            td['hints'][f'hint{i}'] = {
                'check_item': th.get('item', ''),
                'reason': th.get('why', ''),
            }
        return td

    tr = spec.get('tr') or {}
    return {
        'name': name,
        'result': full,
        'translations': {
            'ko': expand_tr(tr.get('ko')),
            'en': expand_tr(tr.get('en')),
            'vi': expand_tr(tr.get('vi')),
        },
    }


def main():
    if len(sys.argv) > 1 and sys.argv[1] != '-':
        with open(sys.argv[1], 'r', encoding='utf-8-sig') as f:
            spec = json.load(f)
    else:
        spec = json.load(sys.stdin)
    payload = build(spec)
    ok = H.commit_payload(payload)
    print('ok' if ok else 'failed')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
