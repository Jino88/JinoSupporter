"""Compact payload builder for AI Batch. Given a list of dataset analyses,
builds the AI_EXCEL_PROC.md JSON, ko/en/vi translations and commits."""
from __future__ import annotations
import json, sys, importlib.util, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
spec = importlib.util.spec_from_file_location('h', os.path.join(os.path.dirname(os.path.abspath(__file__)), '_ai_batch_helper.py'))
h = importlib.util.module_from_spec(spec); spec.loader.exec_module(h)


def build(d: dict) -> dict:
    """`d` is a compact analysis dict. Returns full payload."""
    name = d['name']
    doc_id = ''
    rt = d.get('report_type', 'ng_without_baseline')
    primary = d.get('primary_defect') or 'NG Function'
    related = d.get('related_defects') or []
    parts = d.get('parts') or []
    procs = d.get('processes') or []
    purpose = {'ko': d['purpose_ko'], 'en': d['purpose_en'], 'vi': d['purpose_vi']}
    title = {'ko': d['title_ko'], 'en': d['title_en'], 'vi': d['title_vi']}
    content = {'ko': d.get('content_ko') or [], 'en': d.get('content_en') or [], 'vi': d.get('content_vi') or []}

    conds = d.get('test_conditions') or []
    results = d.get('results') or []
    concls = d.get('conclusions') or []
    hints = d.get('hints') or []
    log = d.get('log') or {}
    conf = d.get('confidence', 0.6)

    # Normalize ids
    for i, c in enumerate(conds):
        c.setdefault('condition_id', f'c{i+1}')
        c.setdefault('source_file', name)
        c.setdefault('sheet_name', d.get('default_sheet', ''))
        c.setdefault('source_cells', [])
    for i, r in enumerate(results):
        r.setdefault('result_id', f'r{i+1}')
        r.setdefault('source_file', name)
        r.setdefault('sheet_name', d.get('default_sheet', ''))
        r.setdefault('source_cells', [])
    for i, c in enumerate(concls):
        c.setdefault('conclusion_id', f'k{i+1}')
        c.setdefault('source_file', name)
        c.setdefault('sheet_name', d.get('default_sheet', ''))
        c.setdefault('source_cells', [])
    for i, hi in enumerate(hints):
        hi.setdefault('hint_id', f'h{i+1}')
        hi.setdefault('source_file', name)
        hi.setdefault('sheet_name', d.get('default_sheet', ''))
        hi.setdefault('source_cells', [])

    document = {
        'document_id': doc_id,
        'source_file': name,
        'source_sheet': d.get('default_sheet', ''),
        'title': d.get('title_en') or name,
        'model': d.get('model', ''),
        'report_date': d.get('report_date', ''),
        'department': d.get('department', ''),
        'marker': d.get('marker', ''),
        'line': d.get('line', ''),
        'report_type': rt,
        'primary_defect': {'canonical_name': primary, 'aliases_in_document': d.get('primary_aliases', [])},
        'related_defects': related,
        'parts': parts,
        'processes': procs,
        'purpose': d.get('purpose_en', ''),
        'content': d.get('content_en') or [],
        'source_cells': d.get('source_cells_doc', {})
    }

    result_obj = {
        'schema_version': '0.1',
        'generated_report_markdown': d.get('generated_report_markdown') or d.get('report_markdown') or '',
        'document': document,
        'test_conditions': conds,
        'results': [{
            'result_id': r['result_id'],
            'condition_id': r.get('condition_id'),
            'measurement_type': r.get('measurement_type', 'Function'),
            'condition_group': r.get('condition_group', ''),
            'date': r.get('date', d.get('report_date', '')),
            'line': r.get('line', ''),
            'input_count': r.get('input_count'),
            'ok_count': r.get('ok_count'),
            'ng_count': r.get('ng_count'),
            'ng_rate_decimal': r.get('ng_rate_decimal'),
            'ng_rate_percent': r.get('ng_rate_percent'),
            'metric_name': r.get('metric_name', ''),
            'metric_value': r.get('metric_value'),
            'unit': r.get('unit'),
            'judgement': r.get('judgement'),
            'ng_breakdown': r.get('ng_breakdown') or {},
            'source_file': r['source_file'],
            'sheet_name': r['sheet_name'],
            'source_cells': r['source_cells'],
        } for r in results],
        'conclusions': [{
            'conclusion_id': c['conclusion_id'],
            'topic': c.get('topic_en', c.get('topic', '')),
            'statement_from_report': c.get('statement_en', c.get('statement', '')),
            'normalized_interpretation': c.get('interp_en', c.get('interp', '')),
            'source_file': c['source_file'],
            'sheet_name': c['sheet_name'],
            'source_cells': c['source_cells'],
        } for c in concls],
        'troubleshooting_index': {
            'defect_name': primary,
            'when_user_asks': d.get('when_user_asks', []),
            'suggested_checks': [{
                'hint_id': hi['hint_id'],
                'check_item': hi.get('check_en', hi.get('check', '')),
                'reason': hi.get('reason_en', hi.get('reason', '')),
                'evidence_strength': hi.get('evidence_strength', 'medium'),
                'related_process': hi.get('related_process', ''),
                'related_part': hi.get('related_part', ''),
                'source_file': hi['source_file'],
                'sheet_name': hi['sheet_name'],
                'source_cells': hi['source_cells'],
            } for hi in hints],
            'limitations': d.get('limitations', [])
        },
        'ai_extraction_log': {
            'confidence': conf,
            'assumptions': log.get('assumptions_en', []),
            'warnings': log.get('warnings_en', []),
            'decision_rationale': log.get('rationale_en', '')
        }
    }

    def trs(lang):
        return {
            'document': {'title': title[lang], 'purpose': purpose[lang], 'content': content[lang]},
            'conclusions': {c['conclusion_id']: {
                'topic': c.get(f'topic_{lang}', c.get('topic_en', '')),
                'statement_from_report': c.get(f'statement_{lang}', c.get('statement_en', '')),
                'normalized_interpretation': c.get(f'interp_{lang}', c.get('interp_en', '')),
            } for c in concls},
            'hints': {hi['hint_id']: {
                'check_item': hi.get(f'check_{lang}', hi.get('check_en', '')),
                'reason': hi.get(f'reason_{lang}', hi.get('reason_en', '')),
            } for hi in hints},
            'log': {
                'assumptions': log.get(f'assumptions_{lang}', log.get('assumptions_en', [])),
                'warnings': log.get(f'warnings_{lang}', log.get('warnings_en', [])),
                'decision_rationale': log.get(f'rationale_{lang}', log.get('rationale_en', '')),
            }
        }

    return {'name': name, 'result': result_obj, 'translations': {'ko': trs('ko'), 'en': trs('en'), 'vi': trs('vi')}}


def commit(d: dict) -> bool:
    payload = build(d)
    return h.commit_payload(payload)
