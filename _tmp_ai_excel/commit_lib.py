# -*- coding: utf-8 -*-
"""AiDocuments commit helper.

Each dataset is committed via commit_payload(payload). The payload dict has the
shape produced inline by the agent (matches AI_EXCEL_PROC.md JSON schema, plus
ko/en/vi narrative bundles). All inserts go through one transaction.
"""
import sqlite3, json, hashlib, os
from datetime import datetime, timezone

DB_PATH = r'D:\000. MyWorks\002. DB\process-review.db'

LANGS = ('ko', 'en', 'vi')


def _now():
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S')


def _id(prefix: str, *parts: str) -> str:
    h = hashlib.sha1('|'.join(parts).encode('utf-8')).hexdigest()[:16]
    return f'{prefix}_{h}'


def _jsons(v):
    return json.dumps(v if v is not None else [], ensure_ascii=False)


def _s(v):
    """Coerce None to empty string for NOT NULL TEXT columns."""
    if v is None:
        return ''
    if isinstance(v, str):
        return v
    return str(v)


def _doc_id(dataset_name: str) -> str:
    return _id('doc', dataset_name)


def commit_payload(payload: dict) -> str:
    """Commit one normalized document. Returns DocumentId."""
    dataset_name = payload['dataset_name']
    doc = payload['document']
    test_conditions = payload.get('test_conditions') or []
    results = payload.get('results') or []
    conclusions = payload.get('conclusions') or []
    troubleshooting = payload.get('troubleshooting') or {}
    log = payload.get('ai_extraction_log') or {}

    document_id = doc.get('document_id') or _doc_id(dataset_name)
    now = _now()

    con = sqlite3.connect(DB_PATH, timeout=30)
    con.execute('PRAGMA foreign_keys=ON')
    cur = con.cursor()
    try:
        cur.execute('BEGIN')

        # Wipe any prior rows for this DocumentId so reruns are idempotent.
        cur.execute('DELETE FROM AiNgBreakdowns WHERE ResultId IN (SELECT ResultId FROM AiResults WHERE DocumentId=?)', (document_id,))
        for tbl in ('AiResults', 'AiTestConditions', 'AiConclusions',
                    'AiTroubleshootingHints', 'AiExtractionLogs'):
            cur.execute(f'DELETE FROM {tbl} WHERE DocumentId=?', (document_id,))
        cur.execute('DELETE FROM AiDocumentTranslations WHERE DocumentId=?', (document_id,))
        cur.execute('DELETE FROM AiConclusionTranslations WHERE ConclusionId LIKE ?',
                    (f'{document_id}_%',))
        cur.execute('DELETE FROM AiHintTranslations WHERE HintId LIKE ?',
                    (f'{document_id}_%',))
        cur.execute('DELETE FROM AiLogTranslations WHERE LogId LIKE ?',
                    (f'{document_id}_%',))
        cur.execute('DELETE FROM AiDocuments WHERE DocumentId=?', (document_id,))

        # AiDocuments base row
        primary_def = doc.get('primary_defect') or {}
        primary_def_name = primary_def.get('canonical_name') if isinstance(primary_def, dict) else primary_def
        cur.execute('''INSERT INTO AiDocuments (
              DocumentId, SourceDataset, SourceFile, Title, Model, ReportDate,
              Department, Marker, Line, ReportType, PrimaryDefect,
              PrimaryDefectJson, RelatedDefectsJson, PartsJson, ProcessesJson,
              Purpose, ContentJson, SourceCellsJson, Confidence, SchemaVersion,
              RawJson, CreatedAt, UpdatedAt
          ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
            document_id,
            dataset_name,
            _s(doc.get('source_file') or dataset_name),
            _s(doc.get('title') or dataset_name),
            _s(doc.get('model')),
            _s(doc.get('report_date')),
            _s(doc.get('department')),
            _s(doc.get('marker')),
            _s(doc.get('line')),
            _s(doc.get('report_type')),
            _s(primary_def_name),
            _jsons(primary_def if isinstance(primary_def, dict) else {}),
            _jsons(doc.get('related_defects')),
            _jsons(doc.get('parts')),
            _jsons(doc.get('processes')),
            _s(doc.get('purpose')),
            _jsons(doc.get('content')),
            _jsons(doc.get('source_cells')),
            float(log.get('confidence') or 0.0),
            payload.get('schema_version', '0.1'),
            _jsons(payload),
            now, now,
        ))

        # AiTestConditions
        for i, c in enumerate(test_conditions):
            cid = c.get('condition_id') or _id('cond', document_id, str(i))
            c['condition_id'] = cid
            cur.execute('''INSERT INTO AiTestConditions (
                  ConditionId, DocumentId, ConditionGroup, Line, Process, ChangedFactor,
                  BeforeValue, AfterValue, Unit, Machine, Jig, MaterialLot, Supplier,
                  DryTimeSec, Temperature, Pressure, BondAmount, UvEnergy,
                  SourceFile, SheetName, SourceCellsJson
              ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                cid, document_id,
                _s(c.get('condition_group')), _s(c.get('line')),
                _s(c.get('process')), _s(c.get('changed_factor')),
                str(c['before_value']) if c.get('before_value') is not None else None,
                str(c['after_value']) if c.get('after_value') is not None else None,
                c.get('unit'), c.get('machine'), c.get('jig'),
                c.get('material_lot'), c.get('supplier'),
                float(c['dry_time_sec']) if c.get('dry_time_sec') is not None else None,
                str(c['temperature']) if c.get('temperature') is not None else None,
                str(c['pressure']) if c.get('pressure') is not None else None,
                str(c['bond_amount']) if c.get('bond_amount') is not None else None,
                str(c['uv_energy']) if c.get('uv_energy') is not None else None,
                _s(c.get('source_file') or dataset_name),
                _s(c.get('sheet_name')),
                _jsons(c.get('source_cells')),
            ))

        # AiResults + AiNgBreakdowns
        for i, r in enumerate(results):
            rid = r.get('result_id') or _id('res', document_id, str(i))
            r['result_id'] = rid
            cur.execute('''INSERT INTO AiResults (
                  ResultId, DocumentId, ConditionId, MeasurementType, ConditionGroup,
                  ResultDate, Line, InputCount, OkCount, NgCount, NgRateDecimal,
                  NgRatePercent, MetricName, MetricValue, Unit, Judgement,
                  SourceFile, SheetName, SourceCellsJson
              ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                rid, document_id, r.get('condition_id'),
                _s(r.get('measurement_type')), _s(r.get('condition_group')),
                _s(r.get('date')), _s(r.get('line')),
                r.get('input_count'), r.get('ok_count'), r.get('ng_count'),
                r.get('ng_rate_decimal'), r.get('ng_rate_percent'),
                _s(r.get('metric_name')),
                float(r['metric_value']) if r.get('metric_value') is not None else None,
                r.get('unit'), r.get('judgement'),
                _s(r.get('source_file') or dataset_name),
                _s(r.get('sheet_name')),
                _jsons(r.get('source_cells')),
            ))
            br = r.get('ng_breakdown') or {}
            if isinstance(br, dict):
                for j, (defect, val) in enumerate(br.items()):
                    bid = _id('br', rid, str(j))
                    if isinstance(val, dict):
                        cnt = val.get('count'); rate = val.get('rate')
                    else:
                        cnt = val; rate = None
                    cur.execute('''INSERT INTO AiNgBreakdowns (
                          BreakdownId, ResultId, DefectName, DefectCount, DefectRate
                      ) VALUES (?,?,?,?,?)''', (
                        bid, rid, defect,
                        float(cnt) if cnt is not None else None,
                        float(rate) if rate is not None else None,
                    ))

        # AiConclusions + translations
        for i, cn in enumerate(conclusions):
            cnid = cn.get('conclusion_id') or _id(f'{document_id}_concl', str(i))
            if not cnid.startswith(document_id):
                cnid = f'{document_id}_concl_{i:02d}'
            cn['conclusion_id'] = cnid
            cur.execute('''INSERT INTO AiConclusions (
                  ConclusionId, DocumentId, Topic, StatementFromReport,
                  NormalizedInterpretation, SourceFile, SheetName, SourceCellsJson
              ) VALUES (?,?,?,?,?,?,?,?)''', (
                cnid, document_id,
                _s(cn.get('topic')), _s(cn.get('statement_from_report')),
                _s(cn.get('normalized_interpretation')),
                _s(cn.get('source_file') or dataset_name),
                _s(cn.get('sheet_name')),
                _jsons(cn.get('source_cells')),
            ))
            tr = cn.get('translations') or {}
            for lang in LANGS:
                t = tr.get(lang) or {}
                cur.execute('''INSERT OR REPLACE INTO AiConclusionTranslations (
                      ConclusionId, Lang, Topic, StatementFromReport,
                      NormalizedInterpretation, UpdatedAt
                  ) VALUES (?,?,?,?,?,?)''', (
                    cnid, lang,
                    _s(t.get('topic') or cn.get('topic')),
                    _s(t.get('statement_from_report') or cn.get('statement_from_report')),
                    _s(t.get('normalized_interpretation') or cn.get('normalized_interpretation')),
                    now,
                ))

        # AiTroubleshootingHints + translations
        defect_name = troubleshooting.get('defect_name') if isinstance(troubleshooting, dict) else None
        suggested = troubleshooting.get('suggested_checks') if isinstance(troubleshooting, dict) else []
        for i, h in enumerate(suggested or []):
            hid = h.get('hint_id') or f'{document_id}_hint_{i:02d}'
            h['hint_id'] = hid
            cur.execute('''INSERT INTO AiTroubleshootingHints (
                  HintId, DocumentId, DefectName, CheckItem, Reason,
                  EvidenceStrength, RelatedProcess, RelatedPart,
                  SourceFile, SheetName, SourceCellsJson
              ) VALUES (?,?,?,?,?,?,?,?,?,?,?)''', (
                hid, document_id, _s(defect_name),
                _s(h.get('check_item')), _s(h.get('reason')),
                _s(h.get('evidence_strength')),
                _s(h.get('related_process')), _s(h.get('related_part')),
                _s(h.get('source_file') or dataset_name),
                _s(h.get('sheet_name')),
                _jsons(h.get('source_cells')),
            ))
            tr = h.get('translations') or {}
            for lang in LANGS:
                t = tr.get(lang) or {}
                cur.execute('''INSERT OR REPLACE INTO AiHintTranslations (
                      HintId, Lang, CheckItem, Reason, UpdatedAt
                  ) VALUES (?,?,?,?,?)''', (
                    hid, lang,
                    _s(t.get('check_item') or h.get('check_item')),
                    _s(t.get('reason') or h.get('reason')),
                    now,
                ))

        # AiExtractionLogs + translations
        log_id = f'{document_id}_log'
        cur.execute('''INSERT INTO AiExtractionLogs (
              LogId, DocumentId, Confidence, AssumptionsJson, WarningsJson,
              DecisionRationale, CreatedAt
          ) VALUES (?,?,?,?,?,?,?)''', (
            log_id, document_id,
            float(log.get('confidence') or 0.0),
            _jsons(log.get('assumptions') or []),
            _jsons(log.get('warnings') or []),
            _s(log.get('decision_rationale')),
            now,
        ))
        log_tr = log.get('translations') or {}
        for lang in LANGS:
            t = log_tr.get(lang) or {}
            cur.execute('''INSERT OR REPLACE INTO AiLogTranslations (
                  LogId, Lang, AssumptionsJson, WarningsJson,
                  DecisionRationale, UpdatedAt
              ) VALUES (?,?,?,?,?,?)''', (
                log_id, lang,
                _jsons(t.get('assumptions') or log.get('assumptions') or []),
                _jsons(t.get('warnings') or log.get('warnings') or []),
                _s(t.get('decision_rationale') or log.get('decision_rationale')),
                now,
            ))

        # AiDocumentTranslations
        doc_tr = doc.get('translations') or {}
        for lang in LANGS:
            t = doc_tr.get(lang) or {}
            cur.execute('''INSERT OR REPLACE INTO AiDocumentTranslations (
                  DocumentId, Lang, Title, Purpose, ContentJson, UpdatedAt
              ) VALUES (?,?,?,?,?,?)''', (
                document_id, lang,
                _s(t.get('title') or doc.get('title') or dataset_name),
                _s(t.get('purpose') or doc.get('purpose')),
                _jsons(t.get('content') or doc.get('content')),
                now,
            ))

        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()
    return document_id


def fetch_paste(dataset_name: str) -> str:
    con = sqlite3.connect(DB_PATH)
    try:
        cur = con.execute(
            "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste' LIMIT 1",
            (dataset_name,))
        row = cur.fetchone()
        return row[0] if row else None
    finally:
        con.close()


if __name__ == '__main__':
    # smoke
    import sys
    name = sys.argv[1] if len(sys.argv) > 1 else None
    if name:
        t = fetch_paste(name)
        print(f'paste len: {len(t) if t else 0}')
