"""Lean payloads — focus on essential data per workbook."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _batch_build as B

D = []

# 13 — TIU C11-20 don't use VP press jig
D.append({
    'name': '19. TIU C11-20  Report TestDont use VP press Jig ( Frame+ VP) 2025.12.24',
    'report_type': 'normal_comparison', 'model': 'TIU C11-20', 'report_date': '2025-12-24', 'department': 'ME', 'marker': 'Thao', 'default_sheet': 'Test',
    'primary_defect': 'Not enough glue VP/Frame', 'related_defects': ['VP separate','NG function'], 'parts': ['VP','Frame'], 'processes': ['VP press jig (Sub3 ass’y)'],
    'title_en': 'TIU C11-20 — Test: Don’t use VP press jig (Frame+VP) (2025-12-24)',
    'title_ko': 'TIU C11-20 — VP press jig 미사용 검증 (2025-12-24)',
    'title_vi': 'TIU C11-20 — Thử không dùng VP press jig (Frame+VP) (24-12-2025)',
    'purpose_en': 'Verify whether removing VP press jig improves VP+Frame damage. Compare ‘don’t use jig’ 50pcs to Normal ‘use jig’ on Frame+VP vision (glue), VP separate, function, and tension.',
    'purpose_ko': 'VP press jig 제거가 VP+Frame damage 를 개선하는지 확인. 미사용 50pcs vs Normal 사용 (Frame+VP vision/glue, VP separate, function, tension).',
    'purpose_vi': 'Kiểm tra việc bỏ VP press jig có cải thiện VP+Frame damage không. So sánh ‘không dùng jig’ 50pcs vs Normal ‘dùng jig’ trên Frame+VP vision/glue, VP separate, function và tension.',
    'content_en': ['Frame+VP vision (Not enough glue): Test 29pcs vs Normal 30pcs.','VP separate: Test 29 vs Normal 30 (both 0).','Function: continues but truncated in extract.'],
    'content_ko': ['Frame+VP vision (glue 부족): Test 29 vs Normal 30.','VP separate: Test 29 vs Normal 30 (모두 0).','Function 표 일부 절단.'],
    'content_vi': ['Frame+VP vision (thiếu glue): Test 29 vs Normal 30.','VP separate: Test 29 vs Normal 30 (đều 0).','Bảng Function bị cắt.'],
    'test_conditions': [{'condition_group': 'press_jig', 'process': 'Sub3 VP press', 'changed_factor': 'VP press jig', 'before_value': 'Use jig (Normal)', 'after_value': "Don't use jig", 'source_cells': ['B8']}],
    'results': [
        {'condition_id': 'c1', 'measurement_type': 'Vision Frame+VP', 'date': '2025-12-24', 'line': 'TIU C11-20R', 'input_count': 29, 'ok_count': 22, 'ng_count': 7, 'ng_rate_decimal': 0.2414, 'ng_rate_percent': 24.14, 'metric_name': "Test don't use VP press jig", 'judgement': 'WORSE', 'ng_breakdown': {'Not enough glue VP/Frame': {'count': 7, 'rate': 0.2414}}, 'source_cells': ['N15']},
        {'condition_id': 'c1', 'measurement_type': 'Vision Frame+VP', 'date': '2025-12-24', 'line': 'TIU C11-20R', 'input_count': 30, 'ok_count': 30, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Normal (use VP press jig)', 'judgement': 'BASELINE', 'source_cells': ['N16']},
        {'condition_id': 'c1', 'measurement_type': 'VP Separate', 'date': '2025-12-24', 'input_count': 29, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': "Test don't use VP press jig", 'judgement': 'SIMILAR', 'source_cells': ['N20']},
        {'condition_id': 'c1', 'measurement_type': 'VP Separate', 'date': '2025-12-24', 'input_count': 30, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Normal (use VP press jig)', 'judgement': 'BASELINE', 'source_cells': ['N21']},
    ],
    'conclusions': [{'topic_en': "VP press jig necessity",'topic_ko':"VP press jig 필요성",'topic_vi':"Tính cần thiết của VP press jig",
        'statement_en': "Without jig 24.14% glue NG vs Normal 0%. VP separate same (0%).", 'statement_ko':"jig 미사용 24.14% glue NG vs Normal 0%. VP separate 동일.", 'statement_vi':"Không jig 24.14% NG glue vs Normal 0%. VP separate như nhau.",
        'interp_en': 'Removing the VP press jig moves Frame+VP glue NG from 0% to 24.14% (infinite relative worsening). VP press jig must remain in production.',
        'interp_ko': 'VP press jig 제거 시 Frame+VP glue NG 가 0% → 24.14% 로 급증 (상대 악화 무한). jig 유지 필수.',
        'interp_vi': 'Bỏ VP press jig đẩy NG glue Frame+VP từ 0% lên 24.14% (xấu đi vô hạn về tương đối). Phải giữ VP press jig trong sản xuất.','source_cells':['N15:N16']}],
    'hints': [{'check_en':'Keep VP press jig in production','reason_en':'24.14% glue NG without jig vs 0% with jig.','check_ko':'VP press jig 유지.','reason_ko':'jig 없으면 glue NG 24.14% vs jig 사용 0%.','check_vi':'Giữ VP press jig trong dây.','reason_vi':'Không có jig 24.14% NG glue vs 0%.','evidence_strength':'strong','related_process':'Sub3 VP press','related_part':'VP/Frame','source_cells':['N15:N16']}],
    'log': {'rationale_en': 'Same-day Test vs Normal pair — normal_comparison. Without-jig is clearly worse.','rationale_ko':'동일일자 Test vs Normal — normal_comparison. jig 미사용 명확히 악화.','rationale_vi':'Cùng ngày Test vs Normal — normal_comparison. Không jig rõ ràng tệ hơn.','warnings_en':['Function table truncated; rerun for full data.'],'warnings_ko':['Function 표 절단 — 필요 시 재실행.'],'warnings_vi':['Bảng Function bị cắt — chạy lại nếu cần.']},
    'when_user_asks': ['VP press jig','TIU C11-20 glue NG','Frame VP damage'],
    'confidence': 0.78,
})

# 17 — VP damage tracking MSU-L20L15-07 (all 0)
D.append({
    'name': '19.BRS-161014 Result checking Problem NG VP damage date 611.2025 -',
    'report_type': 'ng_without_baseline', 'model': 'MSU-L20L15-07', 'report_date': '2025-11-06', 'department': 'ME', 'marker': 'Thao', 'default_sheet': '15.10',
    'primary_defect': 'NG VP damage', 'related_defects': [], 'parts': ['VP'], 'processes': ['Sub1','Main2','Function'],
    'title_en': 'MSU-L20L15-07 — Checking Final Sample NG VP damage (2025-11-06)',
    'title_ko': 'MSU-L20L15-07 — Final 샘플 NG VP damage 점검 (2025-11-06)',
    'title_vi': 'MSU-L20L15-07 — Kiểm tra mẫu Final NG VP damage (06-11-2025)',
    'purpose_en': 'Check the reason of NG VP damage at module by tracking 50pcs through Sub4 (Sub1/Main2/Function) using VP #11 IR 250327005-00001.',
    'purpose_ko': 'VP #11 (IR 250327005-00001) 로 50pcs 를 Sub4 (Sub1/Main2/Function) tracking 하여 module 의 NG VP damage 원인 점검.',
    'purpose_vi': 'Kiểm tra nguyên nhân NG VP damage tại module bằng tracking 50pcs qua Sub4 (Sub1/Main2/Function) dùng VP #11 IR 250327005-00001.',
    'content_en': ['Sub1: VP array / Laser inside / VP array 2nd / UV LED VP+CD / VP+CD Vision / Array Tray — 50pcs each, all 0 NG.','Main2: 8 process steps × 50pcs each, all 0 NG.','Function: Air leak / SPK+Grill / Height / Sigma / Hearing / Marking — truncated but flagged 0.'],
    'content_ko': ['Sub1: VP array / Laser inside / VP array 2nd / UV LED VP+CD / VP+CD Vision / Array Tray — 50pcs 모두 0 NG.','Main2: 8개 공정 × 50pcs 모두 0 NG.','Function: Air leak / SPK+Grill / Height / Sigma / Hearing / Marking — 일부 절단, 0 표시.'],
    'content_vi': ['Sub1: VP array / Laser inside / VP array 2nd / UV LED VP+CD / VP+CD Vision / Array Tray — 50pcs đều 0 NG.','Main2: 8 bước × 50pcs đều 0 NG.','Function: Air leak / SPK+Grill / Height / Sigma / Hearing / Marking — cắt bớt, đánh dấu 0.'],
    'test_conditions': [{'condition_group':'vp_id_tracking','process':'Sub4 tracking','changed_factor':'VP lot #11 (IR 250327005-00001)','before_value':None,'after_value':'VP #11','source_cells':['B16']}],
    'results': [
        {'condition_id':'c1','measurement_type':'Tracking Sub1','date':'2025-11-06','line':'E2-3B','input_count':50,'ng_count':0,'ng_rate_decimal':0.0,'ng_rate_percent':0.0,'metric_name':'Sub1 (6 stations sum)','judgement':'OK','source_cells':['F22:W22']},
        {'condition_id':'c1','measurement_type':'Tracking Main2','date':'2025-11-06','line':'E2-3B','input_count':50,'ng_count':0,'ng_rate_decimal':0.0,'ng_rate_percent':0.0,'metric_name':'Main2 (8 stations sum)','judgement':'OK','source_cells':['F26:AC26']},
    ],
    'conclusions': [{'topic_en':'VP damage tracking result','topic_ko':'VP damage tracking 결과','topic_vi':'Kết quả tracking VP damage',
        'statement_en':'Sub1 (6 stations) + Main2 (8 stations) — 50pcs each, 0 VP damage observed.',
        'statement_ko':'Sub1 (6 station) + Main2 (8 station) — 각 50pcs, VP damage 0건.',
        'statement_vi':'Sub1 (6 trạm) + Main2 (8 trạm) — mỗi trạm 50pcs, không có VP damage.',
        'interp_en':'No same-event baseline available — store as ng_without_baseline. Tracking lot of VP #11 showed zero VP damage across all 14 production stations on 50pcs. Cannot conclude root cause; defect may be intermittent or lot-specific.',
        'interp_ko':'동일 이벤트 baseline 부재 — ng_without_baseline 으로 저장. VP #11 tracking lot 은 50pcs 14 station 모두 VP damage 0. 원인 단정 불가, 간헐적 또는 lot 특이 가능.',
        'interp_vi':'Không có baseline cùng sự kiện — lưu ng_without_baseline. Lô tracking VP #11 không có VP damage trên 50pcs qua 14 trạm. Chưa thể kết luận nguyên nhân — có thể gián đoạn hoặc đặc thù lot.',
        'source_cells':['F22:W22','F26:AC26']}],
    'hints': [{'check_en':'Track VP damage on larger sample / multiple lots','reason_en':'50pcs/lot showed 0 NG — sample too small to confirm root cause.','check_ko':'VP damage 추적을 대량 lot 으로 확대.','reason_ko':'50pcs/lot 으로 0 NG — 근본 원인 확정에 표본 부족.','check_vi':'Theo dõi VP damage với mẫu lớn hơn / nhiều lot.','reason_vi':'50pcs/lot không có NG — mẫu quá nhỏ để khẳng định nguyên nhân.','evidence_strength':'low','related_process':'Sub4 tracking','related_part':'VP','source_cells':['B16']}],
    'log': {'rationale_en':'No same-event baseline; tracking sheet shows 0/50 at every station. Stored as ng_without_baseline — cannot claim improvement.','rationale_ko':'동일 이벤트 baseline 없음, 모든 station 0/50 — ng_without_baseline 으로 저장.','rationale_vi':'Không có baseline cùng sự kiện, mọi trạm 0/50 — ng_without_baseline.','assumptions_en':['Title says MSU-L20L15-07 although filename says BRS-161014 — model taken from sheet title.'],'assumptions_ko':['파일명은 BRS-161014 이나 시트 타이틀 기준 모델 MSU-L20L15-07 로 저장.'],'assumptions_vi':['Tên file BRS-161014 nhưng tiêu đề sheet MSU-L20L15-07 — dùng theo tiêu đề.']},
    'when_user_asks': ['MSU-L20L15-07 VP damage','VP damage tracking module'],
    'confidence': 0.5,
})


if __name__ == '__main__':
    ok = 0; fail = 0
    for d in D:
        try:
            if B.commit(d):
                print(f'OK  {d["name"]}'); ok += 1
            else:
                print(f'FAIL {d["name"]}'); fail += 1
        except Exception as e:
            print(f'ERR {d["name"]}: {e!r}'); fail += 1
    print(f'BATCH RESULT: ok={ok} fail={fail}')
