"""음식과사람 배포내역: observed food/deploy/list grid contract."""
from __future__ import annotations

import re
import random
from datetime import date, datetime
from nmis_slip_automation import go_to_angular_state

MEMBER_TYPES = {"전체": "", "정회원": "1", "비회원": "3"}
SCOPE = "angular.element(document.querySelector('select[name=memberType]')).scope()"


def validate_request(member_type, count, distribution_date, year_month):
    if member_type not in MEMBER_TYPES:
        raise ValueError("회원구분은 정회원·비회원·전체 중 선택하세요.")
    text = str(count).strip()
    if not text.isdigit() or not 1 <= int(text) <= 50000:
        raise ValueError("배포 부수는 1~50000 사이의 정수로 입력하세요.")
    raw_date = re.sub(r"[-./]", "", str(distribution_date).strip())
    raw_month = re.sub(r"[-./]", "", str(year_month).strip())
    if len(raw_date) != 8 or len(raw_month) != 6:
        raise ValueError("배포일자는 YYYYMMDD, 기준년월은 YYYYMM 형식으로 입력하세요.")
    parsed = datetime.strptime(raw_date, "%Y%m%d").date()
    datetime.strptime(raw_month + "01", "%Y%m%d")
    if parsed > date.today():
        raise ValueError("NMIS에서는 미래 배포일자를 등록할 수 없습니다.")
    return int(text), raw_date, raw_month


def choose_rows(rows, member_type, count, *, rng=None, allowed_ids=None):
    allowed = None if allowed_ids is None else set(map(str, allowed_ids))
    eligible = [r for r in rows if (allowed is None or str(r.get("memberId")) in allowed) and r.get("distributionYn") != "Y"
                and not r.get("distributionDate")
                and (member_type == "전체" or r.get("memberTypeName") == member_type)]
    if len(eligible) < count:
        raise ValueError(f"미배포 대상이 부족합니다: 요청 {count}부 / 가능 {len(eligible)}부")
    ids = [str(r.get("memberId") or "") for r in eligible]
    if any(not key for key in ids) or len(set(ids)) != len(ids):
        raise ValueError("선택 대상의 회원 ID가 없거나 중복되어 중단했습니다.")
    return (rng or random.SystemRandom()).sample(ids, count)


def navigate(page):
    if not page.url.startswith("http://nmis.foodservice.or.kr/"):
        raise RuntimeError("NMIS 페이지에 연결되지 않았습니다.")
    go_to_angular_state(page, "food/deploy/list")
    page.locator("select[name='memberType']:visible").wait_for(timeout=15000)


def read_state(page):
    return page.evaluate("""() => {
        const s = angular.element(document.querySelector('select[name=memberType]')).scope();
        return {month:s.searchParams.stdYearMonth, memberType:s.searchParams.memberType || '',
          rows:s.gridDatafooddeploylist.map(r=>({memberId:r.memberId, memberName:r.memberName, ceoMemberName:r.ceoMemberName, memberTypeName:r.memberTypeName,
            distributionYn:r.distributionYn, distributionDate:r.distributionDate, selected:r.selected})),
          total:Number(s.gridDatafooddeploylist[0]?.totCnt || s.gridDatafooddeploylist.length)};
    }""")


APPLY_ROWS_JS = """({ids, rawDate, datesById, rawMonth, memberCode}) => {
    const s = angular.element(document.querySelector('select[name=memberType]')).scope();
    if (s.searchParams.stdYearMonth !== rawMonth || (s.searchParams.memberType || '') !== memberCode)
        throw Error('조회 조건이 바뀌었습니다. 다시 실행하세요.');
    const rows=s.gridDatafooddeploylist, wanted=new Set(ids);
    const targets=rows.filter(r=>wanted.has(String(r.memberId)));
    if(targets.length!==ids.length || targets.some(r=>r.distributionYn==='Y'||r.distributionDate))
        throw Error('대상 문서가 바뀌었습니다. 다시 조회하세요.');
    if(targets.some(r=>!(datesById?.[String(r.memberId)] || rawDate)))
        throw Error('배포일자가 없는 대상이 있습니다.');
    s.$apply(()=>{
        for(const r of rows) r.selected='false';
        for(const r of targets){r.selected='true';r.distributionYn='Y';r.distributionDate=datesById?.[String(r.memberId)] || rawDate;}
    });
    s.gridOptionfooddeploylistApi.core.notifyDataChange('edit');
    return targets.length;
}"""


def lookup_distribution(page, member_type, year_month, *, only_undistributed=False):
    _, _, raw_month = validate_request(member_type, 1, date.today().strftime("%Y%m%d"), year_month)
    navigate(page)
    # Do not discard a manually selected, unfinished batch.
    if any(str(r.get("selected")).lower() == "true" for r in read_state(page)["rows"]):
        raise RuntimeError("사이트에 이미 선택한 행이 있습니다. 저장하거나 선택을 해제한 뒤 실행하세요.")
    page.locator("select[name='memberType']:visible").select_option(label=member_type)
    page.locator("select[name='distributionType']:visible").select_option(label="미배포" if only_undistributed else "전체")
    month = page.locator("input[name='stdYearMonth']:visible")
    month.fill(raw_month)
    month.press("Tab")
    page.wait_for_timeout(500)
    visible_month = re.sub(r"\D", "", month.input_value())
    if visible_month != raw_month:
        raise RuntimeError("기준년월 입력값을 확인하지 못했습니다.")
    # Increase the real grid's page size to avoid selecting only rendered rows.
    page.evaluate("""() => {
        const s=angular.element(document.querySelector('select[name=memberType]')).scope();
        s.__distributionPrevious=s.gridDatafooddeploylist;
        s.$apply(()=>{s.gridOptionfooddeploylist.paginationPageSize=50000;
            s.gridOptionfooddeploylist.paginationCurrentPage=1;s.currentPage=1;});
    }""")
    page.locator("button:visible").filter(has=page.locator("span[lang-code='search']")).click()
    page.wait_for_function("""() => {
        const s=angular.element(document.querySelector('select[name=memberType]')).scope();
        return s.gridDatafooddeploylist!==s.__distributionPrevious &&
            angular.element(document.body).injector().get('$http').pendingRequests.length===0;
    }""", timeout=30000)
    state = read_state(page)
    if state["total"] > len(state["rows"]):
        raise RuntimeError("조회 전체 목록을 읽지 못했습니다. 일부만 등록하지 않고 중단합니다.")
    return state


def prepare_distribution(page, member_type, count, distribution_date, year_month, log=print, *, allowed_ids=None):
    count, raw_date, raw_month = validate_request(member_type, count, distribution_date, year_month)
    state = lookup_distribution(page, member_type, raw_month, only_undistributed=True)
    ids = choose_rows(state["rows"], member_type, count, allowed_ids=allowed_ids)
    applied = page.evaluate(APPLY_ROWS_JS, dict(ids=ids, rawDate=raw_date,
                                               rawMonth=raw_month, memberCode=MEMBER_TYPES[member_type]))
    current = read_state(page)
    verify_prepared(current, ids, raw_date, raw_month, MEMBER_TYPES[member_type])
    log(f"배포 입력 완료: {member_type} {applied}부 / 기준 {raw_month} / 배포일자 {raw_date}")
    return dict(ids=ids, raw_date=raw_date, raw_month=raw_month, member_code=MEMBER_TYPES[member_type])


def saved_groups_plan(groups, dates, selected, member_type, year_month):
    if not selected or len(set(selected)) != len(selected) or any(i not in range(5) for i in selected):
        raise ValueError("처리할 배포1~배포5를 하나 이상 체크하세요.")
    if len(groups) != 5 or len(dates) != 5:
        raise ValueError("저장 명단 데이터를 확인하세요.")
    date_map, details = {}, []
    raw_month = ""
    for index in selected:
        if not groups[index]:
            raise ValueError(f"배포{index+1}에 저장한 명단이 없습니다.")
        try:
            _, raw_date, raw_month = validate_request(member_type, 1, dates[index], year_month)
        except ValueError as exc:
            raise ValueError(f"배포{index+1}: {exc}") from exc
        members = [str(key) for key in groups[index]]
        if any(not key or key in date_map for key in members):
            raise ValueError("선택한 배포 명단에 중복되거나 비어 있는 회원 ID가 있습니다.")
        date_map.update({key: raw_date for key in members})
        details.append(dict(index=index, ids=members, date=raw_date))
    return dict(raw_month=raw_month, dates_by_id=date_map, details=details)


def prepare_saved_distribution(page, member_type, groups, dates, selected, year_month, log=print):
    plan = saved_groups_plan(groups, dates, selected, member_type, year_month)
    state = lookup_distribution(page, member_type, plan["raw_month"], only_undistributed=False)
    candidates = [r for r in state["rows"] if str(r.get("memberId")) in plan["dates_by_id"]
                  and r.get("distributionYn") != "Y" and not r.get("distributionDate")
                  and (member_type == "전체" or r.get("memberTypeName") == member_type)]
    ids = [str(r["memberId"]) for r in candidates]
    if not ids:
        raise ValueError("선택한 저장 명단에 이번 조회 조건의 미배포 대상이 없습니다.")
    if len(set(ids)) != len(ids):
        raise ValueError("조회 목록에 회원 ID가 중복되어 중단했습니다.")
    date_map = {key: plan["dates_by_id"][key] for key in ids}
    details = []
    for group in plan["details"]:
        ready = sum(key in date_map for key in group["ids"])
        skipped = len(group["ids"]) - ready
        line = f"배포{group['index']+1}: {ready}곳 / {group['date']} / 이미 배포·조회 제외 {skipped}곳"
        details.append(line)
        log(line)
    page.evaluate(APPLY_ROWS_JS, dict(ids=ids, rawDate="", datesById=date_map,
                                     rawMonth=plan["raw_month"], memberCode=MEMBER_TYPES[member_type]))
    verify_prepared(read_state(page), ids, date_map, plan["raw_month"], MEMBER_TYPES[member_type])
    return dict(ids=ids, raw_date="", dates_by_id=date_map, raw_month=plan["raw_month"],
                member_code=MEMBER_TYPES[member_type], summary="\n".join(details))


def verify_prepared(state, ids, raw_date, raw_month, member_code):
    selected = [r for r in state["rows"] if str(r.get("selected")).lower() == "true"]
    if (state["month"] != raw_month or state["memberType"] != member_code
        or {str(r["memberId"]) for r in selected} != set(ids)
        or len(selected) != len(ids)
        or any(r.get("distributionYn") != "Y" or r.get("distributionDate") !=
               (raw_date.get(str(r["memberId"])) if isinstance(raw_date, dict) else raw_date) for r in selected)):
        raise RuntimeError("선택 수·배포일자 또는 조회 조건이 바뀌었습니다. 다시 입력하세요.")


def save_distribution(page, prepared, log=print):
    verify_prepared(read_state(page), prepared["ids"], prepared.get("dates_by_id", prepared["raw_date"]),
                    prepared["raw_month"], prepared["member_code"])
    page.locator("button:visible").filter(has=page.locator("span[lang-code='create']")).click()
    # Only accept the observed confirmation modal, never a generic page button.
    confirm = page.locator(".modal:visible button[ng-click*='fnConfirm']")
    confirm.wait_for(timeout=10000)
    confirm.click()
    done = page.locator(".modal:visible button[ng-click*='fnClose']")
    done.wait_for(timeout=15000)
    modal_text = done.locator("xpath=ancestor::*[contains(@class,'modal-content')][1]").inner_text()
    if not any(word in modal_text for word in ("등록되었습니다", "저장되었습니다", "처리되었습니다")):
        raise RuntimeError(f"등록 성공을 확인하지 못했습니다: {modal_text}")
    done.click()
    log(f"배포 등록 완료: {len(prepared['ids'])}부")
