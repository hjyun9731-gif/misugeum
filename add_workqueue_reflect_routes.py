from pathlib import Path

p = Path("main.py")
s = p.read_text(encoding="utf-8")
Path("main.py.bak_workqueue_reflect").write_text(s, encoding="utf-8")

marker = '@app.get("/work"'
pos = s.find(marker)
if pos == -1:
    marker = '# ── 부과대수 관리'
    pos = s.find(marker)
if pos == -1:
    raise SystemExit("ERROR: /work 라우트 위치를 못 찾음")

block = r'''

# ── 업무처리/부과대수 반영 로직 ─────────────────────────────
def _safe_set(obj, name, value):
    try:
        if hasattr(obj, name):
            setattr(obj, name, value)
            return True
    except Exception:
        pass
    return False


def _append_member_note(member, text):
    try:
        old = member.note or ""
        if text and text not in old:
            member.note = (old + (" / " if old else "") + text)[-1000:]
    except Exception:
        pass


def _add_status_event_safe(db, member, wq, before_status, after_status, user_id):
    """
    MemberStatusEvent 컬럼이 프로젝트마다 달라도 최대한 안전하게 기록.
    실패해도 반영 자체는 막지 않음.
    """
    try:
        cols = set(MemberStatusEvent.__table__.columns.keys())
        kwargs = {}

        if "member_id" in cols:
            kwargs["member_id"] = member.id
        if "event_type" in cols:
            kwargs["event_type"] = getattr(wq, "process_type", "") or ""
        if "from_status" in cols:
            kwargs["from_status"] = before_status or ""
        if "to_status" in cols:
            kwargs["to_status"] = after_status or ""
        if "reason" in cols:
            kwargs["reason"] = getattr(wq, "reason", "") or ""
        if "note" in cols:
            kwargs["note"] = getattr(wq, "note", "") or ""
        if "created_by" in cols:
            kwargs["created_by"] = user_id
        if "source" in cols:
            kwargs["source"] = "업무처리반영"
        if "source_id" in cols:
            kwargs["source_id"] = getattr(wq, "id", None)

        if kwargs:
            db.add(MemberStatusEvent(**kwargs))
    except Exception:
        pass


def _reflect_workqueue_item(db: Session, wq, user: User):
    """
    WorkQueue 반영대기 1건을 실제 Member에 반영.
    반영 전까지는 회원 상태를 바꾸지 않고, 여기서 최종 적용한다.
    """
    if not wq:
        return False, "처리대기 항목 없음"

    member = db.query(Member).filter(Member.id == wq.member_id).first()
    if not member:
        return False, "회원 없음"

    if getattr(wq, "status", "") == "반영완료":
        return True, "이미 반영완료"

    ptype_raw = (getattr(wq, "process_type", "") or "").strip()
    ptype = ptype_raw.replace(" ", "")
    before_status = member.status or ""
    before_account = member.account or ""
    before_arrears = int(member.excel_arrears or 0)

    after_status = before_status
    msg = ""

    # 폐업/폐지/관리비폐지 계열
    if ptype in ("폐업", "폐지"):
        member.status = "폐업"
        member.status_source = "work"
        after_status = "폐업"
        msg = "폐업 처리"

    elif ptype in ("관리비폐지", "관리비폐업"):
        member.status = "관리비폐지"
        member.status_source = "work"
        after_status = "관리비폐지"
        msg = "관리비폐지 처리"

    # 양도
    elif ptype in ("양도", "폐업양도", "폐지양도"):
        member.status = "양도"
        member.status_source = "work"
        after_status = "양도"
        msg = "양도 처리"

    # 이관/타도
    elif ptype in ("이관", "타도", "타도이전", "타도전출"):
        member.status = "이관"
        member.status_source = "work"
        after_status = "이관"
        msg = "이관/타도 처리"

    # 탈퇴
    elif ptype == "탈퇴":
        member.status = "탈퇴"
        member.status_source = "work"
        after_status = "탈퇴"
        msg = "탈퇴 처리"

    # 사망
    elif ptype == "사망":
        member.status = "사망"
        member.status_source = "work"
        after_status = "사망"
        msg = "사망 처리"

    # 말소
    elif ptype == "말소":
        member.status = "말소"
        member.status_source = "work"
        after_status = "말소"
        msg = "말소 처리"

    # 택배신규: 관리비/택배관리 대상으로 편입
    elif ptype == "택배신규":
        member.status = "정상"
        member.status_source = "work"
        member.account = "관"
        after_status = "정상"
        msg = "택배신규/관리비 대상 편입"

    # 70세: 부과대상 제외 아님. 협회비 10,000 → 5,000 감액 대상.
    elif ptype in ("70세", "70세감액", "70세이상"):
        member.status = "정상"
        member.status_source = "work"
        _safe_set(member, "age70", True)
        _append_member_note(member, "70세 감액대상: 협회비 10,000원 → 5,000원")
        after_status = "정상"
        msg = "70세 감액 처리"

    # 현역복구/검증해제
    elif ptype in ("현역복구", "검증해제", "정상복구"):
        member.status = "정상"
        member.status_source = "work"
        after_status = "정상"
        msg = "정상 복구"

    # 금액수정은 기존 별도 라우트가 있으면 그쪽 사용. 여기서는 막음.
    elif ptype == "금액수정":
        return False, "금액수정은 금액수정 반영 버튼을 사용하세요"

    else:
        return False, f"알 수 없는 처리구분: {ptype_raw}"

    # 미수금은 삭제하지 않음. 기존 미수금은 정산용으로 보존.
    # 다만 상태가 정상 외로 바뀌면 _clean_filter 기준에서 부과대상 제외됨.
    try:
        member.is_overpay = (member.excel_arrears or 0) < 0
    except Exception:
        pass

    db.add(member)

    # WorkQueue 완료 처리
    wq.status = "반영완료"
    _safe_set(wq, "reflected_by", user.id)
    _safe_set(wq, "reflected_at", datetime.now())
    _safe_set(wq, "result_note", msg)
    db.add(wq)

    _add_status_event_safe(db, member, wq, before_status, after_status, user.id)

    try:
        add_log(
            db,
            user.id,
            "업무처리반영",
            f"{member.name}/{member.vehicle_no}: {ptype_raw} / 상태 {before_status}→{after_status} / 계정 {before_account}→{member.account} / 미수금 {before_arrears:,}원 보존"
        )
    except Exception:
        pass

    return True, msg


@app.post("/workorder/{wid}/reflect")
def workorder_reflect_api(
    wid: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    """
    work.html의 JS가 호출하는 API:
    fetch('/workorder/' + woId + '/reflect', {method:'POST'})
    """
    try:
        wq = db.query(WorkQueue).filter(WorkQueue.id == wid).first()
        ok, msg = _reflect_workqueue_item(db, wq, user)

        if ok:
            db.commit()
            _invalidate_snap(db, "dashboard")
            return {"ok": True, "msg": msg}

        db.rollback()
        return {"ok": False, "msg": msg}

    except Exception as e:
        db.rollback()
        return {"ok": False, "msg": "반영 오류: " + str(e)[:180]}


@app.post("/workorder/reflect_all")
def workorder_reflect_all_api(
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    """
    현재 반영대기 전체 반영.
    금액수정은 별도 반영 버튼이 있으므로 제외.
    """
    try:
        items = (
            db.query(WorkQueue)
            .filter(WorkQueue.status == "반영대기")
            .filter(WorkQueue.process_type != "금액수정")
            .order_by(WorkQueue.id.asc())
            .all()
        )

        ok_count = 0
        fail_count = 0
        messages = []

        for wq in items:
            ok, msg = _reflect_workqueue_item(db, wq, user)
            if ok:
                ok_count += 1
            else:
                fail_count += 1
                messages.append(f"{wq.id}:{msg}")

        db.commit()
        _invalidate_snap(db, "dashboard")

        return {
            "ok": True,
            "msg": f"전체반영 완료 {ok_count}건 / 실패 {fail_count}건" + ((" / " + "; ".join(messages[:5])) if messages else "")
        }

    except Exception as e:
        db.rollback()
        return {"ok": False, "msg": "전체반영 오류: " + str(e)[:180]}


@app.post("/work/{wid}/reflect")
def work_reflect_redirect(
    wid: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    """
    혹시 form 방식 버튼이 있는 경우를 위한 redirect 라우트.
    """
    from urllib.parse import quote

    try:
        wq = db.query(WorkQueue).filter(WorkQueue.id == wid).first()
        ok, msg = _reflect_workqueue_item(db, wq, user)

        if ok:
            db.commit()
            _invalidate_snap(db, "dashboard")
            return RedirectResponse("/work?tab=반영완료&msg=" + quote("반영완료: " + msg), status_code=302)

        db.rollback()
        return RedirectResponse("/work?tab=반영대기&msg=" + quote("반영실패: " + msg), status_code=302)

    except Exception as e:
        db.rollback()
        return RedirectResponse("/work?tab=반영대기&msg=" + quote("반영 오류: " + str(e)[:180]), status_code=302)

'''

if "/workorder/{wid}/reflect" not in s:
    s = s[:pos] + block + "\n" + s[pos:]
    print("OK: workqueue reflect routes inserted")
else:
    print("INFO: workqueue reflect routes already exist")

p.write_text(s, encoding="utf-8")
print("DONE")
