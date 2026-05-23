"""
강원도개인소형화물협회 통합관리시스템 v4
"""
import re, json, math, shutil, os
from pathlib import Path
from datetime import datetime, date
from typing import Optional, List
import tempfile

from fastapi import FastAPI, Request, Depends, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from sqlalchemy.orm import Session
from sqlalchemy import func, or_, and_, text, not_
import pandas as pd

from database import engine, SessionLocal, Base, get_db, IS_SQLITE
from models import (User, UploadBatch, RawImportRow, Member, MonthlyLedger,
                    MemberStatusEvent, WorkQueue, LicenseRecord, BillingPerson,
                    BankTransaction, CollectionTarget, Snapshot, AuditLog, BillingReport)
from auth import hash_pw, verify_pw, ensure_admin, require_user
from core import (
    _s, norm_name, norm_vehicle, veh_last4, normalize_region, parse_amount,
    parse_date_str, detect_status, detect_status_from_mgmt_no,
    build_verify_reasons, phone_clean, address_clean,
    guess_sheet_type, guess_file_year,
    parse_ledger_sheet, parse_status_sheet,
    extract_memo_keys, is_useless_memo, score_bank_match,
    classify_autopay, AUTOPAY_AMOUNTS
)

app = FastAPI(title="강원도개인소형화물협회 v4")
app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SECRET_KEY", "gwf4-secret-2026"),
    max_age=3600 * 12
)
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

CURRENT_YEAR = int(os.getenv("BILLING_YEAR", str(datetime.now().year)))
EXCLUDED_STATUSES = {"폐업","양도","이관","탈퇴","사망","말소","확인필요"}

# ── 스타트업 ──────────────────────────────────────────────────────────────────
@app.on_event("startup")
def startup():
    if os.getenv("RESET_DB") == "1":
        Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    _migrate()
    db = SessionLocal()
    try: ensure_admin(db)
    finally: db.close()

def _migrate():
    with engine.begin() as conn:
        be = engine.url.get_backend_name()
        if be.startswith("postgres"):
            for tbl, col, typ in [
                ("members","is_overpay","BOOLEAN DEFAULT FALSE"),
                ("members","arrears_diff","INTEGER DEFAULT 0"),
                ("members","verify_reason","TEXT"),
                ("members","status_source","VARCHAR(50) DEFAULT 'auto'"),
                ("members","user_confirmed_match","BOOLEAN DEFAULT FALSE"),
                ("upload_batches","file_year","INTEGER"),
                ("license_records","source_sheet","VARCHAR(200)"),
            ]:
                conn.execute(text(f"ALTER TABLE {tbl} ADD COLUMN IF NOT EXISTS {col} {typ}"))
        elif IS_SQLITE:
            def _ac(tbl, col, typ):
                try:
                    cols = [r[1] for r in conn.execute(text(f"PRAGMA table_info({tbl})")).fetchall()]
                    if col not in cols:
                        conn.execute(text(f"ALTER TABLE {tbl} ADD COLUMN {col} {typ}"))
                except: pass
            for tbl, col, typ in [
                ("members","is_overpay","INTEGER DEFAULT 0"),
                ("members","arrears_diff","INTEGER DEFAULT 0"),
                ("members","verify_reason","TEXT"),
                ("members","status_source","TEXT DEFAULT 'auto'"),
                ("members","user_confirmed_match","INTEGER DEFAULT 0"),
                ("upload_batches","file_year","INTEGER"),
                ("license_records","source_sheet","TEXT"),
            ]:
                _ac(tbl, col, typ)

# ── 유틸 ──────────────────────────────────────────────────────────────────────
@app.get("/health")
def health(): return {"status": "ok"}

def add_log(db, user_id, action, detail=""):
    db.add(AuditLog(user_id=user_id, action=action, detail=str(detail)[:500]))
    db.commit()

def fmt_amt(v) -> str:
    if v is None: return "0"
    try:
        iv = int(v)
        if iv < 0: return f"-{abs(iv):,}"
        return f"{iv:,}"
    except: return str(v)


def fmt_date(v) -> str:
    """날짜를 26.05.18. 형식으로 표시"""
    if not v: return ""
    s = str(v).strip()
    import re as _re
    # 2026-05-18 → 26.05.18.
    m = _re.search(r"(20(\d{2}))-(\d{2})-(\d{2})", s)
    if m: return f"{m.group(2)}.{m.group(3)}.{m.group(4)}."
    # 이미 26.05.18. 형식
    if _re.match(r"\d{2}\.\d{2}\.\d{2}\.", s): return s
    return s

def fmt_acc(a) -> str: return {"협":"협회비","관":"관리비"}.get(a or "", a or "")

def _open_xl(path):
    p = Path(path)
    if p.suffix.lower() in {".xlsx",".xlsm"}:
        return pd.ExcelFile(path, engine="openpyxl")
    return pd.ExcelFile(path, engine="xlrd")

# ── 인증 ──────────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def root(request: Request):
    return RedirectResponse("/dashboard" if request.session.get("user_id") else "/login")

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse(request, "login.html", {"request": request, "error": ""})

@app.post("/login")
def login(request: Request, username: str = Form(""), password: str = Form(""),
          db: Session = Depends(get_db)):
    u = db.query(User).filter(User.username == username).first()
    if not u or not verify_pw(password, u.password_hash):
        return templates.TemplateResponse(request, "login.html",
            {"request": request, "error": "아이디 또는 비밀번호가 올바르지 않습니다"})
    request.session["user_id"] = u.id
    return RedirectResponse("/dashboard", status_code=302)

@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login")

# ── 대시보드 ──────────────────────────────────────────────────────────────────
@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db),
              user: User = Depends(require_user)):
    # 스냅샷에서 로드
    snap = _get_snap(db, "dashboard")
    if not snap:
        snap = _build_dashboard_snap(db)
        _set_snap(db, "dashboard", snap)

    recent_logs = db.query(AuditLog).order_by(AuditLog.created_at.desc()).limit(10).all()
    return templates.TemplateResponse(request, "dashboard.html", {
        "request": request, "user": user, "snap": snap,
        "recent_logs": recent_logs, "fmt_amt": fmt_amt,
        "CURRENT_YEAR": CURRENT_YEAR,
    })


SUM_NAMES_DB = {"합계","총계","소계","계","합산","인원수","입금금액"}

def _real_member_q(db: Session):
    """합계행 제외한 실제 회원 쿼리"""
    return db.query(Member).filter(
        ~Member.vehicle_no.in_(list(SUM_NAMES_DB)),
        ~Member.name_key.in_(list(SUM_NAMES_DB)),
        Member.name != None,
        Member.name != "",
    )

def _clean_filter():
    """합계행 제외 + 상태=정상 필터"""
    return and_(
        Member.status == "정상",
        ~Member.vehicle_no.in_(list(SUM_NAMES_DB)),
        ~Member.name_key.in_(list(SUM_NAMES_DB)),
        Member.name != None,
        Member.name != "",
    )

def _build_dashboard_snap(db: Session) -> dict:
    now = datetime.now()
    cf = _clean_filter()
    total      = db.query(Member).filter(cf).count()
    total_arr  = db.query(func.sum(Member.excel_arrears)).filter(cf, Member.excel_arrears > 0).scalar() or 0
    overpay_sum= db.query(func.sum(Member.excel_arrears)).filter(cf, Member.excel_arrears < 0).scalar() or 0
    verify_cnt = db.query(Member).filter(cf, Member.arrears_verified == False).count()
    no_lic     = db.query(Member).filter(cf,
                    Member.user_confirmed_match == False,
                    or_(Member.match_license_id == None, Member.match_status == "전체자미확인")).count()
    coll_cnt   = db.query(CollectionTarget).filter(CollectionTarget.excluded == False).count()
    overpay_cnt= db.query(Member).filter(cf, Member.is_overpay == True).count()
    work_pending = db.query(WorkQueue).filter(WorkQueue.status == "반영대기").count()
    bank_pending = db.query(BankTransaction).filter(
        BankTransaction.applied == False,
        BankTransaction.match_status.in_(["자동매칭","확인필요"])).count()
    status_counts = {s: db.query(Member).filter(Member.status == s).count()
                     for s in ["폐업","양도","이관","탈퇴","사망","말소"]}
    return {
        "base_year": now.year, "base_month": now.month,
        "total_members": total, "total_arrears": int(total_arr),
        "overpay_sum": int(abs(overpay_sum)),
        "verify_cnt": verify_cnt, "no_lic": no_lic, "coll_cnt": coll_cnt,
        "overpay_cnt": overpay_cnt, "work_pending": work_pending,
        "bank_pending": bank_pending, "status_counts": status_counts,
    }

def _get_snap(db: Session, key: str):
    s = db.query(Snapshot).filter(Snapshot.snap_key == key).first()
    if not s: return None
    try: return json.loads(s.snap_value)
    except: return None

def _set_snap(db: Session, key: str, data: dict):
    s = db.query(Snapshot).filter(Snapshot.snap_key == key).first()
    val = json.dumps(data, ensure_ascii=False, default=str)
    if s: s.snap_value = val
    else: db.add(Snapshot(snap_key=key, snap_value=val))
    db.commit()

def _invalidate_snap(db: Session, *keys: str):
    for key in keys:
        db.query(Snapshot).filter(Snapshot.snap_key == key).delete()
    db.commit()

# ── 미수금 명단 ────────────────────────────────────────────────────────────────
@app.get("/arrears", response_class=HTMLResponse)
def arrears_page(request: Request, q: str = "", region: str = "",
                 account: str = "", tab: str = "전체",
                 page: int = 1, per_page: int = 200,
                 db: Session = Depends(get_db), user: User = Depends(require_user)):
    per_page = min(max(per_page, 30), 500)
    page = max(page, 1)
    base_q = db.query(Member).filter(_clean_filter())  # 합계행 제외
    if q:
        like = f"%{q}%"
        l4 = "".join(c for c in q if c.isdigit())[-4:]
        flt = [Member.name.ilike(like), Member.vehicle_no.ilike(like)]
        if l4: flt.append(Member.vehicle_no.ilike(f"%{l4}%"))
        base_q = base_q.filter(or_(*flt))
    if region: base_q = base_q.filter(Member.region == region)
    if account: base_q = base_q.filter(Member.account == account)

    if tab == "문자후보":
        base_q = base_q.filter(Member.excel_arrears > 0, Member.is_overpay == False,
                                Member.is_auto_transfer == False, Member.mobile != None,
                                Member.mobile != "", Member.arrears_verified == True)
    elif tab == "검증필요":
        base_q = base_q.filter(Member.arrears_verified == False)
    elif tab == "연락처없음":
        base_q = base_q.filter(or_(Member.mobile == None, Member.mobile == ""))
    elif tab == "자동이체정기":
        base_q = base_q.filter(Member.is_auto_transfer == True)
    elif tab == "초과납부선납":
        base_q = base_q.filter(Member.is_overpay == True)
    else:
        base_q = base_q.filter(Member.excel_arrears != 0)

    PAGE_SIZE = 200
    total_count = base_q.count()
    total_pages = max(1, (total_count + PAGE_SIZE - 1) // PAGE_SIZE)
    page = max(1, min(int(request.query_params.get("page", 1)), total_pages))
    items = base_q.order_by(Member.region, Member.name).offset((page-1)*PAGE_SIZE).limit(PAGE_SIZE).all()
    total_arr = db.query(func.sum(Member.excel_arrears)).filter(
        Member.status == "정상", Member.excel_arrears > 0).scalar() or 0
    regions = sorted({x[0] for x in db.query(Member.region).distinct().filter(Member.region != None).all() if x[0]})

    tab_counts = {
        "전체":      db.query(Member).filter(Member.status == "정상", Member.excel_arrears != 0).count(),
        "문자후보":  db.query(Member).filter(Member.status == "정상", Member.excel_arrears > 0,
                              Member.is_overpay == False, Member.is_auto_transfer == False,
                              Member.arrears_verified == True).count(),
        "검증필요":  db.query(Member).filter(Member.status == "정상", Member.arrears_verified == False).count(),
        "연락처없음":db.query(Member).filter(Member.status == "정상",
                              or_(Member.mobile == None, Member.mobile == "")).count(),
        "자동이체정기":db.query(Member).filter(Member.status == "정상", Member.is_auto_transfer == True).count(),
        "초과납부선납":db.query(Member).filter(Member.status == "정상", Member.is_overpay == True).count(),
    }

    return templates.TemplateResponse(request, "arrears.html", {
        "request": request, "user": user, "items": items, "q": q,
        "region": region, "account": account, "tab": tab,
        "total_arr": int(total_arr), "regions": regions,
        "tab_counts": tab_counts, "fmt_amt": fmt_amt, "fmt_acc": fmt_acc,
        "msg": request.query_params.get("msg", ""),
        "page": page, "per_page": per_page,
        "total_pages": total_pages, "total_count": total_count,
    })

# 부과대수 관리으로 보내기
@app.post("/arrears/{mid}/to-work")
def to_work_queue(mid: int, process_type: str = Form("폐업"),
                  reason: str = Form(""), event_date: str = Form(""),
                  db: Session = Depends(get_db), user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    db.add(WorkQueue(
        member_id=mid, process_type=process_type, status="반영대기",
        source_screen="미수금명단", reason=reason,
        event_date_raw=event_date, arrears_at_submit=m.excel_arrears or 0,
        submitted_by=user.id
    ))
    db.commit()
    _invalidate_snap(db, "dashboard")
    add_log(db, user.id, "업무처리대기등록", f"{m.name}/{m.vehicle_no} → {process_type}")
    return RedirectResponse(f"/arrears?msg={m.name} 부과대수 관리으로 이동", status_code=302)

# ── 회원 상세 ──────────────────────────────────────────────────────────────────
@app.get("/member/{mid}", response_class=HTMLResponse)
def member_detail(mid: int, request: Request, db: Session = Depends(get_db),
                  user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    ledgers = (db.query(MonthlyLedger).filter(MonthlyLedger.member_id == mid)
               .order_by(MonthlyLedger.year, MonthlyLedger.month).all())
    # 중복 원장 감지
    seen = set(); dups = 0
    deduped = []
    for l in ledgers:
        key = (l.year, l.month, l.source_sheet)
        if key in seen: dups += 1; l.is_duplicate = True
        else: seen.add(key); deduped.append(l)

    events = (db.query(MemberStatusEvent).filter(MemberStatusEvent.member_id == mid)
              .order_by(MemberStatusEvent.created_at).all())
    lic = db.query(LicenseRecord).filter(LicenseRecord.id == m.match_license_id).first() if m.match_license_id else None
    work_history = (db.query(WorkQueue).filter(WorkQueue.member_id == mid)
                    .order_by(WorkQueue.submitted_at.desc()).limit(10).all())

    return templates.TemplateResponse(request, "member_detail.html", {
        "request": request, "user": user, "m": m,
        "ledgers": deduped, "dup_count": dups, "events": events,
        "lic": lic, "work_history": work_history,
        "fmt_amt": fmt_amt, "fmt_acc": fmt_acc,
        "msg": request.query_params.get("msg", ""),
    })

@app.get("/member/{mid}/edit", response_class=HTMLResponse)
def member_edit_page(mid: int, request: Request, db: Session = Depends(get_db),
                     user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    regions = sorted({x[0] for x in db.query(Member.region).distinct().filter(Member.region != None).all() if x[0]})
    return templates.TemplateResponse(request, "edit_member.html",
        {"request": request, "user": user, "m": m, "regions": regions, "fmt_acc": fmt_acc})

@app.post("/member/{mid}/edit")
def member_edit_save(mid: int,
    name: str = Form(""), vehicle_no: str = Form(""), region: str = Form(""),
    account: str = Form("협"), mobile: str = Form(""), phone: str = Form(""),
    address: str = Form(""), official_address: str = Form(""),
    status: str = Form("정상"), note: str = Form(""),
    join_date_raw: str = Form(""), permit_date_raw: str = Form(""),
    cert_issue_date_raw: str = Form(""), cert_no: str = Form(""),
    db: Session = Depends(get_db), user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    m.name = name; m.name_key = norm_name(name)
    m.vehicle_no = vehicle_no; m.vehicle_key = norm_vehicle(vehicle_no)
    m.region = region; m.account = account
    m.mobile = phone_clean(mobile) or mobile; m.phone = phone
    m.address = address; m.official_address = official_address
    m.status = status; m.note = note
    m.join_date_raw = join_date_raw; m.permit_date_raw = permit_date_raw
    m.cert_issue_date_raw = cert_issue_date_raw; m.cert_no = cert_no
    db.commit()
    _invalidate_snap(db, "dashboard")
    add_log(db, user.id, "회원수정", f"id={mid} {name}")
    return RedirectResponse(f"/member/{mid}?msg=저장완료", status_code=302)

# ── 엑셀 업로드 ────────────────────────────────────────────────────────────────
@app.get("/upload", response_class=HTMLResponse)
def upload_page(request: Request, db: Session = Depends(get_db),
                user: User = Depends(require_user)):
    batches = db.query(UploadBatch).order_by(UploadBatch.created_at.desc()).limit(20).all()
    latest = {dt: db.query(UploadBatch).filter(UploadBatch.data_type == dt)
              .order_by(UploadBatch.created_at.desc()).first()
              for dt in ["legacy","license","billing","bank","bank_auto"]}
    return templates.TemplateResponse(request, "upload.html", {
        "request": request, "user": user, "batches": batches, "latest": latest,
        "msg": request.query_params.get("msg", ""),
    })

def _find_or_create_member(db, veh, name, region, account, note, batch_id, src_file, src_sheet, src_row):
    from core import is_sum_row
    if is_sum_row(name, veh): return None, "sum_row"
    vkey = norm_vehicle(veh) if veh else ""
    nkey = norm_name(name) if name else ""
    reg  = normalize_region(region) if region else ""
    acc  = "관" if ("관" in (account or "") or "관리" in (account or "")) else "협"
    is_auto = False  # 자동이체는 전용통장 업로드 후에만 확정

    if vkey and nkey:
        m = db.query(Member).filter(Member.vehicle_key == vkey, Member.name_key == nkey).first()
        if m: return m, "exact"
    if vkey:
        cs = db.query(Member).filter(Member.vehicle_key == vkey).all()
        if len(cs) == 1: return cs[0], "vehicle"
        if len(cs) > 1:  return None, "dup_vehicle"
    if nkey and reg:
        cs = db.query(Member).filter(Member.name_key == nkey, Member.region == reg, Member.account == acc).all()
        if len(cs) == 1: return cs[0], "name_region"
        if len(cs) > 1:  return None, "dup_name"

    m = Member(vehicle_no=veh, vehicle_key=vkey, name=name, name_key=nkey,
               region=reg, account=acc, note=note, is_auto_transfer=is_auto,
               status="정상", status_source="auto",
               source_file=src_file, source_sheet=src_sheet,
               source_row=src_row, source_batch_id=batch_id)
    db.add(m); db.flush()
    return m, "created"

def _recalc_member(db: Session, m: Member):
    """
    회원 미수금 갱신.
    핵심 규칙:
    - excel_arrears = 엑셀 기재 기준월 미수금 (우선값, 음수 그대로 유지)
    - calc_arrears  = excel_arrears (재계산 누적 완전 중단)
    - arrears_diff  = 0
    - is_overpay    = excel_arrears < 0
    - verify_reason = 실제 사유 있을 때만 (차이 0원 금지)
    """
    ledgers = (db.query(MonthlyLedger).filter(MonthlyLedger.member_id == m.id)
               .order_by(MonthlyLedger.year, MonthlyLedger.month).all())

    if ledgers:
        # 기준: 최신월 원장의 arrears_amount (음수 그대로)
        latest = ledgers[-1]
        excel_arr = latest.arrears_amount
        if excel_arr is None:
            excel_arr = m.excel_arrears or 0
        m.excel_arrears = excel_arr
        # 최초 미납월
        m.first_unpaid_month = ""
        for l in ledgers:
            if (l.arrears_amount or 0) > 0:
                m.first_unpaid_month = f"{l.year}-{l.month:02d}"
                break
        # 최근 입금일
        paid_ls = [l for l in reversed(ledgers) if (l.paid_amount or 0) > 0 and l.paid_date]
        if paid_ls:
            m.last_paid_date = paid_ls[0].paid_date
    else:
        excel_arr = m.excel_arrears or 0

    # 재계산 누적 완전 중단 — calc = excel
    m.calc_arrears = excel_arr
    m.arrears_diff = 0
    m.is_overpay   = excel_arr < 0

    # 검증사유: 실제 문제 있을 때만 (차이 0원이므로 금액차이는 없음)
    reasons = build_verify_reasons(
        excel_arr, excel_arr,  # diff = 0 → 금액차이 사유 절대 생성 안 됨
        m.name or "", m.vehicle_no or "", m.account or "",
        m.match_status or "", m.region or "",
        m.mobile or "", m.address or "", m.note or ""
    )
    if reasons:
        m.verify_reason = " | ".join(reasons)
        m.arrears_verified = False
    else:
        m.verify_reason = None
        m.arrears_verified = True

def _full_recalc(db: Session):
    """전체 회원 재계산 후 스냅샷 갱신"""
    for m in db.query(Member).all():
        _recalc_member(db, m)
    db.commit()
    _invalidate_snap(db, "dashboard")

@app.post("/upload/legacy")
async def upload_legacy(request: Request, file: UploadFile = File(...),
                        db: Session = Depends(get_db), user: User = Depends(require_user)):
    tmpdir = Path(tempfile.mkdtemp())
    try:
        dest = tmpdir / (file.filename or "upload.xlsx")
        with open(dest, "wb") as f: shutil.copyfileobj(file.file, f)
        if dest.stat().st_size == 0:
            return RedirectResponse("/upload?msg=파일이 비어 있습니다", status_code=302)

        file_year = guess_file_year(file.filename or "")
        try: xl = _open_xl(dest)
        except Exception as e:
            return RedirectResponse(f"/upload?msg=파일 열기 실패: {e}", status_code=302)

        batch = UploadBatch(file_name=file.filename, data_type="legacy",
                            file_year=file_year, created_by=user.id)
        db.add(batch); db.commit(); db.refresh(batch)

        nm = nl = ne = 0; warns = []
        for sname in xl.sheet_names:
            try: raw = pd.read_excel(xl, sheet_name=sname, header=None, dtype=object)
            except: continue
            if raw is None or raw.empty: continue

            rows_list = raw.values.tolist()
            h = [_s(v) for v in rows_list[0]] if rows_list else []
            stype = guess_sheet_type(sname, h)

            for ridx, row in enumerate(rows_list):
                vals = [_s(v) for v in row]
                if not any(v for v in vals if v): continue
                db.add(RawImportRow(batch_id=batch.id, source_file=file.filename,
                    source_sheet=sname, source_row=ridx+1, sheet_type=stype,
                    raw_data=json.dumps(vals, ensure_ascii=False, default=str)))

            if stype == "ledger":
                parsed = parse_ledger_sheet(raw, sname, file_year)
                for veh,name,region,account,note,carry,monthly,src_row in parsed:
                    m, how = _find_or_create_member(db, veh, name, region, account, note,
                                                    batch.id, file.filename, sname, src_row)
                    if m is None: warns.append(f"{sname}/{src_row}: 중복후보"); continue
                    if how == "created": nm += 1

                    for mo,(charge,paid,arrears,pdate) in monthly.items():
                        exist = db.query(MonthlyLedger).filter(
                            MonthlyLedger.member_id == m.id,
                            MonthlyLedger.year == file_year,
                            MonthlyLedger.month == mo,
                            MonthlyLedger.source_sheet == sname,
                            MonthlyLedger.source_row == src_row,
                        ).first()
                        if exist: continue
                        carry_v = carry if mo == min(monthly.keys()) else 0
                        calc = carry_v + charge - paid
                        db.add(MonthlyLedger(
                            member_id=m.id, batch_id=batch.id,
                            source_file=file.filename, source_sheet=sname, source_row=src_row,
                            raw_vehicle_no=veh, raw_name=name,
                            raw_region=region, raw_account=account, raw_note=note,
                            year=file_year, month=mo,
                            carry_over=carry_v, charge_amount=charge,
                            paid_amount=paid, arrears_amount=arrears,
                            paid_date=pdate,
                            calc_arrears=arrears,   # 재계산 누적 중단: calc = excel
                            verified=True,          # 차이 0이므로 항상 verified
                        ))
                        nl += 1
                    db.flush()

            elif stype == "status":
                parsed = parse_status_sheet(raw, file_year)
                for veh,name,region,account,reason,arrears,src_row,em in parsed:
                    evtype = detect_status(reason)
                    if not evtype: continue
                    m, how = _find_or_create_member(db, veh, name, region,
                        account or "협", "", batch.id, file.filename, sname, src_row)
                    if m:
                        if how == "created": nm += 1
                        if evtype in EXCLUDED_STATUSES:
                            m.status = evtype; m.status_source = "column"
                        db.add(MemberStatusEvent(
                            member_id=m.id, batch_id=batch.id,
                            source_file=file.filename, source_sheet=sname, source_row=src_row,
                            raw_vehicle_no=veh, raw_name=name,
                            raw_region=normalize_region(region), raw_account=account,
                            event_type=evtype, event_date_raw=parse_date_str(reason),
                            reason_raw=reason, arrears_at_event=arrears, event_month=em or "",
                        ))
                        ne += 1
                    db.flush()
            db.commit()

        # 재계산 (excel_arrears 동기화 + is_overpay 갱신)
        for m_obj in db.query(Member).filter(Member.source_batch_id == batch.id).all():
            _recalc_member(db, m_obj)
        db.commit()
        _invalidate_snap(db, "dashboard")
        batch.total_rows = db.query(RawImportRow).filter(RawImportRow.batch_id == batch.id).count()
        batch.saved_rows = nl; batch.warn_rows = len(warns); db.commit()
        add_log(db, user.id, "미수금업로드",
                f"{file.filename}({file_year}): 회원{nm}명 원장{nl}행 이벤트{ne}건")
        msg = f"업로드완료: 신규회원 {nm}명, 원장 {nl}행"
        if warns: msg += f" (경고 {len(warns)}건)"
    except Exception as e:
        msg = f"오류: {e}"
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return RedirectResponse(f"/upload?msg={msg}", status_code=302)

@app.post("/upload/license")
async def upload_license(request: Request, file: UploadFile = File(...),
                          db: Session = Depends(get_db), user: User = Depends(require_user)):
    tmpdir = Path(tempfile.mkdtemp())
    try:
        dest = tmpdir / (file.filename or "license.xlsx")
        with open(dest, "wb") as f: shutil.copyfileobj(file.file, f)
        if dest.stat().st_size == 0:
            return RedirectResponse("/upload?msg=파일이 비어 있습니다", status_code=302)

        # 전체자명단은 utils/excel_parser 사용
        from utils.excel_parser import (read_excel_sheets, valid_person_row, choose,
                                         get_region as _gr, clean_name as _cn,
                                         normalize_vehicle as _nv)
        db.query(LicenseRecord).delete(synchronize_session=False)
        db.query(UploadBatch).filter(UploadBatch.data_type == "license").delete(synchronize_session=False)
        db.commit()
        batch = UploadBatch(file_name=file.filename, data_type="license", created_by=user.id)
        db.add(batch); db.commit(); db.refresh(batch)
        try: sheets = read_excel_sheets(str(dest))
        except ValueError as e:
            return RedirectResponse(f"/upload?msg={e}", status_code=302)
        saved = 0
        for sname, df, hrow, colmap in sheets:
            for idx, row in df.iterrows():
                if not valid_person_row(row, colmap, "license"): continue
                nm_  = _cn(choose(row, colmap, "name"))
                veh_ = choose(row, colmap, "vehicle_no")
                db.add(LicenseRecord(
                    batch_id=batch.id, source_file=file.filename, source_sheet=sname,
                    region=_gr(row,colmap), name=nm_, name_key=norm_name(nm_),
                    vehicle_no=veh_, vehicle_key=_nv(veh_) if veh_ else "",
                    resident_no=choose(row,colmap,"resident_no"),
                    mobile=choose(row,colmap,"mobile"), phone=choose(row,colmap,"phone"),
                    address=choose(row,colmap,"address"),
                    official_address=choose(row,colmap,"official_address"),
                    join_date_raw=choose(row,colmap,"join_date"),
                    permit_date_raw=choose(row,colmap,"permit_date"),
                    cert_issue_date_raw=choose(row,colmap,"cert_issue_date"),
                    cert_no=choose(row,colmap,"cert_no"),
                    note=choose(row,colmap,"note"),
                ))
                saved += 1
        batch.saved_rows = saved; db.commit()
        _reconcile_license(db)
        _full_recalc(db)
        add_log(db, user.id, "전체자업로드", f"{file.filename}: {saved}건")
        msg = f"전체자명단 {saved}건 저장, 대조완료"
    except Exception as e:
        msg = f"오류: {e}"
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return RedirectResponse(f"/upload?msg={msg}", status_code=302)

# ── 전체자 대조 ────────────────────────────────────────────────────────────────
def _reconcile_license(db: Session):
    """
    전체자명단 ↔ 미수금명단 대조.
    정규화: norm_vehicle(강원·공백·호 제거), norm_name(공백제거)
    우선순위:
    1) vkey + 성명 완전일치 → 정상매칭
    2) 뒷자리4 + 성명 일치 → 정상매칭
    3) vkey 단독 단일후보 → 성명확인필요
    4) 뒷자리 단독 단일후보 → 뒷자리매칭
    5) 성명 단독 단일후보 → 차량번호확인필요
    """
    from core import is_sum_row
    lics = db.query(LicenseRecord).all()
    by_vkey: dict = {}
    by_l4: dict = {}
    by_nkey: dict = {}

    for lic in lics:
        # 전체자명단도 현재 정규화 기준으로 재인덱싱
        vk = norm_vehicle(lic.vehicle_no or "")
        if vk: by_vkey.setdefault(vk, []).append(lic)
        l4 = veh_last4(lic.vehicle_no or "")
        if l4 and len(l4) >= 3: by_l4.setdefault(l4, []).append(lic)
        nk = norm_name(lic.name or "")
        if nk: by_nkey.setdefault(nk, []).append(lic)

    for m in db.query(Member).all():
        # 합계행은 대조 생략
        if is_sum_row(m.name or "", m.vehicle_no or ""):
            m.match_status = "해당없음"
            m.match_fail_reason = "합계행 제외"
            continue

        vkey = norm_vehicle(m.vehicle_no or "")
        l4   = veh_last4(m.vehicle_no or "")
        nkey = norm_name(m.name or "")

        vc = list({x.id: x for x in by_vkey.get(vkey, [])}.values())
        lc = list({x.id: x for x in (by_l4.get(l4, []) if l4 and len(l4) >= 3 else [])}.values())
        nc = list({x.id: x for x in by_nkey.get(nkey, [])}.values())

        best = None; status = "전체자미확인"; fail = ""

        # 1순위: vkey + 성명
        exact_vn = [x for x in vc if norm_name(x.name or "") == nkey]
        # 2순위: 뒷자리 + 성명
        exact_l4n = [x for x in lc if norm_name(x.name or "") == nkey]

        if exact_vn:
            best = exact_vn[0]; status = "정상매칭"
        elif exact_l4n:
            best = exact_l4n[0]; status = "정상매칭"
        elif vkey and len(vc) == 1:
            best = vc[0]; status = "성명확인필요"
        elif l4 and len(l4) >= 3 and len(lc) == 1:
            best = lc[0]; status = "뒷자리매칭"
        elif len(nc) == 1:
            best = nc[0]; status = "차량번호확인필요"
        elif len(nc) > 1:
            fail = f"동명이인 {len(nc)}명"
        elif not (m.vehicle_no or "").strip():
            fail = "차량번호 없음"
        elif l4 and len(lc) > 1:
            fail = f"차량뒷자리 후보 {len(lc)}명"
        else:
            fail = "전체자명단에 없음"

        if best:
            m.match_license_id = best.id; m.match_status = status
            m.match_fail_reason = None
            for attr, lattr in [("mobile","mobile"),("phone","phone"),("address","address"),
                                 ("official_address","official_address"),
                                 ("join_date_raw","join_date_raw"),("permit_date_raw","permit_date_raw"),
                                 ("cert_issue_date_raw","cert_issue_date_raw"),("cert_no","cert_no")]:
                if not getattr(m, attr, None) and getattr(best, lattr, None):
                    setattr(m, attr, getattr(best, lattr))
        else:
            m.match_license_id = None; m.match_status = "전체자미확인"
            m.match_fail_reason = fail
    db.commit()

# ── 전체자 대조 미확인 ────────────────────────────────────────────────────────
@app.get("/license-check", response_class=HTMLResponse)
def license_check(request: Request, db: Session = Depends(get_db),
                  user: User = Depends(require_user)):
    items = (db.query(Member)
             .filter(_clean_filter(),
                     Member.user_confirmed_match == False,
                     or_(Member.match_license_id == None,
                         Member.match_status == "전체자미확인"))
             .order_by(Member.region, Member.name).all())
    lic_count = db.query(LicenseRecord).count()
    return templates.TemplateResponse(request, "license_check.html", {
        "request": request, "user": user, "items": items, "lic_count": lic_count,
        "fmt_amt": fmt_amt, "msg": request.query_params.get("msg", ""),
    })

@app.post("/license-check/{mid}/confirm")
def lic_confirm(mid: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    m.user_confirmed_match = True; db.commit()
    return RedirectResponse("/license-check?msg=확인완료", status_code=302)

@app.post("/license-check/{mid}/link")
def lic_link(mid: int, license_id: int = Form(...),
             db: Session = Depends(get_db), user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    lic = db.query(LicenseRecord).filter(LicenseRecord.id == license_id).first()
    if not m or not lic:
        return RedirectResponse("/license-check?msg=레코드 없음", status_code=302)
    m.match_license_id = lic.id; m.match_status = "수동매칭"
    m.match_fail_reason = None; m.user_confirmed_match = True
    for attr, lattr in [("mobile","mobile"),("address","address"),("official_address","official_address")]:
        if not getattr(m, attr) and getattr(lic, lattr, None):
            setattr(m, attr, getattr(lic, lattr))
    db.commit()
    _full_recalc(db)
    return RedirectResponse(f"/license-check?msg={m.name} 연결완료", status_code=302)

@app.post("/license-check/{mid}/to-work")
def lic_to_work(mid: int, process_type: str = Form("폐업"),
                db: Session = Depends(get_db), user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    db.add(WorkQueue(member_id=mid, process_type=process_type, status="반영대기",
                     source_screen="전체자대조", submitted_by=user.id,
                     arrears_at_submit=m.excel_arrears or 0))
    db.commit()
    _invalidate_snap(db, "dashboard")
    return RedirectResponse("/license-check?msg=부과대수 관리으로 이동", status_code=302)

@app.get("/api/license-search")
def api_lic_search(q: str = "", db: Session = Depends(get_db),
                   user: User = Depends(require_user)):
    if not q: return {"items": []}
    l4 = "".join(c for c in q if c.isdigit())[-4:]
    nk = norm_name(q)
    flt = []
    if nk: flt.append(LicenseRecord.name_key.ilike(f"%{nk}%"))
    if l4 and len(l4) >= 3: flt.append(LicenseRecord.vehicle_no.ilike(f"%{l4}%"))
    if not flt: return {"items": []}
    recs = db.query(LicenseRecord).filter(or_(*flt)).limit(10).all()
    return {"items": [{"id":r.id,"name":r.name,"vehicle_no":r.vehicle_no,
        "mobile":r.mobile or "","address":(r.official_address or r.address or "")} for r in recs]}

# ── 부과대수 관리 ──────────────────────────────────────────────────────────────
@app.get("/work", response_class=HTMLResponse)
def work_page(request: Request, tab: str = "전체", q: str = "",
              db: Session = Depends(get_db), user: User = Depends(require_user)):
    TABS = ["전체","반영대기","폐업","양도","이관","탈퇴","사망","말소","현역복구","반영완료"]
    wq = db.query(WorkQueue, Member).join(Member)
    if tab == "반영대기": wq = wq.filter(WorkQueue.status == "반영대기")
    elif tab == "반영완료": wq = wq.filter(WorkQueue.status == "반영완료")
    elif tab != "전체": wq = wq.filter(WorkQueue.process_type == tab, WorkQueue.status == "반영대기")
    if q:
        like = f"%{q}%"
        wq = wq.filter(or_(Member.name.ilike(like), Member.vehicle_no.ilike(like)))
    items = wq.order_by(WorkQueue.submitted_at.desc()).limit(500).all()

    def cnt(t):
        q2 = db.query(WorkQueue)
        if t == "반영대기": return q2.filter(WorkQueue.status == "반영대기").count()
        if t == "반영완료": return q2.filter(WorkQueue.status == "반영완료").count()
        if t == "전체":    return q2.count()
        return q2.filter(WorkQueue.process_type == t, WorkQueue.status == "반영대기").count()
    tab_counts = {t: cnt(t) for t in TABS}

    return templates.TemplateResponse(request, "work.html", {
        "request": request, "user": user, "items": items, "tab": tab, "q": q,
        "tab_counts": tab_counts, "TABS": TABS, "fmt_amt": fmt_amt,
        "msg": request.query_params.get("msg", ""),
    })

@app.post("/work/{wid}/reflect")
def work_reflect(wid: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    wq = db.query(WorkQueue).filter(WorkQueue.id == wid).first()
    if not wq: raise HTTPException(404)
    m = db.query(Member).filter(Member.id == wq.member_id).first()
    if not m: raise HTTPException(404)
    m.status = wq.process_type; m.status_source = "manual"
    wq.status = "반영완료"; wq.reflected_by = user.id
    wq.reflected_at = datetime.now()
    db.commit()
    _full_recalc(db)
    add_log(db, user.id, "업무처리반영", f"{m.name}/{m.vehicle_no} → {wq.process_type}")
    return RedirectResponse(f"/work?msg={m.name} {wq.process_type} 반영완료", status_code=302)

@app.post("/work/{wid}/cancel")
def work_cancel(wid: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    wq = db.query(WorkQueue).filter(WorkQueue.id == wid).first()
    if not wq: raise HTTPException(404)
    wq.status = "취소"; db.commit()
    _invalidate_snap(db, "dashboard")
    return RedirectResponse("/work?msg=취소완료", status_code=302)

# ── 검증필요 ──────────────────────────────────────────────────────────────────
@app.get("/verify", response_class=HTMLResponse)
def verify_page(request: Request, db: Session = Depends(get_db),
                user: User = Depends(require_user)):
    items = (db.query(Member).filter(Member.status == "정상",
                                     Member.arrears_verified == False)
             .order_by(Member.region, Member.name).limit(500).all())
    return templates.TemplateResponse(request, "verify.html", {
        "request": request, "user": user, "items": items, "fmt_amt": fmt_amt,
        "msg": request.query_params.get("msg", ""),
    })

@app.post("/verify/{mid}/clear")
def verify_clear(mid: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    m.arrears_verified = True; m.verify_reason = None; db.commit()
    _invalidate_snap(db, "dashboard")
    return RedirectResponse("/verify?msg=검증해제완료", status_code=302)

# ── 문자대상 ──────────────────────────────────────────────────────────────────
@app.get("/collection", response_class=HTMLResponse)
def collection_page(request: Request, region: str = "", account: str = "",
                    tab: str = "전체", q: str = "", min_amt: int = 0,
                    db: Session = Depends(get_db), user: User = Depends(require_user)):
    from datetime import date as _date
    base_q = db.query(CollectionTarget, Member).join(Member)

    TABS_ALL = ["전체","협회비","관리비","초과납부","연락처없음","주소없음","제외",
                "금액5천+","금액1만+","금액3만+","금액5만+","금액10만+","금액30만+","금액50만+","금액100만+",
                "장기3개월+","장기6개월+","장기12개월+"]

    if tab == "협회비":
        base_q = base_q.filter(CollectionTarget.excluded==False, Member.account=="협")
    elif tab == "관리비":
        base_q = base_q.filter(CollectionTarget.excluded==False, Member.account=="관")
    elif tab == "초과납부":
        base_q = base_q.filter(Member.is_overpay==True)
    elif tab == "연락처없음":
        base_q = base_q.filter(or_(Member.mobile==None, Member.mobile==""))
    elif tab == "주소없음":
        base_q = base_q.filter(and_(
            or_(Member.official_address==None, Member.official_address==""),
            or_(Member.address==None, Member.address=="")
        ))
    elif tab == "제외":
        base_q = base_q.filter(CollectionTarget.excluded==True)
    elif tab.startswith("금액"):
        amt_map = {"금액5천+":5000,"금액1만+":10000,"금액3만+":30000,"금액5만+":50000,
                   "금액10만+":100000,"금액30만+":300000,"금액50만+":500000,"금액100만+":1000000}
        threshold = amt_map.get(tab, 0)
        base_q = base_q.filter(CollectionTarget.excluded==False, Member.excel_arrears>=threshold)
    elif tab.startswith("장기"):
        months_map = {"장기3개월+":3,"장기6개월+":6,"장기12개월+":12}
        mo = months_map.get(tab, 3)
        today = _date.today()
        cutoff_year = today.year - (1 if today.month <= mo else 0)
        cutoff_month = (today.month - mo - 1) % 12 + 1
        cutoff = f"{cutoff_year}-{cutoff_month:02d}"
        base_q = base_q.filter(
            CollectionTarget.excluded==False,
            Member.first_unpaid_month != None,
            Member.first_unpaid_month != "",
            Member.first_unpaid_month <= cutoff,
        )
    else:  # 전체
        base_q = base_q.filter(CollectionTarget.excluded==False)

    if region: base_q = base_q.filter(Member.region==region)
    if account: base_q = base_q.filter(Member.account==account)
    if q:
        like = f"%{q}%"
        base_q = base_q.filter(or_(Member.name.ilike(like), Member.vehicle_no.ilike(like)))

    targets = base_q.order_by(Member.region, Member.name).all()
    total = sum(t.CollectionTarget.arrears for t in targets if not t.CollectionTarget.excluded)
    regions = sorted({x[0] for x in db.query(Member.region).distinct().filter(Member.region!=None).all() if x[0]})

    def _cnt(t):
        bq2 = db.query(CollectionTarget).join(Member)
        if t == "협회비": return bq2.filter(CollectionTarget.excluded==False, Member.account=="협").count()
        if t == "관리비": return bq2.filter(CollectionTarget.excluded==False, Member.account=="관").count()
        if t == "초과납부": return bq2.filter(Member.is_overpay==True).count()
        if t == "연락처없음": return bq2.filter(or_(Member.mobile==None,Member.mobile=="")).count()
        if t == "주소없음": return bq2.filter(and_(or_(Member.official_address==None,Member.official_address==""),or_(Member.address==None,Member.address==""))).count()
        if t == "제외": return bq2.filter(CollectionTarget.excluded==True).count()
        return bq2.filter(CollectionTarget.excluded==False).count()

    tab_counts = {t: _cnt(t) for t in ["전체","협회비","관리비","초과납부","연락처없음","주소없음","제외"]}

    return templates.TemplateResponse(request, "collection.html", {
        "request": request, "user": user, "targets": targets, "total": total,
        "regions": regions, "region": region, "account": account,
        "tab": tab, "q": q, "tab_counts": tab_counts, "TABS_ALL": TABS_ALL,
        "fmt_amt": fmt_amt, "msg": request.query_params.get("msg", ""),
    })

@app.post("/collection/generate")
def collection_generate(db: Session = Depends(get_db), user: User = Depends(require_user)):
    db.query(CollectionTarget).delete(synchronize_session=False); db.commit()
    members = db.query(Member).filter(
        Member.status == "정상",
        Member.excel_arrears > 0,
        Member.is_overpay == False,
    ).all()
    count = 0
    for m in members:
        # 자동이체는 전용통장 업로드 후에만 확정 → 현재는 일반으로 처리
        no_mobile = not (m.mobile or "").strip()
        no_address = not (m.official_address or m.address or "").strip()
        excluded = False
        exc_reason = ""
        # 음수(초과납부) 제외
        if (m.excel_arrears or 0) < 0:
            excluded = True; exc_reason = "초과납부/선납"
        # 검증필요 제외
        elif m.arrears_verified == False and m.verify_reason:
            excluded = True; exc_reason = "검증필요"
        mob_clean = phone_clean(m.mobile or "")
        addr = address_clean(m)
        cat = "연락처없음" if no_mobile else "일반"
        now_ = datetime.now()
        db.add(CollectionTarget(
            member_id=m.id, generated_by=user.id,
            base_year=now_.year, base_month=now_.month,
            arrears=m.excel_arrears, category=cat,
            excluded=excluded, exclude_reason=exc_reason,
            mobile_clean=mob_clean, address_clean=addr,
        ))
        count += 1
    db.commit()
    _invalidate_snap(db, "dashboard")
    add_log(db, user.id, "문자대상생성", f"{count}명")
    return RedirectResponse(f"/collection?msg={count}명 생성완료", status_code=302)

# ── 통장매칭 ──────────────────────────────────────────────────────────────────
@app.get("/bank", response_class=HTMLResponse)
def bank_page(request: Request, status: str = "", q: str = "",
              db: Session = Depends(get_db), user: User = Depends(require_user)):
    bq = db.query(BankTransaction)
    if status: bq = bq.filter(BankTransaction.match_status == status)
    if q:
        like = f"%{q}%"
        bq = bq.filter(or_(BankTransaction.memo.ilike(like)))
    txs = bq.order_by(BankTransaction.id.desc()).limit(500).all()
    import json as _json
    for tx in txs:
        try: tx._candidates = _json.loads(tx.match_candidates_json) if tx.match_candidates_json else []
        except: tx._candidates = []

    counts = {s: db.query(BankTransaction).filter(BankTransaction.match_status == s).count()
              for s in ["자동매칭","확인필요","미매칭","반영완료"]}
    counts["전체"] = db.query(BankTransaction).count()

    return templates.TemplateResponse(request, "bank.html", {
        "request": request, "user": user, "txs": txs, "status": status, "q": q,
        "counts": counts, "fmt_amt": fmt_amt,
        "msg": request.query_params.get("msg", ""),
    })

@app.post("/bank/paste")
def bank_paste(request: Request, pasted_text: str = Form(""),
               db: Session = Depends(get_db), user: User = Depends(require_user)):
    count = _parse_and_match_bank_lines(db, pasted_text.splitlines(), "paste", None)
    add_log(db, user.id, "통장붙여넣기", f"{count}건")
    return RedirectResponse(f"/bank?msg={count}건 파싱/매칭", status_code=302)

def _parse_bank_line(line: str) -> Optional[dict]:
    """
    통장 한 줄 파싱
    컬럼 순서: 날짜 [시간] 입금액 [출금액] 잔액 거래구분 메모
    핵심: 2026같은 4자리 연도를 금액으로 오인하지 않도록 처리
    """
    raw = line.rstrip()
    if not raw.strip(): return None
    if "\t" in raw: parts = [p.strip() for p in raw.split("\t")]
    else: parts = re.split(r"\s{2,}", raw.strip())
    parts = [p for p in parts if p]
    if not parts: return None

    # 1) 날짜 추출
    tx_date = ""
    date_idx = -1
    for i, p in enumerate(parts):
        m = re.match(r"^(20\d{2}[-/.]\d{1,2}[-/.]\d{1,2})", p)
        if m:
            tx_date = m.group(1).replace(".", "-").replace("/", "-")
            date_idx = i; break
    remaining_parts = [p for i, p in enumerate(parts) if i != date_idx] if date_idx >= 0 else parts[:]

    # 2) 금액 컬럼 추출 (쉼표 포함 or 5자리 이상 숫자)
    #    쉼표 없는 4자리 이하 순수 숫자(2026, 1234 등)는 제외
    def is_amount_token(s):
        plain = s.replace(",", "").replace(".", "")
        if not plain.isdigit(): return False
        if "," in s: return True          # 1,000 형태
        if len(plain) >= 5: return True    # 10000 이상
        return False

    amounts = []
    text_parts = []
    for p in remaining_parts:
        if is_amount_token(p):
            try:
                a = parse_amount(p, allow_negative=False)
                if a >= 1000: amounts.append((a, p)); continue
            except: pass
        text_parts.append(p)

    # 컬럼 순서: 첫 번째가 입금액, 두 번째가 잔액 (일반적)
    # 단, 금액이 1개만 있으면 그게 잔액일 수 있으므로 주의
    if len(amounts) >= 2:
        # 첫 번째가 잔액보다 작으면 입금액, 아니면 역순 가능
        deposit = amounts[0][0]
        balance = amounts[-1][0]
    elif len(amounts) == 1:
        # 금액 1개 = 잔액만 있거나 입금액만 있음
        # 잔액이 더 크면 잔액으로 간주 (deposit=0)
        deposit = amounts[0][0]
        balance = 0
    else:
        deposit = balance = 0

    # 3) 나머지에서 거래구분/메모 추출
    tokens = [t for t in text_parts if t and not re.match(r"^\d{1,4}$", t)]
    method = tokens[0] if tokens else ""
    memo   = " ".join(tokens[1:]) if len(tokens) > 1 else ""

    return {"txn_date": tx_date, "deposit_amount": deposit, "balance_amount": balance,
            "method": method, "memo": memo, "raw_text": line}

def _parse_and_match_bank_lines(db: Session, lines, source_type: str, batch_id) -> int:
    count = 0
    for line in lines:
        parsed = _parse_bank_line(str(line))
        if not parsed or parsed["deposit_amount"] <= 0: continue
        if is_useless_memo(parsed["memo"]): continue
        tx = BankTransaction(
            txn_date=parsed["txn_date"], deposit_amount=parsed["deposit_amount"],
            balance_amount=parsed["balance_amount"], method=parsed["method"],
            memo=parsed["memo"], source_type=source_type,
            batch_id=batch_id, raw_text=parsed["raw_text"],
        )
        _do_bank_match(db, tx)
        count += 1
    return count

def _do_bank_match(db: Session, tx: BankTransaction):
    import json as _json
    memo = tx.memo or ""; deposit = tx.deposit_amount or 0
    name_c, vkey, l4 = extract_memo_keys(memo)

    flt = []
    if name_c and len(name_c) >= 2:
        flt.append(func.replace(Member.name," ","").ilike(f"%{norm_name(name_c)}%"))
    if vkey and len(vkey) >= 5: flt.append(Member.vehicle_key == vkey)
    if l4 and len(l4) >= 3: flt.append(Member.vehicle_no.ilike(f"%{l4}%"))

    if not flt:
        tx.match_status = "미매칭"; db.add(tx); db.commit(); db.refresh(tx); return tx

    raw_members = db.query(Member).filter(Member.status == "정상", or_(*flt)).limit(500).all()
    seen: dict = {}
    for m in raw_members:
        sc, rs, tier = score_bank_match(memo, deposit, m)
        if sc <= 0: continue
        if m.id not in seen or sc > seen[m.id][0]:
            seen[m.id] = (sc, rs, tier, {
                "id":m.id,"name":m.name,"vehicle_no":m.vehicle_no,
                "vehicle_last4":veh_last4(m.vehicle_no),
                "region":m.region or "","account":fmt_acc(m.account),
                "amount":m.excel_arrears or 0,"mobile":m.mobile or "",
                "score":sc,"reason":rs,"tier":tier,
            })

    if not seen:
        tx.match_status = "미매칭"; db.add(tx); db.commit(); db.refresh(tx); return tx

    ranked = sorted(seen.values(), key=lambda x: x[0], reverse=True)
    candidates = [x[3] for x in ranked]
    best = ranked[0]
    best_id, best_sc, best_tier = best[3]["id"], best[0], best[2]
    best_member = db.query(Member).filter(Member.id == best_id).first()
    n = len(candidates)
    status = "자동매칭" if (best_tier in {"exact","vkey"} and n == 1) or (best_sc >= 90 and n == 1) else "확인필요"

    tx.matched_member_id = best_id if status == "자동매칭" else None
    tx.match_score = best_sc; tx.match_reason = best[1]
    tx.match_status = status
    tx.match_candidates_json = _json.dumps(candidates, ensure_ascii=False)
    db.add(tx); db.commit(); db.refresh(tx)
    return tx

@app.post("/bank/{txid}/apply")
def bank_apply(txid: int, member_id: Optional[int] = None,
               note: str = Form(""),
               db: Session = Depends(get_db), user: User = Depends(require_user)):
    tx = db.query(BankTransaction).filter(BankTransaction.id == txid).first()
    if not tx: raise HTTPException(404)
    if tx.applied:
        return RedirectResponse("/bank?msg=이미 반영된 건입니다", status_code=302)
    mid = member_id or tx.matched_member_id
    if not mid: return RedirectResponse("/bank?msg=대상자를 선택해주세요", status_code=302)
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: return RedirectResponse("/bank?msg=회원을 찾을 수 없습니다", status_code=302)

    amt = tx.deposit_amount or 0
    before = m.excel_arrears or 0
    after  = before - amt
    overpay = max(0, -after)  # 0 미만이면 초과납부

    # 1. excel_arrears 직접 차감
    m.excel_arrears = after
    m.is_overpay = after < 0
    m.last_paid_date = tx.txn_date or m.last_paid_date

    # 2. monthly_ledgers에 입금 기록 추가 (최신월 원장에 paid_amount 반영)
    if tx.txn_date:
        try:
            tx_year  = int(tx.txn_date[:4])
            tx_month = int(tx.txn_date[5:7])
            # 최신월 원장 찾아서 입금액 추가
            existing_ledger = (db.query(MonthlyLedger)
                .filter(MonthlyLedger.member_id == mid,
                        MonthlyLedger.year == tx_year,
                        MonthlyLedger.month == tx_month)
                .order_by(MonthlyLedger.id.desc()).first())
            if existing_ledger:
                existing_ledger.paid_amount = (existing_ledger.paid_amount or 0) + amt
                existing_ledger.arrears_amount = after
                existing_ledger.paid_date = tx.txn_date
            else:
                # 해당월 원장 없으면 새로 생성
                latest_l = (db.query(MonthlyLedger)
                    .filter(MonthlyLedger.member_id == mid)
                    .order_by(MonthlyLedger.year.desc(), MonthlyLedger.month.desc())
                    .first())
                db.add(MonthlyLedger(
                    member_id=mid, batch_id=None,
                    source_file="통장매칭", source_sheet="bank", source_row=tx.id,
                    year=tx_year, month=tx_month,
                    carry_over=before, charge_amount=0,
                    paid_amount=amt, arrears_amount=after,
                    paid_date=tx.txn_date, calc_arrears=after,
                    verified=True,
                ))
        except Exception as _e:
            pass  # 날짜 파싱 실패 시 무시

    # 3. BankTransaction 상태 업데이트
    tx.applied = True
    tx.applied_amount = amt
    tx.overpay_amount = overpay
    tx.match_status = "반영완료"
    tx.matched_member_id = mid
    tx.applied_at = datetime.now()
    tx.note = note
    db.commit()

    # 4. 해당 회원 재계산 (검증필요, 초과납부, 문자대상 갱신)
    _recalc_member(db, m)
    db.commit()
    _invalidate_snap(db, "dashboard")

    add_log(db, user.id, "통장반영",
            f"{m.name}/{m.vehicle_no}: {before:,}→{after:,}원 (입금 {amt:,}원, 날짜 {tx.txn_date})")
    return RedirectResponse(f"/bank?msg={m.name} 반영완료 ({before:,}→{after:,}원)", status_code=302)

@app.get("/bank/{txid}/preview")
def bank_preview(txid: int, member_id: Optional[int] = None,
                 db: Session = Depends(get_db), user: User = Depends(require_user)):
    tx = db.query(BankTransaction).filter(BankTransaction.id == txid).first()
    if not tx: return JSONResponse({"ok": False, "msg": "거래 없음"})
    mid = member_id or tx.matched_member_id
    if not mid: return JSONResponse({"ok": False, "msg": "대상자 없음"})
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: return JSONResponse({"ok": False, "msg": "회원 없음"})
    amt = tx.deposit_amount or 0
    before = m.excel_arrears or 0
    after = before - amt
    return JSONResponse({"ok":True,"member_id":m.id,"name":m.name,"vehicle_no":m.vehicle_no,
        "account":fmt_acc(m.account),"deposit":amt,"before":before,"after":after,
        "overpay":max(0,amt-before)})

@app.post("/bank/{txid}/hold")
def bank_hold(txid: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    tx = db.query(BankTransaction).filter(BankTransaction.id == txid).first()
    if not tx: raise HTTPException(404)
    tx.match_status = "보류"; db.commit()
    return RedirectResponse("/bank?msg=보류처리", status_code=302)

# ── 설정 ──────────────────────────────────────────────────────────────────────
@app.get("/settings", response_class=HTMLResponse)
def settings_page(request: Request, db: Session = Depends(get_db),
                  user: User = Depends(require_user)):
    users = db.query(User).all()
    return templates.TemplateResponse(request, "settings.html", {
        "request": request, "user": user, "users": users,
        "msg": request.query_params.get("msg", ""),
    })

@app.post("/settings/pw")
def settings_pw(old_pw: str = Form(""), new_pw: str = Form(""),
                db: Session = Depends(get_db), user: User = Depends(require_user)):
    if not verify_pw(old_pw, user.password_hash):
        return RedirectResponse("/settings?msg=현재 비밀번호 틀림", status_code=302)
    user.password_hash = hash_pw(new_pw); db.commit()
    return RedirectResponse("/settings?msg=비밀번호 변경완료", status_code=302)

@app.post("/settings/add-user")
def add_user(username: str = Form(""), name: str = Form(""), password: str = Form(""),
             db: Session = Depends(get_db), user: User = Depends(require_user)):
    if db.query(User).filter(User.username == username).first():
        return RedirectResponse("/settings?msg=이미 존재하는 아이디", status_code=302)
    db.add(User(username=username, name=name, password_hash=hash_pw(password))); db.commit()
    return RedirectResponse("/settings?msg=사용자 추가완료", status_code=302)

@app.post("/settings/rebuild-snapshot")
def rebuild_snap(db: Session = Depends(get_db), user: User = Depends(require_user)):
    _full_recalc(db)
    snap = _build_dashboard_snap(db)
    _set_snap(db, "dashboard", snap)
    add_log(db, user.id, "스냅샷재구축", "전체")
    return RedirectResponse("/settings?msg=스냅샷 재구축 완료", status_code=302)

@app.post("/settings/fix-negatives")
def fix_negatives(db: Session = Depends(get_db), user: User = Depends(require_user)):
    """음수 미수금 자동복구 — raw_data에서 재해석"""
    fixed = 0
    for m in db.query(Member).filter(Member.excel_arrears < 0).all():
        m.is_overpay = True
        fixed += 1
    db.commit()
    _full_recalc(db)
    add_log(db, user.id, "음수미수금복구", f"{fixed}건")
    return RedirectResponse(f"/settings?msg=음수미수금 {fixed}건 처리완료", status_code=302)

@app.post("/admin/reset")
def admin_reset(db: Session = Depends(get_db), user: User = Depends(require_user)):
    for tbl in [CollectionTarget, BankTransaction, BillingPerson, WorkQueue,
                MemberStatusEvent, MonthlyLedger, RawImportRow,
                LicenseRecord, Member, UploadBatch, Snapshot, BillingReport]:
        db.query(tbl).delete(synchronize_session=False)
    db.commit()
    add_log(db, user.id, "데이터초기화", "전체")
    return RedirectResponse("/upload?msg=초기화완료", status_code=302)


@app.get("/member/new", response_class=HTMLResponse)
def add_member_page(request: Request, db: Session = Depends(get_db),
                    user: User = Depends(require_user)):
    regions = sorted({x[0] for x in db.query(Member.region).distinct().filter(Member.region != None).all() if x[0]})
    return templates.TemplateResponse(request, "add_member.html",
        {"request": request, "user": user, "regions": regions})

@app.post("/member/new")
def add_member_save(request: Request,
    name: str = Form(""), vehicle_no: str = Form(""), region: str = Form(""),
    account: str = Form("협"), mobile: str = Form(""),
    amount: str = Form("0"), address: str = Form(""), note: str = Form(""),
    db: Session = Depends(get_db), user: User = Depends(require_user)):
    m = Member(name=name, name_key=norm_name(name),
               vehicle_no=vehicle_no, vehicle_key=norm_vehicle(vehicle_no) if vehicle_no else "",
               region=normalize_region(region), account=account,
               mobile=phone_clean(mobile) or mobile,
               address=address, note=note,
               excel_arrears=parse_amount(amount),
               calc_arrears=parse_amount(amount),
               status="정상", status_source="manual",
               source_file="명단추가")
    db.add(m); db.commit()
    _full_recalc(db)
    add_log(db, user.id, "회원추가", f"{name}/{vehicle_no}")
    return RedirectResponse(f"/member/{m.id}", status_code=302)

@app.get("/api/member-search")
def api_member_search(q: str = "", db: Session = Depends(get_db),
                      user: User = Depends(require_user)):
    if not q: return {"items": []}
    l4 = "".join(c for c in q if c.isdigit())[-4:]
    nk = norm_name(q)
    flt = []
    if nk and len(nk) >= 2: flt.append(func.replace(Member.name," ","").ilike(f"%{nk}%"))
    if l4 and len(l4) >= 3: flt.append(Member.vehicle_no.ilike(f"%{l4}%"))
    if not flt: return {"items": []}
    members = db.query(Member).filter(Member.status == "정상", or_(*flt)).limit(10).all()
    return {"items": [{"id":m.id,"name":m.name,"vehicle_no":m.vehicle_no or "",
        "amount":m.excel_arrears or 0} for m in members]}


@app.get("/bank/auto-analysis", response_class=HTMLResponse)
def bank_auto_analysis(request: Request, db: Session = Depends(get_db),
                       user: User = Depends(require_user)):
    """자동이체 전용통장 1년치 분석: 동일 금액 반복 입금 패턴 감지"""
    import json as _json
    from datetime import date, timedelta

    # 최근 1년치 통장 내역
    txs = (db.query(BankTransaction)
           .filter(BankTransaction.is_auto_account == True)
           .order_by(BankTransaction.txn_date, BankTransaction.memo)
           .all())

    if not txs:
        txs = db.query(BankTransaction).order_by(BankTransaction.txn_date).all()

    # 메모별 입금패턴 분석
    from collections import defaultdict
    memo_groups: dict = defaultdict(list)
    for tx in txs:
        key = tx.memo or "(메모없음)"
        memo_groups[key].append({
            "date": tx.txn_date,
            "amount": tx.deposit_amount or 0,
            "applied": tx.applied,
        })

    # 자동이체/정기납부 후보 분류
    results = []
    for memo, entries in memo_groups.items():
        if len(entries) < 2: continue
        amounts = [e["amount"] for e in entries]
        unique_amts = set(amounts)
        valid_amts = unique_amts & AUTOPAY_AMOUNTS
        if not valid_amts and max(amounts, default=0) > 60000: continue
        pat = classify_autopay([(e["amount"], e["date"] or "") for e in entries])
        if pat == "해당없음": continue
        results.append({
            "memo": memo,
            "pattern": pat,
            "count": len(entries),
            "amounts": sorted(valid_amts or unique_amts),
            "last_date": max(e["date"] or "" for e in entries),
            "entries": entries[-6:],  # 최근 6건만 표시
        })

    results.sort(key=lambda x: (-x["count"], x["memo"]))

    return templates.TemplateResponse(request, "bank_auto_analysis.html", {
        "request": request, "user": user, "results": results,
        "total_txs": len(txs), "fmt_amt": fmt_amt,
        "msg": request.query_params.get("msg", ""),
    })


@app.post("/settings/cleanup-sumrows")
def cleanup_sumrows(db: Session = Depends(get_db), user: User = Depends(require_user)):
    """기존 DB에 잘못 저장된 합계행 삭제"""
    from core import is_sum_row
    members = db.query(Member).all()
    deleted = 0
    for m in members:
        if is_sum_row(m.name or "", m.vehicle_no or ""):
            # 연관 데이터 삭제
            db.query(MonthlyLedger).filter(MonthlyLedger.member_id == m.id).delete(synchronize_session=False)
            db.query(MemberStatusEvent).filter(MemberStatusEvent.member_id == m.id).delete(synchronize_session=False)
            db.query(WorkQueue).filter(WorkQueue.member_id == m.id).delete(synchronize_session=False)
            db.delete(m)
            deleted += 1
    db.commit()
    _invalidate_snap(db, "dashboard")
    add_log(db, user.id, "합계행정리", f"{deleted}건 삭제")
    return RedirectResponse(f"/settings?msg=합계행 {deleted}건 정리완료", status_code=302)


@app.post("/member/{mid}/amount-adjust")
def member_amount_adjust(mid: int,
    adjust_amount: int = Form(0),
    adjust_reason: str = Form(""),
    note: str = Form(""),
    db: Session = Depends(get_db), user: User = Depends(require_user)):
    """금액수정 → 부과대수 반영대기로 등록"""
    m = db.query(Member).filter(Member.id == mid).first()
    if not m: raise HTTPException(404)
    before = m.excel_arrears or 0
    db.add(WorkQueue(
        member_id=mid, process_type="금액수정", status="반영대기",
        source_screen="미수금명단",
        reason=f"수정 전:{before:,}원 → 수정 후:{adjust_amount:,}원 / 사유:{adjust_reason}",
        note=note,
        arrears_at_submit=before,
        submitted_by=user.id,
    ))
    db.commit()
    _invalidate_snap(db, "dashboard")
    add_log(db, user.id, "금액수정요청", f"{m.name}: {before:,}→{adjust_amount:,}원")
    return RedirectResponse(f"/member/{mid}?msg=금액수정 반영대기 등록완료", status_code=302)

@app.post("/work/{wid}/reflect-amount")
def work_reflect_amount(wid: int, db: Session = Depends(get_db),
                        user: User = Depends(require_user)):
    """금액수정 반영대기 → 실제 반영"""
    wq = db.query(WorkQueue).filter(WorkQueue.id == wid).first()
    if not wq or wq.process_type != "금액수정": raise HTTPException(404)
    m = db.query(Member).filter(Member.id == wq.member_id).first()
    if not m: raise HTTPException(404)
    # reason에서 수정 후 금액 파싱
    import re
    match = re.search(r"수정 후:(-?[\d,]+)원", wq.reason or "")
    if match:
        new_amt = int(match.group(1).replace(",",""))
        before = m.excel_arrears or 0
        m.excel_arrears = new_amt
        m.is_overpay = new_amt < 0
        wq.status = "반영완료"
        wq.reflected_by = user.id
        wq.reflected_at = datetime.now()
        db.commit()
        _recalc_member(db, m)
        db.commit()
        _invalidate_snap(db, "dashboard")
        add_log(db, user.id, "금액수정반영", f"{m.name}: {before:,}→{new_amt:,}원")
        return RedirectResponse(f"/work?msg=금액수정 반영완료", status_code=302)
    return RedirectResponse(f"/work?msg=금액 파싱 실패", status_code=302)


# ── 부과대수 관리 (협회 월례보고 항목) ─────────────────────────────────────────
@app.get("/billing-report", response_class=HTMLResponse)
def billing_report_page(request: Request, year: Optional[int] = None, month: Optional[int] = None,
                         db: Session = Depends(get_db), user: User = Depends(require_user)):
    now = datetime.now()
    sel_y = year or now.year; sel_m = month or now.month
    report = db.query(BillingReport).filter(
        BillingReport.year == sel_y, BillingReport.month == sel_m).first()
    history = (db.query(BillingReport)
               .order_by(BillingReport.year.desc(), BillingReport.month.desc())
               .limit(24).all())
    years = list(range(now.year - 2, now.year + 2))
    return templates.TemplateResponse(request, "billing_report.html", {
        "request": request, "user": user, "report": report,
        "sel_y": sel_y, "sel_m": sel_m, "history": history,
        "years": years, "months": list(range(1,13)),
        "msg": request.query_params.get("msg", ""),
    })

@app.post("/billing-report")
def billing_report_save(
    request: Request,
    year: int = Form(...), month: int = Form(...),
    cnt_join: int = Form(0), cnt_transfer: int = Form(0),
    cnt_cross: int = Form(0), cnt_close: int = Form(0),
    cnt_quit: int = Form(0), cnt_delivery_new: int = Form(0),
    cnt_mgmt_close: int = Form(0), cnt_age70: int = Form(0),
    memo: str = Form(""),
    db: Session = Depends(get_db), user: User = Depends(require_user)):
    # 자동계산: 기준대수 = 현재 정상 회원수
    cnt_base = db.query(Member).filter(_clean_filter()).count()
    cnt_delivery = db.query(Member).filter(_clean_filter(), Member.account == "관").count()
    cnt_total = cnt_base

    r = db.query(BillingReport).filter(
        BillingReport.year == year, BillingReport.month == month).first()
    if r:
        r.cnt_join=cnt_join; r.cnt_transfer=cnt_transfer
        r.cnt_cross=cnt_cross; r.cnt_close=cnt_close
        r.cnt_quit=cnt_quit; r.cnt_delivery_new=cnt_delivery_new
        r.cnt_mgmt_close=cnt_mgmt_close; r.cnt_age70=cnt_age70
        r.cnt_base=cnt_base; r.cnt_total=cnt_total
        r.cnt_delivery=cnt_delivery; r.memo=memo
    else:
        db.add(BillingReport(
            year=year, month=month,
            cnt_join=cnt_join, cnt_transfer=cnt_transfer,
            cnt_cross=cnt_cross, cnt_close=cnt_close,
            cnt_quit=cnt_quit, cnt_delivery_new=cnt_delivery_new,
            cnt_mgmt_close=cnt_mgmt_close, cnt_age70=cnt_age70,
            cnt_base=cnt_base, cnt_total=cnt_total,
            cnt_delivery=cnt_delivery, memo=memo, created_by=user.id,
        ))
    db.commit()
    add_log(db, user.id, "부과대수저장", f"{year}년{month}월")
    return RedirectResponse(f"/billing-report?year={year}&month={month}&msg=저장완료", status_code=302)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.getenv("PORT","8080")))
