
from fastapi import APIRouter, Query
from fastapi.responses import StreamingResponse
from io import BytesIO
from datetime import datetime, date
from calendar import monthrange
import re
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

router = APIRouter(prefix="/api/reports", tags=["reports-export"])


def _get_db():
    """
    프로젝트마다 DB 연결 방식이 조금 달라서 최대한 호환되게 처리.
    database.py 안에 SessionLocal이 있으면 그걸 우선 사용.
    """
    try:
        from database import SessionLocal
        db = SessionLocal()
        try:
            yield db
        finally:
            db.close()
    except Exception:
        yield None


def _safe_str(v):
    if v is None:
        return ""
    return str(v).strip()


def _parse_date(v):
    if v is None or v == "":
        return None

    # datetime은 date보다 먼저 처리해야 함.
    # datetime도 date의 하위 타입이라 순서가 중요함.
    if isinstance(v, datetime):
        return v.date()

    if isinstance(v, date):
        return v

    s = str(v).strip()
    if not s:
        return None

    s = s.replace(".", "-").replace("/", "-")
    s = re.sub(r"\s+.*$", "", s)

    for fmt in ("%Y-%m-%d", "%Y-%m", "%Y%m%d", "%Y.%m.%d"):
        try:
            d = datetime.strptime(s, fmt)
            return d.date()
        except Exception:
            pass

    m = re.search(r"(20\d{2}|19\d{2})[-.]?(\d{1,2})[-.]?(\d{1,2})?", s)
    if m:
        y = int(m.group(1))
        mo = int(m.group(2))
        da = int(m.group(3) or 1)
        try:
            return date(y, mo, da)
        except Exception:
            return None

    return None


def _month_range(year, month):
    start = date(year, month, 1)
    end = date(year, month, monthrange(year, month)[1])
    return start, end


def _next_month_after(d):
    if not d:
        return None
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


def _norm_col(name):
    return str(name or "").strip().lower().replace(" ", "").replace("_", "")


def _find_col(columns, candidates):
    norm_map = {_norm_col(c): c for c in columns}
    for cand in candidates:
        key = _norm_col(cand)
        if key in norm_map:
            return norm_map[key]
    for c in columns:
        nc = _norm_col(c)
        for cand in candidates:
            if _norm_col(cand) in nc:
                return c
    return None


def _get_tables_and_rows(db, mode="all", limit=3000):
    """
    DB 전체를 무작정 다 읽으면 Railway에서 500/timeout이 날 수 있어서
    mode에 따라 필요한 테이블만 제한적으로 읽음.
    mode="deposit"이면 입금/통장/거래/납부 관련 테이블만 읽음.
    """
    if db is None:
        return []

    try:
        bind = db.get_bind()
        from sqlalchemy import inspect, text
        inspector = inspect(bind)
        tables = inspector.get_table_names()
        result = []

        deposit_words = [
            "payment", "deposit", "bank", "transaction", "paid", "pay",
            "입금", "통장", "거래", "수납", "납부"
        ]

        for t in tables:
            try:
                low_t = t.lower()
                if low_t.startswith("alembic"):
                    continue

                cols = [c["name"] for c in inspector.get_columns(t)]
                if not cols:
                    continue

                joined_cols = " ".join(cols).lower()

                if mode == "deposit":
                    if not any(w in low_t or w in joined_cols for w in deposit_words):
                        continue

                # 너무 많은 데이터를 한 번에 엑셀화하지 않도록 제한
                sql = text(f'SELECT * FROM "{t}" LIMIT {int(limit)}')
                rows = db.execute(sql).mappings().all()
                result.append({"table": t, "columns": cols, "rows": [dict(r) for r in rows]})
            except Exception as e:
                print("WARN: table read skipped", t, e)
                continue
        return result
    except Exception as e:
        print("WARN: get tables failed", e)
        return []


def _style_sheet(ws):
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="EAF3FF")

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if cell.row == 1:
                cell.font = Font(bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col in ws.columns:
        max_len = 8
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            max_len = max(max_len, len(str(cell.value or "")) + 2)
        ws.column_dimensions[col_letter].width = min(max_len, 35)


def _safe_sheet_title(title):
    title = str(title or "sheet")
    for ch in ['\\', '/', '*', '?', ':', '[', ']']:
        title = title.replace(ch, '_')
    title = title.replace(" ", "_")
    title = title[:31] or "sheet"
    return title


def _add_sheet(wb, title, headers, rows):
    ws = wb.create_sheet(title=_safe_sheet_title(title))
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    _style_sheet(ws)
    return ws


def _is_vehicle_like_row(row):
    keys = list(row.keys())
    return bool(
        _find_col(keys, ["차량번호", "vehicle_number", "car_no", "차량", "번호"]) or
        _find_col(keys, ["성명", "이름", "name", "owner_name"])
    )


def _row_to_member(row):
    keys = list(row.keys())

    col_region = _find_col(keys, ["지역", "region", "시군", "관할"])
    col_vehicle = _find_col(keys, ["차량번호", "vehicle_number", "car_no", "차량"])
    col_name = _find_col(keys, ["성명", "이름", "name", "owner_name"])
    col_phone = _find_col(keys, ["전화번호", "핸드폰", "휴대폰", "phone", "mobile"])
    col_approval = _find_col(keys, ["인가일자", "approval_date", "허가일자"])
    col_join = _find_col(keys, ["가입일자", "join_date", "협회가입일자"])
    col_cert = _find_col(keys, ["자격증명발급일자", "certificate_issue_date", "자격증명일자", "자격증명 발급일자"])
    col_cert_no = _find_col(keys, ["자격증명발급번호", "certificate_no", "자격증명번호"])
    col_status = _find_col(keys, ["상태", "status", "처리구분", "구분"])
    col_process = _find_col(keys, ["처리일자", "process_date", "변경일자", "폐업일자"])
    col_note = _find_col(keys, ["비고", "메모", "note", "remark"])

    return {
        "지역": _safe_str(row.get(col_region)),
        "차량번호": _safe_str(row.get(col_vehicle)),
        "성명": _safe_str(row.get(col_name)),
        "전화번호": _safe_str(row.get(col_phone)),
        "인가일자": _safe_str(row.get(col_approval)),
        "가입일자": _safe_str(row.get(col_join)),
        "자격증명발급일자": _safe_str(row.get(col_cert)),
        "자격증명발급번호": _safe_str(row.get(col_cert_no)),
        "처리구분": _safe_str(row.get(col_status)),
        "처리일자": _safe_str(row.get(col_process)),
        "상태": _safe_str(row.get(col_status)),
        "비고": _safe_str(row.get(col_note)),
        "_approval_date": _parse_date(row.get(col_approval)),
        "_join_date": _parse_date(row.get(col_join)),
        "_cert_date": _parse_date(row.get(col_cert)),
        "_process_date": _parse_date(row.get(col_process)),
    }


def _is_delivery_vehicle(m):
    car = m.get("차량번호", "")
    note = m.get("비고", "")
    status = m.get("상태", "")
    txt = f"{car} {note} {status}"
    return ("배" in car) or ("택배" in txt)


def _is_removed_status(m):
    txt = f"{m.get('처리구분','')} {m.get('상태','')} {m.get('비고','')}"
    removed_words = ["폐업", "폐지", "양도", "이관", "탈퇴", "타도", "사망"]
    return any(w in txt for w in removed_words)


def _is_joined(m):
    return bool(m.get("가입일자"))


def _billing_count_workbook(year, month, tables):
    start, end = _month_range(year, month)
    기준월 = f"{year}-{month:02d}"

    members = []
    for table in tables:
        for row in table["rows"]:
            if _is_vehicle_like_row(row):
                m = _row_to_member(row)
                if m.get("차량번호") or m.get("성명"):
                    members.append(m)

    # 중복 제거: 차량번호 + 성명 기준
    unique = {}
    for m in members:
        key = (m.get("차량번호"), m.get("성명"))
        unique[key] = m
    members = list(unique.values())

    active = [m for m in members if not _is_removed_status(m)]
    joined = [m for m in active if _is_joined(m)]
    delivery = [m for m in active if _is_delivery_vehicle(m)]

    # N월 택배관리: 인가일자 + 자격증명발급일자 둘 다 있고, 다음달부터 반영
    delivery_billable = []
    delivery_new = []
    for m in delivery:
        approval = m.get("_approval_date")
        cert = m.get("_cert_date")
        if approval and cert:
            base = max(approval, cert)
            bill_start = _next_month_after(base)
            if bill_start and bill_start <= start:
                delivery_billable.append(m)
            if bill_start and bill_start.year == year and bill_start.month == month:
                delivery_new.append(m)

    removed = [m for m in members if _is_removed_status(m)]
    transfer = [m for m in removed if "양도" in f"{m.get('처리구분','')} {m.get('상태','')} {m.get('비고','')}"]
    other_area = [m for m in removed if ("타도" in f"{m.get('처리구분','')} {m.get('상태','')} {m.get('비고','')}" or "이관" in f"{m.get('처리구분','')} {m.get('상태','')} {m.get('비고','')}")]
    closed = [m for m in removed if ("폐업" in f"{m.get('처리구분','')} {m.get('상태','')} {m.get('비고','')}" or "폐지" in f"{m.get('처리구분','')} {m.get('상태','')} {m.get('비고','')}")]
    withdrawn = [m for m in removed if "탈퇴" in f"{m.get('처리구분','')} {m.get('상태','')} {m.get('비고','')}"]

    # 70세는 주민번호/생년월일 컬럼이 있을 경우에만 잡힘. 없으면 0.
    age70 = []
    for m in active:
        txt = " ".join(str(v) for v in m.values())
        # 1956년 기준 2026년에 70세 도달. 연도별 자동.
        target_birth_year = year - 70
        if str(target_birth_year) in txt:
            age70.append(m)

    협회기본대수 = len(joined)
    n월택배관리 = len(delivery_billable)
    총부과대수 = 협회기본대수 + n월택배관리

    wb = Workbook()
    ws = wb.active
    ws.title = "부과대수 요약"

    headers = [
        "기준월", "협회기본대수", "협회가입", "양도", "타도", "폐업",
        "탈퇴", "택배신규", "관리비폐지", "70세", "N월 택배관리",
        "총부과대수", "비고"
    ]

    ws.append(headers)
    ws.append([
        기준월,
        협회기본대수,
        len(joined),
        len(transfer),
        len(other_area),
        len(closed),
        len(withdrawn),
        len(delivery_new),
        0,
        len(age70),
        n월택배관리,
        총부과대수,
        "자동 산정. 택배는 인가일자+자격증명발급일자 둘 다 있는 경우 다음 달부터 반영."
    ])
    _style_sheet(ws)

    detail_headers = [
        "지역", "차량번호", "성명", "전화번호", "인가일자", "가입일자",
        "자격증명발급일자", "자격증명발급번호", "처리구분", "처리일자", "상태", "비고"
    ]

    _add_sheet(wb, "협회가입 상세", detail_headers, joined)
    _add_sheet(wb, "양도 상세", detail_headers, transfer)
    _add_sheet(wb, "타도이관 상세", detail_headers, other_area)
    _add_sheet(wb, "폐업 상세", detail_headers, closed)
    _add_sheet(wb, "탈퇴 상세", detail_headers, withdrawn)
    _add_sheet(wb, "택배신규 상세", detail_headers, delivery_new)
    _add_sheet(wb, "관리비폐지 상세", detail_headers, [])
    _add_sheet(wb, "70세 전환 상세", detail_headers, age70)
    _add_sheet(wb, "N월 택배관리 상세", detail_headers, delivery_billable)

    return wb


def _is_payment_like_table(table_name, columns):
    t = table_name.lower()
    text = " ".join(columns).lower()
    words = ["payment", "deposit", "bank", "transaction", "입금", "통장", "거래", "수납", "납부"]
    return any(w in t or w in text for w in words)


def _payment_rows_for_month(year, month, tables):
    start, end = _month_range(year, month)
    out = []

    for table in tables:
        cols = table["columns"]
        if not _is_payment_like_table(table["table"], cols):
            continue

        col_date = _find_col(cols, ["입금일자", "거래일자", "납부일자", "date", "paid_at", "created_at", "transaction_date"])
        col_amount = _find_col(cols, ["입금액", "금액", "amount", "deposit_amount", "paid_amount"])
        col_memo = _find_col(cols, ["입금자명", "보낸분", "이체메모", "메모", "memo", "description", "내용"])
        col_name = _find_col(cols, ["성명", "이름", "name", "member_name"])
        col_vehicle = _find_col(cols, ["차량번호", "vehicle_number", "car_no"])
        col_account = _find_col(cols, ["계정", "구분", "account", "type", "category"])
        col_status = _find_col(cols, ["상태", "매칭상태", "status", "match_status"])

        if not col_date and not col_amount:
            continue

        for row in table["rows"]:
            d = _parse_date(row.get(col_date)) if col_date else None
            if d and not (start <= d <= end):
                continue

            # 날짜 컬럼이 없으면 일단 포함
            out.append({
                "기준월": f"{year}-{month:02d}",
                "자료테이블": table["table"],
                "입금일자": _safe_str(row.get(col_date)),
                "입금액": _safe_str(row.get(col_amount)),
                "입금자명/메모": _safe_str(row.get(col_memo)),
                "성명": _safe_str(row.get(col_name)),
                "차량번호": _safe_str(row.get(col_vehicle)),
                "계정": _safe_str(row.get(col_account)),
                "매칭상태": _safe_str(row.get(col_status)),
                "비고": "",
            })

    return out



def _error_workbook(title, error_message):
    wb = Workbook()
    ws = wb.active
    ws.title = "오류내용"
    ws.append(["구분", "내용"])
    ws.append(["오류", str(error_message)])
    ws.append(["안내", "서버 오류로 다운로드가 실패하지 않도록 오류 내용을 엑셀로 출력했습니다."])
    _style_sheet(ws)
    return wb



def _payment_rows_for_month_direct(db, year, month):
    """
    N? ???? ??.
    bank_transactions ????? ??? ??? ?? ????.
    LIMIT ?? ?? ? ??? ????.
    """
    from sqlalchemy import inspect, text as sql_text
    from datetime import datetime

    bind = db.get_bind()
    inspector = inspect(bind)
    tables = inspector.get_table_names()

    if "bank_transactions" not in tables:
        raise Exception("bank_transactions ???? ????.")

    cols = [c["name"] for c in inspector.get_columns("bank_transactions")]

    col_date = _find_col(cols, [
        "????", "????", "????",
        "transaction_date", "paid_at", "payment_date", "deposit_date",
        "date", "created_at"
    ])
    col_amount = _find_col(cols, [
        "???", "??", "amount", "deposit_amount", "paid_amount"
    ])
    col_memo = _find_col(cols, [
        "????", "???", "????", "??", "memo",
        "description", "??", "sender", "depositor", "payer_name"
    ])
    col_name = _find_col(cols, ["??", "??", "name", "member_name"])
    col_vehicle = _find_col(cols, ["????", "vehicle_number", "car_no"])
    col_account = _find_col(cols, ["??", "??", "account", "type", "category"])
    col_status = _find_col(cols, ["??", "????", "status", "match_status"])

    if not col_date:
        raise Exception("bank_transactions?? ????/????/created_at ??? ?? ?????.")

    start_dt = datetime(year, month, 1)
    if month == 12:
        end_dt = datetime(year + 1, 1, 1)
    else:
        end_dt = datetime(year, month + 1, 1)

    query = (
        'SELECT * FROM bank_transactions '
        'WHERE "' + col_date + '" >= :start_dt '
        'AND "' + col_date + '" < :end_dt '
        'ORDER BY "' + col_date + '" ASC'
    )

    rows = db.execute(
        sql_text(query),
        {"start_dt": start_dt, "end_dt": end_dt}
    ).mappings().all()

    out = []
    for row in rows:
        row = dict(row)
        out.append({
            "???": f"{year}-{month:02d}",
            "?????": "bank_transactions",
            "????": _safe_str(row.get(col_date)),
            "???": _safe_str(row.get(col_amount)),
            "????/??": _safe_str(row.get(col_memo)),
            "??": _safe_str(row.get(col_name)),
            "????": _safe_str(row.get(col_vehicle)),
            "??": _safe_str(row.get(col_account)),
            "????": _safe_str(row.get(col_status)),
            "??": "",
        })

    return out


def _deposit_workbook(year, month, tables):
    rows = _payment_rows_for_month(year, month, tables)

    wb = Workbook()
    ws = wb.active
    ws.title = "N월 입금추출"

    headers = [
        "기준월", "자료테이블", "입금일자", "입금액", "입금자명/메모",
        "성명", "차량번호", "계정", "매칭상태", "비고"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    _style_sheet(ws)

    # 요약 시트
    ws2 = wb.create_sheet("입금요약", 0)
    total = 0
    for r in rows:
        amt = str(r.get("입금액", "")).replace(",", "").replace("원", "").strip()
        try:
            total += int(float(amt))
        except Exception:
            pass

    ws2.append(["기준월", "입금건수", "입금합계", "비고"])
    ws2.append([f"{year}-{month:02d}", len(rows), total, "DB 내 입금/통장/거래/납부 관련 테이블에서 자동 추출"])
    _style_sheet(ws2)

    return wb


def _excel_response(wb, filename):
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    # 한글 파일명은 HTTP 헤더에서 반드시 URL 인코딩해야 함
    safe_filename = quote(filename)
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{safe_filename}"
    }
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )


@router.get("/billing-count/export")
def export_billing_count_excel(
    year: int = Query(..., description="기준년도"),
    month: int = Query(..., ge=1, le=12, description="기준월")
):
    try:
        db_gen = _get_db()
        db = next(db_gen)
        tables = _get_tables_and_rows(db, mode="all", limit=5000)
        wb = _billing_count_workbook(year, month, tables)
        filename = f"{year}년_{month:02d}월_부과대수.xlsx"
        return _excel_response(wb, filename)
    except Exception as e:
        print("ERROR: billing count export failed:", repr(e))
        wb = _error_workbook("부과대수 오류", repr(e))
        filename = f"{year}년_{month:02d}월_부과대수_오류.xlsx"
        return _excel_response(wb, filename)




def _find_recent_payment_rows(db, year, month):
    """
    N? ????:
    /arrears ??? ?? ???? ?????? YYYY-MM? ??? ?? ???.
    ?? ??? ???? / ?? / ?? / ??? ????.
    """
    from sqlalchemy import inspect, text as sql_text

    ym = f"{year}-{month:02d}"

    bind = db.get_bind()
    inspector = inspect(bind)
    tables = inspector.get_table_names()

    result = []

    recent_date_candidates = [
        "?????", "?? ???", "??????", "??? ???",
        "last_payment_date", "recent_payment_date", "latest_payment_date",
        "last_paid_at", "paid_at", "payment_date"
    ]

    vehicle_candidates = [
        "????", "?? ??", "vehicle_number", "car_number", "car_no", "plate_number"
    ]

    name_candidates = [
        "??", "??", "???", "name", "member_name", "owner_name"
    ]

    account_candidates = [
        "??", "??", "????", "account", "account_type", "category", "fee_type"
    ]

    amount_candidates = [
        "??", "???", "????", "?????", "?? ???",
        "amount", "payment_amount", "paid_amount", "deposit_amount",
        "last_payment_amount", "recent_payment_amount"
    ]

    # arrears ?? ??? ??
    preferred_words = [
        "arrears", "member", "dues", "fee", "misu", "??", "??", "??", "??"
    ]

    sorted_tables = sorted(
        tables,
        key=lambda t: 0 if any(w in t.lower() for w in preferred_words) else 1
    )

    for table in sorted_tables:
        try:
            cols = [c["name"] for c in inspector.get_columns(table)]
            if not cols:
                continue

            col_recent = _find_col(cols, recent_date_candidates)
            if not col_recent:
                continue

            col_vehicle = _find_col(cols, vehicle_candidates)
            col_name = _find_col(cols, name_candidates)
            col_account = _find_col(cols, account_candidates)
            col_amount = _find_col(cols, amount_candidates)

            # ??? ????? + ??/??/?? ? ???? ??? ?
            if not (col_vehicle or col_name or col_amount):
                continue

            q = sql_text(
                'SELECT * FROM "' + table + '" '
                'WHERE CAST("' + col_recent + '" AS TEXT) LIKE :ym '
                'ORDER BY "' + col_recent + '" ASC'
            )

            rows = db.execute(q, {"ym": ym + "%"}).mappings().all()

            for row in rows:
                row = dict(row)
                result.append({
                    "????": _safe_str(row.get(col_vehicle)),
                    "??": _safe_str(row.get(col_name)),
                    "??": _safe_str(row.get(col_account)),
                    "??": _safe_str(row.get(col_amount)),
                    "_table": table,
                    "_recent_col": col_recent,
                    "_recent_value": _safe_str(row.get(col_recent)),
                })

        except Exception as e:
            print("WARN: recent payment table skipped", table, repr(e))
            continue

    # ?? ??: ???? + ?? + ?? + ?? ??
    unique = {}
    for r in result:
        key = (
            r.get("????", ""),
            r.get("??", ""),
            r.get("??", ""),
            r.get("??", ""),
        )
        unique[key] = r

    return list(unique.values())


def _make_simple_deposit_workbook(year, month, rows):
    wb = Workbook()

    ws = wb.active
    ws.title = "summary"

    total = 0
    for r in rows:
        amt = str(r.get("??", "")).replace(",", "").replace("?", "").strip()
        try:
            total += int(float(amt))
        except Exception:
            pass

    ws.append(["???", "????", "????", "??"])
    ws.append([
        f"{year}-{month:02d}",
        len(rows),
        total,
        "?????? ???? ???? ??? ?? ? ?? ??"
    ])
    _style_sheet(ws)

    ws2 = wb.create_sheet(title="deposit_list")
    headers = ["????", "??", "??", "??"]
    ws2.append(headers)

    for r in rows:
        ws2.append([
            r.get("????", ""),
            r.get("??", ""),
            r.get("??", ""),
            r.get("??", ""),
        ])

    _style_sheet(ws2)
    return wb



@router.get("/deposit/export")
def export_deposit_excel(
    year: int = Query(..., description="????"),
    month: int = Query(..., ge=1, le=12, description="???")
):
    try:
        db_gen = _get_db()
        db = next(db_gen)

        rows = _find_recent_payment_rows(db, year, month)
        wb = _make_simple_deposit_workbook(year, month, rows)

        filename = f"{year}?_{month:02d}?_????.xlsx"
        return _excel_response(wb, filename)

    except Exception as e:
        print("ERROR: deposit export failed:", repr(e))
        wb = _error_workbook("deposit export error", repr(e))
        filename = f"{year}_{month:02d}_deposit_export_error.xlsx"
        return _excel_response(wb, filename)




# =========================================================
# ???? ???
# - ????? ?? + ???? ?4??? ?? ??? ???? ??
# =========================================================
def _digits_only(v):
    return re.sub(r"\D", "", str(v or ""))

def _compact_text(v):
    return re.sub(r"\s+", "", str(v or "")).strip()

def _vehicle_last4(v):
    d = _digits_only(v)
    return d[-4:] if len(d) >= 4 else d

def _rematch_confirm_needed_name_last4(db):
    from sqlalchemy import inspect, text as sql_text

    bind = db.get_bind()
    inspector = inspect(bind)
    tables = inspector.get_table_names()

    if "bank_transactions" not in tables:
        raise Exception("bank_transactions table not found")

    bank_cols = [c["name"] for c in inspector.get_columns("bank_transactions")]

    bank_pk = None
    try:
        pk_cols = inspector.get_pk_constraint("bank_transactions").get("constrained_columns") or []
        if pk_cols:
            bank_pk = pk_cols[0]
    except Exception:
        pass

    if not bank_pk:
        bank_pk = _find_col(bank_cols, ["id", "transaction_id", "bank_transaction_id"])

    if not bank_pk:
        raise Exception("bank_transactions primary key not found")

    col_memo = _find_col(bank_cols, ["????", "??", "????", "???", "memo", "description", "sender", "depositor", "payer_name"])
    col_status = _find_col(bank_cols, ["??", "????", "status", "match_status"])
    col_target = _find_col(bank_cols, ["????", "??", "matched_target", "match_target"])
    col_reason = _find_col(bank_cols, ["??", "????", "reason", "match_reason"])
    col_amount = _find_col(bank_cols, ["???", "??", "amount", "deposit_amount", "paid_amount"])

    if not col_memo or not col_status:
        raise Exception("bank_transactions memo/status column not found")

    # 1) ??/???/?? ????? ??+???? ?? ??
    member_candidates = []

    for t in tables:
        if t == "bank_transactions":
            continue

        try:
            cols = [c["name"] for c in inspector.get_columns(t)]
            col_name = _find_col(cols, ["??", "??", "???", "name", "member_name", "owner_name"])
            col_vehicle = _find_col(cols, ["????", "?? ??", "vehicle_number", "car_number", "car_no", "plate_number"])
            col_account = _find_col(cols, ["??", "??", "????", "account", "account_type", "category", "fee_type"])

            if not col_name or not col_vehicle:
                continue

            q = sql_text(
                'SELECT "' + col_name + '" AS name, "' + col_vehicle + '" AS vehicle'
                + (', "' + col_account + '" AS account' if col_account else ', NULL AS account')
                + ' FROM "' + t + '"'
            )

            for r in db.execute(q).mappings().all():
                name = _safe_str(r.get("name"))
                vehicle = _safe_str(r.get("vehicle"))
                last4 = _vehicle_last4(vehicle)

                if name and last4:
                    member_candidates.append({
                        "name": name,
                        "vehicle": vehicle,
                        "last4": last4,
                        "account": _safe_str(r.get("account")),
                        "table": t,
                    })

        except Exception as e:
            print("WARN: member candidate table skipped", t, repr(e))
            continue

    # ?? ?? ??
    unique_candidates = {}
    for c in member_candidates:
        key = (c["name"], c["vehicle"], c["last4"], c["account"])
        unique_candidates[key] = c
    member_candidates = list(unique_candidates.values())

    # 2) bank_transactions ? ????? ?? ??? ????
    # ???? ????? ???? DB ???/???? ?? ? ???
    # ???? LIKE ???? ???? ???.
    q = sql_text(
        'SELECT * FROM "bank_transactions" '
        'WHERE COALESCE(CAST("' + col_status + '" AS TEXT), \'\') NOT LIKE :done_status'
    )
    bank_rows = db.execute(q, {"done_status": "%????%"}).mappings().all()

    updated = 0
    skipped_multi = 0
    skipped_none = 0
    examples = []

    for row in bank_rows:
        row = dict(row)
        memo = _compact_text(row.get(col_memo))

        matched = []
        for c in member_candidates:
            name_c = _compact_text(c["name"])
            last4 = c["last4"]

            if name_c and last4 and name_c in memo and last4 in memo:
                matched.append(c)

        # ??+?4??? ? 1?? ??? ?? ??
        if len(matched) == 1:
            c = matched[0]

            set_parts = ['"' + col_status + '" = :new_status']
            params = {
                "new_status": "????",
                "pk": row.get(bank_pk),
            }

            if col_target:
                set_parts.append('"' + col_target + '" = :target')
                params["target"] = f'{c["name"]} {c["vehicle"]}'.strip()

            if col_reason:
                set_parts.append('"' + col_reason + '" = :reason')
                params["reason"] = "??+?????4????"

            update_q = sql_text(
                'UPDATE "bank_transactions" SET '
                + ", ".join(set_parts)
                + ' WHERE "' + bank_pk + '" = :pk'
            )

            db.execute(update_q, params)
            updated += 1

            if len(examples) < 20:
                examples.append({
                    "memo": _safe_str(row.get(col_memo)),
                    "matched": f'{c["name"]} {c["vehicle"]}',
                    "amount": _safe_str(row.get(col_amount)) if col_amount else "",
                })

        elif len(matched) > 1:
            skipped_multi += 1
        else:
            skipped_none += 1

    db.commit()

    return {
        "status": "ok",
        "checked": len(bank_rows),
        "updated": updated,
        "skipped_none": skipped_none,
        "skipped_multi": skipped_multi,
        "examples": examples,
    }


@router.get("/deposit/rematch-name-last4")
def rematch_deposit_name_last4():
    try:
        db_gen = _get_db()
        db = next(db_gen)
        result = _rematch_confirm_needed_name_last4(db)
        return result
    except Exception as e:
        print("ERROR: rematch failed:", repr(e))
        return {
            "status": "error",
            "message": repr(e),
        }




# =========================================================
# ???? ???? ??? v2
# - DB ???? ??? '????'? ???? ?? ???/????? ??
# - ????? ?? + ???? ?4??? ?? ??? ???? ??
# =========================================================
def _find_confirm_needed_transaction_source(db):
    from sqlalchemy import inspect, text as sql_text

    bind = db.get_bind()
    inspector = inspect(bind)
    tables = inspector.get_table_names()

    status_words = ["????", "?"]
    memo_candidates = ["????", "??", "????", "???", "memo", "description", "sender", "depositor", "payer_name"]
    amount_candidates = ["???", "??", "amount", "deposit_amount", "paid_amount"]
    status_candidates = ["??", "????", "status", "match_status", "matching_status"]

    best = None

    for t in tables:
        try:
            cols = [c["name"] for c in inspector.get_columns(t)]
            if not cols:
                continue

            memo_col = _find_col(cols, memo_candidates)
            amount_col = _find_col(cols, amount_candidates)

            # ??/?? ??? ??? ?? ???? ???? ??
            if not memo_col:
                continue

            # ?? ?? ?? ???? ???? ??
            possible_status_cols = []
            for c in cols:
                cn = _norm_col(c)
                if any(_norm_col(x) in cn for x in status_candidates):
                    possible_status_cols.append(c)

            # ??? ??? ?? ??? ?? ??? ??
            if not possible_status_cols:
                possible_status_cols = cols

            for status_col in possible_status_cols:
                try:
                    q = sql_text(
                        'SELECT COUNT(*) AS cnt FROM "' + t + '" '
                        'WHERE CAST("' + status_col + '" AS TEXT) LIKE :kw1 '
                        'OR CAST("' + status_col + '" AS TEXT) LIKE :kw2'
                    )
                    cnt = db.execute(q, {"kw1": "%????%", "kw2": "%?%"}).scalar() or 0

                    if cnt > 0:
                        if best is None or cnt > best["count"]:
                            best = {
                                "table": t,
                                "status_col": status_col,
                                "memo_col": memo_col,
                                "amount_col": amount_col,
                                "count": int(cnt),
                                "columns": cols,
                            }
                except Exception:
                    continue

        except Exception as e:
            print("WARN: confirm source scan skipped", t, repr(e))
            continue

    return best


def _rematch_confirm_needed_name_last4_v2(db):
    from sqlalchemy import inspect, text as sql_text

    bind = db.get_bind()
    inspector = inspect(bind)
    tables = inspector.get_table_names()

    source = _find_confirm_needed_transaction_source(db)

    if not source:
        return {
            "status": "error",
            "message": "DB?? ????? ???? ?? ???/??? ?? ?????.",
            "checked": 0,
            "updated": 0,
            "examples": [],
        }

    tx_table = source["table"]
    col_status = source["status_col"]
    col_memo = source["memo_col"]
    col_amount = source["amount_col"]

    tx_cols = source["columns"]

    tx_pk = None
    try:
        pk_cols = inspector.get_pk_constraint(tx_table).get("constrained_columns") or []
        if pk_cols:
            tx_pk = pk_cols[0]
    except Exception:
        pass

    if not tx_pk:
        tx_pk = _find_col(tx_cols, ["id", "transaction_id", "bank_transaction_id", "payment_id"])

    if not tx_pk:
        raise Exception(f"{tx_table} primary key not found")

    col_target = _find_col(tx_cols, ["????", "??", "matched_target", "match_target"])
    col_reason = _find_col(tx_cols, ["??", "????", "reason", "match_reason"])

    # ?? ?? ??
    member_candidates = []

    for t in tables:
        if t == tx_table:
            continue

        try:
            cols = [c["name"] for c in inspector.get_columns(t)]
            col_name = _find_col(cols, ["??", "??", "???", "name", "member_name", "owner_name"])
            col_vehicle = _find_col(cols, ["????", "?? ??", "vehicle_number", "car_number", "car_no", "plate_number"])
            col_account = _find_col(cols, ["??", "??", "????", "account", "account_type", "category", "fee_type"])

            if not col_name or not col_vehicle:
                continue

            q = sql_text(
                'SELECT "' + col_name + '" AS name, "' + col_vehicle + '" AS vehicle'
                + (', "' + col_account + '" AS account' if col_account else ', NULL AS account')
                + ' FROM "' + t + '"'
            )

            for r in db.execute(q).mappings().all():
                name = _safe_str(r.get("name"))
                vehicle = _safe_str(r.get("vehicle"))
                last4 = _vehicle_last4(vehicle)

                if name and last4:
                    member_candidates.append({
                        "name": name,
                        "vehicle": vehicle,
                        "last4": last4,
                        "account": _safe_str(r.get("account")),
                        "table": t,
                    })

        except Exception as e:
            print("WARN: member candidate table skipped", t, repr(e))
            continue

    # ?? ?? ??
    unique = {}
    for c in member_candidates:
        unique[(c["name"], c["vehicle"], c["last4"], c["account"])] = c
    member_candidates = list(unique.values())

    # ?? ???? ?? ??
    q = sql_text(
        'SELECT * FROM "' + tx_table + '" '
        'WHERE CAST("' + col_status + '" AS TEXT) LIKE :kw1 '
        'OR CAST("' + col_status + '" AS TEXT) LIKE :kw2'
    )
    tx_rows = db.execute(q, {"kw1": "%????%", "kw2": "%?%"}).mappings().all()

    updated = 0
    skipped_none = 0
    skipped_multi = 0
    examples = []

    for row in tx_rows:
        row = dict(row)
        memo = _compact_text(row.get(col_memo))

        matched = []
        for c in member_candidates:
            name_c = _compact_text(c["name"])
            last4 = c["last4"]

            if name_c and last4 and name_c in memo and last4 in memo:
                matched.append(c)

        # ? 1??? ???? ??? ?? ??
        if len(matched) == 1:
            c = matched[0]

            set_parts = ['"' + col_status + '" = :new_status']
            params = {
                "new_status": "????",
                "pk": row.get(tx_pk),
            }

            if col_target:
                set_parts.append('"' + col_target + '" = :target')
                params["target"] = f'{c["name"]} {c["vehicle"]}'.strip()

            if col_reason:
                old_reason = _safe_str(row.get(col_reason))
                set_parts.append('"' + col_reason + '" = :reason')
                params["reason"] = (old_reason + ", " if old_reason else "") + "??+?????4????"

            update_q = sql_text(
                'UPDATE "' + tx_table + '" SET '
                + ", ".join(set_parts)
                + ' WHERE "' + tx_pk + '" = :pk'
            )

            db.execute(update_q, params)
            updated += 1

            if len(examples) < 30:
                examples.append({
                    "memo": _safe_str(row.get(col_memo)),
                    "matched": f'{c["name"]} {c["vehicle"]}',
                    "amount": _safe_str(row.get(col_amount)) if col_amount else "",
                })

        elif len(matched) > 1:
            skipped_multi += 1
        else:
            skipped_none += 1

    db.commit()

    return {
        "status": "ok",
        "source_table": tx_table,
        "status_col": col_status,
        "memo_col": col_memo,
        "checked": len(tx_rows),
        "updated": updated,
        "skipped_none": skipped_none,
        "skipped_multi": skipped_multi,
        "examples": examples,
    }


@router.get("/deposit/rematch-name-last4-v2")
def rematch_deposit_name_last4_v2():
    try:
        db_gen = _get_db()
        db = next(db_gen)
        return _rematch_confirm_needed_name_last4_v2(db)
    except Exception as e:
        print("ERROR: rematch v2 failed:", repr(e))
        return {
            "status": "error",
            "message": repr(e),
        }




# =========================================================
# ???? ??? v3
# - ??? ???? 81? ?? ?? ??
# - ????? ????+??????? ?? ???????+????? ?? ??? ????
# - ?: ???1780, ???2970, 1347???, ??80?1629???
# =========================================================
def _memo_has_korean_name_with_tail_number(memo):
    memo = _compact_text(memo)

    # ?? 2~5? + ?? 3~4??
    # ?: ???1780, ???304, ???1756
    if re.search(r"[?-?]{2,5}\d{3,4}", memo):
        return True

    # ?? 3~4?? + ?? 2~5?
    # ?: 1347???, 1442???
    if re.search(r"\d{3,4}[?-?]{2,5}", memo):
        return True

    # ???? ?? + ??
    # ?: ??80?1629???, ??81?2589?
    if re.search(r"[?-?]{2}\d{2}[?-?]\d{3,4}[?-?]{1,5}", memo):
        return True

    return False


def _rematch_confirm_needed_by_visible_memo_pattern(db):
    from sqlalchemy import inspect, text as sql_text

    bind = db.get_bind()
    inspector = inspect(bind)
    tables = inspector.get_table_names()

    # ?? ????? ???? ???/?? ????? v2 ?? ???
    source = _find_confirm_needed_transaction_source(db)

    if not source:
        return {
            "status": "error",
            "message": "????? ???? ?? ???/??? ?? ?????.",
            "checked": 0,
            "updated": 0,
            "examples": [],
        }

    tx_table = source["table"]
    col_status = source["status_col"]
    col_memo = source["memo_col"]
    col_amount = source["amount_col"]
    tx_cols = source["columns"]

    tx_pk = None
    try:
        pk_cols = inspector.get_pk_constraint(tx_table).get("constrained_columns") or []
        if pk_cols:
            tx_pk = pk_cols[0]
    except Exception:
        pass

    if not tx_pk:
        tx_pk = _find_col(tx_cols, ["id", "transaction_id", "bank_transaction_id", "payment_id"])

    if not tx_pk:
        raise Exception(f"{tx_table} primary key not found")

    col_reason = _find_col(tx_cols, ["??", "????", "reason", "match_reason"])

    q = sql_text(
        'SELECT * FROM "' + tx_table + '" '
        'WHERE CAST("' + col_status + '" AS TEXT) LIKE :kw1 '
        'OR CAST("' + col_status + '" AS TEXT) LIKE :kw2'
    )

    rows = db.execute(q, {"kw1": "%????%", "kw2": "%?%"}).mappings().all()

    checked = len(rows)
    updated = 0
    skipped = 0
    examples = []

    for row in rows:
        row = dict(row)
        memo = _safe_str(row.get(col_memo))
        reason = _safe_str(row.get(col_reason)) if col_reason else ""

        # ????? ?? ? ? ???,
        # ??? ??+??? ??? ??? ????
        # ???? ??? ?? ??? ?? ????
        can_update = False

        if "????" in reason and _memo_has_korean_name_with_tail_number(memo):
            can_update = True

        if "?????????" in reason and _memo_has_korean_name_with_tail_number(memo):
            can_update = True

        if not can_update:
            skipped += 1
            continue

        set_parts = ['"' + col_status + '" = :new_status']
        params = {
            "new_status": "????",
            "pk": row.get(tx_pk),
        }

        if col_reason:
            old_reason = _safe_str(row.get(col_reason))
            set_parts.append('"' + col_reason + '" = :reason')
            params["reason"] = (old_reason + ", " if old_reason else "") + "??? ??+??????? ????"

        update_q = sql_text(
            'UPDATE "' + tx_table + '" SET '
            + ", ".join(set_parts)
            + ' WHERE "' + tx_pk + '" = :pk'
        )

        db.execute(update_q, params)
        updated += 1

        if len(examples) < 50:
            examples.append({
                "memo": memo,
                "amount": _safe_str(row.get(col_amount)) if col_amount else "",
                "reason": reason,
            })

    db.commit()

    return {
        "status": "ok",
        "source_table": tx_table,
        "status_col": col_status,
        "memo_col": col_memo,
        "checked": checked,
        "updated": updated,
        "skipped": skipped,
        "examples": examples,
    }


@router.get("/deposit/rematch-visible-pattern")
def rematch_deposit_visible_pattern():
    try:
        db_gen = _get_db()
        db = next(db_gen)
        return _rematch_confirm_needed_by_visible_memo_pattern(db)
    except Exception as e:
        print("ERROR: rematch visible pattern failed:", repr(e))
        return {
            "status": "error",
            "message": repr(e),
        }

