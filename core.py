"""
core.py - 핵심 파싱/처리 엔진

담당:
- 차량번호/성명 정규화
- 음수 금액 정확 파싱 (-35.000 → -35000)
- 검증사유 생성 (사유 있을 때만 검증필요)
- 상태 우선순위 판단
- 시트 타입 감지
- 월별 원장 가로→세로 변환
- 상태정리 시트 파싱
- 전체자 대조 매칭
- 통장매칭 (안전 규칙 포함)
- 자동이체/정기납부 분류
- 문자대상 번호/주소 정제
"""
import re, math
from typing import Optional, List, Dict, Tuple
import pandas as pd

# ── 지역 정규화 ──────────────────────────────────────────────────────────────
REGION_MAP = {
    "춘천":"춘천시","강릉":"강릉시","원주":"원주시","동해":"동해시",
    "삼척":"삼척시","속초":"속초시","태백":"태백시","홍천":"홍천군",
    "횡성":"횡성군","영월":"영월군","평창":"평창군","정선":"정선군",
    "철원":"철원군","화천":"화천군","양구":"양구군","인제":"인제군",
    "고성":"고성군","양양":"양양군",
}
for k in list(REGION_MAP.keys()):
    REGION_MAP[k+"시"] = REGION_MAP[k]
    REGION_MAP[k+"군"] = REGION_MAP[k]

def normalize_region(r: str) -> str:
    s = _s(r).replace("강원특별자치도","").replace("강원도","").strip()
    s = re.sub(r"[\(（].*$","",s).strip()
    if s in REGION_MAP: return REGION_MAP[s]
    for k, v in REGION_MAP.items():
        if k and k in s: return v
    return s


# ── 합계행/집계행 판별 ──────────────────────────────────────────────────────────
SUM_ROW_NAMES = frozenset({
    "합계","총계","소계","계","합산","인원수","입금금액","누락",
    "부과합계","입금합계","미수합계","지역합계","협회비합계","관리비합계",
    "total","sum","subtotal",
})
SUM_ROW_PATTERNS = re.compile(
    r"^합계$|^총계$|^소계$|^계$|^합산$"
    r"|합계\s*\d*$|^\d*\s*합계|외\s*\d+대|\d+대\s*$"
    r"|^지역$|^성명$|^차량번호$|^이름$"  # 헤더 반복 행
, re.IGNORECASE)

def is_sum_row(name: str, vehicle_no: str) -> bool:
    """합계/집계/헤더 반복 행이면 True → 회원으로 저장하지 않음"""
    n = norm_name(name or "")
    v = _s(vehicle_no or "").replace(" ","")
    if n in SUM_ROW_NAMES: return True
    if v in SUM_ROW_NAMES: return True
    if n and SUM_ROW_PATTERNS.search(name or ""): return True
    if v and SUM_ROW_PATTERNS.search(vehicle_no or ""): return True
    # 이름/차량번호 둘 다 없으면 집계행일 가능성
    if not n and not v: return True
    return False

# ── 기본 문자열 ──────────────────────────────────────────────────────────────
def _s(v) -> str:
    if v is None: return ""
    try:
        if isinstance(v, float) and math.isnan(v): return ""
    except: pass
    s = str(v).strip().replace("\u3000"," ")
    if s.lower() in {"nan","none","nat","#n/a","#value!","#ref!"}: return ""
    return re.sub(r"\s+"," ",s)

def norm_name(n: str) -> str:
    """성명 공백 제거 + 소문자"""
    return re.sub(r"\s+","",_s(n)).lower()

def norm_vehicle(v: str) -> str:
    """
    강원 82배 1330 호 → 82배1330
    공백/강원/호 제거, 대문자
    """
    s = _s(v)
    s = re.sub(r"강원특별자치도|강원도|강원","",s)
    s = re.sub(r"호\s*$","",s.strip())
    s = re.sub(r"\s+","",s)
    return s.upper()

def veh_last4(v: str) -> str:
    """차량번호 끝 4자리 숫자"""
    nums = re.findall(r"\d+", re.sub(r"[가-힣]","", _s(v)))
    j = "".join(nums)
    return j[-4:] if len(j) >= 4 else j

# ── 음수 금액 파싱 (핵심) ────────────────────────────────────────────────────
def parse_amount(v, allow_negative=True) -> int:
    """
    -35.000 → -35000  (소수점 아님, 천단위 구분)
    (35,000) → -35000
    △35,000 / ▲35,000 → -35000
    -35,000 → -35000
    - → 0
    x, X → 0
    """
    s = _s(v)
    if not s: return 0
    if s in {"-","x","X","O","o","N","n","*"}: return 0

    negative = False
    # 음수 기호 처리
    if s.startswith("(") and s.endswith(")"): negative = True; s = s[1:-1]
    if s.startswith("△") or s.startswith("▲"): negative = True; s = s[1:]
    if s.startswith("-"): negative = True; s = s[1:]

    # 소수점과 천단위 구분자 처리
    # 핵심: -35.000 → 35000 (소수점이 아닌 천단위 구분자)
    s = s.strip()
    if re.match(r"^\d+\.\d{3}$", s):
        # n.nnn 형식 → 천단위 구분자
        s = s.replace(".", "")
    elif re.match(r"^\d+,\d{3}$", s):
        s = s.replace(",", "")
    else:
        # 일반 쉼표 제거, 소수점 이하 버림
        s = re.sub(r",","",s)
        s = re.sub(r"\.\d*$","",s)

    s = re.sub(r"[^0-9]","",s)
    if not s: return 0
    try:
        result = int(s)
        return -result if (negative and allow_negative) else result
    except:
        return 0

def parse_date_str(v) -> str:
    s = _s(v)
    if not s: return ""
    if hasattr(v, "strftime"):
        try: return v.strftime("%Y-%m-%d")
        except: pass
    s = s.replace(" 00:00:00","").strip()
    m = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})",s)
    if m: return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r"(\d{2})\.(\d{1,2})\.(\d{1,2})",s)
    if m:
        yy = int(m.group(1))
        yr = 2000+yy if yy <= 40 else 1900+yy
        return f"{yr}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return s

# ── 상태 키워드 (우선순위 순) ────────────────────────────────────────────────
STATUS_KEYWORDS: Dict[str, List[str]] = {
    "사망":        ["사망","별세"],
    "말소":        ["말소","등록말소","직권말소"],
    "탈퇴":        ["탈퇴","자진탈퇴"],
    "폐업":        ["폐지","폐업","폐차"],
    "양도":        ["양도","서울양도","경기양도","전북양도","충남양도",
                    "충북양도","부산양도","인천양도","전남양도","경남양도",
                    "제주양도","광주양도","대전양도","울산양도","대구양도",
                    "세종양도","경북양도","강원양도"],
    "이관":        ["이관","타도","이전","타도이전","타시도이관","타도이관"],
    "관리비폐지":  ["관리비폐지"],
}

def detect_status(reason: str) -> Optional[str]:
    r = _s(reason).replace(" ","")
    for status, kws in STATUS_KEYWORDS.items():
        for kw in kws:
            if kw.replace(" ","") in r:
                return status
    return None

def detect_status_from_mgmt_no(mgmt_no: str) -> Optional[str]:
    """관리번호 접두어 기반 상태 판단"""
    s = _s(mgmt_no).upper()
    if s.startswith("폐-") or s.startswith("폐"): return "폐업"
    if s.startswith("양-") or s.startswith("양"): return "양도"
    if s.startswith("이-") or s.startswith("이관"): return "이관"
    return None

# ── 검증사유 생성 ────────────────────────────────────────────────────────────
def build_verify_reasons(
    excel_arr: int, calc_arr: int,
    name: str, vehicle_no: str, account: str,
    match_status: str, region: str,
    mobile: str, address: str, note: str,
) -> List[str]:
    """
    검증필요 사유 리스트 생성.
    규칙:
    - 재계산 차이는 사유로 쓰지 않음 (calc_arrears = excel_arrears)
    - 법인명/합산행 의심만 실제 사유
    - 비고에 폐업/양도 등 상태충돌 키워드 있을 때만 사유
    - 사유 없으면 빈 리스트 반환 → 검증필요 생성 안 함
    """
    reasons = []

    # 법인명/업체명 의심 (협회 대상이 아닌 법인)
    if name and any(k in name for k in ["(주)","주식회사","유한","합명","법인","회사","조합"]):
        reasons.append("법인명/업체명 의심")

    # 합산행 의심 (n대, 외n대 패턴)
    if vehicle_no and re.search(r"외\s*\d+대|\d+대\s*$", vehicle_no or ""):
        reasons.append("차량 합산행 의심")

    # 비고에 상태충돌 키워드 (폐업인데 비고에 양도 등)
    note_status = detect_status(note or "")
    if note_status:
        reasons.append(f"비고에 [{note_status}] 키워드 있음 — 상태 확인 필요")

    return reasons

# ── 번호 정제 ────────────────────────────────────────────────────────────────
EXCLUDE_PHONE_WORDS = {"대","x","결번","없음","불명","공사","(대)","주민번호","주민","없"}

def phone_clean(raw: str) -> str:
    """010-1234-5678 형식 정제. 유효하지 않으면 ""반환"""
    s = _s(raw).replace("-","").replace(" ","").replace(".","")
    if not s or any(w in s for w in EXCLUDE_PHONE_WORDS): return ""
    if len(s) < 9: return ""
    # 010으로 시작하는 11자리
    if re.match(r"^01[016789]\d{7,8}$", s):
        return f"{s[:3]}-{s[3:7]}-{s[7:]}"
    # 지역번호
    if re.match(r"^0\d{9,10}$", s):
        return s
    return ""

# ── 시트 타입 감지 ───────────────────────────────────────────────────────────
def guess_sheet_type(sheet_name: str, first_row: List[str]) -> str:
    sn = _s(sheet_name).replace(" ","")
    joined = " ".join(_s(v) for v in first_row)

    if re.search(r"\d{4}년회비내역", sn): return "ledger"
    if re.search(r"\d{4}년\(", sn):       return "ledger"   # 2025년(2)
    if re.search(r"\d{4}년$", sn):        return "status"   # 2026년
    if "월별부과대수" in sn:              return "billing_summary"

    # 컬럼 기반
    if "사유" in joined and "차량번호" in joined: return "status"
    if "이월금" in joined and ("부과금" in joined or "미수금" in joined): return "ledger"
    return "unknown"

def guess_file_year(filename: str) -> int:
    m = re.search(r"(20\d{2})", filename)
    if m: return int(m.group(1))
    from datetime import date
    return date.today().year

# ── 회비내역 파싱 ─────────────────────────────────────────────────────────────
def parse_ledger_sheet(raw: pd.DataFrame, sheet_name: str, file_year: int):
    """
    반환: List[(raw_veh, raw_name, raw_region, raw_account, raw_note,
                 carry_over, monthly_dict, source_row)]
    monthly_dict: {month: (charge, paid, arrears, paid_date)}
    """
    if raw is None or raw.empty: return []
    rows = raw.values.tolist()

    # 헤더 행 찾기
    header_idx = 0
    for i, row in enumerate(rows[:15]):
        vals = [_s(v) for v in row]
        joined = " ".join(vals)
        if "차량번호" in joined and ("이월금" in joined or "부과금" in joined):
            header_idx = i; break

    headers = [_s(v) for v in rows[header_idx]]

    def fi(names):
        for nm in names:
            for i,h in enumerate(headers):
                if nm in h: return i
        return None

    i_region  = fi(["지역"])
    i_account = fi(["계정"])
    i_note    = fi(["비고"])
    i_vehicle = fi(["차량번호","차 량"])
    i_name    = fi(["성     명","성명","성 명"])
    i_carry   = fi(["이월금"])

    # 월별 컬럼 인덱스 추출
    monthly_cols: Dict[int, Tuple] = {}
    for i,h in enumerate(headers):
        m = re.search(r"(\d{1,2})월\s*부과금", h)
        if not m: continue
        mo = int(m.group(1))
        ic = i
        ip = id_ = ia = None
        for j in range(i+1, min(i+5, len(headers))):
            hj = headers[j]
            if re.search(r"입금액|입 금", hj) and ip is None: ip = j
            elif "날짜" in hj and id_ is None: id_ = j
            elif re.search(r"(\d{1,2})월\s*미수금|미수금", hj) and ia is None: ia = j
        if ia is None and ip is not None:
            cand = ip+1
            if cand < len(headers) and "날짜" in headers[cand]: cand += 1
            if cand < len(headers): ia = cand
        monthly_cols[mo] = (ic, ip, id_, ia)

    if not monthly_cols: return []

    results = []
    for ridx in range(header_idx+1, len(rows)):
        row = rows[ridx]
        def cell(idx):
            if idx is None or idx >= len(row): return ""
            return _s(row[idx])

        veh  = cell(i_vehicle)
        name = cell(i_name)
        nk   = norm_name(name)
        if not veh and not nk: continue
        if is_sum_row(name, veh): continue  # 합계/집계/헤더 반복 행 제외

        region  = cell(i_region)
        account = cell(i_account)
        note    = cell(i_note)
        carry   = parse_amount(cell(i_carry), allow_negative=True)

        monthly = {}
        for mo,(ic,ip,id_,ia) in monthly_cols.items():
            charge  = parse_amount(cell(ic))
            paid    = parse_amount(cell(ip)) if ip is not None else 0
            arrears = parse_amount(cell(ia), allow_negative=True) if ia is not None else 0
            pdate   = parse_date_str(cell(id_)) if id_ is not None else ""
            if charge == 0 and paid == 0 and arrears == 0: continue
            monthly[mo] = (charge, paid, arrears, pdate)

        if not monthly: continue
        results.append((veh, name, region, account, note, carry, monthly, ridx+1))
    return results

# ── 상태정리 시트 파싱 ────────────────────────────────────────────────────────
def parse_status_sheet(raw: pd.DataFrame, file_year: int):
    """
    반환: List[(raw_veh, raw_name, raw_region, raw_account, reason, arrears, src_row, event_month)]
    """
    if raw is None or raw.empty: return []
    rows = raw.values.tolist()

    header_idx = 0
    for i,row in enumerate(rows[:5]):
        vals = [_s(v) for v in row]
        if "차량번호" in vals and any(k in " ".join(vals) for k in ["사유","성명","성     명"]):
            header_idx = i; break

    headers = [_s(v) for v in rows[header_idx]]
    def fi(names):
        for nm in names:
            for i,h in enumerate(headers):
                if nm in h: return i
        return None

    i_region  = fi(["지역"])
    i_account = fi(["계정"])
    i_vehicle = fi(["차량번호"])
    i_name    = fi(["성     명","성명"])
    i_reason  = fi(["사 유","사유"])
    # 최신 미수금 컬럼
    i_arrears = None
    for i in range(len(headers)-1,-1,-1):
        if "미수금" in headers[i]: i_arrears = i; break

    current_month = None
    results = []
    for ridx in range(header_idx+1, len(rows)):
        row = rows[ridx]
        def cell(idx):
            if idx is None or idx >= len(row): return ""
            return _s(row[idx])

        all_vals = [_s(v) for v in row if _s(v)]
        if len(all_vals) <= 2:
            joined = " ".join(all_vals)
            m = re.search(r"<(\d{1,2})월>|<(\d{1,2})월",joined)
            if m:
                mo = m.group(1) or m.group(2)
                current_month = f"{file_year}-{int(mo):02d}"
            continue

        veh    = cell(i_vehicle)
        name   = cell(i_name)
        if not veh and not name: continue
        nk = norm_name(name)
        if is_sum_row(name, veh): continue

        reason  = cell(i_reason)
        if not reason: continue
        arrears = parse_amount(cell(i_arrears), allow_negative=True) if i_arrears else 0

        results.append((veh, name, cell(i_region), cell(i_account),
                        reason, arrears, ridx+1, current_month))
    return results

# ── 통장매칭 ──────────────────────────────────────────────────────────────────
# 자동매칭 금지 메모 패턴
BANK_SKIP_PATTERNS = re.compile(
    r"PC신한|PC국민|PC하나|PC우리|PC농협|PC기업|스마트뱅킹|인터넷뱅킹|"
    r"자동이체|CMS|수수료|이체수수료|ATM|CD기|텔레뱅킹|모바일뱅킹|"
    r"^E-|^이체$|^출금$"
)

def extract_memo_keys(memo: str):
    """이체메모에서 (성명후보, 차량번호key, 끝4자리) 추출"""
    raw = _s(memo)
    vkey = norm_vehicle(raw)
    nums = re.findall(r"\d+", raw)
    joined = "".join(nums)
    l4 = joined[-4:] if len(joined) >= 4 else joined

    # 차량번호 패턴 제거 후 한글 이름 추출
    name_part = re.sub(r"강원\s*\d{2,3}\s*[가-힣]\s*\d{3,4}\s*호?","",raw)
    name_part = re.sub(r"\d+","",name_part)
    for w in ["협회비","관리비","협","관","입금","자동이체","국민","하나","농협",
              "카카오","스마트","기업","우리","신한","씨티","전북","SC","IBK"]:
        name_part = name_part.replace(w," ")
    hanguls = re.findall(r"[가-힣]{2,5}", name_part)
    name_c = re.sub(r"\s+","","".join(hanguls)) if hanguls else ""
    return name_c, vkey, l4

def is_useless_memo(memo: str) -> bool:
    m = _s(memo).replace(" ","")
    if not m: return True
    if m in {"협회비","관리비","협","관","70세","합계","계","총계"}: return True
    if BANK_SKIP_PATTERNS.search(memo): return True
    name_c, vkey, l4 = extract_memo_keys(memo)
    return not name_c and not vkey and not l4

def score_bank_match(memo: str, deposit: int, member) -> Tuple[int,str,str]:
    """
    반환: (score, reason, tier)
    tier: exact/vkey/l4_name/name_exact/l4/none

    점수 기준:
      성명 정확일치: +40
      차량번호 전체일치: +50
      차량번호 뒷4자리 일치: +35
      입금액 = 현재미수금: +20
      입금액 = 월납부금액(정기): +10
      전화번호 일치: +10
    """
    name_c, memo_vkey, memo_l4 = extract_memo_keys(memo)
    score = 0; reasons = []; tier = "none"

    # 차량번호 전체 일치
    vkey_ok = bool(memo_vkey and len(memo_vkey) >= 5 and member.vehicle_key and memo_vkey == member.vehicle_key)
    if vkey_ok:
        score += 50; reasons.append("차량번호전체일치"); tier = "vkey"

    # 뒷자리 4자리 일치
    ml4 = veh_last4(member.vehicle_no)
    l4_ok = bool(memo_l4 and ml4 and len(memo_l4) >= 3 and memo_l4 == ml4)
    if l4_ok:
        score += 35; reasons.append("차량번호뒷자리일치")

    # 성명 정확일치만 인정 (부분일치는 점수 없음)
    name_ok = False
    if name_c and norm_name(name_c) == norm_name(member.name):
        score += 40; reasons.append("성명일치"); name_ok = True

    # 티어 결정
    if name_ok and vkey_ok:          tier = "exact"
    elif name_ok and l4_ok:          tier = "exact"
    elif vkey_ok:                    tier = "vkey"
    elif name_ok:                    tier = "name_exact"
    elif l4_ok:                      tier = "l4"

    # 금액 일치 보너스 (현재미수금과 일치)
    due = member.excel_arrears or 0
    if deposit and due > 0 and deposit == due:
        score += 20; reasons.append("금액=미수금일치")

    # 월납부 정기금액 일치 보너스 (협회비 30000, 관리비 50000/100000 등 정기금액)
    MONTHLY_AMOUNTS = {30000, 50000, 60000, 100000, 120000, 150000}
    if deposit and deposit in MONTHLY_AMOUNTS:
        score += 10; reasons.append("월납부금액일치")

    return score, ", ".join(reasons), tier

# ── 자동이체/정기납부 분류 ───────────────────────────────────────────────────
AUTOPAY_AMOUNTS = {5000, 10000, 15000, 30000, 60000}

def classify_autopay(payments: List[Tuple[int,str]]) -> str:
    """
    payments: [(amount, date_str), ...]
    반환: '자동이체정상'/'정기납부3개월'/'정기납부6개월'/'정기납부월'/'해당없음'/'입금끊김'
    """
    if len(payments) < 2: return "해당없음"
    amounts = [p[0] for p in payments]
    # 금액이 인정 범위
    unique_amts = set(amounts)
    valid_amounts = unique_amts & AUTOPAY_AMOUNTS
    if not valid_amounts: return "해당없음"

    main_amt = max(valid_amounts, key=lambda a: amounts.count(a))
    same = [p for p in payments if p[0] == main_amt]
    if len(same) < 2: return "해당없음"

    # 날짜 간격 분석
    dates = sorted(p[1] for p in same if p[1])
    if len(dates) < 2: return "해당없음"

    gaps = []
    for i in range(1, len(dates)):
        try:
            from datetime import date
            d1 = date.fromisoformat(dates[i-1][:10])
            d2 = date.fromisoformat(dates[i][:10])
            gaps.append(abs((d2-d1).days))
        except (ValueError, AttributeError, TypeError):
            pass
    if not gaps: return "해당없음"

    avg_gap = sum(gaps)/len(gaps)
    if 25 <= avg_gap <= 35: pat = "정기납부월"
    elif 80 <= avg_gap <= 100: pat = "정기납부3개월"
    elif 170 <= avg_gap <= 200: pat = "정기납부6개월"
    else: return "해당없음"

    # 최근 입금 여부
    from datetime import date, timedelta
    try:
        last = date.fromisoformat(max(dates)[:10])
        if (date.today() - last).days > 200:
            return "입금끊김"
    except (ValueError, AttributeError, TypeError):
        pass

    return pat

# ── 주소 정제 ────────────────────────────────────────────────────────────────
def address_clean(m) -> str:
    """회원 객체에서 우선순위 주소 반환"""
    for addr in [m.official_address, m.address]:
        if addr and addr.strip():
            return addr.strip()
    return ""


# ── 부과대수 파일 파서 ────────────────────────────────────────────────────────

# 항목명 → 정규명 매핑 (공백 제거 후 비교)
BILLING_ITEM_ALIASES: Dict[str, str] = {}
_BILLING_RAW_ALIASES = {
    "협회가입":     ["협회가입","가입","협회가입수"],
    "양도":         ["양도","양도양수","양도건수"],
    "타도":         ["타도","타지역이전","타도이관","타도전출"],
    "폐지":         ["폐지","폐업","폐지건수","차량폐지"],
    "탈퇴":         ["탈퇴","탈퇴건수"],
    "택배신규":     ["택배신규","택배 신규","배신규","택배신규건수"],
    "관리비폐지":   ["관리비폐지","관리비폐지","관리비제외","관리비폐지건수"],
    "70세":         ["70세","70세전환","고령자전환","70세이상","70세이상전환"],
    "협회기본대수": ["협회기본대수","기본대수","협회기본","협회원기본대수"],
    "총부과대수":   ["총부과대수","총부과","총부과건수"],
    "택배관리":     ["택배관리","해당월택배관리","택배관리비건수","당월택배관리"],
}
for _canon, _aliases in _BILLING_RAW_ALIASES.items():
    for _a in _aliases:
        BILLING_ITEM_ALIASES[_a.replace(" ", "")] = _canon

# 처리대기 생성 대상 항목 (WorkQueue/BillingPerson으로 관리해야 하는 항목)
BILLING_WORK_TYPES = {"관리비폐지", "70세", "택배신규", "협회가입", "탈퇴", "폐지", "양도", "타도"}

# 항목 → 계정 추론
BILLING_ACCOUNT_MAP = {
    "협회가입": "협", "탈퇴": "협", "양도": "협", "타도": "협", "폐지": "협",
    "관리비폐지": "관", "70세": "관", "택배신규": "택배",
    "협회기본대수": "", "총부과대수": "", "택배관리": "관",
}

# 항목 → 변경상태 추론
BILLING_TO_STATUS_MAP = {
    "폐지": "폐업", "탈퇴": "탈퇴", "양도": "양도", "타도": "이관",
    "협회가입": "정상", "관리비폐지": "관리비폐지", "70세": "정상",
    "택배신규": "정상",
}


def _nc(v) -> str:
    """셀 값 정규화 — None/NaN → ''"""
    if v is None: return ""
    if isinstance(v, float) and math.isnan(v): return ""
    return str(v).strip()


def _is_vehicle_no(s: str) -> bool:
    """차량번호 패턴 감지"""
    return bool(re.search(r'[가-힣]{1,3}\s*\d{2}\s*[가-힣]\s*\d{4}', s))


def _is_name(s: str) -> bool:
    """한국 성명 패턴 (2~5자 한글)"""
    return bool(re.match(r'^[가-힣]{2,5}$', s.strip()))


def _try_int(v) -> Optional[int]:
    try:
        return int(float(str(v).replace(",", "").strip()))
    except:
        return None


def parse_billing_file(xl_path: str, base_year: int, base_month: int) -> Dict:
    """
    부과대수 엑셀 파일 파싱.
    Returns:
      {
        "counts": {"협회가입": N, "폐지": N, ...},   # 집계 카운트
        "persons": [...],                             # 개인별 처리대기 항목
        "sheets": [...],                              # 처리된 시트명 목록
        "parse_log": [...],                           # 파싱 로그 (디버그용)
      }
    """
    import pandas as _pd
    counts = {k: 0 for k in _BILLING_RAW_ALIASES}
    persons: List[Dict] = []
    parse_log: List[str] = []
    processed_sheets: List[str] = []

    try:
        p = str(xl_path).lower()
        if p.endswith((".xlsx", ".xlsm")):
            xl = _pd.ExcelFile(xl_path, engine="openpyxl")
        else:
            xl = _pd.ExcelFile(xl_path, engine="xlrd")
    except Exception as e:
        return {"counts": counts, "persons": [], "sheets": [], "parse_log": [f"파일 열기 실패: {e}"]}

    for sname in xl.sheet_names:
        try:
            df = xl.parse(sname, header=None, dtype=str)
        except Exception as e:
            parse_log.append(f"{sname}: 파싱실패 {e}")
            continue

        rows = df.values.tolist()
        if not rows:
            continue

        processed_sheets.append(sname)
        current_section: Optional[str] = None  # 현재 섹션 (항목명)

        for ri, raw_row in enumerate(rows):
            cells = [_nc(c) for c in raw_row]
            cells_ns = [c.replace(" ", "") for c in cells]  # 공백제거 버전

            # ── 카운트 추출: "항목명  N" 패턴 ────────────────────────────
            for ci, cv in enumerate(cells_ns):
                if cv in BILLING_ITEM_ALIASES:
                    canon = BILLING_ITEM_ALIASES[cv]
                    # 오른쪽 및 왼쪽 인접 셀에서 숫자 찾기
                    for dc in [1, 2, -1, 3]:
                        nci = ci + dc
                        if 0 <= nci < len(cells):
                            n = _try_int(cells[nci])
                            if n is not None and 0 <= n <= 9999:
                                if n > counts.get(canon, 0):
                                    counts[canon] = n
                                    parse_log.append(f"{sname}[{ri},{ci}] {canon}={n}")
                                break
                    # 섹션 헤더로 지정 (처리대기 항목인 경우)
                    if canon in BILLING_WORK_TYPES:
                        current_section = canon

            # ── 개인행 추출: 차량번호 + 성명이 있는 행 ──────────────────
            veh_found = None
            name_found = None
            region_found = None

            for ci, cv in enumerate(cells):
                cv_ns = cv.replace(" ", "")
                if not veh_found and _is_vehicle_no(cv):
                    veh_found = cv
                if not name_found and _is_name(cv_ns) and 2 <= len(cv_ns) <= 5:
                    # 합계/섹션명 등 제외
                    if cv_ns not in BILLING_ITEM_ALIASES and cv_ns not in {"합계","소계","총계","계","성명","이름"}:
                        name_found = cv
                if not region_found and len(cv_ns) <= 5 and re.match(r'^[가-힣]+[시군]$', cv_ns):
                    region_found = cv

            if veh_found and name_found and current_section:
                raw_dict = {str(i): _nc(v) for i, v in enumerate(raw_row) if _nc(v)}

                # 날짜 컬럼 추출 (신형식: 지역/차량/성명/주민/인가일자/가입일자/자격증명발급일자/발급번호/입금액/항목/비고)
                def _pick_date(idx_list):
                    for idx in idx_list:
                        if idx < len(cells):
                            v = cells[idx].strip()
                            if v and v not in ("x", "X", "O", "o", "-", ""):
                                return v
                    return ""

                # 컬럼 위치 자동 탐지
                permit_date_raw = ""   # 인가일자
                join_date_raw = ""     # 가입일자
                note_raw = ""          # 비고

                # 헤더 기반 위치 탐지 시도
                # 신형식 (6월, 5월, 4월, 3월부과(1)): 인가일자=4, 가입일자=5, 비고=10 또는 11
                # 구형식 (번호/지역/차량/성명/주소/가입일자/자격증명발급일자): 가입일자=5
                for ci2, cv2 in enumerate(cells):
                    cv2s = cv2.replace(" ", "")
                    if ci2 > 0 and len(cv2s) >= 2:
                        # 날짜 패턴 감지: YY.MM.DD 또는 YYYY-MM-DD
                        import re as _re2
                        if _re2.match(r"^\d{2,4}[\.\-/]\s*\d{1,2}[\.\-/]\s*\d{1,2}", cv2s):
                            # 처음 발견된 날짜들을 순서대로 인가일자/가입일자로 배정
                            if not permit_date_raw:
                                permit_date_raw = cv2.strip()
                            elif not join_date_raw:
                                join_date_raw = cv2.strip()

                # 비고: 마지막에서 2번째 또는 마지막 비어있지 않은 셀
                non_empty_cells = [(i, c) for i, c in enumerate(cells) if c.strip()]
                if non_empty_cells:
                    # 항목 컬럼 이후의 마지막 셀을 비고로
                    last_idx, last_val = non_empty_cells[-1]
                    if last_val not in ("x","X","O","o") and not _is_vehicle_no(last_val) and not _is_name(last_val.replace(" ","")):
                        note_raw = last_val.strip()

                # 처리구분별 기준일자 결정
                MGMT_TYPES = {"택배신규", "자격증명발급", "신규관리"}
                ASSOC_TYPES = {"협회가입", "협회가입원"}

                if current_section in MGMT_TYPES:
                    request_date_raw = permit_date_raw or join_date_raw  # 관리비: 인가일자 우선
                elif current_section in ASSOC_TYPES:
                    request_date_raw = join_date_raw or permit_date_raw  # 협회비: 가입일자 우선
                else:
                    # 폐지/양도/관리비폐지 등: 비고란 날짜 우선
                    import re as _re3
                    date_in_note = ""
                    if note_raw:
                        m = _re3.search(r"(\d{2,4}[\.\-/]\s*\d{1,2}[\.\-/]\s*\d{1,2})", note_raw)
                        if m:
                            date_in_note = m.group(1)
                    request_date_raw = date_in_note or permit_date_raw or join_date_raw

                persons.append({
                    "process_type": current_section,
                    "account": BILLING_ACCOUNT_MAP.get(current_section, ""),
                    "name": name_found.strip(),
                    "vehicle_no": veh_found.strip(),
                    "region": (region_found or "").strip(),
                    "source_sheet": sname,
                    "source_row": ri,
                    "raw_data": raw_dict,
                    "to_status": BILLING_TO_STATUS_MAP.get(current_section, ""),
                    "from_status": "정상",
                    "permit_date_raw": permit_date_raw,
                    "join_date_raw": join_date_raw,
                    "note_raw": note_raw,
                    "request_date_raw": request_date_raw,
                })
                parse_log.append(f"{sname}[{ri}] 개인행: {current_section} {name_found} {veh_found} 기준일={request_date_raw}")

            # 빈 행이면 섹션 리셋 (다음 섹션으로 넘어감)
            non_empty = [c for c in cells if c.strip()]
            if len(non_empty) == 0:
                current_section = None

    return {
        "counts": counts,
        "persons": persons,
        "sheets": processed_sheets,
        "parse_log": parse_log,
    }


# FINAL MISUGEUM STRICT SUM PATCH
def is_sum_row(name="", vehicle_no="", *args, **kwargs):
    sums = {"합계", "총계", "소계", "계", "합산", "인원수", "입금금액"}
    n = "" if name is None else str(name).strip().replace(" ", "")
    v = "" if vehicle_no is None else str(vehicle_no).strip().replace(" ", "")
    return (n in sums and (not v or v in sums)) or (v in sums and (not n or n in sums))
# END FINAL MISUGEUM STRICT SUM PATCH
