"""
excel_parser.py – 안전한 엑셀 파싱
- 미수금명단: 5월 미수금 열 최우선
- BadZipFile / 파일오류 안전 처리
- 중복 컬럼명 Series 안전 처리
"""
import re, math, zipfile as _zipfile
from pathlib import Path
from typing import Dict, Iterable
import pandas as pd

BAD_VEHICLE_WORDS = {"번호","1","2","3","4","5","폐지","폐업","탈퇴","양도","합계","계","총계","비고","구분","성명","지역"}
REGIONS = ["춘천시","강릉시","원주시","동해시","삼척시","속초시","태백시","홍천군","횡성군","영월군","평창군","정선군","철원군","화천군","양구군","인제군","고성군","양양군"]
REGION_ALIASES = {
    "춘천":"춘천시","춘천시":"춘천시","강릉":"강릉시","강릉시":"강릉시",
    "원주":"원주시","원주시":"원주시","동해":"동해시","동해시":"동해시",
    "삼척":"삼척시","삼척시":"삼척시","속초":"속초시","속초시":"속초시",
    "태백":"태백시","태백시":"태백시","홍천":"홍천군","홍천군":"홍천군",
    "횡성":"횡성군","횡성군":"횡성군","영월":"영월군","영월군":"영월군",
    "평창":"평창군","평창군":"평창군","정선":"정선군","정선군":"정선군",
    "철원":"철원군","철원군":"철원군","화천":"화천군","화천군":"화천군",
    "양구":"양구군","양구군":"양구군","인제":"인제군","인제군":"인제군",
    "고성":"고성군","고성군":"고성군","양양":"양양군","양양군":"양양군",
}

def normalize_region(value):
    s = norm_text(value).replace(" ","")
    if not s: return ""
    s = re.sub(r"[\(（].*$","",s)
    s = re.sub(r"\d+$","",s)
    s = s.replace("강원특별자치도","").replace("강원도","")
    if s in REGION_ALIASES: return REGION_ALIASES[s]
    for key,region in REGION_ALIASES.items():
        if key and key in s: return region
    for region in REGIONS:
        if region in s: return region
    return norm_text(value)

ALIASES = {
    "region":   ["지역","시군","시·군","시군구","관할","주소지"],
    "name":     ["성명","성 명","이름","대표자","명의자","회원명","운전자","차주","사업자명"],
    "vehicle_no":["차량번호","자동차등록번호","등록번호","차량","차번","자동차번호","차량 번호"],
    "account":  ["계정","구분","항목","계정구분"],
    "mobile":   ["핸드폰","휴대폰","휴대전화","휴대폰번호","핸드폰번호","연락처","휴대폰 번호"],
    "phone":    ["전화번호","전화","일반전화","사무실전화"],
    "address":  ["주소","현주소","주소지","거주지","주민등록주소"],
    "official_address":["공문주소","공문 주소","사무소주소","주사무소","주 사무소","사업장주소","사업장 주소"],
    "join_date":["가입일자","협회가입일자","협회 가입일자","가입일","협회가입"],
    "permit_date":["인가일자","허가일자","인가일","허가일","행정처분일","승인일자"],
    "cert_issue_date":["자격증명발급일자","자격증명 발급일자","자격증명일자","자격증명발급","발급일자"],
    "cert_no":  ["자격증명발급번호","자격증명 발급번호","자격증명번호","발급번호"],
    "resident_no":["주민등록번호","주민번호","생년월일","주민등록","주민"],
    "note":     ["비고","메모","참고","특이사항"],
}

# ── 기본 유틸 ──────────────────────────────────────────────────────────────

def norm_text(x) -> str:
    if x is None: return ""
    try:
        if isinstance(x, float) and math.isnan(x): return ""
    except Exception: pass
    s = str(x).replace("\u3000"," ").strip()
    if s.lower() in {"nan","none","nat"}: return ""
    return re.sub(r"\s+"," ",s)

def clean_key(s) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]","",norm_text(s)).lower()

def clean_name(s) -> str:
    return re.sub(r"\s+","",norm_text(s))

def first_scalar(v):
    if isinstance(v, pd.Series):
        for x in v.tolist():
            sx = norm_text(x)
            if sx: return x
        return ""
    return v

def to_jsonable(v):
    try:
        if pd.isna(v): return ""
    except Exception: pass
    if hasattr(v,"strftime"): return v.strftime("%Y-%m-%d")
    if isinstance(v,(int,float,str,bool)) or v is None:
        if isinstance(v,float) and math.isnan(v): return ""
        return v
    return str(v)

def row_to_raw(row: pd.Series) -> dict:
    return {str(k): to_jsonable(v) for k,v in row.to_dict().items()}

def normalize_vehicle(v) -> str:
    s = norm_text(v)
    s = s.replace("호","")
    s = re.sub(r"\s+","",s)
    s = re.sub(r"^강원","",s)
    return s.upper()

def vehicle_like(v) -> bool:
    s = norm_text(v)
    if not s or s in BAD_VEHICLE_WORDS: return False
    sk = normalize_vehicle(s)
    if sk in BAD_VEHICLE_WORDS: return False
    return bool(re.search(r"\d{2,3}[가-힣]\d{3,4}",sk)
             or re.search(r"[가-힣]\d{3,4}$",sk)
             or re.search(r"\d{4}$",sk))

def parse_amount(v) -> int:
    v = first_scalar(v)
    s = norm_text(v)
    if not s: return 0
    s = re.sub(r"[^0-9\-]","",s)
    if not s or s == "-": return 0
    try: return int(s)
    except Exception: return 0

def _row_col(row, col_name):
    """컬럼명으로 row에서 값 꺼내기. 중복 컬럼 대비 first_scalar 처리."""
    if col_name not in row.index: return None
    val = row[col_name]
    return first_scalar(val)

# ── 금액 추출 (5월 미수금 우선) ────────────────────────────────────────────

def choose_legacy_amount(row, colmap) -> int:
    """
    우선순위:
    1) '5월 미수금' 열
    2) 가장 높은 월의 '미수금' 열
    3) '이월금' 열
    4) 모든 미수금 관련 열 fallback
    """
    cols = list(row.index)

    # 1) 5월 미수금 정확 매칭
    for col in cols:
        c = norm_text(col)
        if "5월" in c and "미수" in clean_key(c) and "부과" not in c and "입금" not in c:
            v = _row_col(row, col)
            return parse_amount(v)

    # 2) 월별 미수금 – 가장 높은 달 우선
    month_candidates = []
    for col in cols:
        c = norm_text(col)
        ck = clean_key(c)
        m = re.search(r"(\d{1,2})월", c)
        if m and "미수" in ck and "입금" not in c and "부과" not in c:
            month = int(m.group(1))
            if 1 <= month <= 12:
                month_candidates.append((month, col))
    if month_candidates:
        month_candidates.sort(key=lambda x: x[0], reverse=True)
        for _, col in month_candidates:
            v = _row_col(row, col)
            r = parse_amount(v)
            if r != 0: return r
        return parse_amount(_row_col(row, month_candidates[0][1]))

    # 3) 이월금
    for col in cols:
        if "이월" in norm_text(col):
            v = _row_col(row, col)
            r = parse_amount(v)
            if r != 0: return r

    # 4) 미수/미납 관련 열
    for col in cols:
        ck = clean_key(norm_text(col))
        if any(k in ck for k in ["미수금","미납금","미수","미납","잔액"]):
            if not any(bad in ck for bad in ["입금","납부","부과","날짜","일자"]):
                v = _row_col(row, col)
                r = parse_amount(v)
                if r != 0: return r

    return 0

# ── 헤더 감지 / 컬럼 매핑 ────────────────────────────────────────────────

def detect_header(raw: pd.DataFrame) -> int:
    best_i, best_score = 0, -1
    max_rows = min(len(raw), 30)
    for i in range(max_rows):
        vals = [norm_text(x) for x in raw.iloc[i].tolist()]
        joined = " ".join(vals)
        score = 0
        for key, aliases in ALIASES.items():
            for a in aliases:
                if a in joined:
                    score += 3 if key in {"name","vehicle_no"} else 1
                    break
        if len([v for v in vals if v]) >= 3: score += 1
        if score > best_score:
            best_score, best_i = score, i
    return best_i

def map_columns(cols: Iterable) -> Dict[str,str]:
    result = {}
    for field, aliases in ALIASES.items():
        for c in cols:
            ck = clean_key(c)
            for a in aliases:
                ak = clean_key(a)
                if ak and (ck == ak or ak in ck or ck in ak):
                    if field not in result:
                        result[field] = c
                    break
            if field in result: break
    return result

# ── 파일 읽기 (안전) ─────────────────────────────────────────────────────

def read_excel_sheets(path: str):
    """
    엑셀 파일 읽기. 오류 발생 시 ValueError 발생.
    """
    p = Path(path)

    if not p.exists():
        raise ValueError("파일이 존재하지 않습니다.")
    if p.stat().st_size == 0:
        raise ValueError("업로드된 파일이 비어 있습니다. 다시 업로드해주세요.")

    ext = p.suffix.lower()
    try:
        if ext in {".xlsx",".xlsm"}:
            excel = pd.ExcelFile(str(path), engine="openpyxl")
        elif ext == ".xls":
            excel = pd.ExcelFile(str(path), engine="xlrd")
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다. (확장자: {ext})")
    except _zipfile.BadZipFile:
        raise ValueError(
            "엑셀 파일 형식이 올바르지 않습니다. "
            "파일을 엑셀에서 열고 [다른 이름으로 저장] → .xlsx 또는 .xlsm으로 다시 저장 후 업로드해주세요."
        )
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"엑셀 파일을 열 수 없습니다: {e}")

    outputs = []
    for sheet in excel.sheet_names:
        try:
            raw = pd.read_excel(excel, sheet_name=sheet, header=None, dtype=object)
            if raw is None or raw.empty: continue
            raw = raw.dropna(how="all")
            if raw.empty: continue
            h = detect_header(raw)
            headers = [norm_text(x) or f"col_{j}" for j,x in enumerate(raw.iloc[h].tolist())]
            df = raw.iloc[h+1:].copy()
            df.columns = headers
            df = df.dropna(how="all")
            outputs.append((sheet, df, h+1, map_columns(headers)))
        except Exception:
            continue

    return outputs

# ── 기타 ──────────────────────────────────────────────────────────────────

def guess_file_type(filename: str, sheet: str = "") -> str:
    text = f"{filename} {sheet}".replace(" ","")
    if any(k in text for k in ["전체면허자","전체자","면허자","현황"]): return "license"
    if any(k in text for k in ["미수금","회비내역"]): return "legacy"
    if any(k in text for k in ["부과대수","부과"]): return "billing_archive"
    return "unknown"

def choose(row, colmap, field) -> str:
    col = colmap.get(field)
    if not col: return ""
    if col not in row.index: return ""
    val = first_scalar(row[col])
    text = norm_text(val)
    if field == "name": return clean_name(text)
    return text

def get_region(row, colmap) -> str:
    r = choose(row, colmap, "region")
    if r: return normalize_region(r)
    raw = " ".join(norm_text(x) for x in row.tolist())
    for region in REGIONS:
        if region in raw: return region
    for key,region in REGION_ALIASES.items():
        if key and key in raw: return region
    return ""

def valid_person_row(row, colmap, mode) -> bool:
    name = choose(row, colmap, "name")
    veh  = choose(row, colmap, "vehicle_no")
    if name in BAD_VEHICLE_WORDS or veh in BAD_VEHICLE_WORDS: return False
    if not name and not vehicle_like(veh): return False
    if mode == "legacy":
        if not name: return False
        # 차량번호 없어도 금액 있으면 유효
        if not vehicle_like(veh):
            amount = choose_legacy_amount(row, colmap)
            if amount <= 0: return False
    if mode == "license":
        if not (name or vehicle_like(veh)): return False
    return True


# FINAL MISUGEUM LENIENT UPLOAD PATCH
_FINAL_SUM_NAMES = {"합계", "총계", "소계", "계", "합산", "인원수", "입금금액"}

def _fx_s(v):
    return "" if v is None else str(v).strip()

def _fx_pick(row, keys):
    if not isinstance(row, dict):
        return ""
    for k in keys:
        if k in row and _fx_s(row.get(k)):
            return _fx_s(row.get(k))
    for rk, rv in row.items():
        nk = _fx_s(rk).replace(" ", "")
        for k in keys:
            if k.replace(" ", "") in nk and _fx_s(rv):
                return _fx_s(rv)
    return ""

def clean_name(v):
    return _fx_s(v)

def normalize_vehicle(v):
    return _fx_s(v).replace(" ", "")

def is_sum_row(name="", vehicle_no="", *args, **kwargs):
    n = _fx_s(name).replace(" ", "")
    v = _fx_s(vehicle_no).replace(" ", "")
    return (n in _FINAL_SUM_NAMES and (not v or v in _FINAL_SUM_NAMES)) or (v in _FINAL_SUM_NAMES and (not n or n in _FINAL_SUM_NAMES))

def valid_person_row(row, *args, **kwargs):
    vals = list(row.values()) if isinstance(row, dict) else (list(row) if isinstance(row, (list, tuple)) else [])
    if not any(_fx_s(v) for v in vals):
        return False
    name = _fx_pick(row, ["성명", "이름", "회원명", "대표자", "차주명", "name"])
    vehicle = _fx_pick(row, ["차량번호", "차량 번호", "차번", "등록번호", "vehicle_no"])
    amount = _fx_pick(row, ["미수금", "미수", "미납", "금액", "잔액", "arrears", "excel_arrears"])
    account = _fx_pick(row, ["계정", "구분", "account"])
    region = _fx_pick(row, ["지역", "시군", "region"])
    if is_sum_row(name, vehicle):
        return False
    return bool(name or vehicle or amount or account or region)
# END FINAL MISUGEUM LENIENT UPLOAD PATCH
