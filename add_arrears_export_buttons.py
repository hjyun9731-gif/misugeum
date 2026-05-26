from pathlib import Path
import re

templates = Path("templates")

candidates = []
for p in templates.rglob("*.html"):
    text = p.read_text(encoding="utf-8", errors="ignore")
    if (
        "/arrears" in text
        or "미수" in text
        or "입금" in text
        or "부과대수" in text
        or "arrears" in p.name.lower()
    ):
        candidates.append(p)

print("후보 템플릿:")
for c in candidates:
    print(" -", c)

if not candidates:
    # 그래도 없으면 가장 가능성 높은 파일명 생성/대상 찾기
    target = templates / "arrears.html"
else:
    # arrears 이름 들어간 파일 우선
    target = None
    for c in candidates:
        if "arrears" in c.name.lower():
            target = c
            break
    if target is None:
        target = candidates[0]

print("선택된 파일:", target)

if not target.exists():
    raise SystemExit("ERROR: arrears 화면 템플릿 파일을 찾지 못했습니다.")

text = target.read_text(encoding="utf-8", errors="ignore")

marker = "report-export-buttons-for-arrears"

block = r'''
<!-- report-export-buttons-for-arrears -->
<div id="report-export-buttons-for-arrears" style="
    margin: 14px 0 18px 0;
    padding: 14px;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    background: #ffffff;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
">
    <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <strong style="font-size:15px;">보고/집계 엑셀 추출</strong>

        <label style="font-size:13px; color:#374151;">
            기준년도
            <input id="reportExportYear" type="number" min="2000" max="2100"
                   style="width:90px; padding:7px 9px; border:1px solid #d1d5db; border-radius:8px;">
        </label>

        <label style="font-size:13px; color:#374151;">
            기준월
            <input id="reportExportMonth" type="number" min="1" max="12"
                   style="width:70px; padding:7px 9px; border:1px solid #d1d5db; border-radius:8px;">
        </label>

        <button type="button" onclick="downloadBillingCountExcel()" style="
            padding: 8px 13px;
            border: 0;
            border-radius: 9px;
            background: #2563eb;
            color: white;
            cursor: pointer;
            font-weight: 600;
        ">N월 부과대수 엑셀</button>

        <button type="button" onclick="downloadDepositExcel()" style="
            padding: 8px 13px;
            border: 0;
            border-radius: 9px;
            background: #16a34a;
            color: white;
            cursor: pointer;
            font-weight: 600;
        ">N월 입금추출 엑셀</button>
    </div>

    <div style="margin-top:8px; font-size:12px; color:#6b7280;">
        선택한 기준월 기준으로 부과대수와 입금내역을 엑셀로 다운로드합니다.
    </div>
</div>

<script>
(function initReportExportMonth() {
    const now = new Date();
    const y = document.getElementById("reportExportYear");
    const m = document.getElementById("reportExportMonth");
    if (y && !y.value) y.value = now.getFullYear();
    if (m && !m.value) m.value = now.getMonth() + 1;
})();

function getReportExportYM() {
    const year = document.getElementById("reportExportYear")?.value || new Date().getFullYear();
    const month = document.getElementById("reportExportMonth")?.value || (new Date().getMonth() + 1);

    if (!year || !month || Number(month) < 1 || Number(month) > 12) {
        alert("기준년도와 기준월을 확인해 주세요.");
        return null;
    }

    return { year, month };
}

function downloadBillingCountExcel() {
    const ym = getReportExportYM();
    if (!ym) return;
    window.location.href = `/api/reports/billing-count/export?year=${ym.year}&month=${ym.month}`;
}

function downloadDepositExcel() {
    const ym = getReportExportYM();
    if (!ym) return;
    window.location.href = `/api/reports/deposit/export?year=${ym.year}&month=${ym.month}`;
}
</script>
<!-- /report-export-buttons-for-arrears -->
'''

if marker in text:
    print("SKIP: 이미 버튼 블록이 들어가 있습니다.")
else:
    # body 바로 뒤에 넣기
    if "<body" in text:
        text = re.sub(r"(<body[^>]*>)", r"\1\n" + block, text, count=1, flags=re.I)
    else:
        # body가 없으면 맨 위에 삽입
        text = block + "\n" + text

    target.write_text(text, encoding="utf-8")
    print("DONE: /arrears 화면에 엑셀 버튼 추가 완료:", target)
