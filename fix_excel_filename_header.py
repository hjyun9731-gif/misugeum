from pathlib import Path

p = Path("report_export_routes.py")
text = p.read_text(encoding="utf-8")

# urllib.parse.quote import 추가
if "from urllib.parse import quote" not in text:
    text = text.replace(
        "import re\n",
        "import re\nfrom urllib.parse import quote\n"
    )

old = '''def _excel_response(wb, filename):
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    quoted = filename.encode("utf-8")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{filename}"
    }
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )
'''

new = '''def _excel_response(wb, filename):
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
'''

if old in text:
    text = text.replace(old, new)
    p.write_text(text, encoding="utf-8")
    print("DONE: 한글 파일명 헤더 인코딩 수정 완료")
else:
    print("WARN: 기존 _excel_response 함수 모양이 달라서 자동 치환 못함")
