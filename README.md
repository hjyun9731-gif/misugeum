# 강원도개인소형화물협회 미수금관리 v4

FastAPI + PostgreSQL 기반 미수금/회원 통합관리 시스템

## 자동배포 구조

```
코드 수정 → git push origin main → Railway 자동배포
```

## Railway 환경변수 설정 필수

| 변수명 | 설명 |
|--------|------|
| `DATABASE_URL` | PostgreSQL 연결 URL (Railway에서 자동 제공) |
| `SECRET_KEY` | 세션 암호화 키 (임의 문자열) |

## 로컬 실행

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

## 기본 로그인

- ID: `admin`
- PW: `admin1234`

## 주요 기능

- 엑셀 업로드 (미수금명단, 전체자명단)
- 미수금 명단 관리 (음수/초과납부 처리)
- 전체자대조 (면허자현황 대조)
- 통장매칭 (입금내역 반영)
- 부과대수 관리
- 문자대상관리
- 검증필요 관리
