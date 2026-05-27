"""
DB 스키마 v4
핵심 테이블:
  members               회원 마스터
  monthly_ledgers       월별 원장 (가로→세로 변환)
  member_status_events  상태 이력
  work_queue            업무처리현황 반영대기함
  bank_transactions     통장 입금내역
  license_records       전체면허자현황
  collection_targets    문자대상 스냅샷
  current_snapshot      대시보드/미수금 캐시 스냅샷
  raw_import_rows       원본 엑셀 보존
  upload_batches        업로드 이력
  users / audit_logs
"""
import json
from sqlalchemy import (Column, Integer, BigInteger, String, DateTime, Numeric, Text,
                        ForeignKey, Boolean, SmallInteger, Index)
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
from database import Base


class User(Base):
    __tablename__ = "users"
    id            = Column(Integer, primary_key=True)
    username      = Column(String(80), unique=True, nullable=False)
    name          = Column(String(80), default="")
    password_hash = Column(String(255), nullable=False)
    role          = Column(String(30), default="admin")
    created_at    = Column(DateTime(timezone=True), server_default=func.now())


class UploadBatch(Base):
    __tablename__ = "upload_batches"
    id            = Column(Integer, primary_key=True)
    file_name     = Column(String(300))
    data_type     = Column(String(50))   # legacy/license/billing/bank/bank_auto
    file_year     = Column(Integer)
    total_rows    = Column(Integer, default=0)
    saved_rows    = Column(Integer, default=0)
    warn_rows     = Column(Integer, default=0)
    created_by    = Column(Integer, ForeignKey("users.id"), nullable=True)
    created_at    = Column(DateTime(timezone=True), server_default=func.now())


class RawImportRow(Base):
    """원본 엑셀 행 전체 보존 — 삭제 금지"""
    __tablename__ = "raw_import_rows"
    id            = Column(Integer, primary_key=True)
    batch_id      = Column(Integer, ForeignKey("upload_batches.id"), index=True)
    source_file   = Column(String(300))
    source_sheet  = Column(String(200))
    source_row    = Column(Integer)
    sheet_type    = Column(String(50))
    raw_data      = Column(Text)        # JSON list
    created_at    = Column(DateTime(timezone=True), server_default=func.now())


class Member(Base):
    """회원 마스터 — 차량번호+성명 기준 1명 1레코드"""
    __tablename__ = "members"
    id                  = Column(Integer, primary_key=True)
    # 식별
    vehicle_no          = Column(String(100))
    vehicle_key         = Column(String(120), index=True)  # 정규화
    name                = Column(String(100), index=True)
    name_key            = Column(String(100), index=True)  # 공백제거
    region              = Column(String(80),  index=True)
    account             = Column(String(20),  index=True)  # 협/관
    # 연락처
    mobile              = Column(String(100))
    phone               = Column(String(100))
    address             = Column(Text)
    official_address    = Column(Text)
    # 부과기준
    join_date_raw       = Column(String(100))
    permit_date_raw     = Column(String(100))
    cert_issue_date_raw = Column(String(100))
    cert_no             = Column(String(100))
    birth_year          = Column(Integer)
    age70               = Column(Boolean, default=False)
    # 상태 (우선순위 계층 반영)
    status              = Column(String(50), default="정상", index=True)
    # 정상/폐업/양도/이관/탈퇴/사망/말소/관리비폐지/확인필요
    status_source       = Column(String(50), default="auto")
    # auto/manual/work_queue/column/note
    is_auto_transfer    = Column(Boolean, default=False)
    note                = Column(Text)
    # 미수금 (스냅샷 값)
    excel_arrears       = Column(BigInteger, default=0)   # 엑셀 기재 최신월 미수금
    calc_arrears        = Column(BigInteger, default=0)   # 재계산값
    arrears_diff        = Column(BigInteger, default=0)   # 차이
    is_overpay          = Column(Boolean, default=False)  # 초과납부/선납
    arrears_verified    = Column(Boolean, default=True)
    verify_reason       = Column(Text)                 # 검증필요 사유 (| 구분)
    first_unpaid_month  = Column(String(10))
    last_paid_date      = Column(String(100))
    # 전체자 대조
    match_license_id    = Column(Integer, ForeignKey("license_records.id"), nullable=True)
    match_status        = Column(String(80), default="전체자미확인")
    match_fail_reason   = Column(Text)
    user_confirmed_match= Column(Boolean, default=False)  # 사용자 확인완료
    # 소스 추적
    source_file         = Column(String(300))
    source_sheet        = Column(String(200))
    source_row          = Column(Integer)
    source_batch_id     = Column(Integer, ForeignKey("upload_batches.id"), nullable=True)
    created_at          = Column(DateTime(timezone=True), server_default=func.now())
    updated_at          = Column(DateTime(timezone=True), onupdate=func.now())

    license_rec = relationship("LicenseRecord", foreign_keys=[match_license_id])

__table_args_member__ = (
    Index("ix_members_vkey_nkey", "vehicle_key", "name_key"),
    Index("ix_members_region_account", "region", "account"),
)


class MonthlyLedger(Base):
    """월별 원장 (가로→세로): 회원×연월 1행"""
    __tablename__ = "monthly_ledgers"
    id              = Column(Integer, primary_key=True)
    member_id       = Column(Integer, ForeignKey("members.id"), index=True)
    batch_id        = Column(Integer, ForeignKey("upload_batches.id"), index=True)
    source_file     = Column(String(300))
    source_sheet    = Column(String(200))
    source_row      = Column(Integer)
    raw_vehicle_no  = Column(String(100))
    raw_name        = Column(String(100))
    raw_region      = Column(String(80))
    raw_account     = Column(String(20))
    raw_note        = Column(Text)
    year            = Column(Integer, index=True)
    month           = Column(SmallInteger, index=True)
    carry_over      = Column(BigInteger, default=0)
    charge_amount   = Column(BigInteger, default=0)
    paid_amount     = Column(BigInteger, default=0)
    arrears_amount  = Column(BigInteger, default=0)   # 엑셀 기재값
    paid_date       = Column(String(100))
    calc_arrears    = Column(BigInteger, default=0)   # 재계산값
    verified        = Column(Boolean, default=True)
    is_duplicate    = Column(Boolean, default=False)
    created_at      = Column(DateTime(timezone=True), server_default=func.now())
    member          = relationship("Member")

Index("ix_ledger_mid_ym", MonthlyLedger.member_id, MonthlyLedger.year, MonthlyLedger.month)
Index("ix_ledger_ym_mid", MonthlyLedger.year, MonthlyLedger.month, MonthlyLedger.member_id)


class MemberStatusEvent(Base):
    """상태 이력 (상태정리 시트 파싱 결과)"""
    __tablename__ = "member_status_events"
    id              = Column(Integer, primary_key=True)
    member_id       = Column(Integer, ForeignKey("members.id"), index=True)
    batch_id        = Column(Integer, ForeignKey("upload_batches.id"), nullable=True)
    source_file     = Column(String(300))
    source_sheet    = Column(String(200))
    source_row      = Column(Integer)
    raw_vehicle_no  = Column(String(100))
    raw_name        = Column(String(100))
    raw_region      = Column(String(80))
    raw_account     = Column(String(20))
    event_type      = Column(String(50), index=True)
    event_date_raw  = Column(String(100))
    reason_raw      = Column(Text)
    arrears_at_event= Column(Integer, default=0)
    event_month     = Column(String(10))
    created_at      = Column(DateTime(timezone=True), server_default=func.now())
    member          = relationship("Member")

Index("ix_mse_mid_type", MemberStatusEvent.member_id, MemberStatusEvent.event_type)


class WorkQueue(Base):
    """업무처리현황 — 반영대기함"""
    __tablename__ = "work_queue"
    id              = Column(Integer, primary_key=True)
    member_id       = Column(Integer, ForeignKey("members.id"), index=True)
    process_type    = Column(String(50), index=True)
    # 폐업/양도/이관/탈퇴/사망/말소/현역복구/검증해제
    status          = Column(String(50), default="반영대기", index=True)
    # 반영대기/반영완료/취소
    source_screen   = Column(String(50))   # 미수금명단/검증필요/전체자대조/수동
    reason          = Column(Text)
    event_date_raw  = Column(String(100))
    arrears_at_submit= Column(BigInteger, default=0)
    submitted_by    = Column(Integer, ForeignKey("users.id"), nullable=True)
    submitted_at    = Column(DateTime(timezone=True), server_default=func.now())
    reflected_by    = Column(Integer, nullable=True)
    reflected_at    = Column(DateTime(timezone=True), nullable=True)
    note            = Column(Text)
    member          = relationship("Member")


class LicenseRecord(Base):
    """전체면허자현황 원본"""
    __tablename__ = "license_records"
    id                  = Column(Integer, primary_key=True)
    batch_id            = Column(Integer, ForeignKey("upload_batches.id"), nullable=True)
    source_file         = Column(String(300))
    source_sheet        = Column(String(200))
    region              = Column(String(80))
    name                = Column(String(100), index=True)
    name_key            = Column(String(100), index=True)
    vehicle_no          = Column(String(100), index=True)
    vehicle_key         = Column(String(120), index=True)
    resident_no         = Column(String(100))
    mobile              = Column(String(100))
    phone               = Column(String(100))
    address             = Column(Text)
    official_address    = Column(Text)
    join_date_raw       = Column(String(100))
    permit_date_raw     = Column(String(100))
    cert_issue_date_raw = Column(String(100))
    cert_no             = Column(String(100))
    note                = Column(Text)
    created_at          = Column(DateTime(timezone=True), server_default=func.now())

Index("ix_lic_vkey", LicenseRecord.vehicle_key)


class BillingPerson(Base):
    """부과대수 파일 — 사람별 처리 대기 항목"""
    __tablename__ = "billing_persons"
    id                  = Column(Integer, primary_key=True)
    batch_id            = Column(Integer, ForeignKey("upload_batches.id"), index=True)
    billing_report_id   = Column(Integer, ForeignKey("billing_reports.id"), nullable=True, index=True)
    source_file         = Column(String(300))
    source_sheet        = Column(String(200))
    source_row          = Column(Integer)
    raw_data            = Column(Text)             # JSON — 원본 행
    year                = Column(Integer, index=True)
    month               = Column(SmallInteger, index=True)
    process_type        = Column(String(50), index=True)
    # 처리구분: 관리비폐지/70세/협회가입/탈퇴/택배신규/양도/타도/폐지
    account             = Column(String(20))       # 협/관/택배
    region              = Column(String(80))
    name                = Column(String(100))
    vehicle_no          = Column(String(100))
    resident_no         = Column(String(100))
    address             = Column(Text)
    permit_date_raw     = Column(String(100))
    join_date_raw       = Column(String(100))
    cert_issue_date_raw = Column(String(100))
    cert_no             = Column(String(100))
    charge_amount       = Column(BigInteger, default=0)
    charge_start_month  = Column(String(10))       # 부과시작월 YYYY-MM
    charge_end_month    = Column(String(10))       # 부과종료월 YYYY-MM
    from_status         = Column(String(50))       # 기존상태
    to_status           = Column(String(50))       # 변경상태
    note                = Column(Text)
    reflect_status      = Column(String(50), default="처리대기", index=True)
    # 처리대기/반영완료/보류/제외
    reflected_member_id = Column(Integer, ForeignKey("members.id"), nullable=True)
    reflected_at        = Column(DateTime(timezone=True), nullable=True)
    reflected_by        = Column(Integer, nullable=True)
    created_at          = Column(DateTime(timezone=True), server_default=func.now())


class BankTransaction(Base):
    """통장 입금내역"""
    __tablename__ = "bank_transactions"
    id                    = Column(Integer, primary_key=True)
    batch_id              = Column(Integer, ForeignKey("upload_batches.id"), nullable=True)
    txn_date              = Column(String(100), index=True)
    deposit_amount        = Column(BigInteger, default=0)
    balance_amount        = Column(BigInteger, default=0)
    method                = Column(String(120))
    memo                  = Column(String(300), index=True)
    source_type           = Column(String(50), default="paste")  # paste/excel/auto
    is_auto_account       = Column(Boolean, default=False)
    match_status          = Column(String(80), default="미매칭", index=True)
    match_score           = Column(Integer, default=0)
    match_reason          = Column(Text)
    match_candidates_json = Column(Text)
    matched_member_id     = Column(Integer, ForeignKey("members.id"), nullable=True)
    applied               = Column(Boolean, default=False, index=True)
    applied_amount        = Column(BigInteger, default=0)
    overpay_amount        = Column(BigInteger, default=0)
    applied_at            = Column(DateTime(timezone=True), nullable=True)
    note                  = Column(Text)
    raw_text              = Column(Text)
    created_at            = Column(DateTime(timezone=True), server_default=func.now())
    matched_member        = relationship("Member", foreign_keys=[matched_member_id])

Index("ix_btx_applied_mid", BankTransaction.applied, BankTransaction.matched_member_id)


class CollectionTarget(Base):
    """문자대상 스냅샷"""
    __tablename__ = "collection_targets"
    id              = Column(Integer, primary_key=True)
    member_id       = Column(Integer, ForeignKey("members.id"), index=True)
    generated_at    = Column(DateTime(timezone=True), server_default=func.now())
    generated_by    = Column(Integer, nullable=True)
    base_year       = Column(Integer)
    base_month      = Column(SmallInteger)
    arrears         = Column(BigInteger, default=0)
    category        = Column(String(50))  # 일반/자동이체/주기납부/초과납부
    excluded        = Column(Boolean, default=False)
    exclude_reason  = Column(String(200))
    mobile_clean    = Column(String(50))   # 정제된 번호
    address_clean   = Column(Text)         # 정제된 주소
    sent            = Column(Boolean, default=False)
    member          = relationship("Member")


class Snapshot(Base):
    """성능 캐시 스냅샷"""
    __tablename__ = "snapshots"
    id          = Column(Integer, primary_key=True)
    snap_key    = Column(String(100), unique=True, index=True)
    snap_value  = Column(Text)   # JSON
    updated_at  = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())


class AuditLog(Base):
    __tablename__ = "audit_logs"
    id         = Column(Integer, primary_key=True)
    user_id    = Column(Integer, nullable=True)
    action     = Column(String(120))
    detail     = Column(Text)
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class AppSetting(Base):
    __tablename__ = "app_settings"
    id    = Column(Integer, primary_key=True)
    key   = Column(String(100), unique=True)
    value = Column(Text)

class BillingReport(Base):
    """부과대수 관리 — 월례보고 기준 항목 (수동입력 + 파일업로드)"""
    __tablename__ = "billing_reports"
    id              = Column(Integer, primary_key=True)
    year            = Column(Integer, index=True)
    month           = Column(SmallInteger, index=True)
    # 수동 입력 항목
    cnt_join        = Column(Integer, default=0)   # 협회가입
    cnt_transfer    = Column(Integer, default=0)   # 양도
    cnt_cross       = Column(Integer, default=0)   # 타도(이관)
    cnt_close       = Column(Integer, default=0)   # 폐지(폐업)
    cnt_quit        = Column(Integer, default=0)   # 탈퇴
    cnt_delivery_new= Column(Integer, default=0)   # 택배신규
    cnt_mgmt_close  = Column(Integer, default=0)   # 관리비폐지
    cnt_age70       = Column(Integer, default=0)   # 70세
    # 자동 계산 (참고값)
    cnt_base        = Column(Integer, default=0)   # 협회기본대수
    cnt_total       = Column(Integer, default=0)   # 총부과대수
    cnt_delivery    = Column(Integer, default=0)   # 해당월 택배관리
    memo            = Column(Text)
    # 파일 업로드 추적
    source_file     = Column(String(300))          # 원본 파일명
    raw_data        = Column(Text)                 # JSON — 파서 결과 전체
    upload_type     = Column(String(20), default="manual")  # manual / file
    created_by      = Column(Integer, nullable=True)
    created_at      = Column(DateTime(timezone=True), server_default=func.now())
    updated_at      = Column(DateTime(timezone=True), onupdate=func.now())


class IncomeLedgerDetail(Base):
    """잡수입/가수금 구조화 상세"""
    __tablename__ = "income_ledger_details"
    id                  = Column(Integer, primary_key=True)
    bank_transaction_id = Column(Integer, ForeignKey("bank_transactions.id"), nullable=True, index=True)
    income_type         = Column(String(20), index=True)   # 잡수입 / 가수금
    work_type           = Column(String(50), index=True)   # 자격증명발급 / 대폐차 / 예금이자 / 가입비 / 특별회비 / 기타
    pending_target      = Column(String(20), default="없음", index=True)  # 없음 / 협회비 / 관리비
    related_vehicle_no  = Column(String(100))
    related_name        = Column(String(100))
    note                = Column(Text)
    next_billing_date   = Column(String(20))
    created_at          = Column(DateTime(timezone=True), server_default=func.now())
    updated_at          = Column(DateTime(timezone=True), onupdate=func.now())
    bank_transaction    = relationship("BankTransaction", foreign_keys=[bank_transaction_id])
