from passlib.context import CryptContext
from sqlalchemy.orm import Session
from models import User
from fastapi import Request, Depends, HTTPException
from database import get_db

_pwd = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

def hash_pw(p: str) -> str: return _pwd.hash(str(p))
def verify_pw(plain: str, hashed: str) -> bool:
    try: return _pwd.verify(str(plain), hashed)
    except Exception: return False

def ensure_admin(db: Session):
    u = db.query(User).filter(User.username == "admin").first()
    if not u:
        db.add(User(username="admin", name="관리자", password_hash=hash_pw("admin1234"), role="admin"))
        db.commit()
    elif not verify_pw("admin1234", u.password_hash):
        pass  # 이미 변경된 비밀번호는 건드리지 않음

def require_user(request: Request, db: Session = Depends(get_db)):
    uid = request.session.get("user_id")
    if not uid:
        raise HTTPException(status_code=302, headers={"Location": "/login"})
    u = db.query(User).filter(User.id == uid).first()
    if not u:
        raise HTTPException(status_code=302, headers={"Location": "/login"})
    return u
