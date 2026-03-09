from fastapi import APIRouter, HTTPException, Depends
from models.schemas import SignUpRequest, LoginRequest
from database.database import SessionLocal, User
import bcrypt
from email_validator import validate_email, EmailNotValidError
import re

router = APIRouter()

def verify_password(plain_password: str, hashed_password: str) -> bool:
    try:
        # bcrypt requires bytes; also truncating password at 72 bytes natively to avoid ValueError
        pwd_bytes = plain_password.encode('utf-8')[:72]
        hash_bytes = hashed_password.encode('utf-8')
        return bcrypt.checkpw(pwd_bytes, hash_bytes)
    except Exception:
        return False

def get_password_hash(password: str) -> str:
    # bcrypt requires bytes; restricting password to 72 bytes natively
    pwd_bytes = password.encode('utf-8')[:72]
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(pwd_bytes, salt).decode('utf-8')

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def is_valid_phone(phone: str):
    # Basic validation for international or 10-digit formats
    pattern = re.compile(r"^\+?[0-9]{10,15}$")
    return bool(pattern.match(phone))

@router.post("/signup")
def signup(request: SignUpRequest, db = Depends(get_db)):
    # 1. Validate Email
    try:
        valid = validate_email(request.email)
        email = valid.normalized
    except EmailNotValidError as e:
        raise HTTPException(status_code=400, detail=str(e))

    # 2. Validate Phone
    if not is_valid_phone(request.phone):
        raise HTTPException(status_code=400, detail="Invalid phone number format. Must be 10-15 digits.")

    # 3. Check if user exists
    existing_user = db.query(User).filter(User.email == email).first()
    if existing_user:
        raise HTTPException(status_code=400, detail="Email already registered. Please sign in.")

    # 4. Hash password and save
    hashed_pw = get_password_hash(request.password)
    
    new_user = User(
        name=request.name.strip(),
        email=email,
        phone=request.phone.strip(),
        hashed_password=hashed_pw
    )
    
    try:
        db.add(new_user)
        db.commit()
        db.refresh(new_user)
        return {
            "success": True, 
            "message": "User registered successfully", 
            "user": {"name": new_user.name, "email": new_user.email}
        }
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"Database error: {str(e)}")

@router.post("/login")
def login(request: LoginRequest, db = Depends(get_db)):
    # 1. Validate Email pattern softly
    try:
        valid = validate_email(request.email)
        email = valid.normalized
    except EmailNotValidError as e:
        raise HTTPException(status_code=400, detail="Invalid email format requested.")
    
    # 2. Find user
    user = db.query(User).filter(User.email == email).first()
    if not user:
        raise HTTPException(status_code=401, detail="Invalid email or password.")
    
    # 3. Verify password
    if not verify_password(request.password, user.hashed_password):
        raise HTTPException(status_code=401, detail="Invalid email or password.")
    
    return {
        "success": True, 
        "message": "Logged in successfully",
        "user": {
            "name": user.name,
            "email": user.email
        }
    }
