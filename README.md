# 보험 수수료 조회 사이트

엑셀 수수료 파일을 업로드하고 회사별, 상품별, 1차년, 2차년, 3차년, 총수수료를 조회하는 내부용 웹사이트입니다.

## 로그인 계정

- 조회 계정: `company / 1234`
- 관리자 계정: `admin / 1234`

조회 계정은 표 조회만 가능하고, 관리자 계정은 엑셀 업로드와 수정/삭제가 가능합니다.

## 로컬 실행

```powershell
pip install -r requirements.txt
python app.py
```

로컬 주소:

```text
http://127.0.0.1:8766/
```

## 배포

Render 같은 Python Web Service 플랫폼에서 다음 설정으로 배포할 수 있습니다.

- Build command: `pip install -r requirements.txt`
- Start command: `python app.py`

서버는 배포 플랫폼의 `PORT` 환경변수를 자동으로 사용합니다.
