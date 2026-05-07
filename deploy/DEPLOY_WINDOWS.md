# PI 검토 툴 — Windows 배포 가이드

## 구성 개요

```
회사 Windows PC (항상 켜짐)
  └── Flask 서버 (localhost:5001)
  └── cloudflared  →  https://pi-review.yourdomain.com
                              ↑
              팀원들이 브라우저로 접속 (설치 불필요)
```

---

## 1. 사전 요구사항

- Windows 10 / 11
- **Python 3.11 이상** — https://www.python.org/downloads/
  - 설치 시 **"Add Python to PATH"** 반드시 체크
- Git (선택) — https://git-scm.com/

---

## 2. 프로젝트 다운로드

### Git 사용 시
```bat
git clone https://github.com/NOROOVIRUZ/pi-review-tool.git
cd pi-review-tool
```

### 수동 다운로드 시
GitHub에서 ZIP 다운로드 후 압축 해제

---

## 3. 설치 (최초 1회)

프로젝트 폴더에서 실행:
```
deploy\setup_windows.bat 더블클릭
```

설치 내용:
- Python 가상환경 (.venv)
- 필요 패키지 (Flask, EasyOCR, reportlab 등)
- poppler 자동 다운로드 (PDF 처리용, ~30MB)

**소요시간:** 5~15분 (인터넷 속도 따라 다름)

---

## 4. 서버 실행

```
deploy\start_windows.bat 더블클릭
```

- 브라우저가 자동으로 열립니다
- 창을 닫으면 서버 종료
- 팀원들은 **같은 사무실 LAN이면** `http://[이 PC의 IP]:5001` 로 바로 접속 가능

### PC IP 확인 방법
```bat
ipconfig
:: IPv4 주소 항목 확인 (예: 192.168.1.105)
```

---

## 5. Windows 시작 시 자동 실행 (선택)

PC를 켤 때마다 자동으로 서버가 백그라운드 실행되게 하려면:

```
deploy\install_autostart_windows.bat 더블클릭
```

---

## 6. 외부망에서도 접속 — Cloudflare Tunnel (무료, 선택)

같은 사무실 LAN이 아닌 재택/외부망에서도 쓰려면 Cloudflare Tunnel 설정.

### 6-1. cloudflared 설치 (Windows)

```powershell
# PowerShell (관리자 권한)
winget install Cloudflare.cloudflared
```

또는 수동 다운로드:
https://github.com/cloudflare/cloudflared/releases/latest 에서
`cloudflared-windows-amd64.exe` 다운로드 → `C:\cloudflared\cloudflared.exe`로 저장

### 6-2. 터널 생성

```bat
cloudflared tunnel login
cloudflared tunnel create pi-review
cloudflared tunnel route dns pi-review pi-review.yourdomain.com
```

### 6-3. 설정 파일

`%USERPROFILE%\.cloudflared\config.yml` 생성:

```yaml
tunnel: <터널-ID>
credentials-file: C:\Users\<사용자명>\.cloudflared\<터널-ID>.json

ingress:
  - hostname: pi-review.yourdomain.com
    service: http://localhost:5001
  - service: http_status:404
```

### 6-4. Windows 서비스 등록 (자동 시작)

```bat
:: 관리자 권한 명령 프롬프트
cloudflared service install
sc start cloudflared
```

---

## 7. 팀원 공유

| 상황 | 접속 주소 |
|------|-----------|
| 같은 사무실 LAN | `http://192.168.x.x:5001` |
| 외부망 / 재택 | `https://pi-review.yourdomain.com` |

→ 브라우저에서 접속하면 끝. 팀원들은 아무것도 설치 안 해도 됩니다.

---

## 8. 업데이트

```bat
git pull
:: 서버 재시작 (start_windows.bat 다시 실행)
```

---

## 문제 해결

| 증상 | 원인 | 해결 |
|------|------|------|
| poppler 오류 | PATH 미등록 | start_windows.bat 통해 실행 확인 |
| 포트 5001 이미 사용 중 | 다른 프로그램 충돌 | config.py에서 PORT 변경 |
| OCR 느림 | 최초 모델 다운로드 | 첫 실행 후 자동 캐시됨 |
| EasyOCR CUDA 경고 | GPU 없음 (정상) | CPU 모드로 자동 전환, 무시 가능 |
