# Oracle Cloud Free + Cloudflare Tunnel 배포 가이드

## 1. Oracle Cloud ARM VM 생성 (무료 영구)

Oracle Cloud 계정 → Compute → Instances → Create Instance
- Shape: VM.Standard.A1.Flex (Ampere, **무료**)
- OCPU: 4, RAM: 24GB (무료 한도 내 최대)
- OS: Ubuntu 22.04 aarch64
- SSH 키 등록

## 2. VM 기본 세팅

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y python3.11 python3.11-venv python3-pip \
                   git poppler-utils
sudo mkdir -p /var/log/pi-review
sudo chown ubuntu:ubuntu /var/log/pi-review
```

## 3. 앱 배포

```bash
cd /opt
sudo git clone https://github.com/NOROOVIRUZ/pi-review-tool.git
sudo chown -R ubuntu:ubuntu pi-review-tool
cd pi-review-tool

python3.11 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## 4. 환경 변수 설정

```bash
sudo nano /etc/systemd/system/pi-review.service
# FLASK_SECRET 값을 안전한 랜덤 문자열로 변경
```

```bash
sudo cp /opt/pi-review-tool/deploy/pi-review.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now pi-review
sudo systemctl status pi-review
```

## 5. Cloudflare Tunnel 설정

```bash
# cloudflared 설치 (ARM64)
curl -L https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-arm64 \
     -o /usr/local/bin/cloudflared
chmod +x /usr/local/bin/cloudflared

# 로그인 + 터널 생성
cloudflared tunnel login
cloudflared tunnel create pi-review
cloudflared tunnel route dns pi-review pi-review.yourdomain.com

# 설정 파일 배치 (tunnel ID와 도메인 수정 후)
sudo cp /opt/pi-review-tool/deploy/cloudflared-tunnel.yml /etc/cloudflared/config.yml
sudo cloudflared service install
sudo systemctl start cloudflared
```

## 6. Oracle VCN 방화벽 (5000 포트 외부 차단)

Oracle Cloud 콘솔에서 5000 포트 인바운드 규칙 추가 **하지 말 것**.
cloudflared가 내부적으로 127.0.0.1:5000 에 연결하므로 외부 노출 불필요.

OS 방화벽으로 이중 차단:
```bash
sudo iptables -A INPUT -p tcp --dport 5000 ! -s 127.0.0.1 -j DROP
```

## 7. 배포 완료 확인

```bash
# 서비스 상태
sudo systemctl status pi-review
sudo systemctl status cloudflared

# 로그 실시간 보기
journalctl -u pi-review -f

# 로컬 헬스체크
curl http://localhost:5000/health
```

브라우저에서 `https://pi-review.yourdomain.com` 접속 확인.

## 업데이트 방법

```bash
cd /opt/pi-review-tool
git pull
sudo systemctl restart pi-review
```
