#!/usr/bin/env bash
# 行程记录与分析智能体 - 腾讯云一键部署脚本（Ubuntu / TencentOS / CentOS 通用）
# 用法：
#   sudo bash deploy.sh
# 可选环境变量：
#   APP_PORT=7860  部署监听端口
#   REPO_URL=...   代码仓库地址（默认 GitHub，失败自动切 ghproxy 镜像）
set -euo pipefail

APP_PORT="${APP_PORT:-7860}"
REPO_URL="${REPO_URL:-https://github.com/weinihao123/itinerary-agent.git}"
REPO_URL_MIRROR="https://ghproxy.net/https://github.com/weinihao123/itinerary-agent.git"
INSTALL_DIR="/opt/itinerary-agent"
SERVICE_NAME="itinerary-agent"
# 域名 HTTPS（可选）：设置 DOMAIN 与 EMAIL 即启用 Nginx 反代 + Let's Encrypt
DOMAIN="${DOMAIN:-}"
EMAIL="${EMAIL:-}"

[ "$(id -u)" -eq 0 ] || { echo "请使用 root 或 sudo 运行此脚本"; exit 1; }

echo "==> 检测操作系统与包管理器"
if command -v apt-get >/dev/null 2>&1; then
  PKG_MGR="apt"; echo "检测到 Debian/Ubuntu 系"
elif command -v dnf >/dev/null 2>&1; then
  PKG_MGR="dnf"; echo "检测到 Fedora/TencentOS 系(dnf)"
elif command -v yum >/dev/null 2>&1; then
  PKG_MGR="yum"; echo "检测到 CentOS 系(yum)"
else
  echo "未能识别包管理器，退出"; exit 1
fi

echo "==> 安装系统依赖（git / python3 / venv / 防火墙）"
if [ "$PKG_MGR" = "apt" ]; then
  apt-get update -y
  DEBIAN_FRONTEND=noninteractive apt-get install -y git python3 python3-pip python3-venv ufw
else
  "$PKG_MGR" install -y git python3 python3-pip firewalld
fi

echo "==> 准备代码"
if [ -f "$INSTALL_DIR/app.py" ]; then
  echo "检测到 $INSTALL_DIR 已有代码，跳过克隆（如需更新请手动 git pull）"
elif [ -d "$INSTALL_DIR/.git" ]; then
  git -C "$INSTALL_DIR" pull --ff-only
else
  git clone "$REPO_URL" "$INSTALL_DIR" \
    || git clone "$REPO_URL_MIRROR" "$INSTALL_DIR"
fi

echo "==> 创建虚拟环境并安装 Python 依赖"
python3 -m venv "$INSTALL_DIR/venv"
"$INSTALL_DIR/venv/bin/pip" install --upgrade pip
"$INSTALL_DIR/venv/bin/pip" install -r "$INSTALL_DIR/requirements.txt"

echo "==> 写入 systemd 服务单元"
cat > "/etc/systemd/system/${SERVICE_NAME}.service" <<EOF
[Unit]
Description=Itinerary Agent Web Service
After=network.target

[Service]
Type=simple
WorkingDirectory=${INSTALL_DIR}
Environment=HOST=0.0.0.0
Environment=PORT=${APP_PORT}
ExecStart=${INSTALL_DIR}/venv/bin/python ${INSTALL_DIR}/app.py
Restart=always
RestartSec=3
User=root

[Install]
WantedBy=multi-user.target
EOF

echo "==> 开放防火墙端口 ${APP_PORT}（OS 层）"
if command -v ufw >/dev/null 2>&1; then
  ufw allow "${APP_PORT}/tcp" || true
  ufw allow 22/tcp || true
  ufw --force enable || true
elif command -v firewall-cmd >/dev/null 2>&1; then
  systemctl enable --now firewalld || true
  firewall-cmd --permanent --add-port=${APP_PORT}/tcp || true
  firewall-cmd --permanent --add-service=ssh || true
  firewall-cmd --reload || true
fi

# ---------- 可选：域名 + HTTPS（Nginx 反代 + Let's Encrypt）----------
if [ -n "${DOMAIN}" ]; then
  [ -n "${EMAIL}" ] || { echo "启用 HTTPS 必须设置 EMAIL（证书注册邮箱）"; exit 1; }
  echo "==> 安装 Nginx 与 Certbot"
  if [ "$PKG_MGR" = "apt" ]; then
    apt-get install -y nginx certbot python3-certbot-nginx
  else
    "$PKG_MGR" install -y epel-release || true
    "$PKG_MGR" install -y nginx certbot python3-certbot-nginx
  fi

  echo "==> 写入 Nginx 反代配置"
  cat > "/etc/nginx/conf.d/${SERVICE_NAME}.conf" <<EOF
server {
    listen 80;
    server_name ${DOMAIN};
    location / {
        proxy_pass http://127.0.0.1:${APP_PORT};
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
EOF

  echo "==> 防火墙放行 80/443"
  if command -v ufw >/dev/null 2>&1; then
    ufw allow 80/tcp || true
    ufw allow 443/tcp || true
  elif command -v firewall-cmd >/dev/null 2>&1; then
    firewall-cmd --permanent --add-service=http || true
    firewall-cmd --permanent --add-service=https || true
    firewall-cmd --reload || true
  fi

  echo "==> 启动 Nginx 并申请证书"
  systemctl enable --now nginx
  certbot --nginx -d "${DOMAIN}" --non-interactive --agree-tos -m "${EMAIL}" --redirect
  systemctl reload nginx
fi

echo "==> 启动并启用开机自启"
systemctl daemon-reload
systemctl enable --now "${SERVICE_NAME}"
sleep 2
systemctl status "${SERVICE_NAME}" --no-pager || true

PUB_IP=$(curl -s --max-time 5 ifconfig.me || hostname -I | awk '{print $1}')
echo "=============================================="
echo "部署完成"
echo "本地访问: http://127.0.0.1:${APP_PORT}"
if [ -n "${DOMAIN}" ]; then
  echo "公网访问: https://${DOMAIN}  （已自动跳转 HTTPS）"
  echo "提示：请确认域名 DNS 的 A 记录已指向 ${PUB_IP}，且已在腾讯云完成 ICP 备案"
else
  echo "公网访问: http://${PUB_IP}:${APP_PORT}"
  echo "提示：若公网打不开，请到腾讯云控制台【安全组】放行入站 ${APP_PORT}/tcp"
fi
echo "查看日志: journalctl -u ${SERVICE_NAME} -f"
echo "=============================================="
