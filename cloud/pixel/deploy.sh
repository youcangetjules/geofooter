#!/usr/bin/env bash
# Set up the AES read-tracking pixel endpoint on an Ubuntu 24.04 VM, either
# fresh or already serving other nginx sites.
#
# Prerequisites: the host name's A record points at this VM, and ports 80/443
# are open to the internet (GCP: "Allow HTTP/HTTPS traffic" on the instance).
#
# On a VM with other sites: packages that are already installed are left alone
# (no nginx upgrade/restart), the default site is only removed on a fresh nginx,
# and the site is enabled as "zz-aes-pixel" so it is included after the others
# and never becomes the catch-all for names it does not serve. Each nginx change
# is checked with nginx -t and applied with a graceful reload.
#
# Usage, on the VM from a copy of cloud/pixel/, with the shared key on stdin
# so it never appears in shell history or the process list:
#   sudo ./deploy.sh track.aliniant.net admin@aliniant.net < key.txt
#
# Safe to re-run: it reinstalls the files, keeps the opens database, and only
# asks certbot for a certificate when there isn't one yet.
set -euo pipefail

HOST=${1:?usage: deploy.sh <host> <letsencrypt-email> < key}
EMAIL=${2:?usage: deploy.sh <host> <letsencrypt-email> < key}
HERE=$(cd "$(dirname "$0")" && pwd)
SITE=/etc/nginx/sites-available/aes-pixel
ENABLED=/etc/nginx/sites-enabled/zz-aes-pixel

[ "$(id -u)" -eq 0 ] || { echo "Run with sudo." >&2; exit 1; }
IFS= read -r KEY || true
[ "${#KEY}" -ge 32 ] || { echo "Pipe the AES_PIXEL_KEY (32+ chars) on stdin." >&2; exit 1; }

render() { sed "s/track\.aliniant\.net/${HOST}/g" "$HERE/$1" > "$SITE"; }

echo "== packages"
MISSING=()
for cmd in nginx certbot python3 curl; do
    command -v "$cmd" >/dev/null 2>&1 || MISSING+=("$cmd")
done
FRESH_NGINX=0
if [ "${#MISSING[@]}" -gt 0 ]; then
    for m in "${MISSING[@]}"; do [ "$m" = nginx ] && FRESH_NGINX=1; done
    apt-get update -q
    DEBIAN_FRONTEND=noninteractive apt-get install -y -q "${MISSING[@]}"
else
    echo "all present, nothing installed"
fi

echo "== pixel service"
install -d -m 755 /opt/aes-pixel /var/www/aes-pixel-acme
install -m 644 "$HERE/aes_pixel_server.py" /opt/aes-pixel/aes_pixel_server.py
( umask 077; printf 'AES_PIXEL_KEY=%s\n' "$KEY" > /etc/aes-pixel.env )
install -m 644 "$HERE/aes-pixel.service" /etc/systemd/system/aes-pixel.service
systemctl daemon-reload
systemctl enable aes-pixel
systemctl restart aes-pixel

echo "== nginx"
if [ "$FRESH_NGINX" -eq 1 ]; then rm -f /etc/nginx/sites-enabled/default; fi
if [ ! -f "/etc/letsencrypt/live/${HOST}/fullchain.pem" ]; then
    render nginx-track.aliniant.net.http.conf
    ln -sf "$SITE" "$ENABLED"
    nginx -t
    systemctl reload nginx
    certbot certonly --webroot -w /var/www/aes-pixel-acme -d "$HOST" \
        -m "$EMAIL" --agree-tos -n --deploy-hook "systemctl reload nginx"
fi
render nginx-track.aliniant.net.conf
ln -sf "$SITE" "$ENABLED"
nginx -t
systemctl reload nginx

echo "== check"
sleep 1
curl -fsS "https://${HOST}/healthz" && echo
systemctl --no-pager --lines=0 status aes-pixel
