#!/usr/bin/env bash
# Set up the AES read-tracking pixel endpoint on a fresh Ubuntu 24.04 VM.
#
# Prerequisites: the host name's A record points at this VM, and ports 80/443
# are open to the internet (GCP: "Allow HTTP/HTTPS traffic" on the instance).
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

[ "$(id -u)" -eq 0 ] || { echo "Run with sudo." >&2; exit 1; }
IFS= read -r KEY || true
[ "${#KEY}" -ge 32 ] || { echo "Pipe the AES_PIXEL_KEY (32+ chars) on stdin." >&2; exit 1; }

render() { sed "s/track\.aliniant\.net/${HOST}/g" "$HERE/$1" > "$SITE"; }

echo "== packages"
apt-get update -q
DEBIAN_FRONTEND=noninteractive apt-get install -y -q nginx certbot python3 curl

echo "== pixel service"
install -d -m 755 /opt/aes-pixel /var/www/aes-pixel-acme
install -m 644 "$HERE/aes_pixel_server.py" /opt/aes-pixel/aes_pixel_server.py
( umask 077; printf 'AES_PIXEL_KEY=%s\n' "$KEY" > /etc/aes-pixel.env )
install -m 644 "$HERE/aes-pixel.service" /etc/systemd/system/aes-pixel.service
systemctl daemon-reload
systemctl enable aes-pixel
systemctl restart aes-pixel

echo "== nginx"
rm -f /etc/nginx/sites-enabled/default
if [ ! -f "/etc/letsencrypt/live/${HOST}/fullchain.pem" ]; then
    render nginx-track.aliniant.net.http.conf
    ln -sf "$SITE" /etc/nginx/sites-enabled/aes-pixel
    nginx -t
    systemctl reload nginx
    certbot certonly --webroot -w /var/www/aes-pixel-acme -d "$HOST" \
        -m "$EMAIL" --agree-tos -n --deploy-hook "systemctl reload nginx"
fi
render nginx-track.aliniant.net.conf
ln -sf "$SITE" /etc/nginx/sites-enabled/aes-pixel
nginx -t
systemctl reload nginx

echo "== check"
sleep 1
curl -fsS "https://${HOST}/healthz" && echo
systemctl --no-pager --lines=0 status aes-pixel
