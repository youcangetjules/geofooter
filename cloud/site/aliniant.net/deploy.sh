#!/usr/bin/env bash
# Publish the aliniant.net holding page on an Ubuntu 24.04 VM, either fresh or
# already serving other nginx sites.
#
# Prerequisites: aliniant.net and www.aliniant.net A records point at this VM,
# and ports 80/443 are open.
#
# Usage, on the VM from a copy of cloud/site/aliniant.net/:
#   sudo ./deploy.sh admin@example.com
#
# Safe to re-run: it republishes the page and only asks certbot for a
# certificate when there isn't one yet. Installs only missing packages (no
# nginx upgrade/restart) and enables the site as "zz-aliniant-net" so it loads
# after other sites and never becomes the catch-all. Each nginx change is
# checked with nginx -t and applied with a graceful reload.
set -euo pipefail

EMAIL=${1:?usage: deploy.sh <letsencrypt-email>}
HERE=$(cd "$(dirname "$0")" && pwd)
SITE=/etc/nginx/sites-available/aliniant-net
ENABLED=/etc/nginx/sites-enabled/zz-aliniant-net
WEBROOT=/var/www/aliniant-net

[ "$(id -u)" -eq 0 ] || { echo "Run with sudo." >&2; exit 1; }

echo "== packages"
MISSING=()
for cmd in nginx certbot curl; do
    command -v "$cmd" >/dev/null 2>&1 || MISSING+=("$cmd")
done
if [ "${#MISSING[@]}" -gt 0 ]; then
    apt-get update -q
    DEBIAN_FRONTEND=noninteractive apt-get install -y -q "${MISSING[@]}"
else
    echo "all present, nothing installed"
fi

echo "== page"
install -d -m 755 "$WEBROOT" /var/www/aliniant-net-acme
install -m 644 "$HERE/index.html" "$HERE/aliniant-shire.jpg" "$HERE/favicon.ico" "$WEBROOT/"

echo "== nginx"
if [ ! -f /etc/letsencrypt/live/aliniant.net/fullchain.pem ]; then
    install -m 644 "$HERE/nginx-aliniant.net.http.conf" "$SITE"
    ln -sf "$SITE" "$ENABLED"
    nginx -t
    systemctl reload nginx
    certbot certonly --webroot -w /var/www/aliniant-net-acme \
        -d aliniant.net -d www.aliniant.net \
        -m "$EMAIL" --agree-tos -n --deploy-hook "systemctl reload nginx"
fi
install -m 644 "$HERE/nginx-aliniant.net.conf" "$SITE"
ln -sf "$SITE" "$ENABLED"
nginx -t
systemctl reload nginx

echo "== check"
sleep 1
curl -fsS -o /dev/null -w "https://aliniant.net/ -> %{http_code}\n" https://aliniant.net/
curl -sS -o /dev/null -w "https://www.aliniant.net/ -> %{http_code} %{redirect_url}\n" https://www.aliniant.net/
