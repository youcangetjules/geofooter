import sys
import os
from pathlib import Path
import html
import json
import urllib.request
import urllib.error
import socket
import dns.resolver
import whois
import logging
import re
from typing import Optional, Dict, List, Tuple, Any
import ipaddress
from dataclasses import dataclass
import time
import pytz
import importlib
from datetime import datetime

# Setup logging
logging.basicConfig(
    filename='C:/Users/julia/Documents/OutlookGeoFooter/geolocate_debug.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

class GeoLocationResult:
    def __init__(self, ip: str):
        self.ip = ip
        self.success = False
        self.city = "Unknown"
        self.country = ""
        self.org = "Unknown"
        self.asn = "Unknown"
        self.blocklist = "Not Listed"
        self.security_data = {
            "Tor": False,
            "Proxy": False,
            "Anonymous": False,
            "VPN": False
        }

class AuthenticationInfo:
    def __init__(self):
        self.auth_lines = []
        self.spf = None
        self.dkim = None
        self.dmarc = None
        self.compauth = None

class EmailSecurityAnalyzer:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.COMPAUTH_REASON_CODES = {
            "102": "SPF/DKIM aligned, no DMARC policy",
            "103": "SPF/DKIM aligned, DMARC policy quarantine",
            "104": "SPF/DKIM aligned, DMARC policy reject",
            "109": "SPF/DKIM aligned, DMARC policy quarantine",
            "110": "SPF/DKIM aligned, DMARC policy reject",
            "601": "SPF/DKIM not aligned, DMARC policy quarantine",
            "602": "SPF/DKIM not aligned, no DMARC policy",
            "610": "SPF/DKIM not aligned, DMARC policy reject",
        }

    def escape_html(self, s: Optional[str]) -> str:
        if s is None:
            return ""
        s = str(s)
        return html.escape(s)  # Use the html module's escape function for proper HTML escaping

    def is_private_ip(self, ip: str) -> bool:
        try:
            ip_parts = list(map(int, ip.split('.')))
            if len(ip_parts) != 4:
                return False
            if ip_parts[0] == 10:
                return True
            if ip_parts[0] == 172 and 16 <= ip_parts[1] <= 31:
                return True
            if ip_parts[0] == 192 and ip_parts[1] == 168:
                return True
            if ip_parts[0] == 169 and ip_parts[1] == 254:
                return True
            if ip_parts[0] == 127:
                return True
            return False
        except (ValueError, IndexError):
            return False

    def is_valid_public_ip(self, ip: str) -> bool:
        if not ip:
            return False
        try:
            socket.inet_aton(ip)
            if self.is_private_ip(ip):
                self.logger.info(f"IP {ip} is_valid_public_ip check: False (is_private: True)")
                return False
            self.logger.info(f"IP {ip} is_valid_public_ip check: True (is_private: False)")
            return True
        except (socket.error, ValueError):
            self.logger.info(f"IP {ip} is_valid_public_ip check: False (invalid format)")
            return False

    def extract_sender_info(self, headers: str) -> Tuple[Optional[str], Optional[str]]:
        sender_match = re.search(r'^From:\s*(?:[^<]*<)?([^>\s]+@[^>\s]+)>?', headers, re.IGNORECASE | re.MULTILINE)
        sender_email = sender_match.group(1) if sender_match else None
        sender_domain = None
        if sender_email:
            sender_domain = sender_email.split('@')[-1].lower()
        self.logger.info(f"Extracted sender: {sender_email}, domain: {sender_domain}")
        return sender_email, sender_domain

    def extract_original_sender_ip(self, headers: str) -> Tuple[Optional[str], Optional[str]]:
        self.logger.info("Attempting to extract original sender IP and hostname.")
        received_lines = re.findall(r'^Received:\s*from\s*([^\s]+)\s*\((?:[^\[]*\[)?([0-9a-fA-F.:]+)(?:\])?\)', headers, re.IGNORECASE | re.MULTILINE)
        for hostname, ip in reversed(received_lines):
            self.logger.info(f"IP {ip} is_private_ip check (custom range): {self.is_private_ip(ip)}")
            if self.is_valid_public_ip(ip):
                self.logger.info(f"Found public IP in Received header: {ip}")
                return ip, hostname
        self.logger.warning("No valid public IP found in Received headers.")
        return None, None

    def extract_authentication_info(self, headers: str) -> AuthenticationInfo:
        auth_info = AuthenticationInfo()
        auth_results_match = re.search(r'^(Authentication-Results(?:-Original)?):\s*([^\n]+(?:\n\s+[^\n]+)*)', headers, re.IGNORECASE | re.MULTILINE)
        if auth_results_match:
            auth_header_content = auth_results_match.group(2).replace('\n', ' ').strip()
            auth_info.auth_lines.append(auth_results_match.group(0).strip())
            self.logger.info(f"Auth-Results header found: {auth_header_content}")

            # Split by semicolon, but only when followed by a key=value pair
            parts = re.split(r';\s*(?=\w+=)', auth_header_content)
            self.logger.debug(f"Split Auth-Results into {len(parts)} parts: {parts}")
            for part in parts:
                part = part.strip()
                if not part:
                    self.logger.debug("Skipping empty part")
                    continue
                # Match key=value, capturing the value more flexibly
                match = re.match(r'(\w+)=([a-zA-Z0-9\-._]+)(?:\s|$)', part, re.IGNORECASE)
                if match:
                    key, value = match.group(1).lower(), match.group(2).lower()
                    self.logger.debug(f"Matched part '{part}': key={key}, value={value}")
                    if key == "spf":
                        auth_info.spf = value
                    elif key == "dkim":
                        auth_info.dkim = value
                    elif key == "dmarc":
                        auth_info.dmarc = value
                        if auth_info.dmarc == "bestguesspass":
                            auth_info.dmarc = "pass"
                    elif key == "compauth":
                        compauth_match = re.search(r'compauth=([a-zA-Z]+)\s+reason=([a-zA-Z0-9_.-]+)', part, re.IGNORECASE)
                        if compauth_match:
                            auth_info.compauth = {"result": compauth_match.group(1).lower(), "reason_code": compauth_match.group(2)}
                        else:
                            auth_info.compauth = {"result": value, "reason_code": "unknown"}
                else:
                    self.logger.debug(f"No key-value match for part: {part}")
        else:
            self.logger.warning("Authentication-Results header not found. Checking fallbacks.")
            spf_simple = re.search(r'Received-SPF:\s*([a-zA-Z]+)', headers, re.IGNORECASE)
            if spf_simple:
                auth_info.spf = spf_simple.group(1).lower()
                self.logger.info(f"Fallback SPF: {auth_info.spf}")

        if not auth_info.compauth:
            auth_info.compauth = {"result": "missing", "reason_code": "not_present"}

        self.logger.info(f"Final Extracted Auth: SPF={auth_info.spf}, DKIM={auth_info.dkim}, DMARC={auth_info.dmarc}, CompAuth={auth_info.compauth}")
        return auth_info

    def geolocate_ip(self, ip: str) -> GeoLocationResult:
        result = GeoLocationResult(ip)
        if not self.is_valid_public_ip(ip):
            result.city = "Invalid IP"
            result.blocklist = "N/A (Invalid IP)"
            return result

        # Try ipwhois.io first
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            self.logger.info(f"Attempting ipwhois.io API call for {ip}, attempt {attempt}")
            try:
                with urllib.request.urlopen(f"https://ipwhois.io/{ip}/json", timeout=5) as response:
                    data = json.loads(response.read().decode('utf-8'))
                    if data.get("success", False):
                        result.success = True
                        result.city = data.get("city", "Unknown") or "Unknown"
                        result.country = data.get("country", "") or ""
                        result.org = data.get("org", "Unknown") or "Unknown"
                        result.asn = data.get("asn", "Unknown") or "Unknown"
                        security = data.get("security", {})
                        result.security_data["Tor"] = security.get("is_tor", False)
                        result.security_data["Proxy"] = security.get("is_proxy", False)
                        result.security_data["Anonymous"] = security.get("is_anonymous", False)
                        result.security_data["VPN"] = security.get("is_vpn", False)
                        break
            except urllib.error.URLError as e:
                self.logger.error(f"Geolocation URL/Network error for {ip} on attempt {attempt}: {e}")
                if attempt == max_retries:
                    self.logger.info(f"ipwhois.io failed for {ip}, trying ip-api.com")
                    break
                time.sleep(2)

        # Fallback to ip-api.com
        if not result.success:
            try:
                with urllib.request.urlopen(f"http://ip-api.com/json/{ip}?fields=66846719", timeout=5) as response:
                    data = json.loads(response.read().decode('utf-8'))
                    if data.get("status", "fail") == "success":
                        result.success = True
                        result.city = data.get("city", "Unknown") or "Unknown"
                        result.country = data.get("country", "") or ""
                        result.org = data.get("org", "Unknown") or "Unknown"
                        result.asn = data.get("as", "Unknown") or "Unknown"
                        result.security_data["Proxy"] = data.get("proxy", False)
                        result.security_data["VPN"] = data.get("hosting", False)
                        self.logger.info(f"Fallback geolocation for {ip}: City={result.city}, Country={result.country}, ORG={result.org}, ASN={result.asn}")
                    else:
                        result.city = data.get("message", "Unknown Error")
                        self.logger.error(f"ip-api.com failed for {ip}: {result.city}")
            except urllib.error.URLError as e:
                result.city = str(e)
                self.logger.error(f"ip-api.com URL/Network error for {ip}: {e}")

        # Blocklist check
        result.blocklist = self.check_multiple_blocklists(ip)
        return result

    def check_multiple_blocklists(self, ip: str) -> str:
        blocklists = [
            "dnsbl.sorbs.net",
            "zen.spamhaus.org",
            "bl.spamcop.net",
            "dnsbl-1.uceprotect.net"
        ]
        listed_in = []
        ip_reversed = ".".join(reversed(ip.split(".")))
        for bl in blocklists:
            query = f"{ip_reversed}.{bl}"
            try:
                dns.resolver.resolve(query, "A")
                listed_in.append(bl)
                self.logger.info(f"IP {ip} listed in blocklist: {bl}")
            except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
                continue
            except Exception as e:
                self.logger.error(f"Blocklist check error for {ip} on {bl}: {e}")
                continue
        if listed_in:
            return f"Listed: {', '.join(listed_in)}"
        return "Not Listed"

    def get_whois_info(self, domain: str) -> Dict[str, Any]:
        try:
            w = whois.whois(domain)
            registrar = w.registrar if w.registrar else "Unknown"
            creation_date = w.creation_date
            expiration_date = w.expiration_date
            status = w.status if w.status else "Unknown"
            if isinstance(creation_date, list):
                creation_date = creation_date[0] if creation_date else "Unknown"
            if isinstance(expiration_date, list):
                expiration_date = expiration_date[0] if expiration_date else "Unknown"
            if isinstance(creation_date, datetime):
                creation_date = creation_date.strftime("%Y-%m-%d")
            if isinstance(expiration_date, datetime):
                expiration_date = expiration_date.strftime("%Y-%m-%d")
            if not registrar or registrar.lower() == "unknown":
                registrar = "Unknown"
            if not creation_date or creation_date.lower() == "unknown":
                creation_date = "Unknown"
            if not expiration_date or expiration_date.lower() == "unknown":
                expiration_date = "Unknown"
            result = {
                "registrar": registrar,
                "creation_date": creation_date,
                "expiration_date": expiration_date,
                "status": status
            }
            self.logger.info(f"WHOIS for {domain}: {result}")
            return result
        except Exception as e:
            self.logger.error(f"WHOIS lookup failed for {domain}: {e}")
            return {"error": f"WHOIS Error: {str(e)}"}

    def assess_security_risk(self, geo_results: List[GeoLocationResult], auth_info: AuthenticationInfo, sender_domain_whois: Dict[str, Any]) -> Dict[str, Any]:
        score = 50
        factors = []
        max_score, min_score = 100, 0

        # Authentication: SPF, DKIM, DMARC
        if auth_info.spf and "pass" in auth_info.spf.lower():
            factors.append("SPF Validated")
            score += 10
        elif auth_info.spf:
            factors.append(f"SPF {auth_info.spf.upper()}")
            score -= 15 if "fail" in auth_info.spf.lower() else 7
        else:
            factors.append("SPF Missing/Problem")
            score -= 7

        if auth_info.dkim and "pass" in auth_info.dkim.lower():
            factors.append("DKIM Validated")
            score += 10
        elif auth_info.dkim:
            factors.append(f"DKIM {auth_info.dkim.upper()}")
            score -= 15 if "fail" in auth_info.dkim.lower() else 7
        else:
            factors.append("DKIM Missing/Problem")
            score -= 7

        if auth_info.dmarc and "pass" in auth_info.dmarc.lower():
            factors.append("DMARC Compliant")
            score += 15
        elif auth_info.dmarc:
            factors.append(f"DMARC {auth_info.dmarc.upper()}")
            score -= 20 if "fail" in auth_info.dmarc.lower() else 10
        else:
            factors.append("DMARC Missing/Problem")
            score -= 10

        # CompAuth
        if auth_info.compauth:
            compauth_result = auth_info.compauth.get("result", "unknown").lower()
            reason_code = auth_info.compauth.get("reason_code", "unknown")
            reason_description = self.COMPAUTH_REASON_CODES.get(reason_code, f"Unknown reason (code {reason_code})")
            factors.append(f"CompAuth {compauth_result.upper()} ({reason_description})")
            if compauth_result == "pass":
                score += 20
                if reason_code in ["104", "110"]:
                    score += 5
                    factors.append("CompAuth Pass with DMARC Reject Policy (+5)")
                elif reason_code in ["103", "109"]:
                    score += 3
                    factors.append("CompAuth Pass with DMARC Quarantine Policy (+3)")
            elif compauth_result == "fail":
                score -= 25
            else:
                score -= 10
        else:
            factors.append("CompAuth Missing")
            score -= 5

        # Geolocation & IP Reputation
        origin_geo = geo_results[0] if geo_results else None
        if origin_geo and origin_geo.success:
            if origin_geo.blocklist and "Listed: " in origin_geo.blocklist:
                factors.append(f"Origin IP Blocklisted ({origin_geo.blocklist.replace('Listed: ', '')})")
                score -= 30
            else:
                factors.append("Origin IP Not Blocklisted")
        elif origin_geo and not origin_geo.success and origin_geo.city:
            factors.append(f"Origin IP Geolocation Failed ({origin_geo.city})")
            score -= 5
        elif origin_geo:
            factors.append("Origin IP Geolocation Failed (Unknown Error)")
            score -= 5

        # Domain Age
        if sender_domain_whois and "creation_date" in sender_domain_whois and sender_domain_whois["creation_date"] != "Unknown":
            try:
                created_str = sender_domain_whois["creation_date"]
                parsed_date = None
                for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%SZ"):
                    try:
                        parsed_date = datetime.strptime(created_str.split(' ')[0], fmt)
                        break
                    except ValueError:
                        continue
                if parsed_date:
                    parsed_date = pytz.utc.localize(parsed_date)
                    now_aware = datetime.now(pytz.utc).astimezone(pytz.timezone('Europe/London'))
                    age_days = (now_aware - parsed_date).days
                    if age_days < 0:
                        factors.append("Domain Creation Date Invalid")
                        score -= 10
                    elif age_days < 30:
                        factors.append(f"Domain Very New ({age_days} days)")
                        score -= 15
                    elif age_days < 90:
                        factors.append(f"Domain New ({age_days} days)")
                        score -= 10
            except Exception as e:
                self.logger.error(f"Error processing domain age: {e}")
        elif sender_domain_whois and sender_domain_whois.get("error"):
            factors.append("Sender Domain WHOIS Error")
            score -= 5

        score = max(min_score, min(max_score, int(score)))
        risk_level = "LOW" if score >= 75 else "MEDIUM" if score >= 40 else "HIGH"
        risk_color = {"LOW": "#4CAF50", "MEDIUM": "#FFC107", "HIGH": "#F44336"}[risk_level]

        self.logger.info(f"Security Assessment: Score={score}, Level={risk_level}, Factors={factors}")
        return {"score": score, "risk_level": risk_level, "risk_color": risk_color, "factors": factors}

    def _build_compact_html_template(self, sender_email: Optional[str], sender_domain: Optional[str],
                                    sender_ip: Optional[str], sender_machine_name: Optional[str],
                                    sender_geo_result: GeoLocationResult, whois_info: Dict[str, Any],
                                    auth_info: AuthenticationInfo, security_assessment: Dict[str, Any]) -> str:
        s = self.escape_html

        analysis_time = datetime.now(pytz.timezone('Europe/London')).strftime('%Y-%m-%d %H:%M BST')
        self.logger.info(f"Analysis time: {analysis_time}")

        risk_level = s(security_assessment['risk_level'])
        risk_score = security_assessment['score']
        header_color = security_assessment['risk_color']
        self.logger.info(f"Header: Risk Level={risk_level}, Score={risk_score}, Color={header_color}")

        ip_display_val = s(sender_ip)
        if sender_machine_name and sender_machine_name.lower() != s(sender_ip).lower():
            ip_display_val = f"{s(sender_ip)} ({s(sender_machine_name)})"
        self.logger.info(f"IP Display: {ip_display_val}")

        geo_city = "Unknown"
        geo_country = ""
        geo_org = "Unknown"
        geo_asn = "Unknown"
        if sender_geo_result.success:
            geo_city = s(sender_geo_result.city)
            geo_country = s(sender_geo_result.country)
            geo_org = s(sender_geo_result.org)
            if sender_geo_result.asn and sender_geo_result.asn != "Unknown":
                asn_val = s(sender_geo_result.asn)
                if asn_val.upper().startswith("AS"):
                    geo_asn = asn_val
                elif asn_val.isdigit():
                    geo_asn = f"AS{asn_val}"
                else:
                    geo_asn = asn_val
        elif sender_geo_result.city:
            geo_city = s(sender_geo_result.city)
        self.logger.info(f"Geo Data: City={geo_city}, Country={geo_country}, Org={geo_org}, ASN={geo_asn}")

        location_display = f"{geo_city}, {geo_country}".strip(", ") if geo_city != "Unknown" or geo_country else "Unknown"
        if not location_display.strip() or location_display == ",":
            location_display = "Unknown"
        self.logger.info(f"Location Display: {location_display}")

        registrar = s(whois_info.get("registrar", "Unknown"))
        created = s(whois_info.get("creation_date", "Unknown"))
        if "error" in whois_info and whois_info["error"] != "Lookup not performed":
            registrar = f"<span style='color: #888888;'>{s(whois_info['error'])}</span>"
            created = "N/A"
        self.logger.info(f"WHOIS: Registrar={registrar}, Created={created}")

        # Security Flags Generation
        sec_flags_parts = []
        if sender_geo_result.success and sender_geo_result.security_data:
            for flag_name, is_active in sender_geo_result.security_data.items():
                if is_active:
                    sec_flags_parts.append(f"<span style='color: #FF0000;'>&#10008; Is {s(flag_name)}</span>")
                else:
                    sec_flags_parts.append(f"<span style='color: #008000;'>&#10004; Not {s(flag_name)}</span>")
        self.logger.info(f"Security Flags Parts: {sec_flags_parts}")

        # Handle blocklist status
        blocklist_status = ""
        if sender_geo_result.blocklist and sender_geo_result.blocklist not in ["N/A (IPv6/Invalid)", "N/A (Invalid IP)"]:
            if "Listed: " in sender_geo_result.blocklist:
                blocklist_short = sender_geo_result.blocklist.replace("Listed: ", "")
                blocklist_status = f"<span style='color: #FF0000;'>&#10008; Blocklisted ({s(blocklist_short)})</span>"
            else:
                blocklist_status = f"<span style='color: #008000;'>&#10004; Not Blocklisted</span>"
        self.logger.info(f"Blocklist Status: {blocklist_status}")

        # Combine security flags
        all_security_flags = sec_flags_parts[:]
        if blocklist_status:
            all_security_flags.append(blocklist_status)
        security_flags_display = "<br />".join(all_security_flags) if all_security_flags else "<span style='color: #008000;'>&#10004; None</span>"
        self.logger.info(f"Security Flags Display: {security_flags_display}")

        # Authentication Status
        spf_status = "&#10004;" if auth_info.spf and "pass" in auth_info.spf.lower() else "&#10008;"
        dkim_status = "&#10004;" if auth_info.dkim and "pass" in auth_info.dkim.lower() else "&#10008;"
        dmarc_status = "&#10004;" if auth_info.dmarc and "pass" in auth_info.dmarc.lower() else "&#10008;"
        compauth_status = "&#10004;" if auth_info.compauth and auth_info.compauth.get("result", "").lower() == "pass" else "&#10008;"

        compauth_display = "None"
        if auth_info.compauth:
            reason_code = auth_info.compauth.get("reason_code", "unknown")
            reason_description = self.COMPAUTH_REASON_CODES.get(reason_code, f"Unknown reason (code {reason_code})")
            compauth_display = f"{s(auth_info.compauth.get('result', 'unknown'))} ({s(reason_description)})"

        self.logger.info(
            f"Auth Status: SPF={spf_status} ({auth_info.spf}), DKIM={dkim_status} ({auth_info.dkim}), "
            f"DMARC={dmarc_status} ({auth_info.dmarc}), CompAuth={compauth_status} ({compauth_display})"
        )

        # Simplified HTML body with even tab alignment and Security Flags formatting
        body_html = f"""
            <table style="font-family: Arial, sans-serif; font-size: 12px;">
                <tr><td colspan="2" style="background-color: {header_color}; color: white; padding: 5px; font-weight: bold;">EMAIL SECURITY ANALYSIS - RISK: {risk_level} ({risk_score}/100)</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">Sender:</td><td>{s(sender_email)}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">Domain:</td><td>{s(sender_domain)}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">IP:</td><td>{ip_display_val}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">Location:</td><td>{location_display}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">Organization:</td><td>{geo_org}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">A.S.N:</td><td>{geo_asn}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">Registrar:</td><td>{registrar}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">Created:</td><td>{created}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">SPF:</td><td><span style='color: {"#008000" if spf_status == "&#10004;" else "#FF0000"}'>{spf_status}</span> {s(auth_info.spf or "None")}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">DKIM:</td><td><span style='color: {"#008000" if dkim_status == "&#10004;" else "#FF0000"}'>{dkim_status}</span> {s(auth_info.dkim or "None")}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">DMARC:</td><td><span style='color: {"#008000" if dmarc_status == "&#10004;" else "#FF0000"}'>{dmarc_status}</span> {s(auth_info.dmarc or "None")}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">CompAuth:</td><td><span style='color: {"#008000" if compauth_status == "&#10004;" else "#FF0000"}'>{compauth_status}</span> {compauth_display}</td></tr>
                <tr><td style="padding: 2px 10px 2px 0;">Security Flags:</td><td></td></tr>
                <tr><td style="padding: 2px 10px 2px 20px;"> </td><td>{security_flags_display}</td></tr>
                <tr><td colspan="2" style="padding: 2px 0;">Analysis: {analysis_time}</td></tr>
            </table>
        """

        return f"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        <title>Email Security Analysis</title>
    </head>
    <body style="margin: 0; padding: 10px; background-color: #FFFFFF; font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #333333;">
        {body_html}
    </body>
    </html>"""

    def process_headers(self, headers: str) -> str:
        sender_email, sender_domain = self.extract_sender_info(headers)
        sender_ip, sender_hostname = self.extract_original_sender_ip(headers)
        auth_info = self.extract_authentication_info(headers)

        geo_results = []
        if sender_ip:
            geo_result = self.geolocate_ip(sender_ip)
            geo_results.append(geo_result)
        else:
            geo_result = GeoLocationResult("Unknown")
            geo_result.city = "No IP Found"
            geo_results.append(geo_result)

        whois_info = {"error": "Lookup not performed"}
        if sender_domain:
            whois_info = self.get_whois_info(sender_domain)

        security_assessment = self.assess_security_risk(geo_results, auth_info, whois_info)
        html_output = self._build_compact_html_template(
            sender_email, sender_domain, sender_ip, sender_hostname,
            geo_results[0], whois_info, auth_info, security_assessment
        )
        return html_output

def check_dependencies():
    logger = logging.getLogger(__name__)
    try:
        import dns.resolver
        logger.info("Dependency dns.resolver is installed.")
    except ImportError:
        logger.error("Dependency 'dnspython' is not installed.")
        raise
    try:
        import whois
        logger.info("Dependency whois is installed.")
    except ImportError:
        logger.error("Dependency 'python-whois' is not installed.")
        raise
    try:
        import pytz
        logger.info("Dependency pytz is installed.")
    except ImportError:
        logger.error("Dependency 'pytz' is not installed.")
        raise
    logger.info("All dependencies (dnspython, python-whois, pytz) are installed.")

def main(input_file: str, output_file: str):
    logger = logging.getLogger(__name__)
    current_time = datetime.now(pytz.timezone('Europe/London')).strftime('%Y-%m-%d %H:%M BST')
    logger.info(f"Script started at {current_time}")

    check_dependencies()

    logger.info(f"Processing started: {input_file} -> {output_file}")
    headers = ""
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            headers = f.read()
    except UnicodeDecodeError as e:
        logger.warning(f"Failed to read with utf-8: {e}")
        with open(input_file, 'r', encoding='utf-16') as f:
            headers = f.read()
        logger.info("Read headers with utf-16.")
    except Exception as e:
        logger.error(f"Failed to read input file: {e}")
        raise

    analyzer = EmailSecurityAnalyzer()
    html_output = analyzer.process_headers(headers)

    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_output)
        logger.info(f"HTML footer generated: {output_file}")
    except Exception as e:
        logger.error(f"Failed to write output file: {e}")
        raise

    logger.info("Script finished successfully.")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python geolocate_headers.py <input_file> <output_file>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])