#!/usr/bin/env python3
"""
Email Security Analysis Tool - Fixed Version
Analyzes email headers for security indicators, geolocation, and authentication.
"""

import sys
import os
import re
import json
import time
import socket
import logging
import ipaddress
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Any, Union
from dataclasses import dataclass, field
from datetime import datetime
from logging.handlers import RotatingFileHandler

import urllib.request
import urllib.error
import dns.resolver
import whois
import requests
import pytz
import html

# Set up logging at the very first step
log_file = os.path.join(os.path.dirname(sys.argv[1]) if len(sys.argv) > 1 else "C:\\GeoFooter", "geolocate_debug.log")
logging.basicConfig(filename=log_file, level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')
logging.debug("Script started with arguments: %s", sys.argv)

# Configuration
CONFIG = {
    "base_path": "C:/GeoFooter",
    "log_file": "geolocate_debug.log",
    "log_max_bytes": 5 * 1024 * 1024,
    "log_backup_count": 5,
    "timezone": "Europe/London",
    "api_keys": {
        "abuseipdb": "",
        "ipinfo": "1a9a711593217b",
        "ipgeolocation": "82dad7d806b54c158da274d5576e7702"
    },
    "blocklists": {
        "ipv4": ["dnsbl.sorbs.net", "zen.spamhaus.org", "bl.spamcop.net", "dnsbl-1.uceprotect.net"],
        "ipv6": ["zen.spamhaus.org"]
    },
    "risk_thresholds": {
        "low": 75,
        "medium": 40
    },
    "domain_age_thresholds": {
        "extremely_new": 10,
        "very_new": 30,
        "new": 180
    }
}

# CompAuth Reason Codes
COMPAUTH_REASON_CODES = {
    "100": "Unknown reason (generic pass or no specific reason provided)",
    "101": "SPF aligned, no DKIM, no DMARC policy",
    "102": "SPF/DKIM aligned, no DMARC policy",
    "103": "SPF/DKIM aligned, DMARC policy quarantine",
    "104": "SPF/DKIM aligned, DMARC policy reject",
    "105": "DKIM aligned, no SPF, no DMARC policy",
    "106": "DKIM aligned, SPF failed, no DMARC policy",
    "107": "SPF failed, DKIM aligned, DMARC policy quarantine",
    "108": "SPF failed, DKIM aligned, DMARC policy reject",
    "109": "SPF/DKIM aligned, DMARC policy quarantine (duplicate for clarity)",
    "110": "SPF/DKIM aligned, DMARC policy reject (duplicate for clarity)",
    "200": "Authentication failed, no SPF/DKIM alignment",
    "201": "SPF failed, DKIM failed, no DMARC policy",
    "202": "SPF failed, DKIM failed, DMARC policy quarantine",
    "203": "SPF failed, DKIM failed, DMARC policy reject",
    "300": "Temporary authentication issue (e.g., timeout or server error)",
    "400": "Invalid or malformed authentication data",
    "601": "SPF/DKIM not aligned, DMARC policy quarantine",
    "602": "SPF/DKIM not aligned, no DMARC policy",
    "610": "SPF/DKIM not aligned, DMARC policy reject",
    "700": "Policy override (e.g., manual whitelist)",
    "800": "Authentication bypassed (e.g., internal relay)",
}

@dataclass
class GeoLocationResult:
    """Represents geolocation information for an IP address."""
    ip: str
    success: bool = False
    city: str = "Unknown"
    country: str = ""
    org: str = "Unknown"
    asn: str = "Unknown"
    blocklist: str = "Not Listed"
    security_data: Dict[str, Optional[bool]] = field(default_factory=lambda: {
        "Tor": None,
        "Proxy": None,
        "Anonymous": None,
        "VPN": None
    })

@dataclass
class AuthenticationInfo:
    """Represents email authentication information."""
    auth_lines: List[str] = field(default_factory=list)
    spf: Optional[str] = None
    dkim: Optional[str] = None
    dmarc: Optional[str] = None
    compauth: Optional[Dict[str, str]] = None
    arc: Optional[Dict[str, Optional[str]]] = None
    arc_seal_valid: bool = False
    arc_chain_info: List[Dict[str, Any]] = field(default_factory=list)
    auth_type: str = "None"

@dataclass
class SecurityAssessment:
    """Represents the security risk assessment results."""
    score: int
    risk_level: str
    risk_color: str
    factors: List[str]
    domain_age_flag: Optional[str] = None

class LoggingSetup:
    """Handles logging configuration."""
    
    @staticmethod
    def setup():
        """Set up logging with rotation."""
        os.makedirs(CONFIG["base_path"], exist_ok=True)
        
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        
        log_path = Path(CONFIG["base_path"]) / CONFIG["log_file"]
        handler = RotatingFileHandler(
            log_path,
            maxBytes=CONFIG["log_max_bytes"],
            backupCount=CONFIG["log_backup_count"]
        )
        
        tz = pytz.timezone(CONFIG["timezone"])
        
        class TimezoneFormatter(logging.Formatter):
            def formatTime(self, record, datefmt=None):
                dt = datetime.fromtimestamp(record.created, tz)
                if datefmt:
                    return dt.strftime(datefmt)
                return dt.strftime('%Y-%m-%d %H:%M:%S,%f')[:-3] + f' {tz.tzname(dt)}'
        
        formatter = TimezoneFormatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S,%f'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

class IPValidator:
    """Handles IP address validation."""
    
    @staticmethod
    def is_private(ip: str) -> bool:
        """Check if IP is private."""
        try:
            ip_obj = ipaddress.ip_address(ip)
            return (ip_obj.is_private or ip_obj.is_loopback or 
                   ip_obj.is_link_local or ip_obj.is_multicast)
        except ValueError:
            return False
    
    @staticmethod
    def is_valid_public(ip: str) -> bool:
        """Check if IP is valid and public."""
        if not ip:
            return False
        try:
            ip_obj = ipaddress.ip_address(ip)
            return (ip_obj.is_global and not ip_obj.is_multicast and
                   not ip_obj.is_reserved and not ip_obj.is_private)
        except ValueError:
            return False

class HeaderParser:
    """Parses email headers for various information."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def extract_sender_info(self, headers: str) -> Tuple[Optional[str], Optional[str]]:
        """Extract sender email and domain from headers."""
        patterns = [
            r'^From:\s*<([^>]+)>',
            r'^From:\s*"[^"]*"\s*<([^>]+)>',
            r'^From:\s*([^<\s]+@[^>\s]+)',
            r'^From:\s*[^<]*<([^>]+)>',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, headers, re.IGNORECASE | re.MULTILINE)
            if match:
                sender_email = match.group(1).strip()
                sender_domain = sender_email.split('@')[-1].lower()
                self.logger.info(f"Extracted sender: {sender_email}, domain: {sender_domain}")
                return sender_email, sender_domain
        
        self.logger.warning("Could not extract sender from headers")
        return None, None
    
    def extract_original_sender_ip(self, headers: str) -> Tuple[Optional[str], Optional[str]]:
        """Extract the original sender's public IP and hostname."""
        self.logger.info("Extracting original sender IP and hostname")
        
        pattern = (r'^Received:\s*from\s+([^\s\(]+)\s*'
                  r'(?:\([^)]*\))?\s*'
                  r'(?:by\s+[^\s]+\s*)?'
                  r'(?:\([^)]*\[([^\]]+)\][^)]*\))?')
        
        simple_pattern = r'^Received:\s*from\s+([^\s]+)\s+\(([^\[]+)\[([^\]]+)\]\)'
        
        received_headers = re.findall(pattern, headers, re.IGNORECASE | re.MULTILINE)
        simple_headers = re.findall(simple_pattern, headers, re.IGNORECASE | re.MULTILINE)
        
        for hostname, hostname2, ip in reversed(simple_headers):
            self.logger.debug(f"Processing (simple): Hostname={hostname}, IP={ip}")
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                return ip, hostname
        
        for hostname, ip in reversed(received_headers):
            if ip:
                self.logger.debug(f"Processing: Hostname={hostname}, IP={ip}")
                if IPValidator.is_valid_public(ip):
                    self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                    return ip, hostname
        
        self.logger.warning("No valid public sender IP found")
        return None, None

class AuthenticationParser:
    """Parses authentication-related headers."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def extract_authentication_info(self, headers: str) -> AuthenticationInfo:
        """Extract authentication information from headers."""
        auth_info = AuthenticationInfo()
        
        auth_pattern = r'^Authentication-Results:\s*([^\n]+(?:\n\s+[^\n]+)*)'
        auth_match = re.search(auth_pattern, headers, re.IGNORECASE | re.MULTILINE)
        
        if auth_match:
            auth_content = auth_match.group(1).replace('\n', ' ').replace('\r', ' ').strip()
            auth_info.auth_lines.append(auth_match.group(0).strip())
            self.logger.info(f"Found Authentication-Results: {auth_content}")
            
            spf_match = re.search(r'spf=(\w+)', auth_content, re.IGNORECASE)
            if spf_match:
                auth_info.spf = spf_match.group(1).lower()
                self.logger.debug(f"SPF: {auth_info.spf}")
            
            dkim_match = re.search(r'dkim=(\w+)', auth_content, re.IGNORECASE)
            if dkim_match:
                auth_info.dkim = dkim_match.group(1).lower()
                self.logger.debug(f"DKIM: {auth_info.dkim}")
            
            dmarc_match = re.search(r'dmarc=(\w+)', auth_content, re.IGNORECASE)
            if dmarc_match:
                auth_info.dmarc = dmarc_match.group(1).lower()
                if auth_info.dmarc == "bestguesspass":
                    auth_info.dmarc = "pass"
                self.logger.debug(f"DMARC: {auth_info.dmarc}")
            
            compauth_match = re.search(r'compauth=(\w+)(?:\s+reason=(\d+))?', auth_content, re.IGNORECASE)
            if compauth_match:
                auth_info.compauth = {
                    "result": compauth_match.group(1).lower(),
                    "reason_code": compauth_match.group(2) if compauth_match.group(2) else "unknown"
                }
                self.logger.debug(f"CompAuth: {auth_info.compauth}")
        else:
            self.logger.warning("No Authentication-Results header found")
        
        self._parse_arc_headers(headers, auth_info)
        if not auth_info.arc:
            auth_info.arc = {"result": "none"}
        
        self._determine_auth_type(auth_info)
        
        if not auth_info.compauth:
            auth_info.compauth = {"result": "missing", "reason_code": "not_present"}
        
        return auth_info
    
    def _parse_arc_headers(self, headers: str, auth_info: AuthenticationInfo):
        """Parse ARC-related headers."""
        pass
    
    def _determine_auth_type(self, auth_info: AuthenticationInfo):
        """Determine the overall authentication type."""
        if auth_info.compauth and auth_info.compauth.get("result") != "missing":
            auth_info.auth_type = "Standard"
        else:
            auth_info.auth_type = "Basic"

class GeoLocationService:
    """Handles IP geolocation and security checks."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def geolocate_ip(self, ip: str) -> GeoLocationResult:
        """Get geolocation and security information for an IP."""
        result = GeoLocationResult(ip)
        
        if not IPValidator.is_valid_public(ip):
            result.city = "Invalid or Private IP"
            result.blocklist = "N/A"
            return result
        
        self._try_abuseipdb(ip, result)
        self._try_ipinfo(ip, result)
        
        if not result.success:
            self._try_fallback_services(ip, result)
        
        result.blocklist = self._check_blocklists(ip)
        
        return result
    
    def _try_abuseipdb(self, ip: str, result: GeoLocationResult) -> bool:
        """Try AbuseIPDB for security flags."""
        try:
            headers = {"Key": CONFIG["api_keys"]["abuseipdb"], "Accept": "application/json"}
            response = requests.get(
                f"https://api.abuseipdb.com/api/v2/check?ipAddress={ip}&maxAgeInDays=90",
                headers=headers,
                timeout=10
            )
            data = response.json()
            self.logger.debug(f"AbuseIPDB response: {data}")
            
            abuse_data = data.get("data", {})
            if abuse_data:
                result.security_data["Tor"] = abuse_data.get("isTor", False)
                result.security_data["Proxy"] = abuse_data.get("isProxy", False) or abuse_data.get("isVpn", False)
                result.security_data["VPN"] = abuse_data.get("isVpn", False)
                result.security_data["Anonymous"] = abuse_data.get("isAnonymous", False)
                return True
        except Exception as e:
            self.logger.error(f"AbuseIPDB error for {ip}: {e}")
        return False
    
    def _try_ipinfo(self, ip: str, result: GeoLocationResult) -> bool:
        """Try ipinfo.io for geolocation and security flags."""
        try:
            response = requests.get(
                f"https://ipinfo.io/{ip}/json?token={CONFIG['api_keys']['ipinfo']}",
                timeout=10
            )
            data = response.json()
            self.logger.debug(f"IPInfo response: {data}")
            
            result.success = True
            result.city = data.get("city", "Unknown") or "Unknown"
            result.country = data.get("country", "") or ""
            result.org = data.get("org", "Unknown") or "Unknown"
            
            if "asn" in data and isinstance(data["asn"], dict):
                result.asn = data["asn"].get("asn", "Unknown") or "Unknown"
            else:
                result.asn = "Unknown"
            
            if result.security_data["VPN"] is None:
                privacy = data.get("privacy", {})
                result.security_data["VPN"] = privacy.get("vpn", False)
                result.security_data["Proxy"] = privacy.get("proxy", False)
                result.security_data["Tor"] = privacy.get("tor", False)
                result.security_data["Anonymous"] = privacy.get("anonymous", False)
            
            return True
        except Exception as e:
            self.logger.error(f"ipinfo.io error for {ip}: {e}")
        return False
    
    def _try_fallback_services(self, ip: str, result: GeoLocationResult):
        """Try fallback geolocation services."""
        try:
            url = f"http://ip-api.com/json/{ip}?fields=66777215"
            with urllib.request.urlopen(url, timeout=10) as response:
                data = json.loads(response.read().decode('utf-8'))
                
                if data.get("status") == "success":
                    result.success = True
                    result.city = data.get("city", "Unknown") or "Unknown"
                    result.country = data.get("country", "") or ""
                    result.org = data.get("org", "Unknown") or "Unknown"
                    result.asn = data.get("as", "Unknown") or "Unknown"
                    return
        except Exception as e:
            self.logger.error(f"ip-api.com error for {ip}: {e}")
        
        result.city = "Geolocation Unavailable"
        result.blocklist = "N/A (Geolocation Failed)"
    
    def _check_blocklists(self, ip: str) -> str:
        """Check IP against multiple blocklists."""
        try:
            ip_obj = ipaddress.ip_address(ip)
            is_ipv6 = ip_obj.version == 6
            
            blocklists = CONFIG["blocklists"]["ipv6" if is_ipv6 else "ipv4"]
            listed_in = []
            
            if not is_ipv6:
                ip_reversed = ".".join(reversed(ip.split(".")))
                for bl in blocklists:
                    if self._check_single_blocklist(f"{ip_reversed}.{bl}"):
                        listed_in.append(bl)
            
            return f"Listed: {', '.join(listed_in)}" if listed_in else "Not Listed"
            
        except ValueError:
            return "N/A (Invalid IP)"
    
    def _check_single_blocklist(self, query: str) -> bool:
        """Check a single blocklist query."""
        try:
            dns.resolver.resolve(query, "A")
            return True
        except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
            return False
        except Exception as e:
            self.logger.error(f"Blocklist check error for {query}: {e}")
            return False

class WhoisService:
    """Handles WHOIS lookups."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def get_whois_info(self, domain: str) -> Dict[str, Any]:
        """Get WHOIS information for a domain."""
        try:
            w = whois.whois(domain)
            
            registrar = self._get_value(w.registrar, "Unknown")
            creation_date = self._normalize_date(w.creation_date)
            expiration_date = self._normalize_date(w.expiration_date)
            status = self._get_value(w.status, "Unknown")
            
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
    
    def _get_value(self, value: Any, default: str) -> str:
        """Get value or return default."""
        if not value or (isinstance(value, str) and value.lower() == "unknown"):
            return default
        return value
    
    def _normalize_date(self, date_value: Any) -> str:
        """Normalize date to string format."""
        if isinstance(date_value, list):
            date_value = date_value[0] if date_value else None
        
        if isinstance(date_value, datetime):
            return date_value.strftime("%Y-%m-%d")
        
        if not date_value or (isinstance(date_value, str) and date_value.lower() == "unknown"):
            return "Unknown"
        
        return str(date_value)

class SecurityAssessor:
    """Assesses security risk based on various factors."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def assess_security_risk(
        self,
        geo_results: List[GeoLocationResult],
        auth_info: AuthenticationInfo,
        sender_domain_whois: Dict[str, Any]
    ) -> SecurityAssessment:
        """Calculate security risk score and assessment."""
        score = 50
        factors = []
        domain_age_flag = None
        
        score = self._assess_authentication(auth_info, score, factors)
        score = self._assess_geolocation(geo_results, score, factors)
        score, domain_age_flag = self._assess_domain_age(sender_domain_whois, score, factors)
        
        score = max(0, min(100, int(score)))
        risk_level = self._calculate_risk_level(score, domain_age_flag)
        risk_color = {"LOW": "#4CAF50", "MEDIUM": "#FFC107", "HIGH": "#F44336"}[risk_level]
        
        return SecurityAssessment(
            score=score,
            risk_level=risk_level,
            risk_color=risk_color,
            factors=factors,
            domain_age_flag=domain_age_flag
        )
    
    def _assess_authentication(
        self,
        auth_info: AuthenticationInfo,
        score: int,
        factors: List[str]
    ) -> int:
        """Assess authentication methods."""
        if auth_info.spf:
            if "pass" in auth_info.spf.lower():
                factors.append("SPF Validated")
                score += 10
            elif "fail" in auth_info.spf.lower():
                factors.append(f"SPF {auth_info.spf.upper()}")
                score -= 15
            else:
                factors.append(f"SPF {auth_info.spf.upper()}")
                score -= 7
        else:
            factors.append("SPF Missing/Problem")
            score -= 7
        
        if auth_info.dkim:
            if "pass" in auth_info.dkim.lower():
                factors.append("DKIM Validated")
                score += 10
            elif "fail" in auth_info.dkim.lower():
                factors.append(f"DKIM {auth_info.dkim.upper()}")
                score -= 15
            else:
                factors.append(f"DKIM {auth_info.dkim.upper()}")
                score -= 7
        else:
            factors.append("DKIM Missing/Problem")
            score -= 7
        
        if auth_info.dmarc:
            if "pass" in auth_info.dmarc.lower():
                factors.append("DMARC Compliant")
                score += 15
            elif "fail" in auth_info.dmarc.lower():
                factors.append(f"DMARC {auth_info.dmarc.upper()}")
                score -= 20
            else:
                factors.append(f"DMARC {auth_info.dmarc.upper()}")
                score -= 10
        else:
            factors.append("DMARC Missing/Problem")
            score -= 10
        
        return score
    
    def _assess_geolocation(
        self,
        geo_results: List[GeoLocationResult],
        score: int,
        factors: List[str]
    ) -> int:
        """Assess geolocation and security flags."""
        if not geo_results:
            return score
        
        origin_geo = geo_results[0]
        
        if origin_geo.success and origin_geo.blocklist and "Listed: " in origin_geo.blocklist:
            blocklists = origin_geo.blocklist.replace('Listed: ', '')
            factors.append(f"Origin IP Blocklisted ({blocklists})")
            score -= 30
        elif origin_geo.success:
            factors.append("Origin IP Not Blocklisted")
        
        if origin_geo.success:
            for flag_name, is_active in origin_geo.security_data.items():
                if is_active is None:
                    continue
                elif is_active:
                    factors.append(f"Security Flag: {flag_name} Active")
                    score -= 5
        
        return score
    
    def _assess_domain_age(
        self,
        whois_info: Dict[str, Any],
        score: int,
        factors: List[str]
    ) -> Tuple[int, Optional[str]]:
        """Assess domain age."""
        domain_age_flag = None
        
        if not whois_info or "error" in whois_info:
            if whois_info and whois_info.get("error"):
                factors.append("Sender Domain WHOIS Error")
                score -= 5
            return score, domain_age_flag
        
        creation_date = whois_info.get("creation_date", "Unknown")
        if creation_date == "Unknown":
            return score, domain_age_flag
        
        try:
            parsed_date = self._parse_date(creation_date)
            if not parsed_date:
                factors.append("Domain Creation Date Invalid")
                score -= 10
                return score, domain_age_flag
            
            parsed_date = pytz.utc.localize(parsed_date)
            now = datetime.now(pytz.utc)
            age_days = (now - parsed_date).days
            
            thresholds = CONFIG["domain_age_thresholds"]
            if age_days < 0:
                factors.append("Domain Creation Date Invalid")
                score -= 10
            elif age_days < thresholds["extremely_new"]:
                factors.append(f"Domain Extremely New ({age_days} days)")
                domain_age_flag = "Domain Less Than 10 Days Old - CAUTION"
                score = 0
            elif age_days < thresholds["very_new"]:
                factors.append(f"Domain Very New ({age_days} days)")
                score -= 30
            elif age_days < thresholds["new"]:
                factors.append(f"Domain New ({age_days} days)")
                score -= 10
            else:
                factors.append(f"Domain Established ({age_days} days)")
                
        except Exception as e:
            self.logger.error(f"Error processing domain age: {e}")
            factors.append("Domain Age Processing Error")
            score -= 5
        
        return score, domain_age_flag
    
    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse date string to datetime."""
        date_formats = ["%Y-%m-%d", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d %H:%M:%S"]
        date_str = date_str.split(' ')[0]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None
    
    def _calculate_risk_level(self, score: int, domain_age_flag: Optional[str]) -> str:
        """Calculate risk level based on score."""
        if domain_age_flag:
            return "HIGH"
        
        thresholds = CONFIG["risk_thresholds"]
        if score >= thresholds["low"]:
            return "LOW"
        elif score >= thresholds["medium"]:
            return "MEDIUM"
        else:
            return "HIGH"

class HTMLReportGenerator:
    """Generates HTML security reports."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def generate_report(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        sender_ip: Optional[str],
        sender_hostname: Optional[str],
        geo_result: GeoLocationResult,
        whois_info: Dict[str, Any],
        auth_info: AuthenticationInfo,
        security_assessment: SecurityAssessment
    ) -> str:
        """Generate HTML security report."""
        display_data = self._prepare_display_data(
            sender_email, sender_domain, sender_ip,
            sender_hostname, geo_result, whois_info,
            auth_info, security_assessment
        )
        return self._build_html(display_data)
    
    def _prepare_display_data(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        sender_ip: Optional[str],
        sender_hostname: Optional[str],
        geo_result: GeoLocationResult,
        whois_info: Dict[str, Any],
        auth_info: AuthenticationInfo,
        security_assessment: SecurityAssessment
    ) -> Dict[str, Any]:
        """Prepare all data for display."""
        tz = pytz.timezone(CONFIG["timezone"])
        analysis_time = datetime.now(tz).strftime('%Y-%m-%d %H:%M BST')
        
        ip_display = self._format_ip_display(sender_ip, sender_hostname)
        location_display = self._format_location(geo_result)
        whois_display = self._format_whois(whois_info)
        auth_display = self._format_authentication(auth_info)
        security_flags = self._format_security_flags(geo_result)
        
        def _auth_type_string() -> str:
            passed = []
            for meth, val in [("SPF", auth_info.spf), ("DKIM", auth_info.dkim),
                              ("DMARC", auth_info.dmarc)]:
                if val and val.lower() == "pass":
                    passed.append(meth)
            if len(passed) == 3:
                return "SPF + DKIM + DMARC (All Passed)"
            if passed:
                return " + ".join(passed) + " (Partial)"
            return "None"

        auth_type_str = _auth_type_string()
        
        return {
            "analysis_time": analysis_time,
            "risk_level": security_assessment.risk_level,
            "risk_score": security_assessment.score,
            "risk_color": security_assessment.risk_color,
            "domain_age_flag": security_assessment.domain_age_flag,
            "sender_email": self._format_email(sender_email),
            "sender_domain": sender_domain or "Unknown",
            "ip_display": ip_display,
            "location": location_display,
            "organization": geo_result.org if geo_result.success else "Unknown",
            "asn": self._format_asn(geo_result.asn),
            "whois": whois_display,
            "auth": auth_display,
            "security_flags": security_flags,
            "auth_type": auth_type_str,
        }
    
    def _format_ip_display(self, ip: Optional[str], hostname: Optional[str]) -> str:
        """Format IP address display."""
        if not ip:
            return "Unknown"
        ip_str = html.escape(ip)
        if hostname and hostname.lower() != ip.lower():
            return f"{ip_str} ({html.escape(hostname)})"
        return ip_str
    
    def _format_location(self, geo_result: GeoLocationResult) -> str:
        """Format location display."""
        if not geo_result.success and geo_result.city:
            return html.escape(geo_result.city)
        city = geo_result.city if geo_result.success else "Unknown"
        country = geo_result.country if geo_result.success else ""
        location = f"{city}, {country}".strip(", ")
        return location if location and location != "," else "Unknown"
    
    def _format_whois(self, whois_info: Dict[str, Any]) -> Dict[str, str]:
        """Format WHOIS information."""
        if "error" in whois_info and whois_info["error"] != "Lookup not performed":
            return {
                "registrar": f"<span style='color: #888888;'>{html.escape(whois_info['error'])}</span>",
                "created": "N/A",
                "age": "N/A"
            }
        registrar = html.escape(str(whois_info.get("registrar", "Unknown")))
        created = html.escape(str(whois_info.get("creation_date", "Unknown")))
        age_days = self._calculate_domain_age(whois_info.get("creation_date"))
        age = f"{age_days} days" if age_days is not None else "N/A"
        return {"registrar": registrar, "created": created, "age": age}
    
    def _calculate_domain_age(self, creation_date: str) -> Optional[int]:
        """Calculate domain age in days."""
        if not creation_date or creation_date == "Unknown":
            return None
        try:
            date_formats = ["%Y-%m-%d", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d %H:%M:%S"]
            date_str = creation_date.split(' ')[0]
            parsed_date = None
            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue
            if parsed_date:
                parsed_date = pytz.utc.localize(parsed_date)
                now = datetime.now(pytz.utc)
                return max(0, (now - parsed_date).days)
        except Exception as e:
            self.logger.error(f"Error calculating domain age: {e}")
        return None
    
    def _format_authentication(self, auth_info: AuthenticationInfo) -> Dict[str, Tuple[str, str, str] | str]:
        """Convert raw auth results to tuples of (symbol, colour, text)."""
        def status(val: Optional[str]) -> Tuple[str, str, str]:
            if not val or val.lower() in {"none", "missing"}:
                return ("–", "#888888", "None")
            val_lc = val.lower()
            if val_lc == "pass":
                return ("&#10003;", "#008000", "Pass")
            if val_lc == "fail":
                return ("&#10007;", "#ff0000", "Fail")
            return ("–", "#888888", val.capitalize())

        spf = status(auth_info.spf)
        dkim = status(auth_info.dkim)
        dmarc = status(auth_info.dmarc)

        if auth_info.compauth and auth_info.compauth.get("result") not in {"missing", None}:
            comp_res = auth_info.compauth["result"]
            reason = auth_info.compauth.get("reason_code", "unknown")
            reason_txt = COMPAUTH_REASON_CODES.get(reason, f"Reason {reason}")
            comp_disp = f"{comp_res.capitalize()} ({reason_txt})"
            compauth = status(comp_res)
        else:
            comp_disp = "Not Applicable (Non-Exchange)"
            compauth = ("–", "#888888", "Not Applicable")

        arc_res = (auth_info.arc or {}).get("result", "none")
        arc = status(arc_res)

        return {
            "spf": spf,
            "dkim": dkim,
            "dmarc": dmarc,
            "compauth": compauth,
            "compauth_display": comp_disp,
            "arc": arc,
        }
    
    def _format_security_flags(self, geo_result: GeoLocationResult) -> str:
        """Format security flags display."""
        flags = []
        for flag_name in ["Tor", "Proxy", "Anonymous", "VPN"]:
            is_active = geo_result.security_data.get(flag_name)
            if is_active is None:
                flags.append(f"<span style='color: #B8860B;' class='symbol'>- Is {flag_name} - No Data</span>")
            else:
                symbol = "✓" if is_active else "✗"
                color = "#008000" if is_active else "#FF0000"
                flags.append(f"<span style='color: {color};' class='symbol'>- Is {flag_name} {symbol}</span>")
        
        if geo_result.blocklist and geo_result.blocklist not in ["N/A (IPv6/Invalid)", "N/A (Invalid IP)"]:
            if "Listed: " in geo_result.blocklist:
                blocklists = geo_result.blocklist.replace("Listed: ", "")
                flags.append(f"<span style='color: #FF0000;' class='symbol' title='Blocklisted'>✗</span> Blocklisted ({html.escape(blocklists)})")
            else:
                flags.append(f"<span style='color: #008000;' class='symbol' title='Not Blocklisted'>✓</span> Not Blocklisted")
        
        return "<br>".join(flags) if flags else "<span style='color: #B8860B;'>- None</span>"
    
    def _format_asn(self, asn: str) -> str:
        """Format ASN display."""
        if not asn or asn == "Unknown":
            return "Unknown"
        asn = html.escape(asn)
        if asn.upper().startswith("AS"):
            return asn
        elif asn.isdigit():
            return f"AS{asn}"
        return asn
    
    def _format_email(self, email: Optional[str]) -> str:
        """Format email address for display."""
        if not email:
            return "Unknown"
        return html.escape(email.replace("@", " at "))
    
    def _build_html(self, data: Dict[str, Any]) -> str:
        """Build the final HTML report."""
        domain_age_warning = ""
        if data["domain_age_flag"]:
            domain_age_warning = (f"<tr><td colspan='2' style='color: #FF0000; font-weight: bold;'>"
                                 f"{html.escape(data['domain_age_flag'])}</td></tr>")
        
        html_content = f"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>Email Security Analysis</title>
    <style>
        .symbol {{ font-family: 'Segoe UI', Arial, sans-serif, 'Apple Color Emoji', 
                 'Segoe UI Emoji', 'Segoe UI Symbol' !important; }}
        table {{ line-height: 1.2 !important; }}
        td {{ padding: 2px 3px 2px 0 !important; }}
        .header {{ text-align: center; }}
        .date {{ text-align: right; }}
        table, th, td {{ border-collapse: collapse; padding: 2px 3px; }}
    </style>
</head>
<body style="margin: 0; padding: 5px; background-color: #FFFFFF; 
    font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; 
    color: #333333; line-height: 1.2;">
    <table style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; width: 100%;">
        <tr>
            <td colspan="2" class="header" style="background-color: {data['risk_color']}; 
                color: white; padding: 2px; font-weight: bold;">
                EMAIL SECURITY ANALYSIS - RISK: {data['risk_level']} ({data['risk_score']}/100)
            </td>
        </tr>
        {domain_age_warning}
        <tr><td style="padding: 2px 5px 2px 0;">Sender:</td><td>{data['sender_email']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">Domain:</td><td>{data['sender_domain']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">IP:</td><td>{data['ip_display']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">Location:</td><td>{data['location']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">Organization:</td><td>{data['organization']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">A.S.N:</td><td>{data['asn']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">Registrar:</td><td>{data['whois']['registrar']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">Created:</td><td>{data['whois']['created']} ({data['whois']['age']})</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">Auth Type:</td><td>{data['auth_type']}</td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">SPF:</td><td><span style='color: {data['auth']['spf'][1]};' class='symbol'>{data['auth']['spf'][0]}</span><span style='color: {data['auth']['spf'][1]};'>{data['auth']['spf'][2]}</span></td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">DKIM:</td><td><span style='color: {data['auth']['dkim'][1]};' class='symbol'>{data['auth']['dkim'][0]}</span><span style='color: {data['auth']['dkim'][1]};'>{data['auth']['dkim'][2]}</span></td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">DMARC:</td><td><span style='color: {data['auth']['dmarc'][1]};' class='symbol'>{data['auth']['dmarc'][0]}</span><span style='color: {data['auth']['dmarc'][1]};'>{data['auth']['dmarc'][2]}</span></td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">CompAuth:</td><td><span style='color: {data['auth']['compauth'][1]};' class='symbol'>{data['auth']['compauth'][0]}</span><span style='color: {data['auth']['compauth'][1]};'>{data['auth']['compauth'][2]}</span></td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">ARC:</td><td><span style='color: {data['auth']['arc'][1]};' class='symbol'>{data['auth']['arc'][0]}</span><span style='color: {data['auth']['arc'][1]};'>{data['auth']['arc'][2]}</span></td></tr>
        <tr><td style="padding: 2px 5px 2px 0;">Security Flags:</td><td></td></tr>
        <tr><td style="padding: 2px 5px 2px 20px;"> </td><td>{data['security_flags']}</td></tr>
        <tr><td colspan="2" class="date" style="padding: 2px 0;">Analysis: {data['analysis_time']}</td></tr>
    </table>
</body>
</html>"""
        return html_content

class EmailSecurityAnalyzer:
    """Main analyzer class that coordinates all services."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.header_parser = HeaderParser()
        self.auth_parser = AuthenticationParser()
        self.geo_service = GeoLocationService()
        self.whois_service = WhoisService()
        self.security_assessor = SecurityAssessor()
        self.report_generator = HTMLReportGenerator()
    
    def process_headers(self, headers: str) -> str:
        """Process email headers and generate security report."""
        try:
            sender_email, sender_domain = self.header_parser.extract_sender_info(headers)
            sender_ip, sender_hostname = self.header_parser.extract_original_sender_ip(headers)
            
            auth_info = self.auth_parser.extract_authentication_info(headers)
            
            geo_result = GeoLocationResult(ip="Unknown", city="No IP Found")
            if sender_ip:
                geo_result = self.geo_service.geolocate_ip(sender_ip)
            
            whois_info = {"error": "Lookup not performed"}
            if sender_domain:
                whois_info = self.whois_service.get_whois_info(sender_domain)
            
            security_assessment = self.security_assessor.assess_security_risk(
                [geo_result], auth_info, whois_info
            )
            
            return self.report_generator.generate_report(
                sender_email, sender_domain, sender_ip, sender_hostname,
                geo_result, whois_info, auth_info, security_assessment
            )
            
        except Exception as e:
            self.logger.critical(f"Failed to process headers: {e}", exc_info=True)
            return self.report_generator.generate_report(
                None, None, None, None,
                GeoLocationResult(ip="Unknown", city="Processing Failed"),
                {"error": f"Processing Failed: {str(e)}"},
                AuthenticationInfo(),
                SecurityAssessment(
                    score=0,
                    risk_level="HIGH",
                    risk_color="#F44336",
                    factors=["Processing Failed"]
                )
            )

def check_dependencies():
    """Check if all required dependencies are installed."""
    logger = logging.getLogger(__name__)
    
    required_packages = {
        "dns.resolver": "dnspython",
        "whois": "python-whois",
        "pytz": "pytz",
        "requests": "requests"
    }
    
    for module, package in required_packages.items():
        try:
            __import__(module)
            logger.info(f"Dependency {package} is installed.")
        except ImportError:
            logger.error(f"Dependency '{package}' is not installed.")
            raise

def main(input_file: str, output_file: str, ip: Optional[str] = None):
    """Main entry point."""
    logger = logging.getLogger(__name__)
    
    tz = pytz.timezone(CONFIG["timezone"])
    current_time = datetime.now(tz).strftime('%Y-%m-%d %H:%M BST')
    logger.info(f"Script started at {current_time}")
    
    try:
        input_path = input_file
        output_path = output_file
        logger.info(f"Processing: {input_path} -> {output_path}")
        
        if not os.path.exists(input_path):
            logger.error(f"Input file does not exist: {input_path}")
            raise FileNotFoundError(f"Input file missing: {input_path}")
        
        check_dependencies()
        
        headers = ""
        encodings = ['utf-16', 'utf-8', 'latin-1', 'iso-8859-1']
        for encoding in encodings:
            try:
                with open(input_path, 'r', encoding=encoding) as f:
                    headers = f.read()
                    logger.info(f"Successfully read headers with {encoding} encoding")
                    logger.debug(f"Header line count: {len(headers.splitlines())}")
                    break
            except Exception as e:
                logger.warning(f"Failed to read with {encoding}: {e}")
                if encoding == encodings[-1]:
                    logger.error(f"Failed to read input file {input_path} with any encoding")
                    raise
        
        analyzer = EmailSecurityAnalyzer()
        if ip:
            logger.info(f"Using provided IP: {ip}")
            analyzer.header_parser.extract_original_sender_ip = lambda h: (ip, None)
        html_output = analyzer.process_headers(headers)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_output)
        logger.info(f"HTML footer generated: {output_path}")
    
    except Exception as e:
        logger.critical(f"Script failed: {str(e)}", exc_info=True)
        try:
            error_html = f"""<!DOCTYPE html>
<html>
<head><title>Email Security Analysis Error</title></head>
<body>
<table style='font-family: "Segoe UI", Arial, sans-serif; font-size: 13px; width: 100%;'>
<tr><td colspan='2' style='background-color: #F44336; color: white; padding: 2px; font-weight: bold;'>
EMAIL SECURITY ANALYSIS - ERROR</td></tr>
<tr><td>Error:</td><td>{html.escape(str(e))}</td></tr>
</table>
</body>
</html>"""
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(error_html)
            logger.info(f"Fallback error HTML written to: {output_path}")
        except Exception as write_e:
            logger.error(f"Failed to write fallback output: {write_e}")
        raise
    
    logger.info("Script finished successfully.")

if __name__ == "__main__":
    LoggingSetup.setup()
    
    if len(sys.argv) not in (3, 4):
        logging.error(f"Invalid arguments: {sys.argv}")
        print("Usage: python geolocate_headers.py <input_file> <output_file> [ip]")
        sys.exit(1)
    
    ip = sys.argv[3] if len(sys.argv) == 4 else None
    main(sys.argv[1], sys.argv[2], ip)