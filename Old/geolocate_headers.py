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
import random
import string
import email.utils
import pycountry
from ipwhois import IPWhois, exceptions as ipwhois_exceptions

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
        
        tz = pytz.utc
        
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
                self.logger.debug(f"DEBUG: Extracted sender: {sender_email}, domain: {sender_domain}")
                return sender_email, sender_domain
        self.logger.warning("Could not extract sender from headers")
        self.logger.debug("DEBUG: Could not extract sender from headers")
        return None, None
    
    def extract_original_sender_ip(self, headers: str) -> Tuple[Optional[str], Optional[str]]:
        """Extract the original sender's public IP and hostname."""
        self.logger.info("Extracting original sender IP and hostname")

        # Preprocess: join continuation lines (lines starting with whitespace)
        lines = headers.splitlines()
        joined_lines = []
        for line in lines:
            if line.startswith((' ', '\t')) and joined_lines:
                joined_lines[-1] += ' ' + line.strip()
            else:
                joined_lines.append(line.strip())
        headers_flat = '\n'.join(joined_lines)

        # Pattern for: Received: from <hostname> ([<ip>])
        pattern_bracket = r'^Received:\s*from\s+([^\s]+)\s*\(\[([\d\.:a-fA-F]+)\]\)'  # e.g., from M1 ([51.179.200.191])
        match_bracket = re.findall(pattern_bracket, headers_flat, re.IGNORECASE | re.MULTILINE)
        for hostname, ip in reversed(match_bracket):
            self.logger.debug(f"Processing (bracket): Hostname={hostname}, IP={ip}")
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                self.logger.debug(f"DEBUG: Found public IP: {ip} (Hostname: {hostname})")
                return ip, hostname

        # Existing patterns
        pattern = (r'^Received:\s*from\s+([^\s\(]+)\s*'
                  r'(?:\([^)]*\))?\s*'
                  r'(?:by\s+[^\s]+\s*)?'
                  r'(?:\([^)]*\[([^\]]+)\][^)]*\))?')
        simple_pattern = r'^Received:\s*from\s+([^\s]+)\s+\(([^\[]+)\[([^\]]+)\]\)'

        simple_headers = re.findall(simple_pattern, headers_flat, re.IGNORECASE | re.MULTILINE)
        for hostname, hostname2, ip in reversed(simple_headers):
            self.logger.debug(f"Processing (simple): Hostname={hostname}, IP={ip}")
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                self.logger.debug(f"DEBUG: Found public IP: {ip} (Hostname: {hostname})")
                return ip, hostname

        received_headers = re.findall(pattern, headers_flat, re.IGNORECASE | re.MULTILINE)
        for hostname, ip in reversed(received_headers):
            if ip:
                self.logger.debug(f"Processing: Hostname={hostname}, IP={ip}")
                if IPValidator.is_valid_public(ip):
                    self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                    self.logger.debug(f"DEBUG: Found public IP: {ip} (Hostname: {hostname})")
                    return ip, hostname

        # Fallback: Try to extract sender IP from Authentication-Results header
        auth_results_pattern = r'^Authentication-Results:.*?(sender IP is ([\d\.]+))'
        auth_results_match = re.search(auth_results_pattern, headers_flat, re.IGNORECASE | re.MULTILINE)
        if auth_results_match:
            ip = auth_results_match.group(2)
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP in Authentication-Results: {ip}")
                self.logger.debug(f"DEBUG: Found public IP in Authentication-Results: {ip}")
                return ip, None

        self.logger.warning("No valid public sender IP found")
        self.logger.debug("DEBUG: No valid public sender IP found")
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
                self.logger.debug(f"DEBUG: SPF: {auth_info.spf}")
            else:
                self.logger.debug("DEBUG: SPF not found")
            
            dkim_match = re.search(r'dkim=(\w+)', auth_content, re.IGNORECASE)
            if dkim_match:
                auth_info.dkim = dkim_match.group(1).lower()
                self.logger.debug(f"DEBUG: DKIM: {auth_info.dkim}")
            else:
                self.logger.debug("DEBUG: DKIM not found")
            
            dmarc_match = re.search(r'dmarc=(\w+)', auth_content, re.IGNORECASE)
            if dmarc_match:
                auth_info.dmarc = dmarc_match.group(1).lower()
                if auth_info.dmarc == "bestguesspass":
                    auth_info.dmarc = "pass"
                self.logger.debug(f"DEBUG: DMARC: {auth_info.dmarc}")
            else:
                self.logger.debug("DEBUG: DMARC not found")
            
            compauth_match = re.search(r'compauth=(\w+)(?:\s+reason=(\d+))?', auth_content, re.IGNORECASE)
            if compauth_match:
                auth_info.compauth = {
                    "result": compauth_match.group(1).lower(),
                    "reason_code": compauth_match.group(2) if compauth_match.group(2) else "unknown"
                }
                self.logger.debug(f"DEBUG: CompAuth: {auth_info.compauth}")
            else:
                self.logger.debug("DEBUG: CompAuth not found")
        else:
            self.logger.warning("No Authentication-Results header found")
            self.logger.debug("DEBUG: No Authentication-Results header found")
        
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
                # Set each flag individually, do not combine
                result.security_data["Tor"] = abuse_data.get("isTor")
                result.security_data["Proxy"] = abuse_data.get("isProxy")
                result.security_data["VPN"] = abuse_data.get("isVpn")
                result.security_data["Anonymous"] = abuse_data.get("isAnonymous")
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
            
            # ASN extraction
            if "asn" in data and isinstance(data["asn"], dict):
                result.asn = data["asn"].get("asn", "Unknown") or "Unknown"
            else:
                # Try to extract ASN from org string if present
                org_str = result.org or ""
                match = re.search(r"AS\d+", org_str)
                if match:
                    result.asn = match.group(0)
                else:
                    result.asn = "Unknown"
            
            # Only set security flags if not already set by AbuseIPDB
            privacy = data.get("privacy", {})
            for flag in ["VPN", "Proxy", "Tor", "Anonymous"]:
                if result.security_data[flag] is None:
                    result.security_data[flag] = privacy.get(flag.lower())
            
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
                    # ASN extraction
                    asn_val = data.get("as", "Unknown") or "Unknown"
                    if asn_val != "Unknown":
                        # asn_val may be like 'AS5482 AllPoints Fibre Limited'
                        match = re.search(r"AS\d+", asn_val)
                        if match:
                            result.asn = match.group(0)
                        else:
                            result.asn = asn_val
                    else:
                        # Try to extract ASN from org string if present
                        org_str = result.org or ""
                        match = re.search(r"AS\d+", org_str)
                        if match:
                            result.asn = match.group(0)
                        else:
                            result.asn = "Unknown"
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
        score = 0
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
        """Assess authentication methods (reversed scoring: higher is worse)."""
        if auth_info.spf:
            if "pass" in auth_info.spf.lower():
                factors.append("SPF Validated")
                score += 0
            elif "fail" in auth_info.spf.lower():
                factors.append(f"SPF {auth_info.spf.upper()}")
                score += 15
            else:
                factors.append(f"SPF {auth_info.spf.upper()}")
                score += 7
        else:
            factors.append("SPF Missing/Problem")
            score += 7
        
        if auth_info.dkim:
            if "pass" in auth_info.dkim.lower():
                factors.append("DKIM Validated")
                score += 0
            elif "fail" in auth_info.dkim.lower():
                factors.append(f"DKIM {auth_info.dkim.upper()}")
                score += 15
            else:
                factors.append(f"DKIM {auth_info.dkim.upper()}")
                score += 7
        else:
            factors.append("DKIM Missing/Problem")
            score += 7
        
        if auth_info.dmarc:
            if "pass" in auth_info.dmarc.lower():
                factors.append("DMARC Compliant")
                score += 0
            elif "fail" in auth_info.dmarc.lower():
                factors.append(f"DMARC {auth_info.dmarc.upper()}")
                score += 20
            else:
                factors.append(f"DMARC {auth_info.dmarc.upper()}")
                score += 10
        else:
            factors.append("DMARC Missing/Problem")
            score += 10
        
        return score
    
    def _assess_geolocation(
        self,
        geo_results: List[GeoLocationResult],
        score: int,
        factors: List[str]
    ) -> int:
        """Assess geolocation and security flags (reversed scoring: higher is worse)."""
        if not geo_results:
            return score
        
        origin_geo = geo_results[0]
        
        if origin_geo.success and origin_geo.blocklist and "Listed: " in origin_geo.blocklist:
            blocklists = origin_geo.blocklist.replace('Listed: ', '')
            factors.append(f"Origin IP Blocklisted ({blocklists})")
            score = 100
            return score
        elif origin_geo.success:
            factors.append("Origin IP Not Blocklisted")
        
        if origin_geo.success:
            any_flag_active = False
            no_data_count = 0
            for flag_name, is_active in origin_geo.security_data.items():
                if is_active is None:
                    no_data_count += 1
                    continue
                elif is_active:
                    factors.append(f"Security Flag: {flag_name} Active")
                    any_flag_active = True
            if any_flag_active:
                score += 10
            if no_data_count > 0:
                score += 5 * no_data_count
        
        return score
    
    def _assess_domain_age(
        self,
        whois_info: Dict[str, Any],
        score: int,
        factors: List[str]
    ) -> Tuple[int, Optional[str]]:
        """Assess domain age (reversed scoring: higher is worse)."""
        domain_age_flag = None
        
        if not whois_info or "error" in whois_info:
            if whois_info and whois_info.get("error"):
                factors.append("Sender Domain WHOIS Error")
                score += 5
            return score, domain_age_flag
        
        creation_date = whois_info.get("creation_date", "Unknown")
        if creation_date == "Unknown":
            return score, domain_age_flag
        
        try:
            parsed_date = self._parse_date(creation_date)
            if not parsed_date:
                factors.append("Domain Creation Date Invalid")
                score += 10
                return score, domain_age_flag
            
            parsed_date = pytz.utc.localize(parsed_date)
            now = datetime.now(pytz.utc)
            age_days = (now - parsed_date).days

            if age_days < 0:
                factors.append("Domain Creation Date Invalid")
                score += 10
            elif age_days < 10:
                factors.append(f"Domain Extremely New ({age_days} days)")
                domain_age_flag = "Domain Less Than 10 Days Old - CAUTION"
                score = 100
            elif age_days < 30:
                factors.append(f"Domain Very New ({age_days} days)")
                score += 50
            elif age_days < 100:
                factors.append(f"Domain New ({age_days} days)")
                score += 30
            elif age_days < 365:
                factors.append(f"Domain Moderately New ({age_days} days)")
                score += 10
            else:
                factors.append(f"Domain Established ({age_days} days)")
                score += 0
                
        except Exception as e:
            self.logger.error(f"Error processing domain age: {e}")
            factors.append("Domain Age Processing Error")
            score += 5
        
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
        """Calculate risk level based on reversed score (higher is worse)."""
        if domain_age_flag:
            return "HIGH"
        
        thresholds = CONFIG["risk_thresholds"]
        if score >= thresholds["low"]:
            return "HIGH"
        elif score >= thresholds["medium"]:
            return "MEDIUM"
        else:
            return "LOW"

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
        security_assessment: SecurityAssessment,
        headers: Optional[str] = None
    ) -> str:
        """Generate HTML security report with a colorful, email-safe, table-based layout."""
        display_data = self._prepare_display_data(
            sender_email, sender_domain, sender_ip,
            sender_hostname, geo_result, whois_info,
            auth_info, security_assessment,
            headers
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
        security_assessment: SecurityAssessment,
        headers: Optional[str] = None
    ) -> Dict[str, Any]:
        """Prepare all data for display."""
        tz = pytz.utc
        analysis_time = datetime.now(tz).strftime('%Y-%m-%d %H:%M BST')
        ip_to_display = sender_ip if sender_ip else geo_result.ip
        hostname_to_display = sender_hostname if sender_ip else None
        ip_display = self._format_ip_display(ip_to_display, hostname_to_display)
        location_display = self._format_location(geo_result)
        whois_display = self._format_whois(whois_info)
        auth_display = self._format_authentication(auth_info)
        security_flags = self._format_security_flags(geo_result)
        org_display = geo_result.org
        if org_display and isinstance(org_display, str):
            org_display = re.sub(r"AS\d+\s*", "", org_display).strip()
        if not geo_result.success:
            org_display = "Unknown"
        def _auth_type_string() -> str:
            passed = []
            for meth, val in [("SPF", auth_info.spf), ("DKIM", auth_info.dkim), ("DMARC", auth_info.dmarc)]:
                if val and val.lower() == "pass":
                    passed.append(meth)
            if len(passed) == 3:
                return "SPF + DKIM + DMARC (All Passed)"
            if passed:
                return " + ".join(passed) + " (Partial)"
            return "None"
        auth_type_str = _auth_type_string()
        # Calculate hops and ToT
        hops, tot = self._calculate_hops_and_tot(headers)
        # Routing info: hostnames and ASNs traversed
        routing_info = self._extract_routing_info(headers)
        per_hop_analysis = self._extract_hop_details(headers)
        return {
            "analysis_time": analysis_time,
            "risk_level": security_assessment.risk_level,
            "risk_score": security_assessment.score,
            "risk_color": security_assessment.risk_color,
            "domain_age_flag": security_assessment.domain_age_flag,
            "sender_email": self._format_email(sender_email),
            "sender_domain": sender_domain or "Unknown",
            "ip_display": ip_display,
            "hops_and_tot": f"{hops} hops; ToT: {tot}",
            "routing_info": routing_info,
            "location": location_display,
            "organization": org_display,
            "asn": self._format_asn(geo_result.asn),
            "whois": whois_display,
            "auth": auth_display,
            "security_flags": security_flags,
            "auth_type": auth_type_str,
            "factors": security_assessment.factors,
            "per_hop_analysis": per_hop_analysis,
        }
    
    def _calculate_hops_and_tot(self, headers: Optional[str]) -> tuple:
        if not headers:
            return ("N/A", "N/A")
        received_lines = re.findall(r'^Received:.*', headers, re.MULTILINE)
        hops = len(received_lines)
        # Extract dates from Received headers
        date_pattern = r';\s*(.+)$'
        times = []
        for line in received_lines:
            m = re.search(date_pattern, line)
            if m:
                try:
                    dt = email.utils.parsedate_to_datetime(m.group(1))
                    times.append(dt)
                except Exception:
                    continue
        if len(times) >= 2:
            tot_seconds = abs((max(times) - min(times)).total_seconds())
            tot_str = f"{tot_seconds:.2f}s"
        else:
            tot_str = "N/A"
        return (hops, tot_str)
    
    def _format_ip_display(self, ip: Optional[str], hostname: Optional[str]) -> str:
        """Format IP address display, always appending terminal/hostname if available and not identical to IP."""
        if not ip:
            return "Unknown"
        ip_str = html.escape(ip)
        if hostname and hostname.strip() and hostname.lower() != ip.lower():
            return f"{ip_str} ({html.escape(hostname.strip())})"
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
        # Ensure we always pass a string, never None
        creation_date_str = str(whois_info.get("creation_date", "Unknown") or "Unknown")
        age_days = self._calculate_domain_age(creation_date_str)
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
            comp_disp = f"{comp_res.capitalize()} (reason {reason}: {reason_txt})"
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
                # Orange for No Data
                flags.append(f"<span style='color: #FFA500;' class='symbol'>- Is {flag_name} - <span style='color: #FFA500;'>No Data</span></span>")
            elif is_active:
                # Red cross for Yes (Detected)
                flags.append(f"<span style='color: #FF0000;' class='symbol'>&#10007;</span> Is {flag_name} - <span style='color: #FF0000;'>Yes</span>")
            else:
                # Green tick for No (Not Detected)
                flags.append(f"<span style='color: #008000;' class='symbol'>&#10003;</span> Is {flag_name} - <span style='color: #008000;'>No</span>")
        
        blocklist_green = False
        if geo_result.blocklist and geo_result.blocklist not in ["N/A (IPv6/Invalid)", "N/A (Invalid IP)"]:
            if "Listed: " in geo_result.blocklist:
                blocklists = geo_result.blocklist.replace("Listed: ", "")
                flags.append(f"<span style='color: #FF0000;' class='symbol' title='Blocklisted'>&#10007;</span> <span style='color: #FF0000;'>Blocklisted ({html.escape(blocklists)})</span>")
            else:
                blocklist_green = True
                flags.append(f"<span style='color: #008000;' class='symbol' title='Not Blocklisted'>&#10003;</span> <span style='color: #008000;'>Not Blocklisted</span>")
        
        # If not blocklisted, make the whole answer green
        if blocklist_green:
            return f"<div style='color: #008000;'>{'<br>'.join(flags)}</div>"
        else:
            return "<br>".join(flags) if flags else "<span style='color: #FFA500;'>- None</span>"
    
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
        """Build a simple, compact, email-safe HTML footer in a 2-column layout, with a colored title banner for the risk score. All rows from Sender to ARC are equally vertically spaced."""
        risk = data["risk_level"]
        score = data["risk_score"]
        def risk_gradient(score):
            s = max(0, min(100, score))
            r = int(76 + (244-76)*s/100)
            g = int(175 + (67-175)*s/100)
            b = int(80 + (54-80)*s/100)
            return f'rgb({r},{g},{b})'
        risk_bg = risk_gradient(score)
        # Security flags as left-aligned rows
        security_flags_lines = data["security_flags"].split('<br>')
        sec_flags_html = ""
        for flag in security_flags_lines:
            sec_flags_html += f"<div style='margin-bottom:2px;'>{flag}</div>"
        # Main data fields as label-value rows
        fields = [
            ("Sender", data['sender_email']),
            ("Domain", data['sender_domain']),
            ("IP", data['ip_display']),
            ("Hops/ToT", data['hops_and_tot']),
            ("Route", data['routing_info']),
            ("Location", data['location']),
            ("Organization", data['organization']),
            ("A.S.N", data['asn']),
            ("Registrar", data['whois']['registrar']),
            ("Created", f"{data['whois']['created']} ({data['whois']['age']})"),
            ("Auth Type", data['auth_type']),
            ("SPF", f"<span style='color: {data['auth']['spf'][1]}; font-family:Arial;'>{data['auth']['spf'][0]}</span><span style='color: {data['auth']['spf'][1]};'>{data['auth']['spf'][2]}</span>"),
            ("DKIM", f"<span style='color: {data['auth']['dkim'][1]}; font-family:Arial;'>{data['auth']['dkim'][0]}</span><span style='color: {data['auth']['dkim'][1]};'>{data['auth']['dkim'][2]}</span>"),
            ("DMARC", f"<span style='color: {data['auth']['dmarc'][1]}; font-family:Arial;'>{data['auth']['dmarc'][0]}</span><span style='color: {data['auth']['dmarc'][1]};'>{data['auth']['dmarc'][2]}</span>"),
            ("CompAuth", f"<span style='color: {data['auth']['compauth'][1]}; font-family:Arial;'>{data['auth']['compauth'][0]}</span><span style='color: {data['auth']['compauth'][1]};'>{data['auth']['compauth'][2]}</span> <span style='color: #888888;'>{data['auth']['compauth_display']}</span>"),
            ("ARC", f"<span style='color: {data['auth']['arc'][1]}; font-family:Arial;'>{data['auth']['arc'][0]}</span><span style='color: {data['auth']['arc'][1]};'>{data['auth']['arc'][2]}</span>"),
        ]
        # Use tighter vertical spacing for all rows from Sender to ARC
        row_style = "padding:2px 4px 2px 0; line-height:1.1; vertical-align:middle;"
        rows = ""
        for label, value in fields:
            rows += f"<tr><td style='{row_style} font-weight:bold; font-size:11px; color:#333;'>{label}</td>"
            rows += f"<td style='{row_style} font-size:11px; color:#333;'>{value}</td></tr>"
        # Security flags row (default spacing)
        rows += f"<tr><td style='padding:1px 4px 1px 0; font-weight:bold; font-size:11px; color:#333;'>Security Flags</td>"
        rows += f"<td style='padding:1px 4px 1px 0; font-size:11px; color:#333;'>{sec_flags_html}</td></tr>"
        # Per-Hop Analysis Table
        per_hop_html = ""
        if data.get('per_hop_analysis'):
            per_hop_html += "<div style='font-weight:bold; font-size:11px; margin:8px 0 2px 0;'>Route Table</div>"
            per_hop_html += "<table cellpadding='0' cellspacing='0' style='width:100%; border-collapse:collapse; font-size:10px;'><tr>"
            headers = [
                "#", "Host (Decoded)", "IP", "ASN", "Organisation", "Prefix", "RIR", "Country/City", "Reverse DNS", "Reputation", "Timestamp", "Delay", "BGP Peers"
            ]
            for h in headers:
                per_hop_html += f"<th style='border-bottom:1px solid #ccc; padding:2px 4px; text-align:left; font-size:10px; color:#333; font-weight:600;'>{h}</th>"
            per_hop_html += "</tr>"
            for idx, hop in enumerate(data['per_hop_analysis'], 1):
                # Determine row color
                abuse_score = None
                if hop['reputation'] and 'Score:' in hop['reputation']:
                    m_score = re.search(r'Score: (\d+)', hop['reputation'])
                    if m_score:
                        try:
                            abuse_score = int(m_score.group(1))
                        except Exception:
                            abuse_score = None
                blocklisted = False
                if hop.get('ip') and hop['ip'] != 'Unknown':
                    geo = GeoLocationService().geolocate_ip(hop['ip'])
                    if geo.blocklist and 'Listed:' in geo.blocklist:
                        blocklisted = True
                if blocklisted or (abuse_score is not None and abuse_score > 10):
                    row_bg = '#F44336'  # Red
                    row_color = '#fff'
                    row_weight = '600'
                elif abuse_score is not None and abuse_score <= 10:
                    row_bg = '#4CAF50'  # Green
                    row_color = '#fff'
                    row_weight = '600'
                else:
                    row_bg = '#FFC107'  # Amber
                    row_color = '#333'
                    row_weight = '400'
                per_hop_html += f"<tr style='background:{row_bg}; color:{row_color}; font-weight:{row_weight};'>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{idx}</td>"
                host_disp = hop['fqdn'] or ''
                if hop['decoded_location'] and hop['decoded_location'] != host_disp:
                    host_disp += f"<br><span style='color:{row_color}; font-size:9px; font-weight:{row_weight};'>({hop['decoded_location']})</span>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{host_disp}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['ip']}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['asn']}</td>"
                # Force clean the organization name
                asn_org_disp = hop['asn_org']
                if asn_org_disp and asn_org_disp != 'Unknown':
                    # Direct cleaning using the exact logic from test script
                    cleaned_org = asn_org_disp
                    # Remove AS numbers
                    cleaned_org = re.sub(r'AS\d+\s*', '', cleaned_org, flags=re.IGNORECASE)
                    # Remove common suffixes
                    suffixes = [' Limited', ' Ltd', ' LLC', ' Inc', ' Corporation', ' Corp', ' Company', ' Co']
                    for suffix in suffixes:
                        if cleaned_org.endswith(suffix):
                            cleaned_org = cleaned_org[:-len(suffix)]
                            break
                    cleaned_org = cleaned_org.strip(' ,.-')
                    if not cleaned_org:
                        cleaned_org = 'Unknown'
                    asn_org_disp = f"<span style='color:{row_color}; font-size:9px; font-weight:{row_weight};'>{cleaned_org}</span>"
                else:
                    asn_org_disp = 'Unknown'
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{asn_org_disp}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['prefix']}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['rir']}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['country']}<br><span style='color:{row_color}; font-size:9px; font-weight:{row_weight};'>{hop['city']}</span></td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['reverse_dns']}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['reputation']}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['timestamp']}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{hop['delay']}</td>"
                per_hop_html += f"<td style='padding:2px 4px; color:{row_color}; font-weight:{row_weight};'>{', '.join(hop['bgp_peers']) if hop['bgp_peers'] else ''}</td>"
                per_hop_html += "</tr>"
            per_hop_html += "</table>"
            # GURI, copyright, datetime at the bottom
            def randhex(n):
                return ''.join(random.choices('0123456789abcdef', k=n))
            guri = f"{randhex(5)}x{randhex(5)}x{randhex(8)}x{randhex(8)}x{randhex(3)}x{randhex(2)}"
            tz = pytz.utc
            now = datetime.now(tz)
            created_zulu_str = now.strftime('%Y %m %d %H:%M') + ' Zulu'
            copyright_year = now.strftime('%Y')
            per_hop_html += (
                "<table style='width:100%; margin-top:6px; font-size:11px; color:#333; border-collapse:collapse;' cellpadding='0' cellspacing='0'>"
                "<tr>"
                f"<td style='text-align:left; padding:2px 4px;'>GURI: {guri}</td>"
                f"<td style='text-align:center; padding:2px 4px;'>(C) Aliniant Labs {copyright_year}</td>"
                f"<td style='text-align:right; padding:2px 4px;'>Created: {created_zulu_str}</td>"
                "</tr></table>"
            )
        html_content = f"""<!DOCTYPE html>
<html><body style='margin:0; padding:2px; background:#fff; font-family:Arial, sans-serif; font-size:11px; color:#333;'>
<table cellpadding='0' cellspacing='0' style='width:100%; border-collapse:collapse; font-size:11px;'>
<tr><td colspan='2' style='background:{risk_bg}; color:#fff; font-weight:bold; font-size:13px; padding:3px 4px 3px 4px; text-align:center;'>EMAIL SECURITY ANALYSIS - RISK {risk} ({score}/100)</td></tr>
{rows}
</table>
{per_hop_html}
</body></html>"""
        return html_content

    def _decode_pod_name(self, segment: str) -> str:
        # Example: SWEP280, GBRP123
        m = re.match(r'^([A-Z]{2,3})([A-Z])P(\d+)$', segment)
        if not m:
            return segment  # Return as-is if not matching
        country_code, region_letter, pod_number = m.groups()
        # Try to map country code to country name
        country = pycountry.countries.get(alpha_2=country_code[:2])
        country_name = country.name if country else country_code
        return f"{country_name} Pod {pod_number}"

    def _extract_routing_info(self, headers: Optional[str]) -> str:
        if not headers:
            return "N/A"
        received_lines = re.findall(r'^Received:.*', headers, re.MULTILINE)
        hops = []
        for line in received_lines:
            # Extract all possible IPs (IPv4/IPv6)
            ips = re.findall(r'\[([\d\.:a-fA-F]+)\]', line)
            host = None
            m_host = re.search(r'from ([^ ]+)', line)
            decoded_location = None
            if m_host:
                host = m_host.group(1)
                # Try to decode pod/region code from the second label in the FQDN
                parts = host.split('.')
                if len(parts) > 1:
                    decoded_location = self._decode_pod_name(parts[1])
            asn = "Unknown"
            found_public_ip = False
            ip_used = None
            for ip in ips:
                if IPValidator.is_valid_public(ip):
                    found_public_ip = True
                    ip_used = ip
                    geo = GeoLocationService().geolocate_ip(ip)
                    if geo.asn and geo.asn != "Unknown":
                        asn = geo.asn
                        break
                    # Try whois lookup for ASN if geolocation fails
                    try:
                        import whois
                        w = whois.whois(ip)
                        if hasattr(w, 'asn') and w.asn:
                            asn = str(w.asn)
                            break
                        for field in ['org', 'netname', 'descr']:
                            val = getattr(w, field, None)
                            if val and isinstance(val, str) and 'AS' in val.upper():
                                m_asn = re.search(r'AS(\d+)', val.upper())
                                if m_asn:
                                    asn = f"AS{m_asn.group(1)}"
                                    break
                    except Exception:
                        pass
            # If no public IP found, try to resolve hostname to IP and use BGPView
            if not found_public_ip and host:
                try:
                    ip_from_host = socket.gethostbyname(host)
                    if IPValidator.is_valid_public(ip_from_host):
                        ip_used = ip_from_host
                        # Query BGPView API
                        try:
                            resp = requests.get(f'https://api.bgpview.io/ip/{ip_from_host}', timeout=5)
                            if resp.status_code == 200:
                                data = resp.json()
                                asn_data = data.get('data', {}).get('prefixes', [])
                                if asn_data:
                                    asn = asn_data[0].get('asn', {}).get('asn', 'Unknown')
                                    if asn != 'Unknown':
                                        asn = f"AS{asn}"
                        except Exception:
                            pass
                except Exception:
                    pass
            # Compose display string
            if not (found_public_ip or (ip_used and asn != "Unknown")):
                if decoded_location:
                    hops.append(f"{host or 'Unknown'} ({decoded_location}, Unknown)")
                else:
                    hops.append(f"{host or 'Unknown'} (Unknown)")
            else:
                if decoded_location:
                    hops.append(f"{host or 'Unknown'} ({decoded_location}, {asn})")
                else:
                    hops.append(f"{host or 'Unknown'} ({asn})")
        return "; ".join(hops) if hops else "N/A"

    def _ipwhois_lookup(self, ip: str) -> dict:
        try:
            data = IPWhois(ip).lookup_rdap(depth=1)
            asn_desc = data.get("asn_description")
            asn_desc_str = str(asn_desc) if asn_desc is not None else ""
            nir_val = data.get("nir", data.get("asn_registry", "—"))
            nir_str = str(nir_val) if nir_val is not None else "—"
            return {
                "asn": data.get("asn") or "n/a",
                "org": asn_desc_str.split(" | ")[0],
                "prefix": data.get("asn_cidr", "—"),
                "rir": nir_str.upper(),
            }
        except (ipwhois_exceptions.IPDefinedError, ipwhois_exceptions.HTTPLookupError, Exception):
            return {"asn": "— (private/bogon)", "org": "", "prefix": "", "rir": ""}

    def _extract_hop_details(self, headers: Optional[str]) -> list:
        import requests
        import socket
        import email.utils
        if not headers:
            return []
        received_lines = re.findall(r'^Received:.*', headers, re.MULTILINE)
        hops = []
        prev_time = None
        for line in received_lines:
            hop = {
                'fqdn': None,
                'decoded_location': None,
                'ip': None,
                'asn': 'Unknown',
                'asn_org': 'Unknown',
                'prefix': '—',
                'rir': '—',
                'country': 'Unknown',
                'city': 'Unknown',
                'reverse_dns': 'Unknown',
                'reputation': 'Unknown',
                'timestamp': 'Unknown',
                'delay': 'N/A',
                'bgp_peers': [],  # Always present, default empty list
            }
            # Hostname
            m_host = re.search(r'from ([^ ]+)', line)
            if m_host:
                hop['fqdn'] = m_host.group(1)
                parts = hop['fqdn'].split('.')
                if len(parts) > 1:
                    hop['decoded_location'] = self._decode_pod_name(parts[1])
            # IP
            ips = re.findall(r'\[([\d\.:a-fA-F]+)\]', line)
            ip_used = None
            for ip in ips:
                if IPValidator.is_valid_public(ip):
                    ip_used = ip
                    break
            if not ip_used and hop['fqdn']:
                try:
                    ip_from_host = socket.gethostbyname(hop['fqdn'])
                    if IPValidator.is_valid_public(ip_from_host):
                        ip_used = ip_from_host
                except Exception:
                    pass
            hop['ip'] = ip_used or 'Unknown'
            # ASN, ASN Org, Country, City, Prefix, RIR
            if ip_used:
                geo = GeoLocationService().geolocate_ip(ip_used)
                hop['asn'] = geo.asn or 'Unknown'
                hop['asn_org'] = geo.org or 'Unknown'
                hop['country'] = geo.country or 'Unknown'
                hop['city'] = geo.city or 'Unknown'
                # ASN enrichment with ipwhois
                ipwhois_info = self._ipwhois_lookup(ip_used)
                if hop['asn'] == 'Unknown' and ipwhois_info['asn']:
                    hop['asn'] = ipwhois_info['asn']
                if hop['asn_org'] == 'Unknown' and ipwhois_info['org']:
                    hop['asn_org'] = ipwhois_info['org']
                hop['prefix'] = ipwhois_info.get('prefix', '—')
                hop['rir'] = ipwhois_info.get('rir', '—')
            # Reverse DNS
            if ip_used:
                try:
                    hop['reverse_dns'] = socket.gethostbyaddr(ip_used)[0]
                except Exception:
                    hop['reverse_dns'] = 'Unknown'
            # Reputation (AbuseIPDB)
            if ip_used and ip_used != 'Unknown':
                try:
                    headers_abuse = {"Key": CONFIG["api_keys"]["abuseipdb"], "Accept": "application/json"}
                    resp = requests.get(f"https://api.abuseipdb.com/api/v2/check?ipAddress={ip_used}&maxAgeInDays=90", headers=headers_abuse, timeout=5)
                    if resp.status_code == 200:
                        data = resp.json().get('data', {})
                        hop['reputation'] = f"Score: {data.get('abuseConfidenceScore', 'N/A')}"
                except Exception:
                    hop['reputation'] = 'Unknown'
            # Timestamp
            m_time = re.search(r';\s*(.+)$', line)
            if m_time:
                try:
                    dt = email.utils.parsedate_to_datetime(m_time.group(1))
                    hop['timestamp'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                    if prev_time:
                        delay = (dt - prev_time).total_seconds()
                        hop['delay'] = f"{delay:.2f}s"
                    prev_time = dt
                except Exception:
                    hop['timestamp'] = 'Unknown'
            hops.append(hop)
        return hops

    def _strip_subdomains_until_resolvable(self, hostname: str) -> str:
        """Strip subdomains from left to right until a resolvable domain is found."""
        if not hostname or '.' not in hostname:
            return hostname
        
        parts = hostname.split('.')
        if len(parts) < 2:
            return hostname
        
        # Try from shortest to longest (right to left)
        for i in range(len(parts) - 1, 0, -1):
            test_domain = '.'.join(parts[i:])
            try:
                socket.gethostbyname(test_domain)
                return test_domain
            except (socket.gaierror, socket.herror):
                continue
        
        # If nothing resolves, return the original
        return hostname

    def _clean_organization_name(self, org_name: str) -> str:
        """Clean organization name by removing any AS numbers and common suffixes."""
        if not org_name or org_name == 'Unknown':
            return org_name
        
        original = org_name
        self.logger.debug(f"Cleaning org name: '{original}'")
        
        # Remove 'AS' followed by numbers and optional whitespace at the start
        org_name = re.sub(r'^AS\d+\s*', '', org_name, flags=re.IGNORECASE)
        self.logger.debug(f"After removing start AS: '{org_name}'")
        
        # Remove 'AS' followed by numbers anywhere else in the string
        org_name = re.sub(r'AS\d+', '', org_name, flags=re.IGNORECASE)
        self.logger.debug(f"After removing all AS: '{org_name}'")
        
        # Remove common suffixes
        suffixes_to_remove = [
            ' Limited', ' Ltd', ' LLC', ' Inc', ' Corporation', ' Corp', ' Company', ' Co',
            ' International', ' Intl', ' Technologies', ' Tech', ' Networks', ' Network',
            ' Communications', ' Comm', ' Services', ' Service', ' Solutions', ' Solution', ' LLC.', ' Inc.'
        ]
        for suffix in suffixes_to_remove:
            if org_name.endswith(suffix):
                org_name = org_name[:-len(suffix)]
                self.logger.debug(f"After removing suffix '{suffix}': '{org_name}'")
                break
        
        # Clean up extra whitespace and punctuation
        org_name = org_name.strip(' ,.-')
        self.logger.debug(f"Final cleaned result: '{org_name}'")
        
        return org_name if org_name else 'Unknown'

    def generate_outlook_signature(self, data: Dict[str, Any]) -> str:
        """Generate a simplified, Outlook-compatible HTML signature."""
        score = data['security_assessment']['score']
        risk = data['security_assessment']['risk_level']
        risk_bg = data['security_assessment']['risk_color']
        
        # Simple text-based signature for Outlook
        signature = f"""
<div style="font-family: Arial, sans-serif; font-size: 10px; color: #333; border-top: 1px solid #ccc; padding-top: 5px; margin-top: 10px;">
<div style="background-color: {risk_bg}; color: white; font-weight: bold; font-size: 11px; padding: 2px 4px; text-align: center; margin-bottom: 5px;">
EMAIL SECURITY ANALYSIS - RISK {risk} ({score}/100)
</div>
<div style="margin-bottom: 3px;"><strong>Sender:</strong> {data['sender_email']}</div>
<div style="margin-bottom: 3px;"><strong>Domain:</strong> {data['sender_domain']}</div>
<div style="margin-bottom: 3px;"><strong>IP:</strong> {data['sender_ip']}</div>
<div style="margin-bottom: 3px;"><strong>Location:</strong> {data['location']}</div>
<div style="margin-bottom: 3px;"><strong>Organization:</strong> {data['organization']}</div>
<div style="margin-bottom: 3px;"><strong>Auth:</strong> {data['auth_type']}</div>
<div style="margin-bottom: 3px;"><strong>SPF:</strong> {data['spf'][1] if isinstance(data['spf'], tuple) else data['spf']}</div>
<div style="margin-bottom: 3px;"><strong>DKIM:</strong> {data['dkim'][1] if isinstance(data['dkim'], tuple) else data['dkim']}</div>
<div style="margin-bottom: 3px;"><strong>DMARC:</strong> {data['dmarc'][1] if isinstance(data['dmarc'], tuple) else data['dmarc']}</div>
<div style="margin-bottom: 3px;"><strong>Blocklist:</strong> {data['security_flags']}</div>
<div style="font-size: 9px; color: #666; margin-top: 5px; border-top: 1px solid #eee; padding-top: 3px;">
GURI: {data.get('guri', 'N/A')} | (C) Aliniant Labs 2025 | Created: {data.get('created', 'N/A')}
</div>
</div>
"""
        return signature

def main():
    """Main execution function."""
    import sys
    import os
    
    # Setup logging
    LoggingSetup.setup()
    logger = logging.getLogger(__name__)
    
    if len(sys.argv) != 2:
        print("Usage: python geolocate_headers.py <header_file>")
        sys.exit(1)
    
    header_file = sys.argv[1]
    if not os.path.exists(header_file):
        print(f"Error: Header file '{header_file}' not found.")
        sys.exit(1)
    
    try:
        # Read header file
        with open(header_file, 'r', encoding='utf-8') as f:
            headers = f.read()
        
        logger.info(f"Processing header file: {header_file}")
        
        # Initialize components
        header_parser = HeaderParser()
        auth_parser = AuthenticationParser()
        geo_service = GeoLocationService()
        whois_service = WhoisService()
        security_assessor = SecurityAssessor()
        html_generator = HTMLReportGenerator()
        
        # Extract sender information
        sender_email, sender_domain = header_parser.extract_sender_info(headers)
        sender_ip, sender_hostname = header_parser.extract_original_sender_ip(headers)
        
        logger.info(f"Sender: {sender_email} ({sender_domain})")
        logger.info(f"IP: {sender_ip} ({sender_hostname})")
        
        # Geolocate sender IP
        geo_result = geo_service.geolocate_ip(sender_ip) if sender_ip else GeoLocationResult(ip="Unknown")
        
        # Get WHOIS information
        whois_info = whois_service.get_whois_info(sender_domain) if sender_domain else {}
        
        # Extract authentication information
        auth_info = auth_parser.extract_authentication_info(headers)
        
        # Assess security risk
        security_assessment = security_assessor.assess_security_risk(
            [geo_result], auth_info, whois_info
        )
        
        # Generate HTML report
        html_content = html_generator.generate_report(
            sender_email, sender_domain, sender_ip, sender_hostname,
            geo_result, whois_info, auth_info, security_assessment, headers
        )
        
        # Ensure output directory exists
        os.makedirs('C:/GeoFooter/output', exist_ok=True)
        
        # Write HTML report
        output_file = 'C:/GeoFooter/output/output_report.html'
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        logger.info(f"Report generated successfully: {output_file}")
        print(f"Report generated: {output_file}")
        
    except Exception as e:
        logger.error(f"Error processing headers: {e}")
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()