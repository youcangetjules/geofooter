#!/usr/bin/env python3
import re

test_header = """ARC-Authentication-Results: i=1; mx.google.com;
       dkim=pass header.i=@community.3ds.com header.s=dassaultsystemes header.b="tCX/g8YD";
       spf=pass (google.com: domain of bounce@community.3ds.com designates 130.248.153.124 as permitted sender) smtp.mailfrom=bounce@community.3ds.com;
       dmarc=pass (p=REJECT sp=REJECT dis=NONE) header.from=3ds.com"""

print("Testing authentication header parsing...")
print("=" * 60)

auth_pattern = r'^(?:ARC-)?Authentication-Results:\s*([^\n]+(?:\n\s+[^\n]+)*)'
match = re.search(auth_pattern, test_header, re.IGNORECASE | re.MULTILINE)

if match:
    content = match.group(1).replace('\n', ' ').replace('\r', ' ').strip()
    print('[OK] Found Authentication-Results header')
    print(f'Content: {content[:100]}...')
    print()
    
    spf_match = re.search(r'spf=(\w+)', content, re.IGNORECASE)
    dkim_match = re.search(r'dkim=(\w+)', content, re.IGNORECASE)
    dmarc_match = re.search(r'dmarc=(\w+)', content, re.IGNORECASE)
    
    print(f'SPF: {spf_match.group(1) if spf_match else "NOT FOUND"}')
    print(f'DKIM: {dkim_match.group(1) if dkim_match else "NOT FOUND"}')
    print(f'DMARC: {dmarc_match.group(1) if dmarc_match else "NOT FOUND"}')
else:
    print('[ERROR] NO MATCH for Authentication-Results header')

print()
print("=" * 60)
print("Testing actual header file...")
print("=" * 60)

# Test with an actual header file
import os
import glob

header_files = sorted(glob.glob('C:/GeoFooter/headers/headers_*.txt'), key=os.path.getmtime, reverse=True)
if header_files:
    with open(header_files[0], 'r', encoding='utf-8') as f:
        headers = f.read()
    
    print(f"Reading: {os.path.basename(header_files[0])}")
    
    match = re.search(auth_pattern, headers, re.IGNORECASE | re.MULTILINE)
    if match:
        content = match.group(1).replace('\n', ' ').replace('\r', ' ').strip()
        print('[OK] Found Authentication-Results in actual file')
        print(f'Content (first 200 chars): {content[:200]}...')
        print()
        
        spf_match = re.search(r'spf=(\w+)', content, re.IGNORECASE)
        dkim_match = re.search(r'dkim=(\w+)', content, re.IGNORECASE)
        dmarc_match = re.search(r'dmarc=(\w+)', content, re.IGNORECASE)
        
        print(f'SPF: {spf_match.group(1) if spf_match else "NOT FOUND"}')
        print(f'DKIM: {dkim_match.group(1) if dkim_match else "NOT FOUND"}')
        print(f'DMARC: {dmarc_match.group(1) if dmarc_match else "NOT FOUND"}')
    else:
        print('[ERROR] NO MATCH for Authentication-Results in actual file')
        print()
        # Try to find any auth-related lines
        for line in headers.split('\n')[:100]:
            if 'auth' in line.lower() or 'spf' in line.lower() or 'dkim' in line.lower():
                print(f"  Found line: {line[:80]}")

