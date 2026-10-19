"""Same-domain mail shows an internal hop address and is not scored as external."""

from aes.geolocate_headers import (
    AuthenticationInfo,
    GeoLocationResult,
    HeaderParser,
    SecurityAssessor,
)

HEADERS = """
Received: from DU0PR08MB8277.eurprd08.prod.outlook.com (2603:10a6:10:40f::11)
 by VE1PR08MB5600.eurprd08.prod.outlook.com with HTTPS; Mon, 19 Oct 2026 11:58:27 +0000
Received: from AS8PR08MB8610.eurprd08.prod.outlook.com (2603:10a6:20b:564::5)
 by DU0PR08MB8277.eurprd08.prod.outlook.com (2603:10a6:10:40f::11) with Microsoft SMTP Server
Received: from AS8PR08MB8610.eurprd08.prod.outlook.com ([fe80::b3a6:db3e:dd89:52fe])
 by AS8PR08MB8610.eurprd08.prod.outlook.com ([fe80::b3a6:db3e:dd89:52fe%7]) with mapi id
From: Graham Friend <graham.friend@coleago.com>
To: Matt Halfmann <matt.halfmann@coleago.com>
Cc: Mike Johnston <mike.johnston@coleago.com>, Julian Garrett <julian.garrett@coleago.com>
Authentication-Results: mx.microsoft.com 1; dkim=none (message not signed)
 header.d=none;dmarc=none action=none header.from=coleago.com;
"""


def test_same_domain_uses_the_first_real_hop_address():
    parser = HeaderParser()
    email, domain = parser.extract_sender_info(HEADERS)
    assert email == "graham.friend@coleago.com"
    assert domain == "coleago.com"
    assert parser.recipients_share_sender_domain(HEADERS, domain)
    assert parser.extract_original_sender_ip(HEADERS) == (None, None)
    ip, _host = parser.extract_internal_hop_ip(HEADERS)
    assert ip == "2603:10a6:20b:564::5"


def test_same_domain_is_not_scored_like_missing_authentication():
    auth = AuthenticationInfo(dkim="none", dmarc="none")
    outside = SecurityAssessor().assess_security_risk(
        [GeoLocationResult(ip="Unknown")],
        auth,
        {},
    )
    internal = SecurityAssessor().assess_security_risk(
        [GeoLocationResult(ip="2603:10a6:20b:564::5", city="Internal within same domain")],
        auth,
        {},
        internal_same_domain=True,
    )
    assert outside.score >= 17
    assert internal.score == 0
    assert internal.score < outside.score - 10
