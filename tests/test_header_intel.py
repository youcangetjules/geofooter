"""Header intelligence checks against trimmed real-world header shapes."""

from aes import header_intel as hi

CARTO = """\
Return-Path: <bounces+21732389-abcd-julian=aliniant.com@em8966.carto.com>
Received: from o1.ptr8966.carto.com (o1.ptr8966.carto.com. [149.72.58.126])
        by mx.google.com with ESMTPS id x12si
        (version=TLS1_3 cipher=TLS_AES_256_GCM_SHA384 bits=256/256);
        Mon, 19 Oct 2026 09:12:01 -0700 (PDT)
Received: by recvd-6c8f-abc with SMTP id abc; Mon, 19 Oct 2026 16:12:00 +0000 (UTC)
Received: from MjE3MzIzODk (unknown) by geopod-ismtpd-11 (SG) with HTTP id xyz;
        Mon, 19 Oct 2026 16:11:59 +0000 (UTC)
DKIM-Signature: v=1; a=rsa-sha256; c=relaxed/relaxed; d=carto.com;
        h=content-type:from:mime-version:subject:reply-to:to:cc;
        s=s1; bh=abc=; b=def=
DKIM-Signature: v=1; a=rsa-sha256; c=relaxed/relaxed; d=sendgrid.info;
        h=content-type:from:mime-version:subject:to:cc;
        s=smtpapi; bh=abc=; b=def=
Authentication-Results: mx.google.com;
       dkim=pass header.i=@carto.com header.s=s1 header.b=abc;
       dkim=pass header.i=@sendgrid.info header.s=smtpapi header.b=def;
       spf=pass (google.com: domain of bounces@em8966.carto.com designates 149.72.58.126 as permitted sender) smtp.mailfrom="bounces+21732389@em8966.carto.com";
       dmarc=pass (p=REJECT sp=REJECT dis=NONE) header.from=carto.com
From: CARTO <support-basemaps@carto.com>
To: julian.garrett@aliniant.com
Subject: Your basemaps usage
Message-ID: <abc123@geopod-ismtpd-11>
X-SG-EID: u001.abcdefg
X-Feedback-ID: 21732389:SG
List-Unsubscribe: <https://u21732389.ct.sendgrid.net/unsub?x=1>, <mailto:unsub@carto.com>
List-Unsubscribe-Post: List-Unsubscribe=One-Click
"""


def _labels(sec):
    return [f["label"] for f in sec.findings]


def test_clean_esp_mail_has_no_identity_findings():
    sec = hi.identity_signals(CARTO, dmarc="pass")
    assert sec.findings == []
    assert any("Return-Path" in r for r in sec.rows)


def test_reply_to_free_webmail_scores_high():
    h = "From: CFO <cfo@aliniant.com>\nReply-To: cfo.aliniant@gmail.com\n"
    sec = hi.identity_signals(h, dmarc="pass")
    assert [f["points"] for f in sec.findings] == [15]


def test_reply_to_other_domain_scores_ten():
    h = "From: Billing <billing@supplier.com>\nReply-To: billing@supp1ier-pay.com\n"
    assert [f["points"] for f in hi.identity_signals(h).findings] == [10]


def test_display_name_with_other_address():
    h = 'From: "ceo@aliniant.com" <ceo.office@outlook.com>\n'
    sec = hi.identity_signals(h)
    assert sec.findings[0]["points"] == 15


def test_brand_display_name_on_unrelated_domain():
    h = "From: PayPal Security <alerts@secure-acct-review.net>\n"
    assert "paypal" in _labels(hi.identity_signals(h))[0].lower()


def test_brand_on_its_own_domain_is_fine():
    h = "From: Microsoft Teams <noreply@email.teams.microsoft.com>\n"
    assert hi.identity_signals(h).findings == []


def test_own_org_name_on_outside_address():
    h = "From: Aliniant Finance <finance.team@gmail.com>\n"
    sec = hi.identity_signals(h, own_domains=["aliniant.com"])
    assert sec.findings and sec.findings[0]["points"] == 10


def test_return_path_mismatch_scored_only_without_dmarc_pass():
    h = "Return-Path: <x@bounce.other.com>\nFrom: A <a@alpha.com>\n"
    assert hi.identity_signals(h, dmarc="pass").findings == []
    assert [f["points"] for f in hi.identity_signals(h, dmarc="fail").findings] == [5]
