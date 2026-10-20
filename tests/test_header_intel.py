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


def _carto_hops():
    return [
        {"fqdn": "MjE3MzIzODk", "by_host": "geopod-ismtpd-11", "asn_org": "Unknown", "asn": "Unknown"},
        {"fqdn": "o1.ptr8966.carto.com", "by_host": "mx.google.com", "asn": "AS11377", "asn_org": "SendGrid, Inc"},
    ]


def test_sendgrid_account_from_api_hop():
    hops = _carto_hops()
    sec = hi.esp_fingerprint(CARTO, hops, tenants={"trusted": [], "blocked": []})
    assert sec.data["esp"] == "sendgrid"
    assert sec.data["tenant"] == "21732389"
    assert "21732389" in hops[0]["decoded_location"]
    assert hops[0]["asn_org"] == "SendGrid (HTTP API)"
    assert sec.findings == []


def test_blocked_tenant_scores():
    sec = hi.esp_fingerprint(CARTO, _carto_hops(), tenants={"trusted": [], "blocked": ["sendgrid:21732389"]})
    assert [f["points"] for f in sec.findings] == [25]


def test_trusted_tenant_marks_esp_hops():
    hops = _carto_hops()
    hi.esp_fingerprint(CARTO, hops, tenants={"trusted": ["sendgrid:21732389"], "blocked": []})
    assert hops[1].get("esp_trusted") is True


def test_b64_digits_rejects_words():
    assert hi._b64_digits("MjE3MzIzODk") == "21732389"
    assert hi._b64_digits("mail") == ""


def test_dual_signing_alignment_and_policy_from_headers():
    sec = hi.signing_signals(CARTO, lite=True)
    text = "\n".join(sec.rows)
    assert "Signed by carto.com and sendgrid.info" in text
    assert "d=carto.com s=s1 a=rsa-sha256 (strict with From)" in text
    assert "d=sendgrid.info s=smtpapi a=rsa-sha256 (not aligned with From)" in text
    assert "SPF domain em8966.carto.com (relaxed with From)" in text
    assert "p=reject" in text
    assert sec.findings == []


def test_unsigned_subject_scores():
    h = (
        "From: A <a@alpha.com>\n"
        "DKIM-Signature: v=1; a=rsa-sha256; d=alpha.com; s=k1; h=to:date; bh=x; b=y\n"
    )
    sec = hi.signing_signals(h)
    assert sec.findings[0]["points"] == 5
    assert "from, subject" in sec.findings[0]["label"]


def _rsa_spki_b64(bits: int) -> str:
    import base64

    n = (1 << (bits - 1)) | 1
    nb = n.to_bytes(bits // 8, "big")

    def tlv(tag, body):
        ln = len(body)
        if ln < 128:
            return bytes([tag, ln]) + body
        lb = ln.to_bytes((ln.bit_length() + 7) // 8, "big")
        return bytes([tag, 0x80 | len(lb)]) + lb + body

    rsa = tlv(0x30, tlv(0x02, b"\x00" + nb) + tlv(0x02, b"\x01\x00\x01"))
    algo = tlv(0x30, bytes.fromhex("06092a864886f70d0101010500"))
    return base64.b64encode(tlv(0x30, algo + tlv(0x03, b"\x00" + rsa))).decode()


def test_dns_key_bits_and_dmarc_on_full_scan():
    zone = {
        "s1._domainkey.alpha.com": [f"v=DKIM1; k=rsa; p={_rsa_spki_b64(512)}"],
        "_dmarc.alpha.com": ["v=DMARC1; p=none; rua=mailto:x@alpha.com"],
    }
    h = (
        "From: A <a@alpha.com>\nSubject: hi\n"
        "DKIM-Signature: v=1; a=rsa-sha256; d=alpha.com; s=s1; h=from:subject; bh=x; b=y\n"
    )
    sec = hi.signing_signals(h, lite=False, resolver=lambda name: zone.get(name, []))
    text = "\n".join(sec.rows)
    assert "RSA 512-bit" in text
    assert "monitoring only" in text
    assert [f["points"] for f in sec.findings] == [5]
    assert hi._der_rsa_bits(__import__("base64").b64decode(_rsa_spki_b64(2048))) == 2048


def test_compact_scan_does_no_dns():
    calls = []
    hi.signing_signals(CARTO, lite=True, resolver=lambda n: calls.append(n) or [])
    assert calls == []


def test_parse_tls_formats():
    assert hi.parse_tls("(version=TLS1_3 cipher=TLS_AES_256_GCM_SHA384 bits=256/256)") == (
        "TLS 1.3", "TLS_AES_256_GCM_SHA384")
    assert hi.parse_tls("(using TLSv1.2 with cipher ECDHE-RSA-AES256-GCM-SHA384 (256/256 bits))") == (
        "TLS 1.2", "ECDHE-RSA-AES256-GCM-SHA384")
    assert hi.parse_tls(
        "with Microsoft SMTP Server (version=TLS1_0, cipher=TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA)"
    )[0] == "TLS 1.0"
    assert hi.parse_tls("with esmtps (TLS1.3) tls TLS_AES_256_GCM_SHA384")[0] == "TLS 1.3"
    assert hi.parse_tls("by mx with SMTP id 12") == ("", "")


def test_transport_scores_weak_and_plain_external_hops():
    hops = [
        {"hop_kind": "public", "protocol": "ESMTPS", "tls_version": "TLS 1.0", "tls": "TLS 1.0"},
        {"hop_kind": "public", "protocol": "SMTP", "fqdn": "mail.bad-host.example"},
        {"hop_kind": "public", "protocol": "SMTP", "fqdn": None},
        {"hop_kind": "internal", "protocol": "SMTP", "fqdn": "relay.local"},
        {"hop_kind": "public", "protocol": "ESMTPS", "tls_version": "TLS 1.3", "tls": "TLS 1.3"},
    ]
    assert sorted(f["points"] for f in hi.transport_signals(hops).findings) == [3, 5]


def test_client_kit_and_windows_host():
    h = "Message-ID: <123@DESKTOP-7Q2K9LM>\nX-Mailer: PHPMailer 5.2.22 (https://github.com/PHPMailer/PHPMailer)\n"
    assert sorted(f["points"] for f in hi.client_signals(h).findings) == [5, 10]
    assert hi.client_signals("Message-ID: <a@mail.gmail.com>\nX-Mailer: Microsoft Outlook 16.0\n").findings == []


def test_bulk_classification_and_unsubscribe_urls():
    sec = hi.bulk_signals(CARTO)
    assert sec.data["kind"] == "Bulk / mailing list"
    assert sec.data["unsubscribe_urls"] == ["https://u21732389.ct.sendgrid.net/unsub?x=1"]
    assert any("one-click" in r for r in sec.rows)
    assert hi.bulk_signals("Auto-Submitted: auto-generated\n").data["kind"] == "Automated"
    assert hi.bulk_signals("From: a@b.com\n").data["kind"] == "Personal"


def _history(days_old=30, seen_ids=(), senders=None):
    from datetime import datetime, timedelta

    stamp = (datetime.now() - timedelta(days=days_old)).strftime("%Y-%m-%d %H:%M:%S")
    base = {"old@friend.com": {"count": 3, "first_seen": stamp}}
    base.update(senders or {})
    return {"senders": base, "seen": {hi.message_id_key(i): stamp for i in seen_ids}}


HIJACK = (
    "From: Accounts <accounts@new-vendor.com>\n"
    "Subject: RE: Invoice PR-14673 due today\n"
    "Message-ID: <fresh1@new-vendor.com>\n"
    "In-Reply-To: <fabricated123@aliniant.com>\n"
    "References: <fabricated123@aliniant.com>\n"
)


def test_unseen_thread_from_first_time_sender_scores():
    sec = hi.thread_signals(HIJACK, sender_email="accounts@new-vendor.com", history=_history())
    assert [f["points"] for f in sec.findings] == [15]


def test_seen_thread_is_fine():
    sec = hi.thread_signals(
        HIJACK, sender_email="accounts@new-vendor.com",
        history=_history(seen_ids=["<fabricated123@aliniant.com>"]),
    )
    assert sec.findings == []


def test_known_contact_or_repeat_sender_is_fine():
    assert hi.thread_signals(
        HIJACK, sender_email="accounts@new-vendor.com", history=_history(),
        known_contacts=["accounts@new-vendor.com"],
    ).findings == []
    h = _history(senders={"accounts@new-vendor.com": {"count": 2, "first_seen": "2026-09-01 10:00:00"}})
    assert hi.thread_signals(HIJACK, sender_email="accounts@new-vendor.com", history=h).findings == []


def test_rescan_still_scores_first_time_sender():
    h = _history(
        seen_ids=["<fresh1@new-vendor.com>"],
        senders={"accounts@new-vendor.com": {"count": 1, "first_seen": "2026-10-20 10:00:00"}},
    )
    sec = hi.thread_signals(HIJACK, sender_email="accounts@new-vendor.com", history=h)
    assert [f["points"] for f in sec.findings] == [15]


def test_fake_re_without_thread_headers():
    h = "From: x@cold.com\nSubject: Re: our call\nMessage-ID: <a@cold.com>\n"
    sec = hi.thread_signals(h, sender_email="x@cold.com", history=_history())
    assert [f["points"] for f in sec.findings] == [10]


def test_young_history_is_report_only():
    sec = hi.thread_signals(HIJACK, sender_email="accounts@new-vendor.com", history=_history(days_old=3))
    assert sec.findings == []
    assert any("report-only" in r for r in sec.rows)


def test_return_path_mismatch_scored_only_without_dmarc_pass():
    h = "Return-Path: <x@bounce.other.com>\nFrom: A <a@alpha.com>\n"
    assert hi.identity_signals(h, dmarc="pass").findings == []
    assert [f["points"] for f in hi.identity_signals(h, dmarc="fail").findings] == [5]
