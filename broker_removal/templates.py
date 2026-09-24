"""Render GDPR / CCPA / generic opt-out letter and email bodies."""

from __future__ import annotations

from datetime import date
from typing import Any, Dict, Optional, Tuple

from broker_removal.profile import IdentityProfile, load_profile


TEMPLATES = {
    "gdpr": "GDPR Art. 17 erasure",
    "ccpa": "CCPA / CPRA delete request",
    "generic": "Generic data-broker opt-out",
}


def _profile_block(profile: IdentityProfile) -> str:
    lines = []
    if profile.full_name:
        lines.append(f"Full name: {profile.full_name}")
    if profile.aliases:
        lines.append("Also known as: " + "; ".join(profile.aliases))
    if profile.emails:
        lines.append("Email addresses: " + "; ".join(profile.emails))
    if profile.phones:
        lines.append("Phone numbers: " + "; ".join(profile.phones))
    if profile.addresses:
        lines.append("Addresses:")
        for addr in profile.addresses:
            lines.append(f"  - {addr}")
    if profile.date_of_birth:
        lines.append(f"Date of birth: {profile.date_of_birth}")
    return "\n".join(lines) if lines else "(identity profile incomplete — fill in GURI → Data Brokers)"


def _subject(broker_name: str, template_id: str) -> str:
    tid = (template_id or "generic").lower()
    if tid == "gdpr":
        return f"GDPR Article 17 erasure request — {broker_name}"
    if tid == "ccpa":
        return f"CCPA / CPRA deletion request — {broker_name}"
    return f"Personal data removal / opt-out request — {broker_name}"


def _body_gdpr(broker_name: str, profile: IdentityProfile) -> str:
    return f"""{date.today().isoformat()}

To: Privacy / Data Protection team, {broker_name}

Subject: Request for erasure of personal data (GDPR Article 17)

Dear Data Protection Officer / Privacy team,

I am writing to exercise my right to erasure under Article 17 of the EU General Data Protection Regulation (GDPR).

Please permanently delete all personal data you hold about me from your databases, websites, affiliate sites, and any third parties to whom you have disclosed my data. Please also confirm in writing when erasure is complete and identify any recipients to whom my data was disclosed.

My identifying details are:

{_profile_block(profile)}

If you refuse this request in whole or in part, please state the legal basis for refusal and how I may lodge a complaint with a supervisory authority.

Please respond within one month of receipt.

Yours faithfully,
{profile.full_name or "[Your name]"}
"""


def _body_ccpa(broker_name: str, profile: IdentityProfile) -> str:
    return f"""{date.today().isoformat()}

To: Privacy team, {broker_name}

Subject: CCPA / CPRA request to delete personal information

Dear Privacy team,

Under the California Consumer Privacy Act (CCPA) as amended by the CPRA, I request that you delete the personal information you have collected about me, and that you direct your service providers to do the same.

Please also disclose the categories of personal information collected, the sources, the business or commercial purpose, and the categories of third parties with whom it was shared.

My identifying details are:

{_profile_block(profile)}

Please confirm completion of this deletion request in writing.

Sincerely,
{profile.full_name or "[Your name]"}
"""


def _body_generic(broker_name: str, profile: IdentityProfile) -> str:
    return f"""{date.today().isoformat()}

To: Privacy / Opt-out team, {broker_name}

Subject: Request to remove my personal information

Dear Privacy team,

I request that you permanently remove all personal information you hold about me from your people-search / data-broker services, directories, marketing lists, and any sites or partners that republish your data.

Please confirm in writing when the removal is complete, and do not sell or share my information further.

My identifying details are:

{_profile_block(profile)}

Thank you for your prompt attention.

Sincerely,
{profile.full_name or "[Your name]"}
"""


def render_removal(
    broker: Dict[str, Any],
    profile: Optional[IdentityProfile] = None,
    template_id: Optional[str] = None,
) -> Tuple[str, str, str]:
    """Return (template_label, email_subject, body_text)."""
    profile = profile or load_profile()
    name = str(broker.get("name") or "Data broker")
    tid = (template_id or broker.get("template_id") or "generic").strip().lower()
    if tid not in TEMPLATES:
        tid = "generic"
    if tid == "gdpr":
        body = _body_gdpr(name, profile)
    elif tid == "ccpa":
        body = _body_ccpa(name, profile)
    else:
        body = _body_generic(name, profile)
    return TEMPLATES[tid], _subject(name, tid), body.strip() + "\n"
