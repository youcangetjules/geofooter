"""One footer per mail: run the VBA footer detection itself, under cscript.

The functions are lifted out of ``VBA\\MSCANModule1.bas`` and mechanically
turned into VBScript (type suffixes and ``As`` clauses dropped), so this
exercises the shipped logic, not a copy of it.
"""

import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
pytestmark = pytest.mark.skipif(
    sys.platform != "win32" or not shutil.which("cscript"), reason="needs Windows Script Host")

FUNCTIONS = [
    "FindAesFooterBounds", "StripBlockBetween", "FindAesHeading", "CountAesHeadings",
    "FindAesAnchorTag", "CountAesStartAnchors", "CountBlocksBetween", "CountAesScanResults",
    "StripAllAesScanResults", "EnclosingTableStart", "RemoveTableStartingAt",
    "RemoveFooterTableAt", "HtmlFragmentIsBlank", "RemoveOrphanFooterAnchors",
]
CONSTANTS = "VISIBLE_HEADING|ANCHOR_START|ANCHOR_END|STRIP_COLOUR|PLAIN_START|PLAIN_END"

DRIVER = r'''
Function ReadUtf8(p)
    Dim s: Set s = CreateObject("ADODB.Stream")
    s.Type = 2: s.Charset = "utf-8": s.Open: s.LoadFromFile p
    ReadUtf8 = s.ReadText: s.Close
End Function
Sub WriteUtf8(p, text)
    Dim s: Set s = CreateObject("ADODB.Stream")
    s.Type = 2: s.Charset = "utf-8": s.Open: s.WriteText text: s.SaveToFile p, 2: s.Close
End Sub
Dim body, stripped
body = ReadUtf8(WScript.Arguments(0))
stripped = StripAllAesScanResults(body)
WScript.Echo CountAesScanResults(body) & " " & CountAesScanResults(stripped)
WriteUtf8 WScript.Arguments(1), stripped
'''


def _vbscript() -> str:
    src = (ROOT / "VBA" / "MSCANModule1.bas").read_bytes().decode("ascii").replace("\r\n", "\n")
    parts = [f"Const AES_{name} = {value}" for name, value in re.findall(
        rf"^Private Const AES_({CONSTANTS}) As String = (.*)$", src, re.M)]
    assert len(parts) == 6
    for name in FUNCTIONS:
        found = re.search(rf"^Private (Function|Sub) {name}\(.*?^End \1$", src, re.M | re.S)
        assert found, name
        parts.append(found.group(0))
    code = "\n\n".join(parts)
    code = re.sub(r"^\s*'.*$", "", code, flags=re.M)
    code = re.sub(r"Optional ByVal (\w+) As \w+ = \w+", r"ByVal \1", code)
    code = re.sub(r"\((\d+) To (\d+)\)", r"(\2)", code)
    code = re.sub(r"\(0 To ([^)]+)\)", r"(\1)", code)
    code = re.sub(r"\s+As\s+(String|Long|Boolean|Object)\b", "", code)
    code = re.sub(r"\b(Left|Mid|Right|LCase|Trim|ChrW)\$\(", r"\1(", code)
    code = re.sub(r"^(\s*)Next \w+$", r"\1Next", code, flags=re.M)
    code = code.replace("On Error GoTo EH", "On Error Resume Next")
    code = re.sub(r"^EH:\s*$", "", code, flags=re.M)
    code = code.replace("Private Function", "Function").replace("Private Sub", "Sub")
    assert " As " not in code
    return "Option Explicit\n" + code + "\n" + DRIVER


@pytest.fixture(scope="module")
def strip(tmp_path_factory):
    work = tmp_path_factory.mktemp("vba")
    script = work / "footer.vbs"
    script.write_text(_vbscript(), encoding="ascii")

    def run(html: str):
        src, out = work / "in.html", work / "out.html"
        src.write_text(html, encoding="utf-8")
        done = subprocess.run(["cscript", "//nologo", str(script), str(src), str(out)],
                              capture_output=True, text=True, timeout=60)
        assert done.returncode == 0 and not done.stderr.strip(), done.stdout + done.stderr
        before, after = (int(n) for n in done.stdout.split())
        return before, after, out.read_text(encoding="utf-8-sig")

    return run


HEADING = "Aliniant&nbsp;AES&nbsp;Scan&nbsp;Result&nbsp;-&nbsp;Basic&nbsp;Scan"

# What the scanner writes.
FRESH = (
    "<!-- AES Start -->\n<a id='aes-footer-start' name='aes-footer-start'></a>\n"
    "<div style='background:#ffffff;'><table bgcolor='#46525c' style='background:#46525c;'>"
    f"<tr><td><table><tr><td><b>{HEADING}</b></td></tr></table></td></tr></table></div>\n"
    "<a id='aes-footer-end' name='aes-footer-end'></a>\n<!-- AES End -->"
)

# The same footer after the mail was replied to in Outlook and came back:
# no comments, no end anchor, start anchor respelled inside a spacer table.
QUOTED = (
    '<table class="MsoNormalTable" style="background:white"><tbody><tr><td>'
    '<p class="MsoNormal"><a name="aes-footer-start"></a><span>&nbsp;</span></p>'
    "</td></tr></tbody></table>\n"
    '<p class="MsoNormal"><span style="display:none">&nbsp;</span></p>\n'
    '<table class="MsoNormalTable" width="100%" style="width:100.0%;background:#46525C">'
    "<tbody><tr><td><table><tbody><tr><td>logo</td></tr></tbody></table>"
    f'<p class="MsoNormal"><b><span>{HEADING}</span></b></p>'
    "<table><tbody><tr><td>QUICK ACTIONS</td></tr></tbody></table>"
    "</td></tr></tbody></table>\n"
)


def page(*chunks: str) -> str:
    return "<html><body>" + "".join(chunks) + "</body></html>"


def test_a_fresh_footer_counts_once_and_strips_clean(strip):
    before, after, out = strip(page("<p>hello</p>", FRESH))
    assert (before, after) == (1, 0)
    assert "hello" in out and "46525c" not in out


def test_a_footer_quoted_through_a_reply_is_seen_and_removed(strip):
    before, after, out = strip(page("<p>new text</p><table><tr><td>signature</td></tr></table>",
                                    QUOTED, FRESH))
    assert before == 2                      # the duplicate that used to go unseen
    assert after == 0
    assert "new text" in out and "signature" in out
    assert "aes-footer-start" not in out and "QUICK ACTIONS" not in out and "Aliniant" not in out


def test_several_quoted_footers_all_go(strip):
    before, after, out = strip(page("<p>top</p>", QUOTED, "<p>middle</p>", QUOTED, "<p>tail</p>", FRESH))
    assert (before, after) == (3, 0)
    assert all(word in out for word in ("top", "middle", "tail"))


def test_text_between_a_quoted_footer_and_the_new_one_survives(strip):
    # Comments stripped (IMAP): the quoted start anchor must not pair with
    # the fresh footer's end anchor and swallow what lies between.
    fresh_no_comments = FRESH.replace("<!-- AES Start -->", "").replace("<!-- AES End -->", "")
    before, after, out = strip(page(QUOTED, "<p>keep this reply</p>", fresh_no_comments))
    assert (before, after) == (2, 0)
    assert "keep this reply" in out


def test_a_typed_product_name_does_not_cost_the_sender_their_table(strip):
    typed = "<table><tr><td>We reviewed the Aliniant AES Scan Result today.</td></tr></table>"
    _before, _after, out = strip(page(typed, FRESH))
    assert "We reviewed the" in out


def test_a_table_that_closed_before_the_footer_is_left_alone(strip):
    before, after, out = strip(page("<table><tr><td>sender table</td></tr></table>", QUOTED))
    assert (before, after) == (1, 0)
    assert "sender table" in out


def test_plain_summary_comment_is_not_a_second_footer(strip):
    plain = ("<!-- AES-Plain-Summary-Start -->\nAliniant AES Scan Result - Basic Scan | x\n"
             "<!-- AES-Plain-Summary-End -->\n")
    before, after, _out = strip(page("<p>hi</p>", plain, FRESH))
    assert (before, after) == (1, 0)


def test_body_without_a_footer_is_untouched(strip):
    html = page("<div><table><tr><td>plain mail</td></tr></table></div>")
    before, after, out = strip(html)
    assert (before, after) == (0, 0)
    assert out == html
