"""The inbox scrape must read headers through a table, not open every message."""

import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import guri.outlook_scraper as sc  # noqa: E402

NOW = datetime.now().replace(microsecond=0)
SINCE = NOW - timedelta(days=14)


class FakeColumns:
    def __init__(self, table):
        self.table = table

    def RemoveAll(self):
        self.table.names = []

    def Add(self, name):
        if name in self.table.refuse:
            raise RuntimeError("cannot add " + name)
        self.table.names.append(name)


class FakeTable:
    def __init__(self, mails, refuse=()):
        self.mails, self.refuse, self.names = list(mails), set(refuse), []
        self.pos, self.sorted, self.calls = 0, None, 0
        self.Columns = FakeColumns(self)

    def Sort(self, prop, descending):
        self.sorted = (prop, descending)
        self.mails.sort(key=lambda m: m["ReceivedTime"], reverse=bool(descending))

    @property
    def EndOfTable(self):
        return self.pos >= len(self.mails)

    def GetArray(self, count):
        self.calls += 1
        page = self.mails[self.pos:self.pos + count]
        self.pos += len(page)
        return tuple(tuple(m.get(n) for n in self.names) for m in page)


class FakeFolder:
    def __init__(self, table):
        self.table, self.filters = table, []

    def GetTable(self, filter_text, contents):
        self.filters.append(filter_text)
        return self.table


def mail(age_hours, **over):
    base = {
        "ReceivedTime": NOW - timedelta(hours=age_hours),
        "Subject": f"subject {age_hours}",
        "MessageClass": "IPM.Note",
        "UnRead": False,
        "To": "me@example.com",
        "CC": "",
        "Importance": 1,
        "FlagStatus": 0,
        "ConversationTopic": "topic",
        "SenderEmailAddress": "sender@example.org",
        "SenderName": "Sender",
        sc._PR_SENDER_SMTP: "sender@example.org",
        sc._PR_CONVERSATION_ID: memoryview(bytes([0xAB]) * 16),
        sc._PR_LONGTERM_ENTRYID: memoryview(bytes([age_hours % 256, 0xEF]) * 35),
    }
    base.update(over)
    return base


CFG = sc.AesAccountScanConfig(store_id="S", smtp="me@example.com")


def rows(table, max_scan=500):
    return list(sc._iter_inbox_rows_since(FakeFolder(table), SINCE, max_scan=max_scan))


def convert(row, unread_only=False):
    return sc._mail_from_row(row, store_id="S", account_smtp="me@example.com", cfg=CFG,
                             since=SINCE, unread_only=unread_only)


def test_rows_come_newest_first_from_an_unfiltered_sorted_table():
    table = FakeTable([mail(5), mail(1), mail(3)])
    folder = FakeFolder(table)
    got = list(sc._iter_inbox_rows_since(folder, SINCE, max_scan=500))
    assert folder.filters == [""]                      # a date filter scans the whole folder
    assert table.sorted == ("[ReceivedTime]", True)
    assert [r["Subject"] for r in got] == ["subject 1", "subject 3", "subject 5"]


def test_reading_stops_at_the_first_row_older_than_the_window():
    table = FakeTable([mail(h) for h in range(1, 200)] + [mail(24 * 20 + h) for h in range(300)])
    assert len(rows(table)) == 199
    assert table.calls == 4                            # 4 pages of 50, not the whole folder


def test_scan_cap_bounds_the_work():
    table = FakeTable([mail(h % 100 + 1) for h in range(1000)])
    assert len(rows(table, max_scan=120)) == 120


def test_missing_optional_column_is_tolerated_but_entry_id_is_required():
    table = FakeTable([mail(1)], refuse={sc._PR_CONVERSATION_ID})
    got = convert(rows(table)[0])
    assert got is not None and got.conversation_id == ""
    try:
        rows(FakeTable([mail(1)], refuse={sc._PR_LONGTERM_ENTRYID}))
    except RuntimeError:
        pass
    else:
        raise AssertionError("a table without entry ids must raise so the caller falls back")


def test_row_becomes_the_same_mail_an_item_would():
    got = convert(mail(2, UnRead=True, Importance=2, FlagStatus=2))
    assert got.entry_id == (bytes([2, 0xEF]) * 35).hex().upper() and len(got.entry_id) == 140
    assert got.conversation_id == "AB" * 16
    assert (got.sender, got.subject, got.unread, got.importance, got.flag_status) == (
        "sender@example.org", "subject 2", True, 2, 2)
    assert got.recipient_role == "to"
    assert got.received == (NOW - timedelta(hours=2)).strftime("%Y-%m-%d %H:%M:%S")
    assert got.body_preview == "" and got.folder == "Inbox"


def test_exchange_sender_uses_the_smtp_column():
    row = mail(1, SenderEmailAddress="/O=EXCHANGELABS/OU=X/CN=RECIPIENTS/CN=abc")
    assert convert(row).sender == "sender@example.org"
    row[sc._PR_SENDER_SMTP] = None
    row["SenderName"] = "Someone"
    assert convert(row).sender.startswith("/O=EXCHANGELABS")


def test_filters_match_the_item_path():
    assert convert(mail(1, MessageClass="REPORT.IPM.Note.NDR")) is None
    assert convert(mail(1, MessageClass="IPM.Schedule.Meeting.Request")).is_meeting_request
    assert convert(mail(1), unread_only=True) is None
    cc_row = mail(1, To="other@example.com", CC="me@example.com")
    assert convert(cc_row).recipient_role == "cc"
    no_cc = sc.AesAccountScanConfig(store_id="S", smtp="me@example.com", in_cc=False)
    assert sc._mail_from_row(cc_row, store_id="S", account_smtp="me@example.com", cfg=no_cc,
                             since=SINCE, unread_only=False) is None
    assert convert(mail(24 * 30)) is None


def _app_with_inbox(inbox):
    class Store:
        StoreID = "S"

        def GetDefaultFolder(self, _kind):
            return inbox

    class Stores:
        Count = 1

        def Item(self, _i):
            return Store()

    class Ns:
        pass

    Ns.Stores = Stores()

    class App:
        def GetNamespace(self, _name):
            return Ns()

    return App()


def test_scrape_uses_the_table_and_never_walks_items(monkeypatch):
    inbox = FakeFolder(FakeTable([mail(h) for h in range(1, 6)]))
    inbox.Items = object()

    def no_walk(*_a, **_k):
        raise AssertionError("the slow item walk ran")

    monkeypatch.setattr(sc, "_outlook_app", lambda: _app_with_inbox(inbox))
    monkeypatch.setattr(sc, "_iter_inbox_since", no_walk)
    got = sc._scrape_outlook_inboxes_impl(days=14, max_per_account=3, account_configs=[CFG])
    assert len(got) == 3 and all(len(m.entry_id) == 140 for m in got)


def test_scrape_falls_back_to_items_when_the_table_fails(monkeypatch):
    walked = []

    class Inbox:
        Items = object()

        def GetTable(self, *_a):
            raise RuntimeError("no table on this store")

    def walk(items, since, max_take):
        walked.append(max_take)
        return iter(())

    monkeypatch.setattr(sc, "_outlook_app", lambda: _app_with_inbox(Inbox()))
    monkeypatch.setattr(sc, "_iter_inbox_since", walk)
    assert sc._scrape_outlook_inboxes_impl(days=14, max_per_account=3, account_configs=[CFG]) == []
    assert walked == [3]
