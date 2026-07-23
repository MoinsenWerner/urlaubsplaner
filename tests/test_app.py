from datetime import date

import pytest

import app as vacation_app


@pytest.fixture()
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(vacation_app, "DB_PATH", tmp_path / "test.sqlite3")
    vacation_app.app.config.update(TESTING=True, SECRET_KEY="test")
    vacation_app.init_db()
    return vacation_app.app.test_client()


def login(client, username="admin", password="admin"):
    return client.post(
        "/login",
        data={"username": username, "password": password},
        follow_redirects=True,
    )


def test_admin_seed_and_login(client):
    response = login(client)
    assert response.status_code == 200
    assert b"Mitglieder verwalten" in response.data


def test_admin_can_create_user_with_multiple_roles(client):
    login(client)
    response = client.post(
        "/members",
        data={
            "action": "create",
            "username": "max",
            "first_name": "Max",
            "last_name": "Muster",
            "password": "secret",
            "roles": ["azubi", "normal"],
        },
        follow_redirects=True,
    )
    assert response.status_code == 200
    with vacation_app.db() as conn:
        roles = [
            row["role"]
            for row in conn.execute(
                "SELECT role FROM user_roles WHERE user_id = 2 ORDER BY role"
            )
        ]
    assert roles == ["azubi", "normal"]


def test_role_based_edit_permissions(client):
    login(client)
    client.post(
        "/members",
        data={
            "action": "create",
            "username": "azubi",
            "first_name": "Ada",
            "last_name": "Azubi",
            "password": "pw",
            "roles": ["azubi"],
        },
    )
    client.get("/logout")
    login(client, "azubi", "pw")
    assert (
        client.post(
            "/entry",
            data={
                "user_id": 2,
                "date": date.today().isoformat(),
                "code": "BS",
                "action": "add",
            },
        ).status_code
        == 302
    )
    assert (
        client.post(
            "/entry",
            data={
                "user_id": 2,
                "date": date.today().isoformat(),
                "code": "KR",
                "action": "add",
            },
        ).status_code
        == 403
    )


def test_weekdays_and_band_ranges(client):
    login(client)
    admin = vacation_app.load_user("1")
    start, end = vacation_app.band_range(admin, 2026)
    assert (start, end) == (date(2025, 7, 2), date(2027, 7, 2))
    days = vacation_app.weekdays_between(date(2026, 1, 1), date(2026, 1, 7))
    assert [day.weekday() for day in days] == [3, 4, 0, 1, 2]


def test_bulk_entry_overwrites_and_deletes_selection(client):
    login(client)
    today = date.today().isoformat()
    add = client.post(
        "/bulk-entry",
        json={"code": "UB", "cells": [{"user_id": 1, "date": today}]},
    )
    assert add.status_code == 200
    overwrite = client.post(
        "/bulk-entry",
        json={"code": "UG", "cells": [{"user_id": 1, "date": today}]},
    )
    assert overwrite.json["updated"] == 1
    with vacation_app.db() as conn:
        codes = [
            row["code"]
            for row in conn.execute(
                "SELECT code FROM entries WHERE user_id = 1 AND entry_date = ?",
                (today,),
            )
        ]
    assert codes == ["UG"]
    delete = client.post(
        "/bulk-entry", json={"delete": True, "cells": [{"user_id": 1, "date": today}]}
    )
    assert delete.status_code == 200
    with vacation_app.db() as conn:
        count = conn.execute(
            "SELECT COUNT(*) FROM entries WHERE user_id = 1 AND entry_date = ?",
            (today,),
        ).fetchone()[0]
    assert count == 0


def test_import_ignores_excel_labels(client, tmp_path):
    login(client)
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["Name", date.today()])
    ws.append(["Monat", None])
    ws.append(["Geplant oder Beantragt", None])
    ws.append(["Erika Mustermann", "UrlbGplntOdrBntrgt"])
    ws.append(["Felix Feiertag", "Frtg"])
    path = tmp_path / "import.xlsx"
    wb.save(path)
    with client:
        vacation_app.import_excel(path)
    with vacation_app.db() as conn:
        names = [
            row["username"]
            for row in conn.execute("SELECT username FROM users ORDER BY id")
        ]
    assert "monat" not in names
    assert "geplant.oder.beantragt" not in names
    assert "erika.mustermann" in names
    assert "felix.feiertag" in names
    with vacation_app.db() as conn:
        imported_codes = [
            row["code"]
            for row in conn.execute(
                "SELECT e.code FROM entries e JOIN users u ON u.id = e.user_id WHERE u.username = ?",
                ("erika.mustermann",),
            )
        ]
        feiertag_count = conn.execute(
            "SELECT COUNT(*) FROM entries e JOIN users u ON u.id = e.user_id WHERE u.username = ?",
            ("felix.feiertag",),
        ).fetchone()[0]
    assert imported_codes == ["UB"]
    assert feiertag_count == 0
