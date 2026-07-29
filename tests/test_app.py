from datetime import date
from pathlib import Path

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


def create_initial_member(client, username="azubi", roles=None):
    client.post(
        "/members",
        data={
            "action": "create",
            "username": username,
            "first_name": username.title(),
            "last_name": "Muster",
            "roles": roles or ["normal"],
        },
    )
    with vacation_app.db() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()
        return row["id"], vacation_app.reveal_initial_password(row)


def activate_member(client, username, initial_password, new_password="new-secret"):
    return client.post(
        "/initial-login",
        data={
            "username": username,
            "initial_password": initial_password,
            "new_password": new_password,
            "password_confirmation": new_password,
        },
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
    with vacation_app.db() as conn:
        user = conn.execute("SELECT * FROM users WHERE id = 2").fetchone()
    password = vacation_app.reveal_initial_password(user)
    assert len(password) == 8
    assert user["must_change_password"] == 1
    assert password.encode() in response.data


def test_role_based_edit_permissions(client):
    login(client)
    _, initial = create_initial_member(client, roles=["azubi"])
    client.get("/logout")
    forced = login(client, "azubi", initial)
    assert b"Initiales Passwort" in forced.data
    activate_member(client, "azubi", initial)
    login(client, "azubi", "new-secret")
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
    ws.append(["Monat", "Juli"])
    ws.append(["KW", 30])
    ws.append(["Ferienzeit", None])
    ws.append(["Name", date.today()])
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


def test_initial_link_and_password_reset(client):
    login(client)
    user_id, initial = create_initial_member(client, "lina", ["normal"])
    link = vacation_app.initial_link("lina", initial)
    client.get("/logout")
    response = client.get(link.replace("https://urlaub.extrahelden.de", ""))
    assert b'value="lina"' in response.data
    assert initial.encode() in response.data

    login(client)
    response = client.post(
        "/members",
        data={"action": "reset-password", "user_id": user_id},
    )
    assert b"Neu erstellte Initialdatenanzeige" in response.data
    blocked = client.get("/", follow_redirects=True)
    assert b"Neu erstellte Initialdatenanzeige" in blocked.data
    acknowledged = client.post(
        "/members/acknowledge-initial-data", json={"user_id": user_id}
    )
    assert acknowledged.json == {"ok": True}
    assert client.get("/").status_code == 200
    with vacation_app.db() as conn:
        reset_user = conn.execute(
            "SELECT * FROM users WHERE id = ?", (user_id,)
        ).fetchone()
    assert vacation_app.reveal_initial_password(reset_user) != initial


def test_bulk_roles_and_ausbilder_may_only_reset_azubi(client):
    login(client)
    azubi_id, _ = create_initial_member(client, "lea", ["azubi"])
    normal_id, _ = create_initial_member(client, "tom", ["normal"])
    trainer_id, trainer_initial = create_initial_member(
        client, "trainer", ["ausbilder"]
    )
    client.post(
        "/members",
        data={
            "action": "roles-bulk",
            "roles_1": ["admin"],
            f"roles_{azubi_id}": ["azubi", "normal"],
            f"roles_{normal_id}": ["putzchef"],
            f"roles_{trainer_id}": ["ausbilder"],
        },
    )
    with vacation_app.db() as conn:
        assert vacation_app.roles_for_users(conn)[normal_id] == ["putzchef"]
    client.get("/logout")
    activate_member(client, "trainer", trainer_initial)
    login(client, "trainer", "new-secret")
    assert (
        client.post(
            "/members", data={"action": "reset-password", "user_id": azubi_id}
        ).status_code
        == 200
    )
    assert (
        client.post(
            "/members", data={"action": "reset-password", "user_id": normal_id}
        ).status_code
        == 403
    )


def test_profile_email_is_required_and_immutable(client):
    login(client)
    missing = client.post("/profile", data={"email": "", "birth_date": ""})
    assert b"Pflichtangabe" in missing.data
    saved = client.post(
        "/profile",
        data={"email": "admin@example.de", "birth_date": "1990-04-12"},
        follow_redirects=True,
    )
    assert b"admin@example.de" in saved.data
    client.post("/profile", data={"email": "changed@example.de", "birth_date": ""})
    with vacation_app.db() as conn:
        row = conn.execute(
            "SELECT email, birth_date FROM users WHERE id = 1"
        ).fetchone()
    assert row["email"] == "admin@example.de"
    assert row["birth_date"] is None


def test_repo_example_excel_imports_entries(client):
    login(client)
    example = (
        Path(__file__).resolve().parents[1] / "Urlaubsübersicht-Detaillierung.xlsx"
    )
    vacation_app.import_excel(example)
    with vacation_app.db() as conn:
        entry_count = conn.execute("SELECT COUNT(*) FROM entries").fetchone()[0]
        imported_user = conn.execute(
            "SELECT id FROM users WHERE username = ?", ("alexander.hälter",)
        ).fetchone()
    assert imported_user is not None
    assert entry_count > 100
