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
