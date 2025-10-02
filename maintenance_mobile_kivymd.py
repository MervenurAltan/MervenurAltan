from __future__ import annotations

"""Maintenance module for KivyMD mobile app.

This module is intentionally written so it can run in two different modes:

* If Kivy/KivyMD is available the UI part of the application can be executed.
* In environments without those heavy dependencies (such as CI) the business
  logic is still fully testable through the unit tests defined at the bottom of
  this file.

The UI implementation is out of scope for the tests and therefore omitted when
running in non-Kivy environments.  The data layer is the interesting part for
unit testing.
"""

import json
import os
import secrets
import hashlib
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

# Optional Kivy imports -----------------------------------------------------
KIVY_AVAILABLE = False
try:  # pragma: no cover - we never execute the UI in tests
    from kivy.lang import Builder  # type: ignore
    from kivy.clock import Clock  # type: ignore
    from kivy.factory import Factory  # type: ignore
    from kivy.properties import StringProperty, BooleanProperty  # type: ignore
    from kivy.uix.screenmanager import Screen, ScreenManager, SlideTransition  # type: ignore

    from kivymd.app import MDApp  # type: ignore
    from kivymd.uix.button import MDRaisedButton, MDFlatButton  # type: ignore
    from kivymd.uix.card import MDCard  # type: ignore
    from kivymd.uix.boxlayout import MDBoxLayout  # type: ignore
    from kivymd.uix.label import MDLabel  # type: ignore
    from kivymd.uix.dialog import MDDialog  # type: ignore
    from kivymd.uix.snackbar import Snackbar  # type: ignore
    from kivymd.uix.textfield import MDTextField  # type: ignore
    from kivymd.uix.list import OneLineListItem  # type: ignore
    from kivymd.uix.datatables import MDDataTable  # type: ignore
    from kivymd.toast import toast  # type: ignore
    from kivy.metrics import dp  # type: ignore

    KIVY_AVAILABLE = True
except Exception:  # pragma: no cover - dependency is optional
    # In the CI environment Kivy/KivyMD is not available which is fine for our
    # business logic tests.
    pass

try:  # pragma: no cover - optional dependency
    from plyer import filechooser  # type: ignore
except Exception:  # pragma: no cover - optional dependency is not required
    filechooser = None


# ------------------------------- Constants ---------------------------------
APP_W, APP_H = 1080, 900
REFRESH_MS = 10_000
FINISHED_KEEP_DAYS = 7

GREEN = "#1f9b4c"
WHITE = "#ffffff"
BG = "#f6f7f9"

USERS_DB_PATH = "users.json"
PERM_KEYS = [
    "duruslu_talep",
    "durussuz_talep",
    "bugun_planli",
    "planli_isler",
    "yardimci",
    "idari",
]


def _hash_password(password: str, salt: Optional[str] = None) -> tuple[str, str]:
    """Return a salted SHA256 hash for ``password``.

    The function returns the hash and the salt used.  A new salt is generated if
    it is not provided.  This is intentionally a small helper that does not aim
    to be cryptographically bullet proof but is sufficient for the app.
    """

    salt = salt or secrets.token_hex(16)
    digest = hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
    return digest, salt


class UserStore:
    """Simple JSON based store that keeps track of the users."""

    def __init__(self, path: str = USERS_DB_PATH):
        self.path = path
        self.data: Dict[str, Any] = {"version": 1, "users": []}
        self.load()
        self.ensure_admin()

    # -- Persistence -----------------------------------------------------
    def load(self) -> None:
        if not os.path.exists(self.path):
            self.data = {"version": 1, "users": []}
            return
        try:
            with open(self.path, "r", encoding="utf-8") as fh:
                self.data = json.load(fh)
        except Exception:
            self.data = {"version": 1, "users": []}

    def save(self) -> None:
        tmp_path = f"{self.path}.tmp"
        with open(tmp_path, "w", encoding="utf-8") as fh:
            json.dump(self.data, fh, ensure_ascii=False, indent=2)
        os.replace(tmp_path, self.path)

    # -- Admin guarantee -------------------------------------------------
    def ensure_admin(self) -> None:
        if any(u.get("is_admin") for u in self.data["users"]):
            return
        pwd_hash, salt = _hash_password("admin")
        self.data["users"].append(
            {
                "username": "admin",
                "email": "admin@local",
                "pwd_hash": pwd_hash,
                "salt": salt,
                "is_admin": True,
                "perms": {k: True for k in PERM_KEYS},
            }
        )
        self.save()

    def _ensure_at_least_one_admin(self) -> None:
        if not any(u.get("is_admin") for u in self.data["users"]):
            raise ValueError("En az bir admin kullanıcı olmalı.")

    # -- Basic helpers ---------------------------------------------------
    def authenticate(self, username: str, password: str) -> Optional[Dict[str, Any]]:
        for user in self.data["users"]:
            if user["username"].lower() != username.lower():
                continue
            hashed, _ = _hash_password(password, user["salt"])
            if hashed == user["pwd_hash"]:
                return {
                    "username": user["username"],
                    "email": user.get("email", ""),
                    "is_admin": bool(user.get("is_admin")),
                    "perms": dict(user.get("perms", {})),
                }
        return None

    def list_users(self) -> List[str]:
        return sorted([u["username"] for u in self.data["users"]], key=str.lower)

    def get_user(self, username: str) -> Optional[Dict[str, Any]]:
        for user in self.data["users"]:
            if user["username"].lower() == username.lower():
                return json.loads(json.dumps(user))  # deep copy
        return None

    def upsert_user(
        self,
        username: str,
        email: str,
        is_admin: bool,
        perms: Dict[str, bool],
        new_password: Optional[str] = None,
        create: bool = False,
    ) -> None:
        if create and self.get_user(username):
            raise ValueError("Bu kullanıcı adı zaten var.")
        if create and (not new_password or not new_password.strip()):
            raise ValueError("Yeni kullanıcı için şifre zorunludur.")

        perms = {k: bool(perms.get(k, False)) for k in PERM_KEYS}
        if is_admin:
            perms = {k: True for k in PERM_KEYS}

        if create:
            pwd_hash, salt = _hash_password(new_password or "")
            self.data["users"].append(
                {
                    "username": username,
                    "email": email or "",
                    "pwd_hash": pwd_hash,
                    "salt": salt,
                    "is_admin": bool(is_admin),
                    "perms": perms,
                }
            )
        else:
            for user in self.data["users"]:
                if user["username"].lower() != username.lower():
                    continue
                user["email"] = email or ""
                user["is_admin"] = bool(is_admin)
                user["perms"] = perms
                if new_password and new_password.strip():
                    user["pwd_hash"], user["salt"] = _hash_password(new_password)
                break
            else:
                raise ValueError("Kullanıcı bulunamadı.")
        self._ensure_at_least_one_admin()
        self.save()

    def reset_password(self, username: str, new_password: str) -> None:
        for user in self.data["users"]:
            if user["username"].lower() != username.lower():
                continue
            user["pwd_hash"], user["salt"] = _hash_password(new_password)
            self.save()
            return
        raise ValueError("Kullanıcı bulunamadı.")

    def delete_user(self, username: str) -> None:
        self.data["users"] = [
            u for u in self.data["users"] if u["username"].lower() != username.lower()
        ]
        self._ensure_at_least_one_admin()
        self.save()


class JobStore:
    """In-memory store for work orders."""

    def __init__(self):
        self.jobs: Dict[str, Dict[str, Any]] = {}
        self._auto_seq = 1

    # -- Seeding helpers -------------------------------------------------
    def seed_from_faults(self, faults: List[Dict[str, Any]], with_stop: Optional[bool] = None) -> None:
        for row in faults:
            job_no = f"WO-{row.get('id', 0)}"
            if job_no in self.jobs:
                continue
            self.jobs[job_no] = {
                "job_no": job_no,
                "machine": row.get("machineId", ""),
                "status": "Bekliyor",
                "created_at": datetime.now().isoformat(),
                "started_at": None,
                "finished_at": None,
                "transferred_at": None,
                "expires_at": None,
                "pm": False,
                "pm_data": None,
                "helper": False,
                "from_helper": False,
                "helper_data": {},
                "transfer": None,
                "admin": False,
                "admin_data": None,
                "assignee": None,
                "with_stop": with_stop,
            }

    def ensure_job(self, job_no: str, machine: str = "") -> None:
        if job_no in self.jobs:
            return
        self.jobs[job_no] = {
            "job_no": job_no,
            "machine": machine,
            "status": "Bekliyor",
            "created_at": datetime.now().isoformat(),
            "started_at": None,
            "finished_at": None,
            "transferred_at": None,
            "expires_at": None,
            "pm": False,
            "pm_data": None,
            "helper": False,
            "from_helper": False,
            "helper_data": {},
            "transfer": None,
            "admin": False,
            "admin_data": None,
            "assignee": None,
        }

    def set_started(self, job_no: str) -> None:
        job = self.jobs.get(job_no)
        if not job:
            self.ensure_job(job_no)
            job = self.jobs[job_no]
        job["status"] = "Devam Ediyor"
        job["started_at"] = datetime.now().isoformat()
        job["finished_at"] = None
        job["transferred_at"] = None
        job["expires_at"] = None

    def set_transferred(self, job_no: str, department: str, note: str = "") -> None:
        job = self.jobs.get(job_no)
        if not job:
            self.ensure_job(job_no)
            job = self.jobs[job_no]
        job["status"] = "Devredildi"
        job["transferred_at"] = datetime.now().isoformat()
        job["transfer"] = {"department": department, "note": note}
        job["started_at"] = None
        job["finished_at"] = None
        job["expires_at"] = None

    def set_finished(self, job_no: str) -> None:
        job = self.jobs.get(job_no)
        if not job:
            self.ensure_job(job_no)
            job = self.jobs[job_no]
        job["status"] = "Bitti"
        finished = datetime.now()
        job["finished_at"] = finished.isoformat()
        job["expires_at"] = (finished + timedelta(days=FINISHED_KEEP_DAYS)).isoformat()

    # -- Planned maintenance --------------------------------------------
    def seed_from_planned(self, plans: List[Dict[str, Any]]) -> None:
        for plan in plans:
            job_no = f"PM-{plan.get('id', 0)}"
            pm_data = self._normalize_pm(plan)
            if job_no in self.jobs:
                self.jobs[job_no]["pm"] = True
                self.jobs[job_no]["pm_data"] = pm_data
                continue

            task_statuses = [task["status"] for task in pm_data["tasks"]]
            if task_statuses and all(status == "Bitti" for status in task_statuses):
                status = "Bitti"
            elif any(status == "Devam Ediyor" for status in task_statuses):
                status = "Devam Ediyor"
            else:
                status = "Bekliyor"

            self.jobs[job_no] = {
                "job_no": job_no,
                "machine": plan.get("machineId", ""),
                "status": status,
                "created_at": datetime.now().isoformat(),
                "started_at": datetime.now().isoformat() if status == "Devam Ediyor" else None,
                "finished_at": datetime.now().isoformat() if status == "Bitti" else None,
                "transferred_at": None,
                "expires_at": None,
                "pm": True,
                "pm_data": pm_data,
                "helper": False,
                "from_helper": False,
                "helper_data": {},
                "transfer": None,
                "admin": False,
                "admin_data": None,
                "assignee": None,
            }

    def _normalize_pm(self, plan: Dict[str, Any]) -> Dict[str, Any]:
        tasks = []
        for task in plan.get("tasks", []):
            dept, ttype, desc, assignee, status = task
            tasks.append(
                {
                    "dept": dept,
                    "type": ttype,
                    "desc": desc,
                    "assignee": assignee,
                    "status": status,
                }
            )
        return {
            "machineId": plan.get("machineId", ""),
            "maintType": plan.get("maintType", ""),
            "maintArea": plan.get("maintArea", ""),
            "plannedStart": plan.get("plannedStart", ""),
            "plannedEnd": plan.get("plannedEnd", ""),
            "isPeriodic": bool(plan.get("isPeriodic", False)),
            "materials": list(plan.get("materials", [])),
            "desc": plan.get("desc", ""),
            "tasks": tasks,
        }

    def pm_take_task(self, job_no: str, idx: int, assignee: str = "Personel") -> None:
        job = self.jobs[job_no]
        job["pm_data"]["tasks"][idx]["assignee"] = assignee
        job["pm_data"]["tasks"][idx]["status"] = "Devam Ediyor"
        job["status"] = "Devam Ediyor"
        job["started_at"] = job["started_at"] or datetime.now().isoformat()
        job["finished_at"] = None
        job["expires_at"] = None
        job["transferred_at"] = None

    def pm_finish_task(self, job_no: str, idx: int) -> None:
        job = self.jobs[job_no]
        job["pm_data"]["tasks"][idx]["status"] = "Bitti"
        if all(task["status"] == "Bitti" for task in job["pm_data"]["tasks"]):
            self.set_finished(job_no)
        else:
            job["status"] = "Devam Ediyor"

    # -- Helper module --------------------------------------------------
    def seed_from_helpers(self, helpers: List[Dict[str, Any]]) -> None:
        for helper in helpers:
            job_no = f"HL-{helper.get('id', 0)}"
            if job_no in self.jobs:
                continue
            self.jobs[job_no] = {
                "job_no": job_no,
                "machine": helper.get("workCenter", ""),
                "status": "Bekliyor",
                "created_at": datetime.now().isoformat(),
                "started_at": None,
                "finished_at": None,
                "transferred_at": None,
                "expires_at": None,
                "pm": False,
                "pm_data": None,
                "helper": True,
                "from_helper": False,
                "helper_data": {
                    "workCenter": helper.get("workCenter", ""),
                    "helperType": helper.get("helperType", ""),
                    "checks": helper.get("checks", []),
                    "contractor": bool(helper.get("contractor", False)),
                },
                "transfer": None,
                "admin": False,
                "admin_data": None,
                "assignee": None,
            }

    def create_fault_from_helper(
        self,
        work_center: str,
        reason: str,
        contractor: bool = False,
        failed_checks: Optional[List[str]] = None,
        source_job: Optional[str] = None,
    ) -> str:
        stamp = int(datetime.now().timestamp())
        job_no = f"WO-H{stamp}{self._auto_seq:02d}"
        self._auto_seq += 1
        self.jobs[job_no] = {
            "job_no": job_no,
            "machine": work_center,
            "status": "Bekliyor",
            "created_at": datetime.now().isoformat(),
            "started_at": None,
            "finished_at": None,
            "transferred_at": None,
            "expires_at": None,
            "pm": False,
            "pm_data": None,
            "helper": False,
            "from_helper": True,
            "helper_data": {
                "workCenter": work_center,
                "faultReason": reason,
                "contractor": bool(contractor),
                "failed_checks": failed_checks or [],
                "source_job": source_job,
            },
            "transfer": None,
            "admin": False,
            "admin_data": None,
            "assignee": None,
        }
        return job_no

    def seed_from_admins(self, admins: List[Dict[str, Any]]) -> None:
        for admin_job in admins:
            job_no = f"ADM-{admin_job.get('id', 0)}"
            if job_no in self.jobs:
                continue
            self.jobs[job_no] = {
                "job_no": job_no,
                "machine": "İdari",
                "status": "Bekliyor",
                "created_at": datetime.now().isoformat(),
                "started_at": None,
                "finished_at": None,
                "transferred_at": None,
                "expires_at": None,
                "pm": False,
                "pm_data": None,
                "helper": False,
                "from_helper": False,
                "helper_data": {},
                "transfer": None,
                "admin": True,
                "admin_data": {
                    "title": admin_job.get("title", "İdari İş"),
                    "requestDate": admin_job.get("requestDate", ""),
                },
                "assignee": None,
            }

    # -- Queries --------------------------------------------------------
    def cleanup_expired(self) -> None:
        now = datetime.now()
        to_delete: List[str] = []
        for job_no, job in self.jobs.items():
            if job["status"] != "Bitti" or not job.get("expires_at"):
                continue
            try:
                expires = datetime.fromisoformat(job["expires_at"])
            except ValueError:
                continue
            if expires < now:
                to_delete.append(job_no)
        for job_no in to_delete:
            self.jobs.pop(job_no, None)

    def get_active(self) -> List[Dict[str, Any]]:
        return [job for job in self.jobs.values() if job["status"] != "Bitti"]

    def get_finished(self) -> List[Dict[str, Any]]:
        now = datetime.now()
        results = []
        for job in self.jobs.values():
            if job["status"] != "Bitti":
                continue
            expires_at = job.get("expires_at")
            if not expires_at:
                results.append(job)
                continue
            try:
                if datetime.fromisoformat(expires_at) >= now:
                    results.append(job)
            except ValueError:
                results.append(job)
        return results

    def get_helper_list_rows(self) -> List[Dict[str, Any]]:
        rows = [job for job in self.jobs.values() if job.get("helper") or job.get("from_helper")]
        rows = [job for job in rows if job["status"] != "Bitti"]
        rows.sort(key=lambda job: (job["status"] != "Devam Ediyor", job["created_at"]))
        return rows


# --------------------------- Fetch helpers ---------------------------------
@dataclass
class CardConfig:
    key: str
    title: str
    color: str
    open_on_click: bool = True


CARDS = [
    CardConfig("duruslu_talep", "Duruşlu Servis Talepleri", GREEN),
    CardConfig("durussuz_talep", "Duruşsuz İş Talepleri", GREEN),
    CardConfig("bugun_planli", "Bugünün Planlı Bakımları", GREEN),
    CardConfig("planli_isler", "Planlı Bakım İşleri", GREEN),
    CardConfig("yardimci", "Yardımcı Ekipmanlar", WHITE),
    CardConfig("idari", "İdari İşler", WHITE),
]


def fetch_faults(with_stop: bool) -> List[Dict[str, Any]]:
    base = 1000 if with_stop else 2000
    return [
        {
            "id": base + 11,
            "machineId": "Film 10",
            "faultArea": "Sarici A",
            "department": "Mekanik",
            "requestDate": "2025-09-19T09:12:00Z",
            "faultReason": "Kavrama gecikmesi",
            "requester": "Operatör A",
            "faultType": "Mekanik",
        },
        {
            "id": base + 12,
            "machineId": "Kesim 3",
            "faultArea": "Bıçak",
            "department": "Mekanik",
            "requestDate": "2025-09-19T10:05:00Z",
            "faultReason": "Bıçak körleşmesi",
            "requester": "Operatör B",
            "faultType": "Mekanik",
        },
        {
            "id": base + 13,
            "machineId": "Baskı 1",
            "faultArea": "Elektrik",
            "department": "Elektrik",
            "requestDate": "2025-09-19T11:22:00Z",
            "faultReason": "Sensör arızası",
            "requester": "Operatör C",
            "faultType": "Elektrik",
        },
    ]


def fetch_planned() -> List[Dict[str, Any]]:
    base = 3000
    return [
        {
            "id": base + 1,
            "machineId": "Film 10",
            "maintType": "Periyodik - Rulman",
            "maintArea": "Mekanik",
            "plannedStart": "2025-09-21T08:00:00Z",
            "plannedEnd": "2025-09-21T16:30:00Z",
            "isPeriodic": True,
            "materials": ["Rulman 6205", "Gres Yağı NLGI-2"],
            "desc": "Rulman değişimi ve yağlama.",
            "tasks": [
                ("Mekanik", "Rulman Değişimi", "Açıklama", "—", "Bekliyor"),
                ("Elektrik", "Kablo Kontrolü", "Açıklama", "—", "Devam Ediyor"),
                ("Elektrik", "Pano Temizliği", "Açıklama", "—", "Bitti"),
            ],
        },
        {
            "id": base + 2,
            "machineId": "Baskı 1",
            "maintType": "Kalibrasyon",
            "maintArea": "Elektrik",
            "plannedStart": "2025-09-22T09:00:00Z",
            "plannedEnd": "2025-09-22T12:00:00Z",
            "isPeriodic": False,
            "materials": ["Kalibrasyon Seti"],
            "desc": "Sensör kalibrasyonları.",
            "tasks": [
                ("Elektrik", "Kalibrasyon", "Hat sensörleri", "—", "Bekliyor"),
            ],
        },
        {
            "id": base + 3,
            "machineId": "Kesim 3",
            "maintType": "Periyodik - Bıçak",
            "maintArea": "Mekanik",
            "plannedStart": "2025-09-23T08:00:00Z",
            "plannedEnd": "2025-09-23T17:00:00Z",
            "isPeriodic": True,
            "materials": ["Bıçak seti"],
            "desc": "Bıçakların kontrolü ve değişimi.",
            "tasks": [
                ("Mekanik", "Bıçak Kontrol", "Açıklama", "—", "Bekliyor"),
            ],
        },
    ]


def fetch_helpers() -> List[Dict[str, Any]]:
    base = 5000
    return [
        {
            "id": base + 1,
            "workCenter": "Su Deposu",
            "helperType": "Günlük Kontrol",
            "contractor": False,
            "checks": [
                {"label": "Ham Su Depo Seviyesi", "range": "> 30", "kind": "gt", "min": 30, "max": None},
                {"label": "Arıtılmış Su Depo Seviyesi", "range": "> 50", "kind": "gt", "min": 50, "max": None},
                {"label": "pH (Arıtılmış Su)", "range": "5 < x < 7", "kind": "between", "min": 5, "max": 7},
                {"label": "UV Led", "range": "EVET/HAYIR", "kind": "bool", "min": None, "max": None},
            ],
        },
        {
            "id": base + 2,
            "workCenter": "Kompresör",
            "helperType": "Vardiya Kontrol",
            "contractor": True,
            "checks": [
                {"label": "Hat Basıncı (bar)", "range": "6-8", "kind": "between", "min": 6, "max": 8},
                {"label": "Yağ Seviyesi", "range": "EVET/HAYIR", "kind": "bool", "min": None, "max": None},
            ],
        },
    ]


def fetch_admins() -> List[Dict[str, Any]]:
    base = 7000
    return [
        {"id": base + 1, "title": "Kapı kolu değişimi", "requestDate": "2025-09-18T10:20:00Z"},
        {"id": base + 2, "title": "Ofis düzeni – masa taşınması", "requestDate": "2025-09-19T08:45:00Z"},
        {"id": base + 3, "title": "Toplantı odası projektör kurulumu", "requestDate": "2025-09-19T13:10:00Z"},
    ]


# ------------------------------ Utilities ----------------------------------
def fmt_dt_iso_to_tr(value: str) -> str:
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except Exception:
        return value or "-"
    return dt.strftime("%d.%m.%Y %H:%M")


def user_can(user: Optional[Dict[str, Any]], key: str) -> bool:
    if not user:
        return False
    if user.get("is_admin"):
        return True
    return bool(user.get("perms", {}).get(key, False))


# ------------------------------ Unit tests ----------------------------------
if not KIVY_AVAILABLE:
    import tempfile
    import unittest

    class UserStoreTests(unittest.TestCase):
        def setUp(self) -> None:
            self.tmpdir = tempfile.TemporaryDirectory()
            self.path = os.path.join(self.tmpdir.name, "users.json")

        def tearDown(self) -> None:
            self.tmpdir.cleanup()

        def test_admin_is_created_automatically(self) -> None:
            store = UserStore(self.path)
            self.assertEqual(store.list_users(), ["admin"])
            admin = store.get_user("admin")
            self.assertTrue(admin["is_admin"])
            self.assertTrue(all(admin["perms"].values()))
            # authenticate with default password
            self.assertIsNotNone(store.authenticate("admin", "admin"))

        def test_create_and_authenticate_user(self) -> None:
            store = UserStore(self.path)
            perms = {key: (key == "yardimci") for key in PERM_KEYS}
            store.upsert_user(
                "technician",
                "tech@example.com",
                is_admin=False,
                perms=perms,
                new_password="secret",
                create=True,
            )
            user = store.authenticate("technician", "secret")
            self.assertIsNotNone(user)
            assert user is not None  # help mypy
            self.assertFalse(user["is_admin"])
            self.assertTrue(user["perms"]["yardimci"])
            self.assertFalse(user["perms"]["idari"])

        def test_delete_last_admin_is_not_allowed(self) -> None:
            store = UserStore(self.path)
            with self.assertRaises(ValueError):
                store.delete_user("admin")

    class JobStoreTests(unittest.TestCase):
        def setUp(self) -> None:
            self.store = JobStore()

        def test_seed_from_faults(self) -> None:
            faults = fetch_faults(with_stop=True)
            self.store.seed_from_faults(faults, with_stop=True)
            job_no = f"WO-{faults[0]['id']}"
            self.assertIn(job_no, self.store.jobs)
            job = self.store.jobs[job_no]
            self.assertEqual(job["machine"], faults[0]["machineId"])
            self.assertEqual(job["status"], "Bekliyor")

        def test_start_and_finish_job(self) -> None:
            self.store.ensure_job("WO-1", machine="Test")
            self.store.set_started("WO-1")
            job = self.store.jobs["WO-1"]
            self.assertEqual(job["status"], "Devam Ediyor")
            self.assertIsNotNone(job["started_at"])
            self.store.set_finished("WO-1")
            job = self.store.jobs["WO-1"]
            self.assertEqual(job["status"], "Bitti")
            self.assertIsNotNone(job["finished_at"])
            self.assertIsNotNone(job["expires_at"])

        def test_cleanup_expired_jobs(self) -> None:
            self.store.ensure_job("WO-1")
            self.store.set_finished("WO-1")
            job = self.store.jobs["WO-1"]
            job["expires_at"] = (datetime.now() - timedelta(days=1)).isoformat()
            self.store.cleanup_expired()
            self.assertNotIn("WO-1", self.store.jobs)

        def test_pm_flow(self) -> None:
            plans = fetch_planned()
            self.store.seed_from_planned(plans)
            job_no = f"PM-{plans[0]['id']}"
            self.store.pm_take_task(job_no, 0, assignee="Ali")
            job = self.store.jobs[job_no]
            self.assertEqual(job["status"], "Devam Ediyor")
            self.assertEqual(job["pm_data"]["tasks"][0]["assignee"], "Ali")
            self.store.pm_finish_task(job_no, 0)
            self.assertEqual(job["pm_data"]["tasks"][0]["status"], "Bitti")

        def test_helper_to_fault_conversion(self) -> None:
            helper_jobs = fetch_helpers()
            self.store.seed_from_helpers(helper_jobs)
            created = self.store.create_fault_from_helper(
                work_center="Su Deposu",
                reason="Düşük seviye",
                contractor=False,
                failed_checks=["Ham Su Depo Seviyesi"],
                source_job="HL-5001",
            )
            self.assertIn(created, self.store.jobs)
            job = self.store.jobs[created]
            self.assertTrue(job["from_helper"])
            self.assertEqual(job["helper_data"]["faultReason"], "Düşük seviye")

        def test_getters(self) -> None:
            self.store.ensure_job("WO-1")
            self.store.ensure_job("WO-2")
            self.store.set_started("WO-1")
            self.store.set_finished("WO-2")
            finished = self.store.get_finished()
            active = self.store.get_active()
            self.assertEqual(len(finished), 1)
            self.assertEqual(len(active), 1)

    if __name__ == "__main__":  # pragma: no cover - direct execution entry
        unittest.main()
