"""Local encrypted accounts and persistent, disjoint distribution groups."""
import copy
import json
import uuid
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit


def assign_distribution_members(groups, target_index, members):
    """Move explicitly checked members to one group, preserving its other members."""
    if len(groups) != 5 or not 0 <= target_index < 5:
        raise ValueError("저장할 배포 탭을 확인하세요.")
    updated = copy.deepcopy(groups)
    for group in updated:
        for key in members:
            group.pop(key, None)
    updated[target_index].update(copy.deepcopy(members))
    return updated


class ProfileStore:
    def __init__(self, path, protect=None, unprotect=None):
        from receipt_register import protect_windows_secret, unprotect_windows_secret
        self.path = Path(path)
        self.protect = protect or protect_windows_secret
        self.unprotect = unprotect or unprotect_windows_secret
        self.data = {"accounts": [], "active_id": None, "groups": {}}
        if self.path.is_file():
            self.data = json.loads(self.path.read_text(encoding="utf-8"))
            if len(self.data["accounts"]) > 3:
                raise ValueError("계정 파일에 3개를 초과한 계정이 있습니다.")
        if "checked_ids" not in self.data:
            self.data["checked_ids"] = [self.data["active_id"]] if self.data["active_id"] else []

    def persist(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_suffix(".tmp")
        temp.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")
        temp.replace(self.path)

    @property
    def accounts(self):
        return copy.deepcopy(self.data["accounts"])

    @property
    def active(self):
        return next((a.copy() for a in self.data["accounts"] if a["id"] == self.data["active_id"] and a["id"] in self.data["checked_ids"]), None)

    @property
    def login_choices(self):
        return [a for a in self.accounts if a["id"] in self.data["checked_ids"]]

    def set_checked(self, account_id, checked):
        if not any(a["id"] == account_id for a in self.accounts):
            raise ValueError("먼저 계정을 저장하세요.")
        ids = set(self.data["checked_ids"])
        if checked:
            ids.add(account_id)
        else:
            ids.discard(account_id)
            if self.data["active_id"] == account_id:
                self.data["active_id"] = None
        self.data["checked_ids"] = list(ids)
        self.persist()

    def credentials(self):
        a = self.active
        return (a["user_id"], self.unprotect(a["password"])) if a else ("", "")

    def save_account(self, slot, user_id, password):
        user_id = user_id.strip()
        if not user_id or not 0 <= slot < 3:
            raise ValueError("계정 아이디와 계정 위치(1~3)를 확인하세요.")
        accounts = self.data["accounts"]
        if slot > len(accounts):
            raise ValueError("앞쪽 계정부터 차례로 저장하세요.")
        if any(a["user_id"].strip().casefold() == user_id.casefold() and i != slot for i, a in enumerate(accounts)):
            raise ValueError("이미 저장된 아이디입니다.")
        old = accounts[slot] if slot < len(accounts) else None
        same = old and old["user_id"].strip().casefold() == user_id.casefold()
        if not password and not same:
            raise ValueError("새 계정의 비밀번호를 입력하세요.")
        record = {"id": old["id"] if same else uuid.uuid4().hex, "user_id": user_id,
                  "password": self.protect(password) if password else old["password"]}
        if old:
            accounts[slot] = record
            if not same:
                self.data["checked_ids"] = [key for key in self.data["checked_ids"] if key != old["id"]]
                if self.data["active_id"] == old["id"]:
                    self.data["active_id"] = None
        else:
            accounts.append(record)
        self.persist()
        return record.copy()

    def activate(self, account_id):
        if not any(a["id"] == account_id for a in self.data["accounts"]):
            raise ValueError("저장된 계정을 찾지 못했습니다.")
        self.data["active_id"] = account_id
        if account_id not in self.data["checked_ids"]:
            self.data["checked_ids"].append(account_id)
        self.persist()

    def deactivate(self, account_id):
        self.set_checked(account_id, False)

    def migrate(self, user_id, password):
        if not self.accounts and user_id and password:
            record = self.save_account(0, user_id, password)
            self.activate(record["id"])

    def groups(self, account_id):
        return copy.deepcopy(self.data["groups"].get(account_id, [{} for _ in range(5)]))

    def group_dates(self, account_id):
        return list(self.data.get("group_dates", {}).get(account_id, ["" for _ in range(5)]))

    def save_distribution_data(self, account_id, groups, dates):
        if not any(a["id"] == account_id for a in self.accounts) or len(groups) != 5 or len(dates) != 5:
            raise ValueError("계정과 배포1~배포5 데이터를 확인하세요.")
        claimed = set()
        for group in groups:
            if any(not str(key) for key in group) or claimed.intersection(group):
                raise ValueError("회원 ID가 없거나 여러 배포 탭에 중복된 업소가 있습니다.")
            claimed.update(group)
        normalized = []
        for i, value in enumerate(dates):
            raw = re.sub(r"[-./]", "", str(value).strip())
            if raw:
                if len(raw) != 8 or not raw.isdigit():
                    raise ValueError(f"배포{i+1} 날짜는 YYYYMMDD로 입력하세요.")
                try:
                    datetime.strptime(raw, "%Y%m%d")
                except ValueError as exc:
                    raise ValueError(f"배포{i+1} 날짜가 올바르지 않습니다.") from exc
            normalized.append(raw)
        previous = copy.deepcopy(self.data)
        try:
            self.data["groups"][account_id] = copy.deepcopy(groups)
            self.data.setdefault("group_dates", {})[account_id] = normalized
            self.persist()
        except Exception:
            self.data = previous
            raise

    def reload_distribution_data(self, account_id):
        if not any(a["id"] == account_id for a in self.accounts):
            raise ValueError("저장된 계정을 찾지 못했습니다.")
        if self.path.is_file():
            saved = json.loads(self.path.read_text(encoding="utf-8"))
            groups = copy.deepcopy(saved.get("groups", {}).get(account_id, [{} for _ in range(5)]))
            dates = list(saved.get("group_dates", {}).get(account_id, ["" for _ in range(5)]))
            if len(groups) != 5 or len(dates) != 5:
                raise ValueError("저장한 배포 데이터 형식을 확인하세요.")
            self.data["groups"][account_id] = groups
            self.data.setdefault("group_dates", {})[account_id] = dates
        return self.groups(account_id), self.group_dates(account_id)

    def save_group(self, account_id, index, members):
        if not any(a["id"] == account_id for a in self.accounts) or not 0 <= index < 5:
            raise ValueError("계정 또는 배포 탭을 확인하세요.")
        groups = self.groups(account_id)
        claimed = {key for i, group in enumerate(groups) if i != index for key in group}
        if claimed.intersection(members):
            raise ValueError("다른 배포 탭에서 이미 선택한 업소가 포함되어 있습니다.")
        if any(not str(key) for key in members):
            raise ValueError("회원 ID가 없는 업소는 저장할 수 없습니다.")
        groups[index] = copy.deepcopy(members)
        self.data["groups"][account_id] = groups
        self.persist()


def require_session_account(page, user_id, password, log=print, *, switch=False):
    from membership_fee_register import ensure_nmis_logged_in
    if not user_id or not password:
        raise RuntimeError("계정을 체크한 뒤 Chrome 연결 열기에서 로그인할 계정을 선택하세요.")
    target = urlsplit(page.url)
    if target.hostname != "nmis.foodservice.or.kr" or target.scheme not in {"http", "https"}:
        raise RuntimeError("NMIS 주소가 아닌 페이지에서는 계정을 사용할 수 없습니다.")
    if page.locator("#userId:visible").count():
        ensure_nmis_logged_in(page, user_id, password, log)
    identity = """() => {
        const roots=[document.querySelector('[ng-app]'),document.body].filter(Boolean);
        for(const root of roots){const s=window.angular && angular.element(root).scope();
            if(s?.userInfo?.userId) return String(s.userInfo.userId);}
        return '';
    }"""
    normalize = lambda value: str(value or "").strip().casefold()
    current = normalize(page.evaluate(identity))
    expected = normalize(user_id)
    if current != expected and switch:
        if not current:
            raise RuntimeError("NMIS 로그인 계정을 확인하지 못했습니다.")
        page.locator("a[ng-click*='fnLogout']:visible").click()
        # Logout may show a normal confirmation dialog.
        for _ in range(20):
            if page.locator("#userId:visible").count():
                break
            confirm = page.locator(".modal:visible button[ng-click*='fnConfirm']")
            if confirm.count():
                confirm.first.click()
            page.wait_for_timeout(250)
        ensure_nmis_logged_in(page, user_id, password, log)
        current = normalize(page.evaluate(identity))
    if not current:
        raise RuntimeError("NMIS 로그인 계정을 확인하지 못했습니다. Chrome 연결 열기로 다시 연결하세요.")
    if current != expected:
        raise RuntimeError("Chrome에서 로그인 계정이 변경되었습니다. Chrome 연결 열기에서 사용할 계정을 선택하세요.")
