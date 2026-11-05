"""Entry point for the Windows portable build."""
import json
import sys
import tempfile
import traceback
from pathlib import Path


def self_test(report_path, *, gui=True):
    report = {"ok": False}
    try:
        import pythoncom
        import win32com.client
        from playwright.sync_api import sync_playwright
        from app_paths import APP_ROOT, RESOURCE_ROOT
        from local_profiles import ProfileStore
        import nmis_slip_ui as ui
        from app_version import APP_VERSION
        assert ui.SETTINGS_FILE.parent == APP_ROOT
        assert (RESOURCE_ROOT / "assets/nmis_automation.ico").is_file()
        assert (RESOURCE_ROOT / "hancom_security/FilePathCheckerModuleExample.dll").is_file()
        with sync_playwright() as playwright:
            assert playwright.chromium
        with tempfile.TemporaryDirectory() as folder:
            store = ProfileStore(Path(folder) / "profiles.json")
            account = store.save_account(0, "portable-probe", "temporary-probe-password")
            store.activate(account["id"])
            assert store.credentials()[1] == "temporary-probe-password"
            store.save_distribution_data(account["id"], [{"probe": {}}] + [{} for _ in range(4)], ["20261105"] * 5)
            assert ProfileStore(store.path).groups(account["id"])[0] == {"probe": {}}
        if gui:
            ui.ModernSlipUI.check_browser_connection = lambda self: None
            app = ui.ModernSlipUI()
            app._select_tab("distribution")
            app.update_idletasks()
            app.after(1500, app.quit)
            app.mainloop()
            app.destroy()
        report = {"ok": True, "version": APP_VERSION, "frozen": bool(getattr(sys, "frozen", False)),
                  "checks": (["GUI"] if gui else []) + ["Playwright driver", "COM libraries", "DPAPI", "profile persistence", "assets", "Hancom module", "writable paths"]}
    except Exception:
        report["error"] = traceback.format_exc()
    Path(report_path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] in {"--self-test", "--update-probe"}:
        raise SystemExit(self_test(sys.argv[2], gui=sys.argv[1] == "--self-test"))
    from nmis_slip_ui import main
    main()
