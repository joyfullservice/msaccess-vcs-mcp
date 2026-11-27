"""User-level setting decisions: the resolver, the file, vcs_configure, and the CLI.

The conftest points ``ACCESS_VCS_SETTINGS_PATH`` at a temporary file for every
test, so nothing here touches the real ``~/.msaccess-vcs-mcp/settings.json``.
"""

from __future__ import annotations

import asyncio
import json
import os
import time
from types import SimpleNamespace

import pytest
from mcp.server.elicitation import (
    AcceptedElicitation,
    CancelledElicitation,
    DeclinedElicitation,
)
from mcp.types import (
    ClientCapabilities,
    ElicitationCapability,
    FormElicitationCapability,
    UrlElicitationCapability,
)

import msaccess_vcs_mcp.config as config
import msaccess_vcs_mcp.screenshot_store as store
import msaccess_vcs_mcp.settings_form as settings_form
import msaccess_vcs_mcp.tools as tools_module
import msaccess_vcs_mcp.user_settings as us
from msaccess_vcs_mcp.access_gate import EXEMPT_TOOLS
from msaccess_vcs_mcp.cli import build_parser, run_configure

KEY = us.SCREENSHOT_RETENTION.key
NAME = us.SCREENSHOT_RETENTION.name


def _write(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data), encoding="utf-8")
    stamp = time.time() + 2
    os.utime(path, (stamp, stamp))


def _unwrap(fn):
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


# ---------------------------------------------------------------------------
# Resolver
# ---------------------------------------------------------------------------


def test_undecided_setting_uses_the_default():
    assert us.get_setting(KEY) == "30"
    assert us.get_setting_source(KEY) == us.SOURCE_DEFAULT
    assert us.pending_settings() == [NAME]
    assert store.retention_days() == 30.0


def test_saved_value_is_used_and_recorded_with_who_decided(settings_file):
    assert us.save({NAME: "90"}, "agent") == [NAME]
    assert us.get_setting(KEY) == "90"
    assert us.get_setting_source(KEY) == us.SOURCE_FILE
    assert us.pending_settings() == []
    assert store.retention_days() == 90.0
    entry = json.loads(settings_file.read_text(encoding="utf-8"))[KEY]
    assert entry["value"] == "90" and entry["decided_by"] == "agent"
    assert entry["decided_at"]


def test_environment_wins_over_the_file(monkeypatch):
    us.save({NAME: "90"}, "cli")
    monkeypatch.setenv(KEY, "7")
    assert us.get_setting(KEY) == "7"
    assert us.get_setting_source(KEY) == us.SOURCE_ENV
    assert store.retention_days() == 7.0


def test_workspace_roots_env_file_wins_over_the_file(monkeypatch, tmp_path):
    """The lazy roots load uses override=False; nothing here may pre-fill os.environ."""
    monkeypatch.setenv(KEY, "placeholder")
    monkeypatch.delenv(KEY)
    us.save({NAME: "90"}, "cli")
    store.retention_days()
    assert KEY not in os.environ

    workspace = tmp_path / "workspace"
    workspace.mkdir()
    (workspace / ".env").write_text(f"{KEY}=14\n", encoding="utf-8")
    monkeypatch.setattr(config, "_project_root", None)
    config._load_env_from_directory(workspace)

    assert us.get_setting(KEY) == "14"
    assert us.get_setting_source(KEY) == us.SOURCE_ENV


def test_choosing_the_default_follows_a_later_change(monkeypatch, settings_file):
    us.save({NAME: "default"}, "form")
    assert json.loads(settings_file.read_text(encoding="utf-8"))[KEY]["decision"] == "default"
    assert us.is_decided(KEY)
    assert us.get_setting_source(KEY) == us.SOURCE_DEFAULT
    changed = us.DecisionSetting(**{**us.SCREENSHOT_RETENTION.__dict__, "default": "60"})
    monkeypatch.setattr(us, "DECISION_SETTINGS", (changed,))
    assert us.get_setting(KEY) == "60"


def test_a_write_by_another_process_is_picked_up(settings_file):
    us.save({NAME: "90"}, "cli")
    assert us.get_setting(KEY) == "90"
    _write(settings_file, {KEY: {"value": "120", "decided_by": "cli"}})
    assert us.get_setting(KEY) == "120"


def test_merge_keeps_entries_this_version_does_not_know(settings_file):
    _write(settings_file, {"ACCESS_VCS_FUTURE_SETTING": {"value": "x", "decided_by": "cli"}})
    us.save({NAME: "45"}, "cli")
    data = json.loads(settings_file.read_text(encoding="utf-8"))
    assert data["ACCESS_VCS_FUTURE_SETTING"]["value"] == "x"
    assert data[KEY]["value"] == "45"
    assert not list(settings_file.parent.glob("*.tmp"))
    assert not settings_file.with_name("settings.json.lock").exists()


def test_a_held_lock_times_out_and_saves_nothing(monkeypatch, settings_file):
    monkeypatch.setattr(us, "LOCK_TIMEOUT_SEC", 0.2)
    lock = settings_file.with_name("settings.json.lock")
    lock.parent.mkdir(parents=True, exist_ok=True)
    lock.write_text("123", encoding="ascii")
    with pytest.raises(TimeoutError):
        us.save({NAME: "45"}, "cli")
    assert not settings_file.exists()


def test_a_stale_lock_is_cleared(settings_file):
    lock = settings_file.with_name("settings.json.lock")
    lock.parent.mkdir(parents=True, exist_ok=True)
    lock.write_text("123", encoding="ascii")
    old = time.time() - us.STALE_LOCK_SEC - 5
    os.utime(lock, (old, old))
    us.save({NAME: "45"}, "cli")
    assert us.get_setting(KEY) == "45"


def test_a_corrupt_file_is_set_aside(settings_file):
    settings_file.parent.mkdir(parents=True, exist_ok=True)
    settings_file.write_text("{not json", encoding="utf-8")
    assert us.get_setting(KEY) == "30"
    assert not settings_file.exists()
    assert len(list(settings_file.parent.glob("settings.json.corrupt-*"))) == 1
    us.save({NAME: "45"}, "cli")
    assert us.get_setting(KEY) == "45"


def test_an_invalid_saved_value_counts_as_undecided(settings_file):
    _write(settings_file, {KEY: {"value": "soon", "decided_by": "cli"}})
    assert us.get_setting(KEY) == "30"
    assert us.pending_settings() == [NAME]


@pytest.mark.parametrize("value", ["-1", "nope", "nan", "inf", "0", "0.0", "1.5", "always"])
def test_invalid_values_save_nothing(value, settings_file):
    with pytest.raises(ValueError):
        us.save({NAME: value}, "cli")
    assert not settings_file.exists()


@pytest.mark.parametrize("value", [0, -1])
def test_numbers_never_stand_for_a_policy(value):
    with pytest.raises(ValueError, match='"never", "forever"'):
        us.save({NAME: value}, "agent")


@pytest.mark.parametrize("value, word, days", [
    ("forever", "forever", None), (" FOREVER ", "forever", None),
    ("never", "never", 0.0), ("Never", "never", 0.0),
])
def test_never_and_forever_are_saved_as_words(value, word, days, settings_file):
    us.save({NAME: value}, "cli")
    assert json.loads(settings_file.read_text(encoding="utf-8"))[KEY]["value"] == word
    assert us.get_setting(KEY) == word
    assert us.describe(us.SCREENSHOT_RETENTION)["display"] == word
    assert store.retention_days() == days
    assert store.saving_allowed() is (word != "never")


@pytest.mark.parametrize("env, expected", [("forever", "forever"), ("never", "never"), ("0", "0")])
def test_the_environment_takes_the_same_values(monkeypatch, env, expected):
    monkeypatch.setenv(KEY, env)
    assert us.get_setting(KEY) == expected
    assert us.get_setting_source(KEY) == us.SOURCE_ENV


def test_a_stored_zero_from_the_old_form_is_asked_again(settings_file):
    _write(settings_file, {KEY: {"value": "0", "decided_by": "form"}})
    assert us.pending_settings() == [NAME]
    assert us.get_setting(KEY) == "30"


def test_reset_makes_the_setting_pending_again():
    us.save({NAME: "45"}, "cli")
    assert us.reset(NAME) is True
    assert us.pending_settings() == [NAME]
    assert us.reset(NAME) is False


def test_audit_logging_is_not_a_decision_setting():
    assert "ACCESS_VCS_ENABLE_LOGGING" not in {s.key for s in us.DECISION_SETTINGS}


# ---------------------------------------------------------------------------
# vcs_configure
# ---------------------------------------------------------------------------


class _Ctx:
    def __init__(self, elicitation=None, *, action="accept", data=None, delay=0.0, error=None):
        caps = ClientCapabilities(elicitation=elicitation)
        self.session = SimpleNamespace(client_params=SimpleNamespace(capabilities=caps))
        self.action, self.data, self.delay, self.error = action, data or {}, delay, error
        self.asked = 0
        self.schema = None

    async def elicit(self, message, schema):
        self.asked += 1
        self.schema = schema
        if self.delay:
            await asyncio.sleep(self.delay)
        if self.error:
            raise self.error
        if self.action == "accept":
            return AcceptedElicitation(data=schema(**self.data))
        if self.action == "decline":
            return DeclinedElicitation()
        return CancelledElicitation()


FORM = ElicitationCapability(form=FormElicitationCapability())


def _configure(ctx, **values):
    return asyncio.run(settings_form.configure(ctx, values))


def test_vcs_configure_is_gate_exempt_and_registered():
    assert "vcs_configure" in EXEMPT_TOOLS
    assert "vcs_configure" in {t.name for t in tools_module.mcp._tool_manager.list_tools()}


@pytest.mark.parametrize("elicitation, expected", [
    (None, False),
    (FORM, True),
    (ElicitationCapability(), True),
    (ElicitationCapability(url=UrlElicitationCapability()), False),
    (ElicitationCapability(form=FormElicitationCapability(), url=UrlElicitationCapability()), True),
])
def test_form_support_is_read_from_the_modes(elicitation, expected):
    assert settings_form.client_supports_form(_Ctx(elicitation)) is expected


def test_no_session_means_no_form():
    assert settings_form.client_supports_form(None) is False


def test_accepted_form_saves_with_decided_by_form():
    ctx = _Ctx(FORM, data={NAME: "7 days"})
    result = _configure(ctx)
    assert result["form"] == "accepted"
    assert result["saved"] == [NAME]
    assert result["pending_settings"] == []
    assert us.describe(us.SCREENSHOT_RETENTION)["decided_by"] == "form"
    assert us.get_setting(KEY) == "7"
    assert "agent_guidance" not in result


def _form_property():
    model = settings_form._form_model(list(us.DECISION_SETTINGS))
    return model.model_json_schema()["properties"][NAME]


def test_form_is_one_dropdown_that_passes_the_sdk_validation():
    from mcp.server.elicitation import _validate_elicitation_schema

    _validate_elicitation_schema(settings_form._form_model(list(us.DECISION_SETTINGS)))
    prop = _form_property()
    assert prop["type"] == "string"
    assert prop["enum"] == [
        "Never (don't save screenshots)", "7 days", "30 days (default)", "Forever",
    ]
    assert prop["default"] == "30 days (default)"
    assert "oneOf" not in prop and "anyOf" not in prop


@pytest.mark.parametrize("label, value", [
    ("Forever", "forever"), ("Never (don't save screenshots)", "never"), ("7 days", "7"),
])
def test_each_choice_saves_its_value_and_is_selected_next_time(label, value):
    _configure(_Ctx(FORM, data={NAME: label}))
    assert us.get_setting(KEY) == value
    assert _form_property()["default"] == label


def test_the_default_label_saves_a_default_decision(settings_file):
    _configure(_Ctx(FORM, data={NAME: "30 days (default)"}))
    assert json.loads(settings_file.read_text(encoding="utf-8"))[KEY]["decision"] == "default"


def test_a_custom_saved_value_is_offered_and_selected():
    us.save({NAME: "45"}, "cli")
    prop = _form_property()
    assert prop["enum"] == [
        "Never (don't save screenshots)", "7 days", "30 days (default)",
        "45 days (current)", "Forever",
    ]
    assert prop["default"] == "45 days (current)"
    us.save({NAME: "30"}, "cli")
    assert _form_property()["default"] == "30 days (default)"


def test_free_text_from_a_client_without_dropdowns_is_parsed():
    _configure(_Ctx(FORM, data={NAME: "14"}))
    assert us.get_setting(KEY) == "14"


def test_unreadable_free_text_saves_nothing_and_asks_through_the_agent(settings_file):
    result = _configure(_Ctx(FORM, data={NAME: "0"}))
    assert result["form"] == "failed"
    assert result["agent_guidance"] == settings_form.ASK_USER_GUIDANCE
    assert not settings_file.exists()


@pytest.mark.parametrize("action", ["decline", "cancel"])
def test_declined_or_cancelled_form_saves_nothing(action, settings_file):
    result = _configure(_Ctx(FORM, action=action))
    assert result["form"] == ("declined" if action == "decline" else "cancelled")
    assert "saved" not in result
    assert result["pending_settings"] == [NAME]
    assert not settings_file.exists()
    assert "agent_guidance" not in result


def test_form_timeout_saves_nothing(monkeypatch, settings_file):
    monkeypatch.setattr(settings_form, "FORM_TIMEOUT_SEC", 0.05)
    result = _configure(_Ctx(FORM, delay=1.0))
    assert result["form"] == "timeout"
    assert not settings_file.exists()


def test_form_error_falls_back_to_asking_through_the_agent(settings_file):
    result = _configure(_Ctx(FORM, error=RuntimeError("client refused")))
    assert result["form"] == "failed"
    assert result["agent_guidance"] == settings_form.ASK_USER_GUIDANCE
    assert not settings_file.exists()


def test_no_form_support_tells_the_agent_to_ask():
    ctx = _Ctx(None)
    result = _configure(ctx)
    assert result["form"] == "unavailable"
    assert ctx.asked == 0
    assert result["agent_guidance"] == settings_form.ASK_USER_GUIDANCE
    assert "Do not choose for them" in result["agent_guidance"]


def test_no_form_when_the_environment_sets_everything(monkeypatch):
    monkeypatch.setenv(KEY, "7")
    ctx = _Ctx(FORM)
    result = _configure(ctx)
    assert result["form"] == "not_needed"
    assert ctx.asked == 0
    assert result["pending_settings"] == []


def test_values_are_saved_with_decided_by_agent():
    ctx = _Ctx(FORM)
    result = _configure(ctx, **{NAME: 14})
    assert result["saved"] == [NAME]
    assert ctx.asked == 0
    assert us.describe(us.SCREENSHOT_RETENTION)["decided_by"] == "agent"
    assert us.get_setting(KEY) == "14"


def test_default_value_is_stored_as_a_decision():
    result = _configure(_Ctx(None), **{NAME: "default"})
    assert result["pending_settings"] == []
    assert us.describe(us.SCREENSHOT_RETENTION)["value"] == "30"


@pytest.mark.parametrize("value", ["-3", 0, "0"])
def test_invalid_value_is_refused_and_the_agent_is_told_to_ask(value, settings_file):
    result = _configure(_Ctx(None), **{NAME: value})
    assert result["success"] is False
    assert result["error_pattern"] == "invalid_argument"
    assert result["agent_guidance"] == settings_form.ASK_USER_GUIDANCE
    assert not settings_file.exists()


@pytest.mark.parametrize("word", ["forever", "never"])
def test_words_are_accepted_from_the_agent(word):
    result = _configure(_Ctx(None), **{NAME: word})
    assert result["saved"] == [NAME]
    assert us.describe(us.SCREENSHOT_RETENTION)["display"] == word


def test_saving_under_an_environment_value_warns(monkeypatch):
    monkeypatch.setenv(KEY, "7")
    result = _configure(_Ctx(None), **{NAME: 14})
    assert any(KEY in w for w in result["warnings"])


def test_the_tool_maps_its_argument():
    tool = _unwrap(tools_module.vcs_configure)
    result = asyncio.run(tool(screenshot_retention=21, ctx=None))
    assert result["saved"] == [NAME]
    assert us.get_setting(KEY) == "21"


def test_version_info_reports_settings(monkeypatch):
    import msaccess_vcs_mcp.validation as validation

    monkeypatch.setattr(validation, "get_version_info_safe", lambda: {"success": True})
    result = asyncio.run(_unwrap(tools_module.vcs_get_version_info)(ctx=None))
    assert result["pending_settings"] == [NAME]
    assert result["settings"][0]["name"] == NAME
    assert result["settings_path"].endswith("settings.json")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _cli(*argv):
    return build_parser().parse_args(["configure", *argv])


def test_cli_set_saves_with_decided_by_cli(capsys):
    assert run_configure(_cli("--set", f"{NAME}=60")) == 0
    assert us.get_setting(KEY) == "60"
    assert us.describe(us.SCREENSHOT_RETENTION)["decided_by"] == "cli"
    assert f"{NAME} = 60 days" in capsys.readouterr().out


def test_cli_words_and_zero(capsys):
    assert run_configure(_cli("--set", f"{NAME}=forever")) == 0
    assert f"{NAME} = forever" in capsys.readouterr().out
    assert run_configure(_cli("--set", f"{NAME}=never")) == 0
    assert f"{NAME} = never" in capsys.readouterr().out
    assert run_configure(_cli("--set", f"{NAME}=0")) == 2
    assert '"never", "forever"' in capsys.readouterr().err
    assert us.get_setting(KEY) == "never"


def test_cli_interactive_accepts_forever():
    prompts = []

    def answer(prompt):
        prompts.append(prompt)
        return "forever"

    assert run_configure(_cli(), input_fn=answer, is_tty=True) == 0
    assert us.get_setting(KEY) == "forever"
    assert '"never", "forever"' in prompts[0] and "30 days" in prompts[0]


def test_cli_accepts_the_environment_variable_name():
    assert run_configure(_cli("--set", f"{KEY}=default")) == 0
    assert us.pending_settings() == []


def test_cli_bad_set_saves_nothing_and_resets_nothing(capsys):
    us.save({NAME: "45"}, "cli")
    assert run_configure(_cli("--reset", NAME, "--set", f"{NAME}=-1")) == 2
    assert us.get_setting(KEY) == "45"
    assert run_configure(_cli("--set", "unknown_setting=1")) == 2
    assert "Unknown setting" in capsys.readouterr().err


def test_cli_reset():
    us.save({NAME: "45"}, "cli")
    assert run_configure(_cli("--reset", NAME)) == 0
    assert us.pending_settings() == [NAME]


def test_cli_show_writes_nothing(capsys, settings_file):
    assert run_configure(_cli("--show")) == 0
    out = capsys.readouterr().out
    assert "not chosen yet" in out
    assert not settings_file.exists()


def test_cli_without_a_terminal_refuses_to_prompt(capsys, settings_file):
    def no_input(prompt):
        raise AssertionError("prompted without a terminal")

    assert run_configure(_cli(), input_fn=no_input, is_tty=False) == 2
    assert "needs a terminal" in capsys.readouterr().err
    assert not settings_file.exists()


def test_cli_interactive_enter_chooses_the_default():
    assert run_configure(_cli(), input_fn=lambda prompt: "", is_tty=True) == 0
    assert us.pending_settings() == []
    assert us.describe(us.SCREENSHOT_RETENTION)["source"] == us.SOURCE_DEFAULT


def test_cli_interactive_retries_an_invalid_answer(capsys):
    answers = iter(["soon", "10"])
    assert run_configure(_cli(), input_fn=lambda prompt: next(answers), is_tty=True) == 0
    assert us.get_setting(KEY) == "10"
    assert "not valid" in capsys.readouterr().out


def test_cli_interactive_enter_keeps_a_decided_value():
    us.save({NAME: "45"}, "agent")
    assert run_configure(_cli(), input_fn=lambda prompt: "", is_tty=True) == 0
    assert us.describe(us.SCREENSHOT_RETENTION)["decided_by"] == "agent"
