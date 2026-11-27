"""``vcs_configure``: record setting decisions, by form or from values the agent relays.

The form opens only when this tool is called on purpose, never inside an
unrelated call: an agent or a client hook can answer a form, some clients
declare form support and decline every request unseen, and headless runs have
nobody to answer. Decline, cancel, and timeout therefore save nothing; only an
accepted form or an explicit value is a decision.
"""

from __future__ import annotations

import asyncio
from typing import Any

from pydantic import Field, create_model

from . import user_settings
from .user_settings import SCREENSHOT_RETENTION

FORM_TIMEOUT_SEC = 120.0

FORM_MESSAGE = "msaccess-vcs-mcp: how long should screenshots saved for the chat be kept?"

ASK_USER_GUIDANCE = (
    "Ask the user how long to keep screenshots saved for the chat: \"never\" (save none), "
    "7 or 30 days, or \"forever\". Then call vcs_configure(screenshot_retention=...) with "
    "their answer. Do not choose for them."
)

NEVER_LABEL = "Never (don't save screenshots)"
FOREVER_LABEL = "Forever"
RETENTION_PRESET_DAYS = (7, 30)


def client_supports_form(ctx: Any) -> bool:
    """True when the client declared form elicitation.

    ``ServerSession.check_client_capability`` accepts any elicitation
    capability, including a URL-only one, so the modes are read directly. An
    empty ``elicitation: {}`` is the 2025-06-18 form, which meant form mode.
    """
    try:
        params = ctx.session.client_params
    except Exception:
        return False
    capabilities = getattr(params, "capabilities", None)
    elicitation = getattr(capabilities, "elicitation", None)
    if elicitation is None:
        return False
    return elicitation.form is not None or elicitation.url is None


def retention_choices() -> tuple[dict[str, str], str]:
    """Form labels mapped to the value each saves, and the label to preselect.

    A plain ``enum`` of labels, not titled ``oneOf``: Cursor has been reported
    to show ``oneOf`` as a text box. A day count saved through the agent or the
    CLI that is not a preset is offered as an extra "(current)" choice.
    """
    setting = SCREENSHOT_RETENTION
    current = user_settings.describe(setting)
    saved = (
        current["value"]
        if current["decided"] and current["source"] == user_settings.SOURCE_FILE
        else None
    )
    days = set(RETENTION_PRESET_DAYS) | {int(setting.default)}
    if saved and saved.isdigit():
        days.add(int(saved))
    choices = {NEVER_LABEL: user_settings.NEVER}
    selected = f"{setting.default} days (default)"
    for count in sorted(days):
        if str(count) == setting.default:
            choices[f"{count} days (default)"] = user_settings.DEFAULT_CHOICE
            continue
        label = f"{count} days" + ("" if count in RETENTION_PRESET_DAYS else " (current)")
        choices[label] = str(count)
    choices[FOREVER_LABEL] = user_settings.FOREVER
    for label, value in choices.items():
        if saved == value:
            selected = label
    return choices, selected


def _form_model(settings: list[user_settings.DecisionSetting]):
    fields: dict[str, Any] = {}
    for setting in settings:
        if setting is SCREENSHOT_RETENTION:
            choices, selected = retention_choices()
            fields[setting.name] = (
                str,
                Field(
                    default=selected,
                    title="Keep saved screenshots for",
                    description=(
                        "Screenshots saved for the chat are deleted after this period, and a "
                        "chat older than that shows them as broken. Never saves none: the "
                        "image appears only in the reply, and any already saved are deleted."
                    ),
                    json_schema_extra={"enum": list(choices)},
                ),
            )
    return create_model("SettingsForm", **fields)


def _answer_values(answers: dict[str, Any]) -> dict[str, str]:
    """Form labels to stored values; free text from a client without dropdowns is parsed."""
    choices, _ = retention_choices()
    values: dict[str, str] = {}
    for name, answer in answers.items():
        setting = user_settings.find_setting(name)
        values[name] = choices.get(str(answer)) or user_settings.normalize(setting, answer)
    return values


async def _ask(ctx: Any, settings: list[user_settings.DecisionSetting]) -> tuple[str, dict[str, Any]]:
    from .usage_logging import log_diagnostic_event

    try:
        answer = await asyncio.wait_for(
            ctx.elicit(message=FORM_MESSAGE, schema=_form_model(settings)),
            FORM_TIMEOUT_SEC,
        )
    except asyncio.TimeoutError:
        return "timeout", {}
    except Exception as exc:
        log_diagnostic_event("settings_form_failed", error_type=type(exc).__name__)
        return "failed", {}
    if answer.action == "accept":
        try:
            return "accepted", _answer_values(answer.data.model_dump())
        except ValueError:
            return "failed", {}
    return ("declined" if answer.action == "decline" else "cancelled"), {}


def _env_overrides(names: list[str]) -> list[str]:
    return [
        f"{d['env_var']} is set in the environment, which wins over the saved value."
        for d in user_settings.describe_all()
        if d["name"] in names and d["source"] == user_settings.SOURCE_ENV
    ]


async def configure(ctx: Any, values: dict[str, Any]) -> dict[str, Any]:
    if values:
        try:
            saved = user_settings.save(values, "agent")
        except (KeyError, ValueError) as exc:
            return {
                "success": False,
                "error": str(exc).strip("'\""),
                "error_pattern": "invalid_argument",
                "agent_guidance": ASK_USER_GUIDANCE,
                "settings": user_settings.describe_all(),
            }
        result: dict[str, Any] = {"success": True, "saved": saved}
        warnings = _env_overrides(saved)
        if warnings:
            result["warnings"] = warnings
    else:
        askable = [
            s for s in user_settings.DECISION_SETTINGS
            if user_settings.get_setting_source(s.key) != user_settings.SOURCE_ENV
        ]
        if not askable:
            form = "not_needed"
        elif not client_supports_form(ctx):
            form = "unavailable"
        else:
            form, answers = await _ask(ctx, askable)
        result = {"success": True, "form": form}
        if form == "accepted":
            result["saved"] = user_settings.save(answers, "form")

    result["settings_path"] = str(user_settings.settings_path())
    result["settings"] = user_settings.describe_all()
    pending = user_settings.pending_settings()
    result["pending_settings"] = pending
    if pending and result.get("form") in ("unavailable", "failed"):
        result["agent_guidance"] = ASK_USER_GUIDANCE
    return result
