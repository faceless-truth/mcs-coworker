"""
MCS CoWorker — Multi-Agent Event Wiring (Tier 3A)
==================================================
This module wires the existing plugins together via the Event Bus so they
react to each other's outcomes, not just run on independent timers.

APEX ALIGNMENT
--------------
APEX's "orchestration layer" routes outcomes from one agent to trigger
actions in another. This module implements that pattern for CoWorker.

EVENT TAXONOMY
--------------
Published by existing plugins (after this wiring is applied):

  email.triage.complete
    payload: {category, sender, subject, msg_id, confidence}
    → triggers: NOA Processor (if DOCUMENTS_RECEIVED)
    → triggers: Engagement Letter (if NEW_CLIENT)
    → triggers: Meeting Prep (if MEETING_REQUEST)
    → triggers: Correspondence Logger (always)

  noa.processed
    payload: {client_name, client_email, outcome, tax_year}
    → triggers: Correspondence Logger
    → triggers: Memory store

  asic.processed
    payload: {company_name, client_email, outcome}
    → triggers: Correspondence Logger

  fusesign.completed
    payload: {envelope_id, client_name, client_email, document_name}
    → triggers: XPM job note
    → triggers: Teams notification

  annual_review.prompted
    payload: {client_name, email}
    → triggers: Meeting Prep (pre-loads client context)

  morning_briefing.sent
    payload: {date, jobs_count, inbox_count}
    → triggers: nothing (terminal event)

  wip_summary.sent
    payload: {date, total_jobs, overdue_count}
    → triggers: Teams alert if overdue_count > threshold

USAGE
-----
Call `wire_all(event_bus, plugin_registry)` once at startup from
plugin_loader.py after all plugins are loaded.
"""

from __future__ import annotations
import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from event_bus import EventBus

logger = logging.getLogger(__name__)


# ── Event constants ───────────────────────────────────────────────────────────

class Events:
    """Canonical event names used across all plugins."""
    # Email triage
    EMAIL_TRIAGE_COMPLETE   = "email.triage.complete"
    EMAIL_TRIAGE_BATCH_DONE = "email.triage.batch_done"

    # Document processing
    NOA_PROCESSED           = "noa.processed"
    ASIC_PROCESSED          = "asic.processed"
    DOCUMENTS_RECEIVED      = "email.category.documents_received"

    # Client lifecycle
    NEW_CLIENT_DETECTED     = "email.category.new_client"
    MEETING_REQUESTED       = "email.category.meeting_request"
    ANNUAL_REVIEW_PROMPTED  = "annual_review.prompted"

    # FuseSign
    FUSESIGN_COMPLETED      = "fusesign.completed"
    FUSESIGN_OVERDUE        = "fusesign.overdue"

    # Briefings
    MORNING_BRIEFING_SENT   = "morning_briefing.sent"
    WIP_SUMMARY_SENT        = "wip_summary.sent"

    # Heartbeat
    HEARTBEAT_TICK          = "heartbeat.tick"

    # App lifecycle
    APP_STARTED             = "app.started"
    APP_STOPPING            = "app.stopping"

    # Plugin outcomes
    PLUGIN_RUN_COMPLETE     = "plugin.run.complete"
    PLUGIN_RUN_FAILED       = "plugin.run.failed"


# ── Wiring logic ──────────────────────────────────────────────────────────────

def wire_all(event_bus: "EventBus", plugin_registry: dict) -> None:
    """
    Subscribe all cross-plugin event handlers.
    Called once at startup after all plugins are loaded.

    Args:
        event_bus:       The global EventBus singleton.
        plugin_registry: Dict of {plugin_id: plugin_instance} from plugin_loader.
    """
    _wire_triage_to_processors(event_bus, plugin_registry)
    _wire_wip_to_teams_alert(event_bus, plugin_registry)
    _wire_fusesign_to_xpm(event_bus, plugin_registry)
    _wire_annual_review_to_meeting_prep(event_bus, plugin_registry)
    logger.info("[EventWiring] All cross-plugin event handlers registered.")


def _wire_triage_to_processors(event_bus: "EventBus", plugins: dict) -> None:
    """
    When email triage completes, route to the appropriate processor
    based on the email category — without waiting for the next scheduler tick.
    """

    def on_triage_complete(event):
        payload  = event.get("payload", {})
        category = payload.get("category", "").upper()
        msg_id   = payload.get("msg_id", "")

        logger.debug(f"[EventWiring] Triage complete: category={category}")

        # DOCUMENTS_RECEIVED → wake NOA Processor immediately
        if category == "DOCUMENTS_RECEIVED":
            noa = plugins.get("plugin_noa_processor")
            if noa and hasattr(noa, "_pending_msg_ids"):
                noa._pending_msg_ids.add(msg_id)
                logger.info(f"[EventWiring] NOA Processor notified of msg {msg_id}")
            event_bus.publish(Events.DOCUMENTS_RECEIVED,
                              source="event_wiring",
                              payload=payload)

        # NEW_CLIENT → wake Engagement Letter Generator
        elif category in ("NEW_CLIENT", "ONBOARDING"):
            eng = plugins.get("plugin_engagement_letter")
            if eng and hasattr(eng, "_pending_emails"):
                eng._pending_emails.add(msg_id)
                logger.info(f"[EventWiring] Engagement Letter notified of msg {msg_id}")
            event_bus.publish(Events.NEW_CLIENT_DETECTED,
                              source="event_wiring",
                              payload=payload)

        # MEETING_REQUEST → wake Meeting Prep
        elif category == "MEETING_REQUEST":
            meeting = plugins.get("plugin_meeting_prep")
            if meeting and hasattr(meeting, "_pending_emails"):
                meeting._pending_emails.add(msg_id)
                logger.info(f"[EventWiring] Meeting Prep notified of msg {msg_id}")
            event_bus.publish(Events.MEETING_REQUESTED,
                              source="event_wiring",
                              payload=payload)

    event_bus.subscribe(Events.EMAIL_TRIAGE_COMPLETE, on_triage_complete,
                        subscriber_id="wiring.triage_router")


def _wire_wip_to_teams_alert(event_bus: "EventBus", plugins: dict) -> None:
    """
    When a WIP summary is sent, check if overdue count exceeds threshold
    and send a Teams alert if so.
    """

    def on_wip_summary(event):
        payload      = event.get("payload", {})
        overdue      = payload.get("overdue_count", 0)
        total        = payload.get("total_jobs", 0)
        date_str     = payload.get("date", "")

        # Default threshold: alert if more than 10% of jobs are overdue 90+ days
        threshold = 5
        if overdue >= threshold:
            logger.info(f"[EventWiring] WIP alert: {overdue} overdue jobs (threshold={threshold})")
            # Find a plugin with gateway access to send the Teams alert
            for plugin_id, plugin in plugins.items():
                if hasattr(plugin, "_last_context") and plugin._last_context:
                    ctx = plugin._last_context
                    if ctx.gateway and ctx.gateway.is_available("teams"):
                        try:
                            ctx.gateway.teams.send_alert(
                                title="⚠️ WIP Overdue Alert",
                                message=(f"{overdue} jobs are 90+ days overdue "
                                         f"(out of {total} total). "
                                         f"Review required — {date_str}"),
                                color="warning")
                        except Exception as e:
                            logger.warning(f"[EventWiring] Teams WIP alert failed: {e}")
                        break

    event_bus.subscribe(Events.WIP_SUMMARY_SENT, on_wip_summary,
                        subscriber_id="wiring.wip_alert")


def _wire_fusesign_to_xpm(event_bus: "EventBus", plugins: dict) -> None:
    """
    When FuseSign reports a completed envelope, add a note to the
    corresponding XPM job automatically.
    """

    def on_fusesign_completed(event):
        payload       = event.get("payload", {})
        client_name   = payload.get("client_name", "")
        document_name = payload.get("document_name", "")
        envelope_id   = payload.get("envelope_id", "")

        if not client_name:
            return

        logger.info(f"[EventWiring] FuseSign completed: {document_name} for {client_name}")

        for plugin_id, plugin in plugins.items():
            if hasattr(plugin, "_last_context") and plugin._last_context:
                ctx = plugin._last_context
                if ctx.gateway and ctx.gateway.is_available("xpm"):
                    try:
                        clients = ctx.gateway.xpm.list_clients(
                            search=client_name, limit=1)
                        if clients:
                            client_id = clients[0].get("id") or clients[0].get("client_id")
                            if client_id:
                                ctx.gateway.xpm.add_client_note(
                                    client_id=client_id,
                                    note=(f"Document signed via FuseSign: "
                                          f"{document_name} (envelope {envelope_id})"))
                                logger.info(f"[EventWiring] XPM note added for {client_name}")
                    except Exception as e:
                        logger.warning(f"[EventWiring] XPM note failed: {e}")
                    break

    event_bus.subscribe(Events.FUSESIGN_COMPLETED, on_fusesign_completed,
                        subscriber_id="wiring.fusesign_xpm")


def _wire_annual_review_to_meeting_prep(event_bus: "EventBus", plugins: dict) -> None:
    """
    When an annual review invitation is sent, pre-load the client's
    context into the Meeting Prep plugin's cache so it's ready when
    the client replies to book a meeting.
    """

    def on_annual_review_prompted(event):
        payload     = event.get("payload", {})
        client_name = payload.get("client_name", "")
        email       = payload.get("email", "")

        if not client_name:
            return

        logger.info(f"[EventWiring] Annual review prompted for {client_name} — "
                    f"pre-loading Meeting Prep context.")

        meeting = plugins.get("plugin_meeting_prep")
        if meeting and hasattr(meeting, "_context_cache"):
            meeting._context_cache[email] = {
                "client_name": client_name,
                "source":      "annual_review_prompt",
                "preloaded":   True}

    event_bus.subscribe(Events.ANNUAL_REVIEW_PROMPTED, on_annual_review_prompted,
                        subscriber_id="wiring.annual_review_meeting_prep")


# ── Patch helpers for existing plugins ───────────────────────────────────────

def patch_email_triage_plugin(plugin, event_bus: "EventBus") -> None:
    """
    Monkey-patch the EmailTriagePlugin to publish events after each
    email is classified. Called by plugin_loader after loading.
    """
    original_run = plugin.run

    def patched_run(context):
        result = original_run(context)
        # Publish a batch-done event with summary stats
        event_bus.publish(
            Events.EMAIL_TRIAGE_BATCH_DONE,
            source="plugin_email_triage",
            payload={
                "actions_taken": result.actions_taken,
                "drafts_created": result.drafts_created,
            })
        return result

    plugin.run = patched_run
    logger.info("[EventWiring] EmailTriagePlugin patched to publish events.")


def patch_noa_processor_plugin(plugin, event_bus: "EventBus") -> None:
    """
    Monkey-patch the NOAProcessorPlugin to publish events after each
    NOA is processed.
    """
    original_run = plugin.run

    def patched_run(context):
        result = original_run(context)
        if result.actions_taken > 0:
            event_bus.publish(
                Events.NOA_PROCESSED,
                source="plugin_noa_processor",
                payload={"actions_taken": result.actions_taken})
        return result

    plugin.run = patched_run
    logger.info("[EventWiring] NOAProcessorPlugin patched to publish events.")


def apply_all_patches(plugins: dict, event_bus: "EventBus") -> None:
    """
    Apply all monkey-patches to existing plugins to make them
    publish events. Called by plugin_loader after loading.
    """
    if "plugin_email_triage" in plugins:
        patch_email_triage_plugin(plugins["plugin_email_triage"], event_bus)
    if "plugin_noa_processor" in plugins:
        patch_noa_processor_plugin(plugins["plugin_noa_processor"], event_bus)
    logger.info("[EventWiring] All plugin patches applied.")
