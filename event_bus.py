"""
MC & S Desktop Agent — Event Bus (Stream 3 — APEX Upgrade)
===========================================================
Provides a lightweight publish/subscribe event system so plugins and the
app can react to things happening in real time, rather than relying solely
on the polling scheduler.

DESIGN
------
  EventBus is a module-level singleton.  Any component can:

    1. Subscribe to an event type:
         EventBus.subscribe("email.received", my_handler)

    2. Publish an event:
         EventBus.publish(Event("email.received", {"subject": "...", "from": "..."}))

  Handlers are called synchronously in the publishing thread by default.
  For long-running handlers, set async_dispatch=True on subscribe() to
  dispatch in a background thread.

BUILT-IN EVENT TYPES
--------------------
  email.received          — new email arrived in inbox
  email.sent              — email was sent (draft approved and dispatched)
  noa.processed           — NOA plugin finished processing an NOA
  outreach.drafted        — outreach draft created for a client
  plugin.run.complete     — any plugin finished a run (payload: plugin_id, result)
  plugin.run.failed       — a plugin run raised an unhandled exception
  heartbeat.tick          — fired every N seconds by the HeartbeatPlugin
  app.started             — application fully initialised
  app.stopping            — application is shutting down

USAGE IN PLUGINS
----------------
    from event_bus import EventBus, Event

    # Subscribe (usually in plugin.load())
    EventBus.subscribe("email.received", self._on_email_received)

    # Publish
    EventBus.publish(Event("noa.processed", {
        "client_email": "john@example.com",
        "outcome": "REFUND",
        "amount": "$2,400",
    }))

    # Unsubscribe (usually in plugin.stop())
    EventBus.unsubscribe("email.received", self._on_email_received)
"""

import threading
import time
import traceback
from dataclasses import dataclass, field
from typing import Any, Callable


# ── Event dataclass ───────────────────────────────────────────────────────────

@dataclass
class Event:
    """
    A single event published on the bus.

    Attributes
    ----------
    type     : Dot-separated event type string, e.g. "email.received"
    payload  : Arbitrary dict of event data.
    source   : Optional string identifying the publisher.
    timestamp: Unix timestamp of when the event was created.
    """
    type: str
    payload: dict = field(default_factory=dict)
    source: str = ""
    timestamp: float = field(default_factory=time.time)


# ── EventBus ──────────────────────────────────────────────────────────────────

class _EventBus:
    """
    Internal singleton implementation.  Use the module-level `EventBus`
    instance rather than instantiating this class directly.
    """

    def __init__(self):
        # { event_type: [(handler, async_dispatch), ...] }
        self._handlers: dict[str, list[tuple[Callable, bool]]] = {}
        self._lock = threading.Lock()
        self._history: list[Event] = []
        self._history_max = 200  # keep last N events for diagnostics

    # ── Subscription management ───────────────────────────────────────────────

    def subscribe(
        self,
        event_type: str,
        handler: Callable[[Event], None],
        async_dispatch: bool = False,
    ) -> None:
        """
        Register a handler for an event type.

        Parameters
        ----------
        event_type     : The event type string to listen for.
                         Use "*" to subscribe to ALL event types.
        handler        : Callable that accepts a single Event argument.
        async_dispatch : If True, the handler is called in a daemon thread
                         so it doesn't block the publisher.
        """
        with self._lock:
            self._handlers.setdefault(event_type, [])
            # Avoid duplicate subscriptions
            existing = [h for h, _ in self._handlers[event_type]]
            if handler not in existing:
                self._handlers[event_type].append((handler, async_dispatch))

    def unsubscribe(
        self,
        event_type: str,
        handler: Callable[[Event], None],
    ) -> None:
        """Remove a handler for an event type."""
        with self._lock:
            if event_type in self._handlers:
                self._handlers[event_type] = [
                    (h, a) for h, a in self._handlers[event_type]
                    if h is not handler
                ]

    def unsubscribe_all(self, handler: Callable[[Event], None]) -> None:
        """Remove a handler from ALL event types it is subscribed to."""
        with self._lock:
            for event_type in list(self._handlers.keys()):
                self._handlers[event_type] = [
                    (h, a) for h, a in self._handlers[event_type]
                    if h is not handler
                ]

    # ── Publishing ────────────────────────────────────────────────────────────

    def publish(self, event: Event) -> int:
        """
        Publish an event to all registered handlers.

        Returns the number of handlers that were called.
        """
        # Collect handlers under lock, then dispatch outside lock
        with self._lock:
            specific = list(self._handlers.get(event.type, []))
            wildcard = list(self._handlers.get("*", []))
            # Record in history
            self._history.append(event)
            if len(self._history) > self._history_max:
                self._history = self._history[-self._history_max:]

        all_handlers = specific + wildcard
        for handler, async_dispatch in all_handlers:
            if async_dispatch:
                t = threading.Thread(
                    target=self._safe_call, args=(handler, event), daemon=True
                )
                t.start()
            else:
                self._safe_call(handler, event)

        return len(all_handlers)

    def emit(self, event_type: str, payload: dict = None, source: str = "") -> int:
        """Convenience wrapper: create and publish an Event in one call."""
        return self.publish(Event(
            type=event_type,
            payload=payload or {},
            source=source,
        ))

    # ── Diagnostics ───────────────────────────────────────────────────────────

    def get_history(self, event_type: str = None, limit: int = 50) -> list[Event]:
        """Return recent event history, optionally filtered by type."""
        with self._lock:
            history = list(self._history)
        if event_type:
            history = [e for e in history if e.type == event_type]
        return history[-limit:]

    def subscriber_count(self, event_type: str = None) -> int:
        """Return the number of registered handlers (optionally for a specific type)."""
        with self._lock:
            if event_type:
                return len(self._handlers.get(event_type, []))
            return sum(len(v) for v in self._handlers.values())

    def clear_history(self) -> None:
        """Clear the event history buffer."""
        with self._lock:
            self._history.clear()

    # ── Internal ──────────────────────────────────────────────────────────────

    @staticmethod
    def _safe_call(handler: Callable, event: Event) -> None:
        """Call a handler, swallowing exceptions so one bad handler can't break others."""
        try:
            handler(event)
        except Exception:
            # Print to stderr so it's visible in logs without crashing the bus
            traceback.print_exc()


# Module-level singleton — import and use this directly
EventBus = _EventBus()


# ── HeartbeatPlugin ───────────────────────────────────────────────────────────

class HeartbeatPlugin:
    """
    Fires a "heartbeat.tick" event on the EventBus at a configurable interval.

    This gives plugins a way to perform lightweight periodic checks without
    needing their own scheduler thread.  The heartbeat runs in a single
    daemon thread shared across all subscribers.

    Usage
    -----
        from event_bus import EventBus, Event, HeartbeatPlugin

        # Start the heartbeat (call once at app startup)
        HeartbeatPlugin.start(interval_seconds=60)

        # Subscribe to ticks
        def on_tick(event: Event):
            print("Tick!", event.payload["tick_count"])

        EventBus.subscribe("heartbeat.tick", on_tick)

        # Stop (call at app shutdown)
        HeartbeatPlugin.stop()
    """

    _thread: threading.Thread | None = None
    _running: bool = False
    _interval: int = 60
    _tick_count: int = 0
    _lock = threading.Lock()

    @classmethod
    def start(cls, interval_seconds: int = 60) -> None:
        """Start the heartbeat thread.  Safe to call multiple times."""
        with cls._lock:
            if cls._running:
                return
            # Allow intervals as low as 1 second (useful for tests and fast-tick scenarios)
            cls._interval = max(1, interval_seconds)
            cls._running = True
            cls._tick_count = 0
            cls._thread = threading.Thread(
                target=cls._loop, daemon=True, name="HeartbeatPlugin"
            )
            cls._thread.start()

    @classmethod
    def stop(cls) -> None:
        """Stop the heartbeat thread."""
        with cls._lock:
            cls._running = False

    @classmethod
    def set_interval(cls, interval_seconds: int) -> None:
        """Change the heartbeat interval (takes effect on the next tick)."""
        with cls._lock:
            cls._interval = max(5, interval_seconds)

    @classmethod
    def is_running(cls) -> bool:
        return cls._running

    @classmethod
    def _loop(cls) -> None:
        """
        Main heartbeat loop.
        Sleeps in 0.1-second increments so stop() takes effect quickly.
        This also means the minimum effective interval is ~0.1s (useful in tests).
        """
        while cls._running:
            # Sleep in small slices so we can respond to stop() promptly
            with cls._lock:
                target_interval = cls._interval
            elapsed = 0.0
            slice_size = min(0.1, target_interval)
            while elapsed < target_interval and cls._running:
                time.sleep(slice_size)
                elapsed += slice_size

            if not cls._running:
                break

            with cls._lock:
                cls._tick_count += 1
                tick = cls._tick_count
                interval = cls._interval
            EventBus.emit(
                "heartbeat.tick",
                payload={
                    "tick_count": tick,
                    "interval_seconds": interval,
                    "timestamp": time.time(),
                },
                source="HeartbeatPlugin",
            )
