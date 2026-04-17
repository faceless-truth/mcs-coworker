"""Debug script to trace the heartbeat tick issue."""
import time
import threading
from event_bus import EventBus, HeartbeatPlugin

# Stop any running heartbeat
HeartbeatPlugin.stop()
time.sleep(0.5)

print(f"HeartbeatPlugin._running before start: {HeartbeatPlugin._running}")
print(f"HeartbeatPlugin._interval: {HeartbeatPlugin._interval}")

ticks = []
done = threading.Event()

def on_tick(e):
    print(f"TICK received: {e.type} payload={e.payload}")
    ticks.append(e)
    done.set()

EventBus.subscribe("heartbeat.tick", on_tick)
print(f"Subscribers on heartbeat.tick: {EventBus.subscriber_count('heartbeat.tick')}")

HeartbeatPlugin.start(interval_seconds=1)
print(f"HeartbeatPlugin._running after start: {HeartbeatPlugin._running}")
print(f"HeartbeatPlugin._interval after start: {HeartbeatPlugin._interval}")
print(f"HeartbeatPlugin._thread: {HeartbeatPlugin._thread}")
print(f"Thread alive: {HeartbeatPlugin._thread.is_alive() if HeartbeatPlugin._thread else 'No thread'}")

print("Waiting 3 seconds for tick...")
time.sleep(3)

print(f"Ticks received: {len(ticks)}")
print(f"HeartbeatPlugin._running: {HeartbeatPlugin._running}")
print(f"HeartbeatPlugin._tick_count: {HeartbeatPlugin._tick_count}")

HeartbeatPlugin.stop()
EventBus.unsubscribe("heartbeat.tick", on_tick)
