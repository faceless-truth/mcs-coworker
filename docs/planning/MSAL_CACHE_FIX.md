# MSAL_CACHE_FIX.md — Claude Code Task List

Fix the MSAL token cache so it persists after Microsoft sign-in. Currently the cache file is never written, which means every app restart requires re-authentication and SharePoint integration cannot work.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Ensure MSAL token cache persists after every authentication event

**Files:** `graph_client.py`

**Problem:** The MSAL token cache file (`.msal_cache.bin`) is never created. After the user signs into Microsoft through the app, the token is held in memory but `_save_token_cache()` is never called — or it's called but `cache.has_state_changed` is False so it skips the write. On the next app restart, the user has to sign in again. SharePoint test connection also fails because the command-line test can't find a cached token.

**Root cause investigation:**

The `_save_token_cache` function likely has this guard:
```python
def _save_token_cache(cache):
    if cache.has_state_changed:
        _get_cache_path().write_text(cache.serialize())
```

This is correct in theory — MSAL sets `has_state_changed = True` when tokens are acquired. But if `_save_token_cache` is never called after token acquisition, the flag is irrelevant.

**Changes required:**

### Part A — Add save calls after every token acquisition path

1. Find the `authenticate()` method in `GraphClient`. This is the interactive sign-in flow. After a successful `acquire_token_by_authorization_code` (or whatever method is used), add:
   ```python
   # After successful token acquisition:
   _save_token_cache(self._cache)
   logger.info("MSAL token cache saved after authentication")
   ```

2. Find the `_get_token()` method. This does silent token acquisition/refresh. After a successful `acquire_token_silent`, add:
   ```python
   # After successful silent acquisition:
   _save_token_cache(self._cache)
   ```

3. Find the `is_authenticated()` method. If it calls `acquire_token_silent`, add the same save call after success.

4. Search for ANY other place in the file that calls `acquire_token_*` on the MSAL app:
   ```
   grep -n "acquire_token" graph_client.py
   ```
   Add `_save_token_cache(self._cache)` after every successful acquisition.

### Part B — Fix _save_token_cache to create the file on first run

5. Update `_save_token_cache` to write the file even on first run when no cache exists yet:
   ```python
   def _save_token_cache(cache):
       """Write the MSAL token cache to disk."""
       cache_path = _get_cache_path()
       try:
           # Write if state changed OR if the file doesn't exist yet (first run)
           if cache.has_state_changed or not cache_path.exists():
               cache_path.parent.mkdir(parents=True, exist_ok=True)
               cache_path.write_text(cache.serialize())
               logger.info(f"MSAL token cache saved to {cache_path} (state_changed={cache.has_state_changed})")
           else:
               logger.debug(f"MSAL token cache unchanged, skipping save")
       except Exception as e:
           logger.error(f"Failed to save MSAL token cache: {e}")
   ```

### Part C — Verify atexit handler

6. In `GraphClient.__init__`, verify there's an atexit handler that saves the cache on shutdown:
   ```python
   import atexit
   
   class GraphClient:
       def __init__(self):
           self._cache = _load_token_cache()
           # ... existing init code ...
           
           # Save cache on shutdown
           atexit.register(lambda: _save_token_cache(self._cache))
   ```

   If the atexit handler exists but uses `self._persist_cache()` or similar, make sure that method actually calls `_save_token_cache(self._cache)`.

### Part D — Add debug logging

7. Add a log line at the start of `_save_token_cache`:
   ```python
   logger.debug(f"_save_token_cache called: has_state_changed={cache.has_state_changed}, path={_get_cache_path()}, path_exists={_get_cache_path().exists()}")
   ```

8. Add a log line in `_load_token_cache`:
   ```python
   logger.info(f"MSAL cache loaded from {cache_path}: {'found' if cache_path.exists() else 'not found (fresh start)'}")
   ```

### Part E — Verify the authentication flow end-to-end

9. Check the full `authenticate()` method flow. In CoWorker, the auth typically works like this:
   - Flask serves an `/auth/callback` or `/oauth/callback` route
   - User is redirected to Microsoft login
   - Microsoft redirects back with an auth code
   - The callback handler calls `acquire_token_by_authorization_code(code, scopes)`
   - The result contains `access_token`
   
   Find this callback handler (it may be in `api_server.py` rather than `graph_client.py`). If the token acquisition happens in `api_server.py`, that's where `_save_token_cache` needs to be called:
   ```python
   # In the OAuth callback handler in api_server.py:
   result = graph_client.msal_app.acquire_token_by_authorization_code(code, scopes, redirect_uri)
   if "access_token" in result:
       # THIS IS WHERE THE SAVE MUST HAPPEN:
       from graph_client import _save_token_cache
       _save_token_cache(graph_client._cache)
   ```

10. Search both files for the auth code exchange:
    ```
    grep -n "acquire_token_by_authorization_code\|authorization_code" graph_client.py api_server.py
    ```
    Add `_save_token_cache` after every successful result.

**Test:**

After the fix:
```cmd
del "C:\Users\Elio\AppData\Local\Programs\MCS CoWorker\data\.msal_cache.bin" 2>nul
python main.py
```

Sign into Microsoft when prompted. Then check:
```cmd
dir /s /b C:\Users\Elio\*.msal_cache.bin
```

The file should now exist. Then stop the app and verify the token persists:
```cmd
python -c "from graph_client import GraphClient; g = GraphClient(); print('Authenticated:', g.is_authenticated())"
```

Should print `Authenticated: True` without prompting for sign-in.

**Commit message:** `fix: ensure MSAL token cache persists after every authentication event`

---

## Post-fix checklist

- [ ] Delete old MSAL cache files
- [ ] Start CoWorker, sign into Microsoft
- [ ] `.msal_cache.bin` file is created
- [ ] Restart CoWorker — no sign-in prompt (cached token works)
- [ ] Settings → SharePoint → Test Connection → green tick
- [ ] Command-line test: `python -c "from graph_client import GraphClient; g = GraphClient(); print(g.is_authenticated())"` → True
