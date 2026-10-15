# write_lock.py
"""One lock per Excel file, so a manual batch and the auto-watcher never write the
same file at the same time (the second writer would overwrite the first one's rows).

Both sides ask for the lock with blocking=False: if the file is busy they say so and
try again later instead of waiting. The key is the local path of the file, which is
the same for both because they share LOCAL_BATCH_DIR."""
import os
import threading

_locks = {}
_guard = threading.Lock()


def lock_for(path):
    key = os.path.normcase(os.path.abspath(path)).casefold()
    with _guard:
        if key not in _locks:
            _locks[key] = threading.Lock()
        return _locks[key]
