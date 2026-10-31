"""One run at a time. A lock file is created at the start of a run and removed at the end.
A lock is taken over at once when the process that made it is gone (a crash, a power cut, a restart), and in any case when it is
older than the run time limit (+15 minutes)."""
import os
import re
import time


class LockBusy(Exception):
    pass


def pid_alive(pid):
    """Is a process with this id running? Any doubt means True (the age rule then decides)."""
    if pid <= 0:
        return False
    if pid == os.getpid():
        return True
    try:
        if os.name == "nt":
            import ctypes
            k = ctypes.WinDLL("kernel32", use_last_error=True)
            handle = k.OpenProcess(0x1000, False, pid)                  # PROCESS_QUERY_LIMITED_INFORMATION
            if not handle:
                return ctypes.get_last_error() == 5                      # 5 = access denied: it exists, we may not look at it
            code = ctypes.c_ulong()
            ok = k.GetExitCodeProcess(handle, ctypes.byref(code))
            k.CloseHandle(handle)
            return (not ok) or code.value == 259                         # 259 = STILL_ACTIVE
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except Exception:
        return True


class RunLock:
    def __init__(self, path, stale_minutes=60):
        self.path, self.stale = path, stale_minutes * 60
        self.held = False
        self.took_over_from = None                  # the process id of a run that did not finish, when its lock was taken over

    def __enter__(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        for _attempt in range(2):
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, f"pid={os.getpid()} started={time.strftime('%Y-%m-%d %H:%M:%S')}".encode())
                os.close(fd)
                self.held = True
                return self
            except FileExistsError:
                try:
                    age = time.time() - os.path.getmtime(self.path)
                except OSError:
                    continue
                owner = self._owner()
                if age > self.stale or (owner and not pid_alive(owner)):
                    try:
                        os.remove(self.path)                  # left over from a run that did not finish
                        self.took_over_from = owner or 0
                    except OSError:
                        pass
                    continue
                raise LockBusy(f"another run is in progress (lock file {self.path}, {int(age // 60)} min old)")
        raise LockBusy("could not take the run lock")

    def _owner(self):
        try:
            with open(self.path, encoding="utf-8", errors="replace") as f:
                m = re.search(r"pid=(\d+)", f.read())
            return int(m.group(1)) if m else None
        except OSError:
            return None

    def __exit__(self, *exc):
        if self.held:
            try:
                os.remove(self.path)
            except OSError:
                pass
        return False
