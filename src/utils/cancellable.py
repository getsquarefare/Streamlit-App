import threading
import time

from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx


class CancellableTask:
    """Runs a long-running function in a background thread with cooperative cancel.

    The target function must accept a `cancel_event` keyword argument (a
    threading.Event) and check it at loop boundaries. When `.cancel()` is
    called, the event is set; it's the target's responsibility to stop work.
    """

    def __init__(self, target, *args, **kwargs):
        self.cancel_event = threading.Event()
        self.result = None
        self.error = None
        self.done = False
        self.start_time = None
        self.end_time = None

        self._target = target
        self._args = args
        self._kwargs = kwargs
        self._thread = threading.Thread(target=self._run, daemon=True)

        ctx = get_script_run_ctx()
        if ctx is not None:
            add_script_run_ctx(self._thread, ctx)

    def _run(self):
        try:
            self._kwargs["cancel_event"] = self.cancel_event
            self.result = self._target(*self._args, **self._kwargs)
        except Exception as e:
            self.error = e
        finally:
            self.end_time = time.time()
            self.done = True

    def start(self):
        self.start_time = time.time()
        self._thread.start()

    def cancel(self):
        self.cancel_event.set()

    def is_done(self):
        return self.done

    def is_cancelled(self):
        return self.cancel_event.is_set()

    def elapsed(self):
        end = self.end_time if self.end_time is not None else time.time()
        start = self.start_time if self.start_time is not None else end
        return end - start
