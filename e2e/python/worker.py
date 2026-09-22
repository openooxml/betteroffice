"""Execute timed requests in a persistent Python interpreter."""

import base64
import json
import sys
import time
import traceback


class Timed:
    def __init__(self):
        self.ms = {}

    def __call__(self, stage):
        return _Span(self, stage)


class _Span:
    def __init__(self, timed, stage):
        self.timed = timed
        self.stage = stage

    def __enter__(self):
        self.start = time.perf_counter()
        return self

    def __exit__(self, *exc):
        elapsed = (time.perf_counter() - self.start) * 1000.0
        self.timed.ms[self.stage] = self.timed.ms.get(self.stage, 0.0) + elapsed
        return False


state = {}
for line in sys.stdin:
    request = json.loads(line)
    timed = Timed()
    scope = {"state": state, "base64": base64, "json": json}
    try:
        exec(request["script"], scope)
        result = scope["run"](state, request["input"], timed)
        response = {"id": request["id"], "ok": True, "result": result, "timings": timed.ms}
    except Exception:
        response = {"id": request["id"], "ok": False, "error": traceback.format_exc(), "timings": timed.ms}
    sys.stdout.write(json.dumps(response) + "\n")
    sys.stdout.flush()
