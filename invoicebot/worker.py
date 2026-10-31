"""Reader processes. Each runs extractor.analyse() on one file at a time.

They are separate OS processes (not threads) so several documents are read at the same time on
several CPU cores, and so memory is returned to Windows when a worker is recycled."""
import os
import time

from . import extractor, memory, ocr, settings


def init_worker(snapshot, ocr_config, home):
    """Called once in every reader process."""
    if home:
        os.environ["BOT_HOME"] = home
    settings.use_snapshot(snapshot)
    ocr.configure(**ocr_config)


def process_job(job):
    """job = {'job_id', 'path', 'ctx'}  ->  the analysis dict, plus job_id, seconds and peak memory."""
    t = time.time()
    try:
        res = extractor.analyse(job["path"], job["ctx"])
    except Exception as e:                                   # a bad file must never stop the run
        res = extractor._blank()
        res["fatal"] = True
        res["problems"].append(f"unexpected error while reading ({type(e).__name__}): {str(e)[:150]}")
    res["job_id"] = job["job_id"]
    res["seconds"] = round(time.time() - t, 2)
    res["peak_mb"] = memory.peak_mb()
    return res
