"""Run manifests separate current successful artifacts from archived files."""

import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path


class FetchError(RuntimeError):
    status = "fetch_failed"


class ParseError(RuntimeError):
    status = "parse_failed"


class NoReleasesError(RuntimeError):
    status = "no_releases"


class IncompleteRunError(RuntimeError):
    status = "partial"


class RunStatus:
    def __init__(self, path, target):
        self.path = Path(path)
        self.data = {
            "run_id": uuid.uuid4().hex,
            "requested_date": target.isoformat(),
            "effective_date": target.isoformat(),
            "started_at": datetime.now(timezone.utc).isoformat(),
            "status": "running", "files": {},
        }

    def update(self, **values):
        self.data.update(values)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_name(self.path.name + "." + self.data["run_id"] + ".tmp")
        try:
            temp.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(temp, self.path)
        finally:
            temp.unlink(missing_ok=True)

    def __enter__(self):
        # Invalidate any previous run before doing network or generation work.
        self.update()
        return self

    def __exit__(self, kind, error, traceback):
        if error:
            self.update(status=getattr(error, "status", "failed"), error=str(error))


def completed_files(manifest_path, target_date):
    """Return one complete run's exact files; never discover artifacts by glob."""
    data = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    if data.get("status") != "complete" or data.get("effective_date") != target_date:
        raise IncompleteRunError("No complete run is available for the requested date.")
    files = data.get("files", {})
    if not all(files.get(key) and Path(files[key]).is_file()
               and Path(files[key]).stat().st_size > 0 for key in ("docx", "xlsx", "email_html")):
        raise IncompleteRunError("The completed run is missing required artifacts.")
    return {key: Path(value) for key, value in files.items()}
