"""Persistent parsed-data cache for expensive Excel reference loaders."""

from __future__ import annotations

import hashlib
import pickle
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable


CACHE_VERSION = "reference-loader-v3"
CACHE_DIR = Path(__file__).parent.parent / ".cache" / "dpmptsp-report-generator"


@dataclass(frozen=True)
class CacheLoadResult:
    """Result wrapper for a cached or freshly parsed reference file."""

    data: Any
    status: str
    elapsed_seconds: float
    path: Path


def get_cache_key(file_type: str, file_content: bytes, year: int, version: str = CACHE_VERSION) -> str:
    """Build a stable cache key from file content, file type, year, and loader version."""
    file_hash = hashlib.sha256(file_content).hexdigest()
    metadata = f"{version}|{file_type}|{year}|{file_hash}".encode("utf-8")
    return hashlib.sha256(metadata).hexdigest()


def get_cache_path(file_type: str, file_content: bytes, year: int, version: str = CACHE_VERSION) -> Path:
    """Return the on-disk cache path for a parsed reference file."""
    safe_type = "".join(ch.lower() if ch.isalnum() else "_" for ch in file_type).strip("_")
    return CACHE_DIR / f"{safe_type}-{get_cache_key(file_type, file_content, year, version)}.pickle"


def load_or_build(
    file_type: str,
    file_content: bytes,
    filename: str,
    year: int,
    builder: Callable[[bytes, str, int], Any],
    version: str = CACHE_VERSION,
) -> CacheLoadResult:
    """Load parsed data from disk cache or build and cache it."""
    cache_path = get_cache_path(file_type, file_content, year, version)
    start = time.perf_counter()

    if cache_path.exists():
        try:
            with cache_path.open("rb") as handle:
                data = pickle.load(handle)
            return CacheLoadResult(data=data, status="cache", elapsed_seconds=time.perf_counter() - start, path=cache_path)
        except Exception:
            cache_path.unlink(missing_ok=True)

    data = builder(file_content, filename, year)
    elapsed = time.perf_counter() - start

    if data is not None:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp_path = cache_path.with_suffix(".tmp")
        with tmp_path.open("wb") as handle:
            pickle.dump(data, handle, protocol=pickle.HIGHEST_PROTOCOL)
        tmp_path.replace(cache_path)

    return CacheLoadResult(data=data, status="parsed", elapsed_seconds=elapsed, path=cache_path)
