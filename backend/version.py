from __future__ import annotations

import os


def get_app_version() -> str:
    manual_version = os.getenv("APP_VERSION", "").strip()
    if manual_version:
        return manual_version

    for variable_name in (
        "RENDER_GIT_COMMIT",
        "SOURCE_VERSION",
        "GIT_COMMIT",
        "HEROKU_SLUG_COMMIT",
        "COMMIT_REF",
    ):
        value = os.getenv(variable_name, "").strip()
        if value:
            return value[:12]

    return "local"