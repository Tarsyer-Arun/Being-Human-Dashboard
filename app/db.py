"""MongoDB connection for the Being Human dashboard.

Every credential is read from the environment (see .env.example). Nothing
sensitive is hard-coded here.
"""

import os
from urllib.parse import quote_plus

from pymongo import MongoClient

_client = None
_db_name = "beinghumanServer"

_ds_client = None
_ds_db_name = "devices_summary"


def _mongo_uri():
    return os.environ.get("MONGO_URI", "")


def init_db(app):
    global _client, _db_name
    uri = _mongo_uri()
    if not uri:
        raise RuntimeError(
            "MONGO_URI is not set. Copy .env.example to .env and fill in the values."
        )
    _db_name = os.environ.get("DB_NAME", "beinghumanServer")
    _client = MongoClient(uri, serverSelectionTimeoutMS=5000)

    # Optional: device heartbeat DB (devices_summary), used for the NVR /
    # Device Status monitoring tab. Read-only — this app never writes here.
    global _ds_client, _ds_db_name
    ds_host = os.environ.get("DS_HOST")
    if ds_host:
        ds_port = os.environ.get("DS_PORT", "27018")
        ds_user = os.environ.get("DS_USER", "")
        ds_pass = os.environ.get("DS_PASS", "")
        _ds_db_name = os.environ.get("DS_NAME", "devices_summary")
        if ds_user and ds_pass:
            ds_uri = (
                f"mongodb://{quote_plus(ds_user)}:{quote_plus(ds_pass)}"
                f"@{ds_host}:{ds_port}/{_ds_db_name}?authSource=admin"
            )
        else:
            ds_uri = f"mongodb://{ds_host}:{ds_port}/{_ds_db_name}"
        _ds_client = MongoClient(ds_uri, serverSelectionTimeoutMS=5000)


def get_bh_db():
    if _client is None:
        raise RuntimeError("init_db() must be called before get_bh_db()")
    return _client[_db_name]


def get_devices_summary_db():
    """Read-only handle to the device heartbeat DB. None if not configured."""
    if _ds_client is None:
        return None
    return _ds_client[_ds_db_name]


def get_devices_heartbeat_db():
    """Read-only handle to the devices_heartbeat DB (recurring_data collection).

    Same MongoDB server/cluster as devices_summary, so it reuses that client —
    only the database name differs. None if devices_summary isn't configured.
    """
    if _ds_client is None:
        return None
    db_name = os.environ.get("DS_HEARTBEAT_DB_NAME", "devices_heartbeat")
    return _ds_client[db_name]
