"""MongoDB connection for the Being Human dashboard.

Every credential is read from the environment (see .env.example). Nothing
sensitive is hard-coded here.
"""

import os

from pymongo import MongoClient

_client = None
_db_name = "beinghumanServer"


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


def get_bh_db():
    if _client is None:
        raise RuntimeError("init_db() must be called before get_bh_db()")
    return _client[_db_name]
