"""Signs GCS object URLs so private NVR snapshots can be viewed in the browser.

nvr_monitoring.image_url points at objects in the tarsyer_client_storage
bucket, which is not publicly readable. A short-lived v4 signed URL is
generated on each request using a read-only service account key.
"""

import os
from datetime import timedelta

_GCS_PREFIX = "https://storage.googleapis.com/"
_client = None
_init_attempted = False


def _get_client():
    global _client, _init_attempted
    if _init_attempted:
        return _client
    _init_attempted = True
    key_path = os.environ.get("GCS_KEY_PATH") or os.path.join(
        os.path.dirname(os.path.dirname(__file__)), "gcs-key.json"
    )
    if os.path.exists(key_path):
        try:
            from google.cloud import storage
            _client = storage.Client.from_service_account_json(key_path)
        except Exception as e:
            print(f"[GCS] ERROR initialising client: {e}", flush=True)
    else:
        print(f"[GCS] key file not found at {key_path}", flush=True)
    return _client


def signed_url(raw_url: str, expires_minutes: int = 20) -> str:
    """Turn a public-form GCS URL into a short-lived v4 signed URL."""
    if not raw_url or not raw_url.startswith(_GCS_PREFIX):
        return raw_url
    client = _get_client()
    if client is None:
        return raw_url
    path = raw_url[len(_GCS_PREFIX):]
    bucket_name, _, blob_name = path.partition("/")
    try:
        blob = client.bucket(bucket_name).blob(blob_name)
        return blob.generate_signed_url(
            version="v4", expiration=timedelta(minutes=expires_minutes), method="GET"
        )
    except Exception as e:
        print(f"[GCS] ERROR signing URL: {e}", flush=True)
        return raw_url


def sign_field(docs, field: str = "image_url"):
    """Replace `field` in each dict of `docs` with a signed URL, in place."""
    for d in docs:
        if d.get(field):
            d[field] = signed_url(d[field])
    return docs
