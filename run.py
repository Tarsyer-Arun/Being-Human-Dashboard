import os

from app import create_app

app = create_app()

if __name__ == "__main__":
    # debug defaults to off: the Werkzeug debugger allows arbitrary code
    # execution and must never be reachable on a bound interface.
    debug = os.environ.get("FLASK_DEBUG", "").lower() in ("1", "true", "yes")
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", 21582))
    app.run(debug=debug, host=host, port=port)
