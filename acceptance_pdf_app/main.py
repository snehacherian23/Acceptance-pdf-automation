"""
main.py
-------
Entry point. Run with:  python main.py

Keeps this file minimal on purpose — all real setup logic lives in the
app factory (app/__init__.py), config.py, and the extension/blueprint
modules. This file's only job is to build the app and start the
dev server.
"""

from app import create_app

app = create_app()

if __name__ == "__main__":
    # threaded=True is required: processing runs on a background thread
    # (see processing_routes.py) so the dashboard can poll
    # /processing/status/<job_id> for live progress while it works --
    # the dev server needs to handle those requests concurrently.
    app.run(host="0.0.0.0", port=5000, debug=app.config.get("DEBUG", False), threaded=True)
