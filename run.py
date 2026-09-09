#!/usr/bin/env python3
"""Start the AutoGeneration Platform.

    python run.py                 -> http://127.0.0.1:8000
    python run.py --port 9000
    python run.py --host 0.0.0.0  -> reachable from other machines on the network

For more than a handful of people, run it behind a real WSGI server instead:

    waitress-serve --port=8000 --call platform_app.app:create_app
"""

from __future__ import annotations

import argparse
import webbrowser
from threading import Timer

from platform_app import config
from platform_app.app import create_app
from platform_app.jobs import store


def main() -> int:
    parser = argparse.ArgumentParser(description=f"Run the {config.APP_NAME}.")
    parser.add_argument("--host", default=config.HOST)
    parser.add_argument("--port", type=int, default=config.PORT)
    parser.add_argument("--debug", action="store_true", help="Reload on code changes.")
    parser.add_argument("--no-browser", action="store_true", help="Do not open a browser.")
    args = parser.parse_args()

    config.ensure_dirs()
    swept = store.sweep()
    if swept:
        print(f"Cleared {swept} expired run(s).")

    app = create_app()
    url = f"http://{'127.0.0.1' if args.host in ('0.0.0.0', '') else args.host}:{args.port}"

    print(f"{config.APP_NAME} ready at {url}")
    print("Press Ctrl+C to stop.")

    if not args.no_browser and not args.debug:
        Timer(1.0, lambda: webbrowser.open(url)).start()

    # threaded=True keeps the page responsive while a comparison is running.
    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True,
            use_reloader=args.debug)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
