"""Command line interface for uploading chapters to rulate.ru."""

import argparse
from typing import List

from rulate_uploader import upload_chapters


def main(argv: List[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Upload chapter files to rulate.ru")
    parser.add_argument("book_url", help="Base URL of the book on rulate.ru")
    parser.add_argument("files", nargs="+", help="Chapter files to upload")
    parser.add_argument("--username", help="Username for authentication")
    parser.add_argument("--password", help="Password for authentication")
    parser.add_argument("--deferred", action="store_true", help="Upload chapters as deferred")
    parser.add_argument("--subscription", action="store_true", help="Require subscription to read")
    parser.add_argument("--volume", type=int, help="Volume number for uploaded chapters")
    parser.add_argument("--publish-at", dest="publish_at", help="Schedule publication time")
    parser.add_argument(
        "--no-headless",
        dest="headless",
        action="store_false",
        help="Run browser with a GUI instead of headless mode",
    )
    parser.set_defaults(headless=True)

    args = parser.parse_args(argv)

    results = upload_chapters(
        args.book_url,
        args.files,
        username=args.username,
        password=args.password,
        deferred=args.deferred,
        subscription=args.subscription,
        volume=args.volume,
        publish_at=args.publish_at,
        headless=args.headless,
    )

    for file, status in results.items():
        outcome = "uploaded" if status else "failed"
        print(f"{file}: {outcome}")


if __name__ == "__main__":  # pragma: no cover - manual execution
    main()
