#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import UTC, date, datetime
from pathlib import Path
from urllib import error, parse, request

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from scripts.lib.linkedin_posts import read_post_archive  # noqa: E402


DEFAULT_ARCHIVE_ROOT = REPO_ROOT / "social" / "linkedin-posts"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Publish a generated LinkedIn post archive.")
    parser.add_argument("--date", default=date.today().isoformat(), help="Publish date in YYYY-MM-DD format.")
    parser.add_argument(
        "--archive-root",
        default=str(DEFAULT_ARCHIVE_ROOT),
        help="Directory where dated LinkedIn post archives are stored.",
    )
    parser.add_argument("--force", action="store_true", help="Publish even if a publication marker already exists.")
    return parser.parse_args()


def set_output(name: str, value: str) -> None:
    output_path = os.environ.get("GITHUB_OUTPUT")
    if not output_path:
        return
    with open(output_path, "a", encoding="utf-8") as handle:
        handle.write(f"{name}={value}\n")


def linkedin_version_candidates(explicit_version: str | None = None) -> list[str]:
    explicit = (explicit_version or "").strip()
    if explicit:
        return [explicit]

    now = datetime.now(UTC)
    current = f"{now.year}{now.month:02d}"
    previous_month = 12 if now.month == 1 else now.month - 1
    previous_year = now.year - 1 if now.month == 1 else now.year
    previous = f"{previous_year}{previous_month:02d}"
    return [current] if current == previous else [current, previous]


def is_inactive_version_error(message: str) -> bool:
    return "NONEXISTENT_VERSION" in message


def linkedin_request(
    *,
    method: str,
    url: str,
    token: str,
    version: str | None = None,
    body: bytes | None = None,
    content_type: str | None = None,
) -> tuple[dict[str, str], bytes]:
    headers = {
        "Authorization": f"Bearer {token}",
        "X-Restli-Protocol-Version": "2.0.0",
    }
    if version:
        headers["Linkedin-Version"] = version
    if content_type:
        headers["Content-Type"] = content_type

    req = request.Request(url, data=body, method=method, headers=headers)
    try:
        with request.urlopen(req) as response:
            response_headers = {key.lower(): value for key, value in response.headers.items()}
            return response_headers, response.read()
    except error.HTTPError as exc:
        payload = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"LinkedIn API {method} {url} failed with {exc.code}: {payload}") from exc


def resolve_author_urn(token: str) -> str:
    explicit = os.environ.get("LINKEDIN_AUTHOR_URN", "").strip()
    if explicit:
        return explicit

    _, payload = linkedin_request(
        method="GET",
        url="https://api.linkedin.com/v2/userinfo",
        token=token,
    )
    profile = json.loads(payload.decode("utf-8"))
    member_id = str(profile.get("sub", "")).strip()
    if not member_id:
        raise RuntimeError("LinkedIn userinfo response did not include a member identifier.")
    return f"urn:li:person:{member_id}"


def enforce_author_policy(author_urn: str) -> None:
    required_type = os.environ.get("LINKEDIN_REQUIRED_AUTHOR_TYPE", "organization").strip().lower()
    if required_type == "organization" and not author_urn.startswith("urn:li:organization:"):
        raise RuntimeError(
            "Refusing to publish with a personal LinkedIn author. "
            "Set LINKEDIN_AUTHOR_URN to an organization URN such as urn:li:organization:... "
            "or override LINKEDIN_REQUIRED_AUTHOR_TYPE if personal posting is explicitly intended."
        )
    if required_type == "person" and not author_urn.startswith("urn:li:person:"):
        raise RuntimeError("Refusing to publish because LINKEDIN_AUTHOR_URN is not a person URN.")


def publish_post(token: str, author_urn: str, commentary: str) -> tuple[str, str]:
    last_error: RuntimeError | None = None
    for version in linkedin_version_candidates(os.environ.get("LINKEDIN_API_VERSION")):
        try:
            headers, _ = linkedin_request(
                method="POST",
                url="https://api.linkedin.com/rest/posts",
                token=token,
                version=version,
                content_type="application/json",
                body=json.dumps(
                    {
                        "author": author_urn,
                        "commentary": commentary,
                        "visibility": "PUBLIC",
                        "distribution": {
                            "feedDistribution": "MAIN_FEED",
                            "targetEntities": [],
                            "thirdPartyDistributionChannels": [],
                        },
                        "lifecycleState": "PUBLISHED",
                        "isReshareDisabledByAuthor": False,
                    }
                ).encode("utf-8"),
            )
            return version, headers.get("x-restli-id", "")
        except RuntimeError as exc:
            last_error = exc
            if not is_inactive_version_error(str(exc)):
                raise

    if last_error is not None:
        raise last_error
    raise RuntimeError("LinkedIn publication failed without returning an error.")


def main() -> int:
    args = parse_args()
    publish_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    archive_root = Path(args.archive_root).resolve()

    resolved = read_post_archive(archive_root, publish_date)
    if resolved is None:
        print(f"No generated LinkedIn post archive exists for {publish_date.isoformat()}.")
        set_output("publish_status", "missing")
        return 0

    archive_dir, post = resolved
    publication_path = archive_dir / "linkedin-publication.json"

    if publication_path.exists() and not args.force:
        print(f"LinkedIn archive {archive_dir.name} is already marked as published.")
        set_output("publish_status", "exists")
        return 0

    token = os.environ.get("LINKEDIN_ACCESS_TOKEN", "").strip()
    if not token:
        raise SystemExit("Missing required environment variable: LINKEDIN_ACCESS_TOKEN")

    author_urn = resolve_author_urn(token)
    enforce_author_policy(author_urn)
    version, post_id = publish_post(token, author_urn, post["social"]["linkedin"])

    publication = {
        "published_at": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "linkedin_version": version,
        "author": author_urn,
        "post_id": post_id,
    }
    publication_path.write_text(json.dumps(publication, indent=2) + "\n", encoding="utf-8")

    print(f"Published LinkedIn archive {archive_dir.name} as {post_id or 'unknown post id'}.")
    set_output("publish_status", "created")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
