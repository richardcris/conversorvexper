from __future__ import annotations

import json
import os
import shutil
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import app


ROOT = Path(__file__).resolve().parent
INSTALLER_PATH = ROOT / "installer" / app.UPDATE_INSTALLER_NAME
DEFAULT_FEED = ROOT / app.UPDATE_FEED_DIRNAME
GITHUB_API = "https://api.github.com"


def update_feed_dir() -> Path:
    configured = os.environ.get("VEXPER_UPDATE_FEED_DIR", "").strip()
    if configured:
        return Path(configured)
    return DEFAULT_FEED


def github_repository() -> str:
    return os.environ.get("VEXPER_GITHUB_REPO", os.environ.get("GITHUB_REPOSITORY", "")).strip().strip("/")


def github_token() -> str:
    return os.environ.get("VEXPER_GITHUB_TOKEN", os.environ.get("GITHUB_TOKEN", "")).strip()


def github_headers(extra: dict[str, str] | None = None) -> dict[str, str]:
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {github_token()}",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "vexper-updater",
    }
    if extra:
        headers.update(extra)
    return headers


def github_request(method: str, url: str, data: bytes | None = None, headers: dict[str, str] | None = None) -> str:
    request = Request(url, data=data, headers=github_headers(headers), method=method)
    with urlopen(request, timeout=60) as response:
        return response.read().decode("utf-8")


def ensure_github_release(repo: str, tag: str) -> dict[str, object]:
    release_url = f"{GITHUB_API}/repos/{repo}/releases/tags/{tag}"
    try:
        return json.loads(github_request("GET", release_url))
    except HTTPError as error:
        if error.code != 404:
            raise

    payload = {
        "tag_name": tag,
        "name": f"{app.APP_TITLE} {app.APP_VERSION}",
        "body": f"Atualizacao automatica {app.APP_VERSION}",
        "draft": False,
        "prerelease": False,
        "generate_release_notes": False,
    }
    created = github_request(
        "POST",
        f"{GITHUB_API}/repos/{repo}/releases",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    return json.loads(created)


def upload_release_asset(repo: str, release: dict[str, object], asset_path: Path, content_type: str) -> None:
    for asset in release.get("assets", []) or []:
        if str(asset.get("name", "")) == asset_path.name and asset.get("id") is not None:
            github_request("DELETE", f"{GITHUB_API}/repos/{repo}/releases/assets/{asset['id']}")

    upload_url = str(release["upload_url"]).split("{")[0]
    query = urlencode({"name": asset_path.name})
    github_request(
        "POST",
        f"{upload_url}?{query}",
        data=asset_path.read_bytes(),
        headers={"Content-Type": content_type},
    )


def publish_github_release(manifest_path: Path) -> None:
    repo = github_repository()
    token = github_token()
    if not repo or not token:
        print("Publicacao no GitHub ignorada: defina VEXPER_GITHUB_REPO e VEXPER_GITHUB_TOKEN.")
        return

    tag = f"v{app.APP_VERSION}"
    release = ensure_github_release(repo, tag)
    upload_release_asset(repo, release, INSTALLER_PATH, "application/octet-stream")
    upload_release_asset(repo, release, manifest_path, "application/json")
    print(f"Release do GitHub atualizada em {repo} com a tag {tag}")


def main() -> None:
    if not INSTALLER_PATH.exists():
        raise SystemExit(f"Instalador nao encontrado em {INSTALLER_PATH}")

    feed_dir = update_feed_dir()
    feed_dir.mkdir(parents=True, exist_ok=True)

    published_installer = feed_dir / app.UPDATE_INSTALLER_NAME
    shutil.copy2(INSTALLER_PATH, published_installer)

    manifest = {
        "app_name": app.APP_TITLE,
        "version": app.APP_VERSION,
        "installer_name": app.UPDATE_INSTALLER_NAME,
        "installer_location": app.UPDATE_INSTALLER_NAME,
        "notes": f"Atualizacao automatica publicada em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        "published_at": datetime.now().isoformat(timespec="seconds"),
    }
    manifest_path = feed_dir / app.UPDATE_MANIFEST_NAME
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=True), encoding="utf-8")
    print(f"Feed de atualizacao publicado em {feed_dir}")
    publish_github_release(manifest_path)


if __name__ == "__main__":
    main()