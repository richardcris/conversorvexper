from pathlib import Path

import app


ROOT = Path(__file__).resolve().parent


def main() -> None:
    version = app.APP_VERSION.strip()
    if not version:
        raise SystemExit("APP_VERSION vazio em app.py")

    installer_include = ROOT / "installer_version.iss"
    installer_include.write_text(f'#define MyAppVersion "{version}"\n', encoding="utf-8")

    version_file = ROOT / "build_version.txt"
    version_file.write_text(version + "\n", encoding="utf-8")

    print(f"Metadados de build gerados para a versao {version}")


if __name__ == "__main__":
    main()