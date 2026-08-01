#!/usr/bin/env python3
"""Valida que el sitio generado contenga todas las diapositivas y recursos locales."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--site", default="dist")
    args = parser.parse_args()

    site = Path(args.site).resolve()
    errors: list[str] = []
    required = ["index.html", "404.html", "manifest.json", "search-index.json", ".nojekyll"]
    for name in required:
        if not (site / name).exists():
            errors.append(f"Falta {name}")

    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1

    manifest = json.loads((site / "manifest.json").read_text(encoding="utf-8"))
    slides = manifest.get("slides", [])
    if not slides:
        errors.append("manifest.json no contiene diapositivas")

    for slide in slides:
        for key in ("image", "thumbnail"):
            relative = slide.get(key)
            if not relative or not (site / relative).exists():
                errors.append(f"Diapositiva {slide.get('number')}: recurso ausente {key}={relative}")
        embed = slide.get("embed")
        if embed and not embed.startswith(("http://", "https://")) and not (site / embed).exists():
            errors.append(f"Diapositiva {slide.get('number')}: embed ausente {embed}")

    html = (site / "index.html").read_text(encoding="utf-8")
    card_count = len(re.findall(r'class="slide-card"', html))
    if card_count != len(slides):
        errors.append(f"index.html contiene {card_count} tarjetas, manifest contiene {len(slides)}")

    for match in re.finditer(r'(?:src|href)="([^"]+)"', html):
        value = match.group(1)
        if value.startswith(("http://", "https://", "#", "mailto:", "javascript:")):
            continue
        if not (site / value).exists():
            errors.append(f"Referencia local ausente en index.html: {value}")

    if errors:
        print("Validación fallida:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    print(f"Sitio válido: {site}")
    print(f"Diapositivas verificadas: {len(slides)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
