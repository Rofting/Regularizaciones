"""Compara el OCR de escaneos reales con y sin preparación de imagen.

Uso:
    python scripts/comparar_ocr.py CARPETA [--paginas 1]

Para cada PDF de la carpeta (y subcarpetas) renderiza la portada como lo hace
la aplicación, lee el original y la versión preparada (enderezado, ruido,
contraste) y muestra cuántos datos útiles (fechas, importes, CIF) salen en
cada caso y cuál elegiría la aplicación. No modifica ni copia ningún archivo.
Los PDF con texto digital se indican y no necesitan OCR.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("carpeta", type=Path)
    parser.add_argument("--paginas", type=int, default=1, help="páginas por documento (por defecto, la portada)")
    args = parser.parse_args(argv)

    import lector_pdf
    from pdf2image import convert_from_path
    from scan_preprocessing import assess, ocr_score, prepare

    files = sorted(path for path in args.carpeta.rglob("*.pdf") if path.is_file())
    if not files:
        print("No hay PDF en la carpeta.")
        return 1
    better = same = worse = digital = 0
    print(f"{'archivo':40} {'problemas':32} {'original':>8} {'preparada':>9}  elegida")
    for path in files:
        if not lector_pdf._pdf_probablemente_escaneado(path):
            digital += 1
            print(f"{path.name[:40]:40} {'PDF digital: sin OCR':32}")
            continue
        try:
            pages = convert_from_path(str(path), dpi=200, poppler_path=lector_pdf._poppler_path(),
                                      first_page=1, last_page=args.paginas)
        except Exception as error:
            print(f"{path.name[:40]:40} no se pudo renderizar: {error}")
            continue
        for number, page in enumerate(pages, start=1):
            quality = assess(page)
            original = ocr_score(lector_pdf._rapidocr_text(str(path), page))
            prepared = prepare(page)
            treated = ocr_score(lector_pdf._rapidocr_text(str(path), prepared.image)) if prepared.changed else None
            if treated is None or treated == original:
                same += 1
                chosen = "original (igual)" if treated is not None else "original (no lo necesita)"
            elif treated > original:
                better += 1
                chosen = "preparada: " + ", ".join(prepared.steps)
            else:
                worse += 1
                chosen = "original (la preparada leía peor)"
            label = f"{path.name[:36]} p{number}" if args.paginas > 1 else path.name[:40]
            print(f"{label:40} {', '.join(quality.problems) or '—':32} {original:8.0f} "
                  f"{'—' if treated is None else f'{treated:.0f}':>9}  {chosen}")
    print(f"\nMejoradas: {better} · iguales: {same} · la preparada leía peor (se usa el original): {worse}"
          f" · PDF digitales: {digital}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
