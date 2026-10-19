"""Renderiza cada slide de um .pptx em PNG usando o PowerPoint (COM), para revisão visual.

Uso: python deck/render.py output/Apresentacao_Omie_Canal_Contadores.pptx <pasta_saida>
"""
import sys
from pathlib import Path

import win32com.client as win32


def render(pptx: str, out_dir: str) -> None:
    out = Path(out_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("slide-*.png"):
        f.unlink()
    app = win32.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(Path(pptx).resolve()), ReadOnly=True, WithWindow=False)
        for i, slide in enumerate(pres.Slides, start=1):
            path = out / f"slide-{i}.png"
            slide.Export(str(path), "PNG", 1600, 900)
            print(path)
        pres.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    render(sys.argv[1], sys.argv[2])
