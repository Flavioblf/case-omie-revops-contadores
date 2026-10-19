"""Extrai o logo Omie do print do site e gera PNGs transparentes em alta resolução (ciano, branco, escuro),
mais o símbolo 'play' decorativo (motivo do hero do site) para os slides escuros.

O logo é ciano chapado sobre fundo escuro, então a transparência sai da distância de cada pixel até o fundo.
Uso: python deck/design_system/build_logo.py "<caminho do print>"
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

FUNDO = (0x00, 0x1E, 0x27)
CIANO = (0x00, 0xE2, 0xF4)
CAIXA = (14, 74, 150, 122)  # recorte do logo no print (x0, y0, x1, y1)
ESCALA = 8
CORES = {"ciano": CIANO, "branco": (0xFF, 0xFF, 0xFF), "escuro": FUNDO}


def main(print_path: str) -> None:
    out = Path(__file__).parent
    rec = Image.open(print_path).convert("RGB").crop(CAIXA)
    # alfa = quanto o pixel se aproxima do ciano, partindo do fundo (canal G tem o maior contraste)
    g = rec.getchannel("G")
    alfa = g.point(lambda v: max(0, min(255, int((v - FUNDO[1]) / (CIANO[1] - FUNDO[1]) * 255))))
    w, h = alfa.size
    alfa = alfa.resize((w * ESCALA, h * ESCALA), Image.LANCZOS).filter(ImageFilter.GaussianBlur(ESCALA * 0.35))
    # reforça a borda: curva em S deixa o contorno nítido sem serrilhar
    alfa = alfa.point(lambda v: 0 if v < 70 else 255 if v > 185 else int((v - 70) / 115 * 255))
    alfa = alfa.crop(alfa.getbbox())
    for nome, cor in CORES.items():
        img = Image.new("RGBA", alfa.size, cor + (0,))
        img.putalpha(alfa)
        path = out / f"logo_omie_{nome}.png"
        img.save(path)
        print(path, img.size)
    simbolo_play(out / "simbolo_play.png")


def simbolo_play(path: Path, lado: int = 1600) -> None:
    """Círculo petrol com o triângulo 'play' vazado, como o grafismo do hero do site."""
    ss = 2  # supersampling para bordas suaves
    n = lado * ss
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((0, 0, n - 1, n - 1), fill=(0x00, 0x53, 0x5E, 90))  # sutil: decoração, não compete com dados
    cx, cy, r = n * 0.47, n * 0.5, n * 0.27
    d.polygon([(cx - r * 0.55, cy - r), (cx - r * 0.55, cy + r), (cx + r * 1.05, cy)], fill=(0, 0, 0, 0))
    img = img.resize((lado, lado), Image.LANCZOS)
    img.save(path)
    print(path, img.size)


if __name__ == "__main__":
    main(sys.argv[1])
