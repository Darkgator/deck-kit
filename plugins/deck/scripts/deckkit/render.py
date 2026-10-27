"""Prévia: um PNG por slide, para o agente (e você) olharem antes de entregar.

Ordem de tentativa: PowerPoint no Windows (fiel ao que o usuário verá), depois LibreOffice + pdftoppm.
Sem nenhum dos dois, devolve lista vazia e o agente deve avisar que não conseguiu olhar os slides.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def _powerpoint(pptx: Path, saida: Path) -> bool:
    if os.name != "nt":
        return False
    ps = ("$p=New-Object -ComObject PowerPoint.Application; try { "
          f"$d=$p.Presentations.Open('{str(pptx).replace(chr(39), chr(39) * 2)}',$true,$true,$false); "
          f"$d.Export('{str(saida).replace(chr(39), chr(39) * 2)}','PNG',1600,900); $d.Close() }} finally {{ $p.Quit() }}")
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True, timeout=300)
        return True
    except (OSError, subprocess.SubprocessError):
        return False


def _libreoffice(pptx: Path, saida: Path) -> bool:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    ppm = shutil.which("pdftoppm")
    if not (soffice and ppm):
        return False
    with tempfile.TemporaryDirectory() as tmp:
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(pptx)], check=True,
                           capture_output=True, timeout=300)
            pdf = Path(tmp) / (pptx.stem + ".pdf")
            subprocess.run([ppm, "-png", "-r", "120", str(pdf), str(saida / "Slide")], check=True, capture_output=True,
                           timeout=300)
        except (OSError, subprocess.SubprocessError):
            return False
    return True


def _numero(p: Path) -> int:
    dig = "".join(c for c in p.stem if c.isdigit())
    return int(dig) if dig else 0


def prever(pptx: Path, saida: Path = None) -> list[Path]:
    pptx = Path(pptx).resolve()
    saida = Path(saida).resolve() if saida else pptx.with_name(pptx.stem + " (previa)")
    if saida.exists():
        shutil.rmtree(saida)
    saida.mkdir(parents=True)
    if not (_powerpoint(pptx, saida) or _libreoffice(pptx, saida)):
        return []
    # nomes uniformes: slide-01.png, slide-02.png ...
    imgs = sorted((p for p in saida.iterdir() if p.suffix.lower() == ".png"), key=_numero)
    novos = []
    for i, p in enumerate(imgs, 1):
        destino = saida / f"slide-{i:02d}.png"
        p.rename(destino)
        novos.append(destino)
    return novos


def folha(imagens: list[Path], destino: Path, colunas: int = 4, largura: int = 480) -> Path:
    """Folha de contato com todos os slides, para uma olhada geral."""
    from PIL import Image
    ims = [Image.open(p).convert("RGB") for p in imagens]
    h = round(largura * ims[0].height / ims[0].width)
    linhas = (len(ims) + colunas - 1) // colunas
    folha_img = Image.new("RGB", (colunas * (largura + 8) + 8, linhas * (h + 8) + 8), "white")
    for k, im in enumerate(ims):
        folha_img.paste(im.resize((largura, h)), (8 + (k % colunas) * (largura + 8), 8 + (k // colunas) * (h + 8)))
    folha_img.save(destino)
    return destino
