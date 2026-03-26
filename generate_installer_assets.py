from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent
LOGO = ROOT / "logo.png"
WIZARD = ROOT / "installer_wizard.bmp"
SMALL = ROOT / "installer_small.bmp"


def build_background(size: tuple[int, int], accent: str) -> Image.Image:
    width, height = size
    image = Image.new("RGB", size, "#07141F")
    draw = ImageDraw.Draw(image)

    for y in range(height):
        ratio = y / max(height - 1, 1)
        r = int(7 + (20 * ratio))
        g = int(20 + (35 * ratio))
        b = int(31 + (55 * ratio))
        draw.line((0, y, width, y), fill=(r, g, b))

    draw.ellipse((-40, height - 120, width // 2, height + 80), fill="#0F2E44")
    draw.ellipse((width // 3, -30, width + 90, height // 2), fill=accent)
    return image.filter(ImageFilter.GaussianBlur(0.6))


def paste_logo(canvas: Image.Image, max_size: tuple[int, int]) -> None:
    if not LOGO.exists():
        return
    logo = Image.open(LOGO).convert("RGBA")
    logo.thumbnail(max_size, Image.LANCZOS)
    x = (canvas.width - logo.width) // 2
    y = 28
    glow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    glow.paste(logo.getchannel("A"), (x, y))
    glow = glow.filter(ImageFilter.GaussianBlur(10))
    tint = Image.new("RGBA", canvas.size, (118, 228, 247, 0))
    tint.putalpha(glow.getchannel("A"))
    canvas.alpha_composite(tint)
    canvas.alpha_composite(logo, (x, y))


def add_text(canvas: Image.Image, title: str, subtitle: str) -> None:
    draw = ImageDraw.Draw(canvas)
    draw.text((18, canvas.height - 92), title, fill="#F8FBFF")
    draw.text((18, canvas.height - 66), subtitle, fill="#C6D3E1")


def main() -> None:
    wizard = build_background((164, 314), "#123D5B")
    wizard_rgba = wizard.convert("RGBA")
    paste_logo(wizard_rgba, (120, 120))
    add_text(wizard_rgba, "CONVERSOR - VEXPER", "Instalacao guiada")
    wizard_rgba.convert("RGB").save(WIZARD, format="BMP")

    small = build_background((55, 55), "#1A6D78")
    small_rgba = small.convert("RGBA")
    paste_logo(small_rgba, (36, 36))
    small_rgba.convert("RGB").save(SMALL, format="BMP")


if __name__ == "__main__":
    main()