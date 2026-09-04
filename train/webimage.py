from PIL import Image, ImageDraw, ImageFont
import os
import sys
import io

# === KONFIGURASI — EDIT 2 BARIS INI SESUAI STRUKTUR PROJECT ANDA ===
# Jika svg/ ada di /Users/sinitygs/Projects2/svg, set:
SVG_DIR = '/Users/sinitygs/Projects2/mainprogram/svg'
# Output ke folder paper_assets di root project:
OUTPUT_DIR = '/Users/sinitygs/Projects2/train/paper_assets'
# ====================================================================

SVG_FILES = {
    'a': '01_landing.svg',
    'b': '02_form.svg',
    'c': '03_result.svg',
    'd': '04_history.svg',
}

LABELS = {
    'a': 'Landing Page',
    'b': 'Input Form',
    'c': 'Prediction Result',
    'd': 'History & Retrain',
}

OUTPUT = os.path.join(OUTPUT_DIR, 'figure7_stuntor_platform.png')

# Per-panel dimensions
PANEL_W = 800
PANEL_H = 500
GAP = 30
LABEL_HEIGHT = 40
DPI = 300

# === DEPENDENCY CHECK ===
try:
    import cairosvg
except ImportError:
    print("[!] cairosvg belum terinstall. Jalankan dulu:")
    print("    pip install cairosvg")
    print("    (Mac: brew install cairo pango libffi gdk-pixbuf jika error native lib)")
    sys.exit(1)


def svg_to_pil(svg_path, target_w):
    """Convert SVG to PIL Image at given width (preserving aspect ratio)."""
    if not os.path.exists(svg_path):
        raise FileNotFoundError(f"SVG tidak ditemukan: {svg_path}")
    png_bytes = cairosvg.svg2png(url=svg_path, output_width=target_w * 2)  # render 2× for downsampling
    return Image.open(io.BytesIO(png_bytes)).convert('RGB')


def fit_to_panel(img, w, h, bg='white'):
    """Resize preserving aspect ratio + center on white canvas."""
    img.thumbnail((w, h), Image.Resampling.LANCZOS)
    canvas = Image.new('RGB', (w, h), bg)
    x = (w - img.width) // 2
    y = (h - img.height) // 2
    canvas.paste(img, (x, y))
    return canvas


def get_font(size):
    """Try platform fonts in order, fall back to default."""
    candidates = [
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/HelveticaNeue.ttc",
        "/Library/Fonts/Arial.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def main():
    # Validate input
    print(f"[INFO] SVG_DIR: {SVG_DIR}")
    if not os.path.isdir(SVG_DIR):
        print(f"[!] Folder SVG tidak ada: {SVG_DIR}")
        print(f"    Edit variabel SVG_DIR di awal script.")
        sys.exit(1)

    # Load + render all panels
    panels = {}
    for key, fname in SVG_FILES.items():
        full_path = os.path.join(SVG_DIR, fname)
        print(f"[ {key} ] Rendering {fname} ...")
        img = svg_to_pil(full_path, PANEL_W)
        panels[key] = fit_to_panel(img, PANEL_W, PANEL_H)

    # Build composite (2×2 grid)
    total_w = PANEL_W * 2 + GAP
    total_h = (PANEL_H + LABEL_HEIGHT) * 2 + GAP
    canvas = Image.new('RGB', (total_w, total_h), 'white')
    draw = ImageDraw.Draw(canvas)
    font = get_font(22)

    positions = {
        'a': (0, 0),
        'b': (PANEL_W + GAP, 0),
        'c': (0, PANEL_H + LABEL_HEIGHT + GAP),
        'd': (PANEL_W + GAP, PANEL_H + LABEL_HEIGHT + GAP),
    }

    for key, (x, y) in positions.items():
        # Border halus untuk memisahkan panel dari background
        draw.rectangle(
            [x, y, x + PANEL_W - 1, y + PANEL_H - 1],
            outline='#cccccc', width=1
        )
        canvas.paste(panels[key], (x, y))

        # Label di bawah image
        label_text = f'({key}) {LABELS[key]}'
        bbox = draw.textbbox((0, 0), label_text, font=font)
        text_w = bbox[2] - bbox[0]
        text_x = x + (PANEL_W - text_w) // 2
        text_y = y + PANEL_H + 8
        draw.text((text_x, text_y), label_text, fill='black', font=font)

    # Save
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    canvas.save(OUTPUT, dpi=(DPI, DPI), optimize=True)

    print()
    print(f"[OK] Saved: {OUTPUT}")
    print(f"     Size: {total_w}×{total_h} px @ {DPI} dpi")
    print(f"     File size: {os.path.getsize(OUTPUT) / 1024:.1f} KB")


if __name__ == '__main__':
    main()