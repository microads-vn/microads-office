"""Builds the blog images in server/office/assets/blog/ (run once, results are committed).

  - Cover cards 1200x630 per topic pillar and language (Open Graph size): <pillar>-<lang>.jpg
  - Product screenshots resized to 1200 px wide: shot-<name>.jpg
usage: python server/dev/make_blog_images.py
Fonts are read from C:\\Windows\\Fonts (read only).
"""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'server', 'office', 'assets', 'blog')
SHOTS = os.path.join(ROOT, 'release', 'store', 'screenshots')
ASSETS = os.path.join(ROOT, 'server', 'office', 'assets')
os.makedirs(OUT, exist_ok=True)
FB = r'C:\Windows\Fonts\segoeuib.ttf'
FR = r'C:\Windows\Fonts\segoeui.ttf'

# pillar → (icon, accent colour, {lang: (kicker, title)})
PILLARS = {
    'excel-formulas': ('excel.png', (43, 168, 74), {'en': ('EXCEL FORMULAS', 'Formulas that just work'), 'vi': ('CÔNG THỨC EXCEL', 'Hàm Excel dễ hiểu, dùng ngay')}),
    'free-office': ('office.png', (232, 49, 47), {'en': ('FREE OFFICE SUITE', 'Word, Excel & PDF — free forever'), 'vi': ('PHẦN MỀM VĂN PHÒNG MIỄN PHÍ', 'Word, Excel, PDF — miễn phí vĩnh viễn')}),
    'pdf-signing': ('pdf.png', (232, 49, 47), {'en': ('PDF & DIGITAL SIGNATURES', 'Sign and verify PDFs'), 'vi': ('PDF & CHỮ KÝ SỐ', 'Ký số và xác thực PDF')}),
    'documents': ('word.png', (37, 99, 235), {'en': ('DOCUMENTS', 'Write clean, professional documents'), 'vi': ('SOẠN THẢO VĂN BẢN', 'Văn bản chuẩn, đẹp, nhanh')}),
    'ai-office': ('office.png', (124, 58, 237), {'en': ('AI IN THE OFFICE', 'Your AI assistant, inside your documents'), 'vi': ('TRỢ LÝ AI', 'AI ngay trong văn bản và bảng tính')}),
    'office-tips': ('office.png', (43, 168, 74), {'en': ('OFFICE TIPS', 'Get everyday work done faster'), 'vi': ('MẸO VĂN PHÒNG', 'Làm việc văn phòng nhanh hơn')}),
}

def wrap(draw, text, font, width):
    words, lines, cur = text.split(), [], ''
    for w in words:
        t = (cur + ' ' + w).strip()
        if draw.textlength(t, font=font) <= width: cur = t
        else: lines.append(cur); cur = w
    lines.append(cur)
    return lines

def card(pillar, lang):
    icon, accent, texts = PILLARS[pillar]
    kicker, title = texts[lang]
    im = Image.new('RGB', (1200, 630), (16, 33, 23))
    d = ImageDraw.Draw(im)
    for y in range(630):  # subtle vertical gradient
        c = int(16 - 4 * y / 630), int(33 - 7 * y / 630), int(23 - 5 * y / 630)
        d.line([(0, y), (1200, y)], fill=c)
    d.rectangle([0, 0, 14, 630], fill=accent)
    ic = Image.open(os.path.join(ASSETS, icon)).convert('RGBA').resize((132, 132), Image.LANCZOS)
    im.paste(ic, (90, 96), ic)
    d.text((90, 268), kicker, font=ImageFont.truetype(FB, 30), fill=accent if accent != (16, 33, 23) else (255, 255, 255))
    f = ImageFont.truetype(FB, 64)
    y = 318
    for line in wrap(d, title, f, 1010):
        d.text((88, y), line, font=f, fill=(255, 255, 255)); y += 78
    d.text((90, 548), 'Microads Office  ·  microads.vn/office', font=ImageFont.truetype(FR, 28), fill=(255, 255, 255, 170))
    free = 'Miễn phí · Không cần tài khoản' if lang == 'vi' else 'Free · No account needed'
    ff = ImageFont.truetype(FB, 26)
    w = d.textlength(free, font=ff)
    d.rounded_rectangle([1110 - w - 30, 540, 1110, 590], radius=25, fill=(43, 168, 74))
    d.text((1110 - w - 15, 548), free, font=ff, fill=(255, 255, 255))
    im.save(os.path.join(OUT, f'{pillar}-{lang}.jpg'), quality=86, optimize=True, progressive=True)

for p in PILLARS:
    for l in ('en', 'vi'):
        card(p, l)
for name in ('01-word-ai', '02-excel-ai', '03-ai-settings', '04-excel-explain'):
    im = Image.open(os.path.join(SHOTS, name + '.png')).convert('RGB')
    im = im.resize((1200, round(im.height * 1200 / im.width)), Image.LANCZOS)
    im.save(os.path.join(OUT, 'shot-' + name[3:] + '.jpg'), quality=84, optimize=True, progressive=True)
print('\n'.join(sorted(f'{f} {os.path.getsize(os.path.join(OUT, f)) // 1024} KB' for f in os.listdir(OUT))))
