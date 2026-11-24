"""สร้างรูปริชเมนูจาก menu.html → jpg (ต้องมี Google Chrome + playwright: pip install playwright)
   python3 build.py   → yanghaus-richmenu-staff.jpg (2500×843) + yanghaus-richmenu-admin.jpg (2500×1686)"""
import pathlib
from playwright.sync_api import sync_playwright
here = pathlib.Path(__file__).resolve().parent
with sync_playwright() as p:
    b = p.chromium.launch(channel='chrome', args=['--no-sandbox', '--allow-file-access-from-files'])
    for name, h in (('staff', 843), ('admin', 1686)):
        pg = b.new_page(viewport={'width': 2500, 'height': h})
        pg.goto((here / 'menu.html').as_uri() + '?set=' + name)
        pg.evaluate('document.fonts.ready')
        pg.wait_for_timeout(300)
        out = here / f'yanghaus-richmenu-{name}.jpg'
        pg.screenshot(path=str(out), type='jpeg', quality=88)
        print(out.name, out.stat().st_size)
    b.close()
