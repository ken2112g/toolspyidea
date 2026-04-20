"""Build helper - tạo icon và wizard images cho installer"""
import sys

def create_icon():
    from PIL import Image, ImageDraw
    img = Image.new('RGBA', (256, 256), (0,0,0,0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([8,8,248,248], radius=48, fill=(132,85,239))
    d.ellipse([78,78,178,178], outline='white', width=6)
    d.line([128,40,128,100], fill=(83,221,252), width=5)
    d.line([128,156,128,216], fill=(83,221,252), width=5)
    d.line([40,128,100,128], fill=(83,221,252), width=5)
    d.line([156,128,216,128], fill=(83,221,252), width=5)
    d.ellipse([118,118,138,138], fill=(83,221,252))
    img.save('app_icon.ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])
    print("    Icon created OK")

def create_wizard_images():
    from PIL import Image, ImageDraw, ImageFont
    img = Image.new('RGB', (164, 314), (6,14,32))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([12,12,152,152], radius=24, fill=(132,85,239))
    d.ellipse([57,57,107,107], outline='white', width=4)
    d.line([82,35,82,70], fill=(83,221,252), width=3)
    d.line([82,94,82,130], fill=(83,221,252), width=3)
    d.line([35,82,70,82], fill=(83,221,252), width=3)
    d.line([94,82,130,82], fill=(83,221,252), width=3)
    d.ellipse([76,76,88,88], fill=(83,221,252))
    try:
        font = ImageFont.truetype('arial.ttf', 18)
        sfont = ImageFont.truetype('arial.ttf', 11)
    except:
        font = sfont = ImageFont.load_default()
    d.text((82, 180), 'Tool Spy', fill='white', anchor='mt', font=font)
    d.text((82, 205), 'Idea', fill=(186,158,255), anchor='mt', font=font)
    d.text((82, 240), 'v1.0.0', fill=(64,72,93), anchor='mt', font=sfont)
    d.text((82, 260), 'by ChThanh', fill=(64,72,93), anchor='mt', font=sfont)
    img.save('wizard_image.bmp')
    
    img2 = Image.new('RGB', (55, 55), (132,85,239))
    d2 = ImageDraw.Draw(img2)
    d2.ellipse([14,14,41,41], outline='white', width=2)
    d2.ellipse([25,25,30,30], fill=(83,221,252))
    img2.save('wizard_small.bmp')
    print("    Wizard images created OK")

def get_playwright_path():
    import playwright, os
    print(os.path.dirname(playwright.__file__))

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "all"
    if cmd == "icon":
        create_icon()
    elif cmd == "wizard":
        create_wizard_images()
    elif cmd == "pw_path":
        get_playwright_path()
    else:
        create_icon()
        create_wizard_images()
