from PIL import Image, ImageDraw

img = Image.new('RGB', (400, 300), color = (73, 109, 137))
d = ImageDraw.Draw(img)
d.text((10,10), "Hello World", fill=(255,255,0))
img.save('test_image.png')
