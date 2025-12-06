from PIL import Image, ImageDraw, ImageFont

# Create a new image with transparency
img = Image.new('RGBA', (200, 200), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Draw outer circle (dark blue border)
draw.ellipse([5, 5, 195, 195], fill='#1565c0', outline='#0d47a1', width=3)

# Draw inner circle (lighter blue)
draw.ellipse([25, 25, 175, 175], fill='#2196F3', outline='#1976D2', width=2)

# Add 'U' text
try:
    font = ImageFont.truetype('C:\\Windows\\Fonts\\arial.ttf', 110)
except:
    font = ImageFont.load_default()

# Draw the 'U' in white
draw.text((50, 35), 'U', fill='white', font=font)

# Save the image
img.save('static/images/logo.png')
print('✓ Logo created successfully at static/images/logo.png')
