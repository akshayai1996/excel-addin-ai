import os
from PIL import Image, ImageDraw

os.makedirs("excel-ai-addin/assets", exist_ok=True)

def create_icon(size):
    # Create image with transparent background
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Draw rounded background badge (Excel green / indigo gradient feel)
    # Excel green: #107C41 -> Dark Teal: #0B5A2F
    radius = max(2, int(size * 0.22))
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=(16, 124, 65, 255))
    
    # Draw AI sparkles / star in center
    cx, cy = size / 2, size / 2
    r_outer = size * 0.32
    r_inner = size * 0.12
    
    points = []
    # 4-pointed star
    for i in range(8):
        angle = i * 3.14159265 / 4
        r = r_outer if (i % 2 == 0) else r_inner
        import math
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        points.append((x, y))
        
    draw.polygon(points, fill=(255, 255, 255, 255))
    
    # Add a small top-right secondary sparkle
    if size >= 32:
        s_cx = cx + size * 0.25
        s_cy = cy - size * 0.25
        sr_out = size * 0.10
        sr_in = size * 0.04
        s_points = []
        for i in range(8):
            angle = i * 3.14159265 / 4
            r = sr_out if (i % 2 == 0) else sr_in
            x = s_cx + r * math.cos(angle)
            y = s_cy + r * math.sin(angle)
            s_points.append((x, y))
        draw.polygon(s_points, fill=(255, 235, 120, 255))
        
    return img

for sz in [16, 32, 64, 80]:
    icon = create_icon(sz)
    icon.save(f"excel-ai-addin/assets/icon-{sz}.png")
    print(f"Generated icon-{sz}.png", flush=True)
