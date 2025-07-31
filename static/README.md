# Static Files for Attendance Report System

## Favicon Files

This directory contains the favicon files for the browser tab icon. Currently, there are placeholder files that need to be replaced with actual icon files.

### Required Files:

1. **favicon.ico** - Traditional favicon format (16x16, 32x32 pixels)
2. **favicon.png** - PNG format favicon (32x32 pixels recommended)
3. **icon.svg** - SVG format icon (scalable vector graphics)

### How to Create Real Favicon Files:

#### Option 1: Online Tools
1. Go to [favicon.io](https://favicon.io/) or [realfavicongenerator.net](https://realfavicongenerator.net/)
2. Upload your image or create a new one
3. Download the generated favicon files
4. Replace the placeholder files in this directory

#### Option 2: Convert Existing Image
1. Use an image editing tool like GIMP, Photoshop, or online converters
2. Create a 32x32 pixel image
3. Save as .ico, .png, and .svg formats
4. Replace the placeholder files

#### Option 3: Use the SVG Icon
The `icon.svg` file contains a clock icon that represents the attendance system. You can:
1. Open the SVG file in a browser or SVG editor
2. Export it to different formats
3. Use online tools to convert SVG to ICO/PNG

### Recommended Icon Design:
- **Theme**: Clock or time-related icon (represents attendance/time tracking)
- **Colors**: Use the system's primary colors (#4195F4, #1054BE)
- **Style**: Simple, recognizable, professional
- **Size**: 32x32 pixels for best compatibility

### File Structure:
```
static/
├── favicon.ico    # Traditional favicon
├── favicon.png    # PNG format favicon
├── icon.svg       # SVG format icon
└── README.md      # This file
```

### Browser Support:
- **Modern browsers**: SVG favicon (best quality)
- **Older browsers**: ICO/PNG favicon (fallback)
- **Mobile devices**: PNG favicon for app icons

### Testing:
After replacing the files, refresh your browser and check:
1. Browser tab icon appears
2. Bookmark icon appears
3. Mobile home screen icon (if added to home screen) 