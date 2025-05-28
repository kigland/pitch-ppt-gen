# Image to PowerPoint Converter

A Python tool that automatically converts JPG images in a directory to a PowerPoint presentation (PPTX) with proper ordering and formatting.

## Features

- **Smart Sorting**: Automatically sorts images by numerical order (e.g., 1.jpg, 2.jpg, 10.jpg, 11.jpg)
- **16:9 Aspect Ratio**: Creates presentations in standard widescreen format
- **Full-Screen Images**: Each image fills the entire slide for maximum visual impact
- **Batch Processing**: Processes all JPG/JPEG files in a directory at once
- **Simple Interface**: Easy-to-use command-line interface

## Requirements

- Python 3.6+
- PIL (Pillow)
- python-pptx

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip3 install python-pptx pillow
```

## Usage

1. Run the script:
```bash
python3 main.py
```

2. When prompted, enter the path to the directory containing your JPG images:
```
请输入包含JPG图片的目录路径: /path/to/your/images
```

3. The script will:
   - Scan the directory for JPG/JPEG files
   - Sort them in natural numerical order
   - Create a PowerPoint presentation with one image per slide
   - Save the presentation as "图片演示文稿.pptx" in the same directory

## How It Works

1. **File Discovery**: Scans the specified directory for files with `.jpg` or `.jpeg` extensions
2. **Natural Sorting**: Uses a custom sorting algorithm that properly handles numerical sequences in filenames
3. **Presentation Creation**: Creates a new PowerPoint presentation with 16:9 aspect ratio
4. **Image Processing**: Adds each image to a separate slide, scaling to fill the entire slide
5. **Output**: Saves the final presentation in the source directory

## Example

If your directory contains:
```
images/
├── 1.jpg
├── 2.jpg
├── 10.jpg
├── 11.jpg
└── photo.jpg
```

The script will create slides in this order: 1.jpg → 2.jpg → 10.jpg → 11.jpg → photo.jpg

## Notes

- Only JPG and JPEG formats are supported
- Images are automatically scaled to fit the slide dimensions
- The output file is named "图片演示文稿.pptx" (Chinese: "Image Presentation.pptx")
- If no JPG files are found, the script will display an appropriate message
