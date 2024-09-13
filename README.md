# DOCX Image Converter

This is a simple tool for converting and resizing images within a DOCX file. The tool allows users to replace images in a DOCX document with compressed JPEG versions, reducing file size while maintaining image quality.

## Features

- Convert images in `.docx` files to JPEG format.
- Resize images to a specified width and DPI.
- Maintain the document structure and replace original images seamlessly.
- Display file size reduction after conversion.

## How to Use

1. **Download or Clone the Repository:**
   ```bash
   git clone https://github.com/yourusername/docx-image-converter.git
   ```
2. **Run the Tool:**
   - Ensure you have Python installed, along with the required libraries:
     - `Pillow`
     - `Tkinter`
   - Run the `docx_image_converter.py` file:
     ```bash
     python docx_image_converter.py
     ```

3. **Select the DOCX File:**
   - In the GUI window, click "Browse" and select the DOCX file you want to convert.

4. **Set the Image Quality and DPI:**
   - Input the desired image quality (1-100) and DPI for resizing the images. Defaults are set to `85` for quality and `200` DPI.

5. **Convert the File:**
   - Click "Convert" to start the image conversion process.
   - After conversion, the tool will show the new file size and the space saved.

6. **View Output:**
   - The converted DOCX file will be saved in the `output` folder with a `converted_` prefix.

7. **Exit the Tool:**
   - Click the "Exit" button to close the application.

## Requirements

- Python 3.x
- Required Python packages:
  - `Pillow`
  - `Tkinter`
