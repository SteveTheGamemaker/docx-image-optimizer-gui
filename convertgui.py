import zipfile
import os
import shutil
from PIL import Image
import logging
import tkinter as tk
from tkinter import filedialog, messagebox

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Set the upload folder and output folder paths
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
ALLOWED_EXTENSIONS = {'docx'}
ALLOWED_IMAGE_EXTENSIONS = {'.png', '.bmp', '.tiff', '.webp'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def has_transparency(image):
    if image.mode != 'RGBA':
        image = image.convert('RGBA')
    alpha_channel = image.getchannel('A')
    return alpha_channel.getextrema()[0] < 255

def resize_image(image, target_width, dpi):
    target_height = int((target_width / image.width) * image.height)
    resized_image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)
    return resized_image

def update_rels_for_image_conversion(rels_filepath, original_image, converted_image):
    with open(rels_filepath, 'r', encoding='utf-8') as f:
        rels_content = f.read()
    updated_rels_content = rels_content.replace(original_image, converted_image)
    with open(rels_filepath, 'w', encoding='utf-8') as f:
        f.write(updated_rels_content)

def update_all_rels_files(rels_dir, original_image, converted_image):
    for filename in os.listdir(rels_dir):
        if filename.endswith('.rels'):
            rels_filepath = os.path.join(rels_dir, filename)
            update_rels_for_image_conversion(rels_filepath, original_image, converted_image)

def convert_docx_images(docx_filepath, output_filepath, quality, dpi):
    temp_dir = 'temp_docx'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    try:
        with zipfile.ZipFile(docx_filepath, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
    except zipfile.BadZipFile:
        logging.error(f"Error: The file '{docx_filepath}' is not a valid DOCX or is corrupted.")
        return False

    rels_dir = os.path.join(temp_dir, 'word/_rels')
    if not os.path.exists(rels_dir):
        logging.error(f"Error: The '_rels' directory does not exist in the DOCX structure.")
        shutil.rmtree(temp_dir)
        return False

    media_dir = os.path.join(temp_dir, 'word/media')
    if not os.path.exists(media_dir):
        logging.error(f"Error: The 'media' directory does not exist in the DOCX structure.")
        shutil.rmtree(temp_dir)
        return False

    image_count = 0
    for filename in os.listdir(media_dir):
        file_extension = os.path.splitext(filename)[1].lower()
        if file_extension in ALLOWED_IMAGE_EXTENSIONS:
            filepath = os.path.join(media_dir, filename)
            try:
                with Image.open(filepath) as image:
                    if not has_transparency(image):
                        target_width = 6 * dpi
                        resized_image = resize_image(image.convert('RGB'), target_width, dpi=dpi)
                        new_filename = filename.rsplit('.', 1)[0] + '.jpg'
                        new_filepath = os.path.join(media_dir, new_filename)
                        resized_image.save(new_filepath, format='JPEG', quality=quality, dpi=(dpi, dpi))
                        os.remove(filepath)
                        update_all_rels_files(rels_dir, f'media/{filename}', f'media/{new_filename}')
                        image_count += 1
            except Exception as e:
                logging.error(f"Error processing image {filename}: {e}")

    if image_count == 0:
        logging.info("No valid images were found to convert.")
    else:
        logging.info(f"Converted {image_count} images successfully.")

    with zipfile.ZipFile(output_filepath, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, temp_dir)
                docx_zip.write(file_path, arcname)

    shutil.rmtree(temp_dir)
    return True

def display_file_size_reduction(original_filepath, converted_filepath):
    original_size = os.path.getsize(original_filepath)
    converted_size = os.path.getsize(converted_filepath)
    size_difference = original_size - converted_size
    percentage_reduction = (size_difference / original_size) * 100
    return f"Original file size: {original_size / 1024:.2f} KB\n" \
           f"Converted file size: {converted_size / 1024:.2f} KB\n" \
           f"Space saved: {size_difference / 1024:.2f} KB ({percentage_reduction:.2f}% reduction)"

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

def convert_file():
    input_path = entry_file_path.get()
    quality = entry_quality.get()
    dpi = entry_dpi.get()

    if not os.path.exists(input_path):
        messagebox.showerror("Error", f"The file '{input_path}' does not exist.")
        return

    if not allowed_file(input_path):
        messagebox.showerror("Error", "Please select a valid .docx file.")
        return

    if not quality.isdigit() or not dpi.isdigit():
        messagebox.showerror("Error", "Please enter valid numeric values for quality and DPI.")
        return

    quality = int(quality)
    dpi = int(dpi)

    if quality < 1 or quality > 100:
        messagebox.showerror("Error", "Quality must be between 1 and 100.")
        return

    output_filename = f"converted_{os.path.basename(input_path)}"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    label_status.config(text="Converting images...")
    success = convert_docx_images(input_path, output_path, quality=quality, dpi=dpi)

    if success:
        reduction_info = display_file_size_reduction(input_path, output_path)
        label_status.config(text=f"Conversion successful! File saved as:\n{output_path}")
        label_reduction.config(text=reduction_info)
    else:
        label_status.config(text="Conversion failed.")

# Create a GUI window
root = tk.Tk()
root.title("DOCX Image Converter")

# File selection
frame = tk.Frame(root)
frame.pack(pady=10)

label_select_file = tk.Label(frame, text="Select .docx file:")
label_select_file.grid(row=0, column=0, padx=10, pady=5)

entry_file_path = tk.Entry(frame, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)

button_browse = tk.Button(frame, text="Browse", command=select_file)
button_browse.grid(row=0, column=2, padx=10, pady=5)

# Quality input
label_quality = tk.Label(root, text="Image Quality (1-100):")
label_quality.pack(pady=5)
entry_quality = tk.Entry(root, width=10)
entry_quality.pack(pady=5)
entry_quality.insert(0, "85")  # Default value

# DPI input
label_dpi = tk.Label(root, text="Image DPI:")
label_dpi.pack(pady=5)
entry_dpi = tk.Entry(root, width=10)
entry_dpi.pack(pady=5)
entry_dpi.insert(0, "200")  # Default value

# Convert button
button_convert = tk.Button(root, text="Convert", command=convert_file)
button_convert.pack(pady=10)

# Status label
label_status = tk.Label(root, text="", fg="green")
label_status.pack(pady=10)

# Reduction information
label_reduction = tk.Label(root, text="", fg="blue")
label_reduction.pack(pady=10)

# Exit button
button_exit = tk.Button(root, text="Exit", command=root.quit)
button_exit.pack(pady=20)

# Start the Tkinter event loop
root.mainloop()