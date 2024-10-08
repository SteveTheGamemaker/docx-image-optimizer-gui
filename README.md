# 'Smush' DOCX Image Converter
Smush is a tool that reduces the size of `.docx` files by converting and resizing images within the document, saving space while maintaining image quality. This significantly reduces the tedium of performing the process manually.    
**Notes:**    
- This process works on both locked and unlocked documents in SVN (since Smush only saves copies of the files to a different location than the repo). 
- This guide is specific to Windows.
   
   
## Features
- Converts images in `.docx` files to JPEG format.
- Resizes images to a user-specified width and DPI.
- Maintains document structure while compressing images.
- Displays file size reduction after conversion.


## Prerequisites
1. Install Python (if not already).   
		1.1. Check if Python is already installed (many computers already have Python).   
			1.1.1. Press the Windows + R keys to open the Run window.   
			1.1.2. Type `cmd.exe` and press enter to open a Command Prompt window.   
			1.1.3. Type `python --version` and press enter.   
			1.1.4. If the result does not display *Python 3.#.#*, proceed to 1.2.   
		1.2. Install Python from https://www.python.org/   
	     
2. Install the Pillow package to allow the tool to process images.   
		2.1. Open a Command Prompt window (if not already).    
		2.2. Type `pip install pillow` and press enter.    
		2.3. A successful installation message ends with, *Successfully installed pillow-10.#.#.*   


## How to Use   
1. **Run the Application:**   
		1.1. Make sure all Prerequisites are met.    
		1.2. Verify the Command Prompt is set to the same directory Smush is saved to.   
			1.2.1. Change the directory if necessary by typing `cd`, followed by the file path to Smush, then press enter.    
			- Example: Smush is saved/was extracted to the Desktop, in the *smush-main* folder, but the Command Prompt is set to *C:\WINDOWS\system32*   
			Type, `cd C:\Users\[myusername]\Desktop\smush-main` and press enter.     
		1.3. Type `python convertgui.py` to run Smush.    
        
2. **Select a DOCX File:**   
		2.1. In the DOCX Image Converter window, click *Browse* and select the DOCX file you want to process.    
       
3. **Adjust Image Quality and DPI:**    
		3.1. Enter the desired image quality (1-100) and DPI. Defaults are pre-set to 85 (quality) and 200 (DPI), but you can adjust them as needed.    
     
4. **Convert the File:**    
		4.1. Click *Convert* to start the image conversion process. Smush will resize and compress the images.    
       
5. **Check the Results:**    
		5.1. Once the conversion is complete, the application will show how much space was saved and where the converted file is located.    
   
6. **Exit the Application:**   
		6.1. Click *Exit* to close the tool when finished.    
        		
      		
## Revision History   
Changes to tool use/README instructions.   
   
01 || 2024-SEP-14    
	- First release.    
