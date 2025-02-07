import os
import requests
import pdfkit
import subprocess
import tempfile
from PyPDF2 import PdfMerger, PdfReader
from readability import Document
import validators
from urllib.parse import urlparse
from PIL import Image

# ---------------------------- Common Functions ------------------------------
def clear_screen():
    """Clear terminal screen"""
    os.system('cls' if os.name == 'nt' else 'clear')

def show_menu():
    """Display main menu"""
    clear_screen()
    print("=== 📚 PDF Master ===")
    print("1. Merge PDF Files")
    print("2. Blog to PDF Converter")
    print("3. Create PDF")
    print("4. Exit")
    choice = input("\nChoose an option (1-4): ")
    return choice

def create_pdf_menu():
    """Display Create PDF sub-menu"""
    clear_screen()
    print("=== 🛠️ Create PDF ===")
    print("1. Image to PDF")
    print("2. Word to PDF")
    print("3. Excel to PDF")
    print("4. PowerPoint to PDF")
    print("5. Return to Main Menu")
    sub_choice = input("\nChoose an option (1-5): ")
    return sub_choice

def get_output_filename(default_name):
    """Get validated output filename"""
    while True:
        name = input(f"\n💾 Enter output filename ({default_name}): ").strip()
        if not name:
            name = default_name
        if not name.lower().endswith('.pdf'):
            name += '.pdf'
        if os.path.dirname(name):
            print("❌ Please enter a filename only, not a path")
            continue
        return name

# ---------------------------- PDF Merger Feature ----------------------------
def pdf_merger():
    """Handle PDF merging functionality"""
    clear_screen()
    print("=== 🖨️ PDF Merger ===")
    files = get_file_names()
    
    if not files:
        print("\n⚠️  No valid PDF files selected. Returning to menu.")
        return
    
    output_name = get_output_filename("merged.pdf")
    merge_pdfs(files, output_name)

def get_file_names():
    """Collect and validate PDF filenames"""
    file_names = []
    print("\n📁 Enter PDF filenames from current directory (type 'q' to finish):")
    
    while True:
        file_name = input("> ").strip()
        
        if file_name.lower() == 'q':
            break
            
        try:
            validate_pdf_file(file_name)
            file_names.append(file_name)
            print(f"✅ Validated: {file_name}")
        except Exception as e:
            print(f"❌ Error: {str(e)}")

    return file_names

def validate_pdf_file(filename):
    """Validate PDF file"""
    if os.path.dirname(filename):
        raise ValueError("Please enter filename only")
    if not os.path.exists(filename):
        raise FileNotFoundError(f"File '{filename}' not found")
    if not filename.lower().endswith('.pdf'):
        raise ValueError("Not a PDF file")
    with open(filename, 'rb') as f:
        PdfReader(f)

def merge_pdfs(file_paths, output_filename):
    """Merge PDF files with error handling"""
    merger = PdfMerger()
    error_count = 0

    for file_path in file_paths:
        try:
            merger.append(file_path)
            print(f"✅ Added: {file_path}")
        except Exception as e:
            print(f"❌ Failed {file_path}: {str(e)}")
            error_count += 1

    try:
        with open(output_filename, 'wb') as f:
            merger.write(f)
        print(f"\n🎉 Successfully created: {output_filename}")
        print(f"📊 Stats: {len(file_paths)-error_count} succeeded, {error_count} failed")
    except Exception as e:
        print(f"\n🔥 Critical error: {str(e)}")
    finally:
        merger.close()

# ------------------------- Blog to PDF Feature ------------------------------
def blog_to_pdf():
    """Handle blog to PDF conversion"""
    clear_screen()
    print("=== ✨ Blog to PDF Converter ===")
    url = get_valid_url()
    
    if not url:
        print("\n⚠️  Invalid URL. Returning to menu.")
        return
    
    output_name = get_output_filename("blog_post.pdf")
    convert_blog(url, output_name)

def get_valid_url():
    """Get and validate blog URL"""
    print("\n🌐 Enter blog post URL (type 'q' to return):")
    while True:
        url = input("> ").strip()
        
        if url.lower() == 'q':
            return None
            
        if not validators.url(url):
            print("❌ Invalid URL format")
            continue
            
        if not is_blog_url(url):
            print("⚠️  This doesn't appear to be a blog post URL")
            proceed = input("Continue anyway? (y/n): ").lower()
            if proceed != 'y':
                continue
                
        return url

def is_blog_url(url):
    """Basic heuristic to detect blog URLs"""
    parsed = urlparse(url)
    path = parsed.path.lower()
    
    blog_indicators = [
        '/blog/', '/post/', '/article/',
        'medium.com', 'wordpress.com', 'blogspot'
    ]
    
    return any(indicator in path or indicator in parsed.netloc 
              for indicator in blog_indicators)

def convert_blog(url, output_filename):
    """Convert blog post to PDF"""
    try:
        print("\n⏳ Downloading content...")
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        print("🔍 Processing content...")
        doc = Document(response.text)
        content = doc.summary()
        
        print("🛠  Generating PDF...")
        styled_content = add_pdf_styles(content)
        pdfkit.from_string(styled_content, output_filename, options={
            'encoding': 'UTF-8',
            'quiet': ''
        })
        
        print(f"\n🎉 Successfully created: {output_filename}")
        print(f"📄 Article title: {doc.title()}")
        
    except Exception as e:
        print(f"\n🔥 Conversion failed: {str(e)}")
        if os.path.exists(output_filename):
            os.remove(output_filename)

def add_pdf_styles(html_content):
    """Add PDF-friendly CSS styles"""
    styles = """
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 2cm; }
        img { max-width: 100%; height: auto; }
        pre { background: #f4f4f4; padding: 10px; overflow: auto; }
        code { font-family: Monaco, Consolas, monospace; }
        h1, h2, h3 { color: #2c3e50; }
        a { color: #3498db; text-decoration: none; }
    </style>
    """
    return f"{styles}\n{html_content}"

# ------------------------ Image to PDF Feature ------------------------------
def image_to_pdf():
    """Handle image to PDF conversion"""
    clear_screen()
    print("=== 🖼️ Image to PDF Converter ===")
    files = get_image_files()
    
    if not files:
        print("\n⚠️  No valid image files selected. Returning to menu.")
        return
    
    output_name = get_output_filename("images.pdf")
    convert_images_to_pdf(files, output_name)

def get_image_files():
    """Collect and validate image filenames"""
    file_names = []
    print("\n📷 Enter image filenames (type 'q' to finish):")
    
    while True:
        file_name = input("> ").strip()
        
        if file_name.lower() == 'q':
            break
            
        try:
            validate_image_file(file_name)
            file_names.append(file_name)
            print(f"✅ Validated: {file_name}")
        except Exception as e:
            print(f"❌ Error: {str(e)}")

    return file_names

def validate_image_file(filename):
    """Validate image file"""
    if os.path.dirname(filename):
        raise ValueError("Please enter filename only")
    if not os.path.exists(filename):
        raise FileNotFoundError(f"File '{filename}' not found")
    valid_extensions = ('.jpg', '.jpeg', '.png', '.webp', '.bmp')
    if not filename.lower().endswith(valid_extensions):
        raise ValueError("Not a supported image file")
    try:
        with Image.open(filename) as img:
            img.verify()
    except Exception as e:
        raise ValueError(f"Invalid image file: {str(e)}")

def convert_images_to_pdf(image_paths, output_filename):
    """Convert images to PDF with error handling"""
    images = []
    try:
        for path in image_paths:
            try:
                img = Image.open(path)
                if img.mode == 'RGBA':
                    img = img.convert('RGB')
                images.append(img)
                print(f"✅ Processed: {path}")
            except Exception as e:
                print(f"❌ Failed {path}: {str(e)}")
        
        if not images:
            print("⚠️  No valid images to convert")
            return
        
        images[0].save(
            output_filename,
            save_all=True,
            append_images=images[1:],
            resolution=100.0
        )
        print(f"\n🎉 Successfully created: {output_filename}")
        print(f"📸 Number of images converted: {len(images)}")
        
    except Exception as e:
        print(f"\n🔥 Conversion failed: {str(e)}")
        if os.path.exists(output_filename):
            os.remove(output_filename)
    finally:
        for img in images:
            img.close()

# ------------------------ Office to PDF Features ----------------------------
def validate_office_file(filename, extensions, file_type):
    """Validate office file"""
    if os.path.dirname(filename):
        raise ValueError("Please enter filename only")
    if not os.path.exists(filename):
        raise FileNotFoundError(f"File '{filename}' not found")
    if not filename.lower().endswith(extensions):
        raise ValueError(f"Not a supported {file_type} file")

def get_office_files(extensions, file_type):
    """Collect and validate office filenames"""
    file_names = []
    print(f"\n📁 Enter {file_type} filenames (type 'q' to finish):")
    
    while True:
        file_name = input("> ").strip()
        
        if file_name.lower() == 'q':
            break
            
        try:
            validate_office_file(file_name, extensions, file_type)
            file_names.append(file_name)
            print(f"✅ Validated: {file_name}")
        except Exception as e:
            print(f"❌ Error: {str(e)}")

    return file_names

def convert_office_to_pdf(input_path, temp_dir, file_type):
    """Convert office file to PDF using LibreOffice"""
    try:
        subprocess.run([
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', temp_dir,
            input_path
        ], check=True)
        
        base_name = os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
        generated_pdf = os.path.join(temp_dir, base_name)
        
        if os.path.exists(generated_pdf):
            return generated_pdf
        raise RuntimeError(f"{file_type} conversion failed")
    except Exception as e:
        print(f"❌ Error converting {input_path}: {str(e)}")
        return None

def handle_office_conversion(file_type, extensions, default_output):
    """Generic handler for office conversions"""
    clear_screen()
    print(f"=== 📄 {file_type} to PDF Converter ===")
    files = get_office_files(extensions, file_type)
    
    if not files:
        print(f"\n⚠️  No valid {file_type} files selected. Returning to menu.")
        return
    
    output_name = get_output_filename(default_output)
    merger = PdfMerger()
    temp_dir = tempfile.TemporaryDirectory()
    error_count = 0
    
    for file_path in files:
        temp_pdf = convert_office_to_pdf(file_path, temp_dir.name, file_type)
        if temp_pdf:
            try:
                merger.append(temp_pdf)
                print(f"✅ Converted: {file_path}")
            except Exception as e:
                print(f"❌ Failed to merge {file_path}: {str(e)}")
                error_count += 1
        else:
            error_count += 1
    
    if len(merger.pages) == 0:
        print("⚠️  No valid PDFs generated. Exiting.")
        return
    
    try:
        with open(output_name, 'wb') as f:
            merger.write(f)
        print(f"\n🎉 Successfully created: {output_name}")
        print(f"📊 Stats: {len(files)-error_count} succeeded, {error_count} failed")
    except Exception as e:
        print(f"\n🔥 Critical error: {str(e)}")
    finally:
        merger.close()
        temp_dir.cleanup()

def word_to_pdf():
    """Handle Word to PDF conversion"""
    handle_office_conversion(
        "Word", 
        ('.docx', '.doc'), 
        "word_merged.pdf"
    )

def excel_to_pdf():
    """Handle Excel to PDF conversion"""
    handle_office_conversion(
        "Excel", 
        ('.xlsx', '.xls'), 
        "excel_merged.pdf"
    )

def powerpoint_to_pdf():
    """Handle PowerPoint to PDF conversion"""
    handle_office_conversion(
        "PowerPoint", 
        ('.pptx', '.ppt'), 
        "powerpoint_merged.pdf"
    )

# ------------------------ Main Program Flow -------------------------------
if __name__ == "__main__":
    while True:
        choice = show_menu()
        
        if choice == '1':
            pdf_merger()
        elif choice == '2':
            blog_to_pdf()
        elif choice == '3':
            while True:
                sub_choice = create_pdf_menu()
                if sub_choice == '1':
                    image_to_pdf()
                elif sub_choice == '2':
                    word_to_pdf()
                elif sub_choice == '3':
                    excel_to_pdf()
                elif sub_choice == '4':
                    powerpoint_to_pdf()
                elif sub_choice == '5':
                    break
                else:
                    print("\n⚠️  Invalid choice!")
                input("\nPress Enter to continue...")
        elif choice == '4':
            print("\n👋 Exiting...")
            break
        else:
            print("\n⚠️  Invalid choice!")
        
        input("\nPress Enter to continue...")