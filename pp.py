import streamlit as st
import os
import tempfile
import zipfile
import shutil
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Preformatted, Flowable  # Added Flowable
from reportlab.lib.colors import Color, HexColor, black, white
from pygments import highlight
from pygments.lexers import get_lexer_for_filename, PythonLexer, guess_lexer
from pygments.formatters import HtmlFormatter
from pygments.styles import get_style_by_name
import html2text
import itertools
from tqdm import tqdm
import math
from streamlit.components.v1 import html
import base64

class ColoredPreformattedText(Flowable):
    def __init__(self, text, style, color_map):
        Flowable.__init__(self)
        self.text = text.rstrip()
        self.style = style
        self.color_map = color_map
        self.split_lines = []
    
    def wrap(self, awidth, aheight):
        # Calculate lines that will fit on current page
        available_height = aheight - self.style.fontSize
        max_lines = int(available_height / self.style.leading)
        
        lines = self.text.split('\n')
        if not self.split_lines:
            self.split_lines = lines
        
        # Get lines that will fit
        current_lines = self.split_lines[:max_lines]
        # Store remaining lines for next page
        self.split_lines = self.split_lines[max_lines:]
        
        self.width = awidth
        self.height = len(current_lines) * self.style.leading + self.style.fontSize
        self.current_text = '\n'.join(current_lines)
        
        return self.width, self.height
    
    def draw(self):
        self.canv.saveState()
        self.canv.setFont(self.style.fontName, self.style.fontSize)
        
        # Draw background with padding
        padding = 5
        self.canv.setFillColor(HexColor('#1e1e1e'))
        self.canv.rect(-padding, -padding, 
                      self.width + (2 * padding), 
                      self.height + (2 * padding), 
                      fill=True)
        
        # Add line numbers
        line_number_width = 30
        self.canv.setFillColor(HexColor('#666666'))
        self.canv.setFont(self.style.fontName, self.style.fontSize - 1)
        
        y = self.height - self.style.fontSize
        for i, line in enumerate(self.current_text.split('\n'), 1):
            self.canv.drawRightString(line_number_width - 5, y, str(i))
            
            # Draw line number background
            self.canv.setFillColor(HexColor('#2a2a2a'))
            self.canv.rect(0, y - 2, line_number_width, self.style.leading, fill=True)
            
            # Draw actual code
            x = line_number_width + 5
            parts = line.split('█')
            for j, part in enumerate(parts):
                if j % 2 == 1 and part[:2] in self.color_map:
                    color_code = part[:2]
                    text = part[2:]
                    self.canv.setFillColor(self.color_map[color_code])
                else:
                    text = part
                    self.canv.setFillColor(white)
                
                self.canv.drawString(x, y, text)
                x += self.canv.stringWidth(text, self.style.fontName, self.style.fontSize)
            y -= self.style.leading
        
        self.canv.restoreState()
    
    def split(self, awidth, aheight):
        # If there are remaining lines, create a new flowable
        if self.split_lines:
            return [self]
        return []

def clean_line_numbers(text):
    """Remove line number artifacts and clean up the text"""
    import re
    # Remove line number spans and styling
    text = re.sub(r'<span[^>]*style="color: #37474F[^>]*>\s*\d+\s*</span>', '', text)
    # Remove continued on next page artifacts
    text = re.sub(r'\(continued on next page\.\.\.\)', '', text)
    # Remove any remaining line number artifacts
    text = re.sub(r'^\s*\d+\s', '', text, flags=re.MULTILINE)
    return text

def add_color_markers(code_text):
    """Add color markers to code text based on token types and clean up HTML artifacts"""
    # First clean up line numbers
    code_text = clean_line_numbers(code_text)
    
    # Clean up HTML structure more thoroughly
    html_cleanup = [
        ('<div[^>]*class="source"[^>]*>', ''),
        ('<div[^>]*class="linenodiv"[^>]*>', ''),
        ('<div[^>]*class="highlight"[^>]*>', ''),
        ('<table[^>]*class="sourcetable"[^>]*>', ''),
        ('</div>', ''),
        ('</table>', ''),
        ('<td[^>]*class="linenos"[^>]*>', ''),
        ('<td[^>]*class="code"[^>]*>', ''),
        ('</td>', ''),
        ('<tr[^>]*>', ''),
        ('</tr>', ''),
        ('<pre[^>]*>', ''),
        ('</pre>', '')
    ]
    
    import re
    for pattern, replacement in html_cleanup:
        code_text = re.sub(pattern, replacement, code_text)
    
    # Enhanced color markers with more token types
    color_replacements = {
        'class="c1"': '█cm',     # Comments - Green
        'class="s1"': '█st',     # Strings - Yellow
        'class="s2"': '█st',     # Strings - Yellow
        'class="k"': '█kw',      # Keywords - Pink
        'class="kc"': '█kw',     # Keywords constant - Pink
        'class="kd"': '█kw',     # Keywords declaration - Pink
        'class="kt"': '█kw',     # Keywords type - Pink
        'class="n"': '█nm',      # Names - Light blue
        'class="na"': '█nm',     # Name attribute - Light blue
        'class="nb"': '█bi',     # Built-in - Lime
        'class="nf"': '█fn',     # Functions - Lime
        'class="nc"': '█cl',     # Classes - Orange
        'class="nn"': '█ns',     # Namespace - Light blue
        'class="o"': '█op',      # Operators - Orange
        'class="ow"': '█ow',     # Operator words - Pink
        'class="p"': '█pu',      # Punctuation - Light gray
        'class="mi"': '█nu',     # Numbers - Purple
        'class="mf"': '█nu',     # Float numbers - Purple
    }
    
    for old, new in color_replacements.items():
        code_text = code_text.replace(old, new)
    
    # Clean up any remaining spans and HTML
    code_text = re.sub(r'<[^>]+>', '', code_text)
    
    # Remove extra whitespace while preserving indentation
    lines = []
    for line in code_text.splitlines():
        if line.strip():
            lines.append(line.rstrip())
    
    return '\n'.join(lines)

def get_file_extension(filename):
    return os.path.splitext(filename)[1].lower()

def is_code_file(filename):
    code_extensions = {
        '.py', '.java', '.cpp', '.js', '.html', '.css', '.php', '.rb', '.go', '.ts', 
        '.jsx', '.tsx', '.json', '.xml', '.yaml', '.yml', '.sql', '.sh', '.bat',
        '.c', '.h', '.cs', '.swift', '.kt', '.rs', '.m', '.mm', '.scala'
    }
    return get_file_extension(filename) in code_extensions

def extract_code_content(content):
    try:
        text = content.decode('utf-8')
    except UnicodeDecodeError:
        try:
            text = content.decode('latin-1')
        except Exception:
            text = content.decode('utf-8', errors='ignore')
    
    lines = text.splitlines()
    cleaned_lines = []
    prev_line_empty = False
    
    for line in lines:
        is_empty = not line.strip()
        if is_empty and prev_line_empty:
            continue
        cleaned_lines.append(line)
        prev_line_empty = is_empty
    
    return '\n'.join(cleaned_lines)

def format_size(size):
    """Convert bytes to human readable format"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024.0:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} TB"

def get_html_preview(code, lexer):
    """Generate HTML preview with syntax highlighting"""
    formatter = HtmlFormatter(
        style='monokai',  # Changed to monokai for more vibrant colors
        linenos='table',  # Changed to 'table' for better structure
        linenostart=1,
        lineanchors='line',
        linespans='line',
        cssclass="source",
        noclasses=True,
        wrapcode=True,
        nobackground=True  # Prevent background color conflicts
    )
    highlighted = highlight(code, lexer, formatter)
    
    # Clean up HTML output
    highlighted = highlighted.replace('<div class="linenodiv">', '<div class="linenodiv" style="user-select: none;">')
    highlighted = highlighted.replace('<table class="sourcetable">', '<table class="sourcetable" style="width: 100%;">')
    
    return highlighted

def show_preview(file_path, code):
    """Show preview of the code file"""
    try:
        lexer = get_lexer_for_filename(file_path)
    except:
        try:
            lexer = guess_lexer(code)
        except:
            lexer = PythonLexer()
    
    html_content = get_html_preview(code, lexer)
    
    # Add custom CSS for better preview
    preview_css = """
    <style>
        .preview-container { 
            border: 1px solid #ccc;
            border-radius: 5px;
            padding: 10px;
            background-color: #1e1e1e;
            margin: 10px 0;
            overflow-x: auto;
        }
        .source { font-size: 14px; font-family: 'Courier New', monospace; }
        .source pre { margin: 0; padding: 10px; tab-size: 4; }
        .source .linenodiv { 
            color: #888;
            background: #2a2a2a;
            padding: 0 10px;
            border-right: 1px solid #404040;
            text-align: right;
            user-select: none;
        }
        .source .code { padding-left: 10px; }
        
        /* Enhanced syntax highlighting colors */
        .source .c1 { color: #7CC379 !important; }  /* Comments - Green */
        .source .s1, .source .s2 { color: #E6DB74 !important; }  /* Strings - Yellow */
        .source .k, .source .kc, .source .kd { color: #FF6188 !important; }  /* Keywords - Pink */
        .source .n { color: #78DCE8 !important; }  /* Names - Light blue */
        .source .o { color: #FF9D00 !important; }  /* Operators - Orange */
        .source .p { color: #F8F8F2 !important; }  /* Punctuation - Light gray */
        .source .nb { color: #A9DC76 !important; }  /* Built-in - Lime */
        .source .nf { color: #A9DC76 !important; }  /* Functions - Lime */
        .source .nc { color: #FFB86C !important; }  /* Classes - Orange */
        .source .nn { color: #78DCE8 !important; }  /* Namespace - Light blue */
        .source .ow { color: #FF6188 !important; }  /* Operator words - Pink */
    </style>
    """
    
    st.markdown(preview_css, unsafe_allow_html=True)
    st.markdown(f"<div class='preview-container'>{html_content}</div>", unsafe_allow_html=True)

def create_pdf_from_code(folder_path, output_pdf):
    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    
    styles = getSampleStyleSheet()
    heading_style = ParagraphStyle(
        'HeadingStyle',
        parent=styles['Heading2'],
        spaceBefore=10,
        spaceAfter=5
    )
    
    code_style = ParagraphStyle(
        'CodeStyle',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=10,  # Increased from 8
        spaceAfter=5,
        leading=12,   # Increased from 8
        spaceBefore=0
    )
    
    content = []
    
    files_to_process = []
    for root, _, files in os.walk(folder_path):
        for file in sorted(files):
            if is_code_file(file):
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, folder_path)
                files_to_process.append((file_path, relative_path))
    
    total_size = sum(os.path.getsize(os.path.join(root, file)) 
                    for root, _, files in os.walk(folder_path) 
                    for file in files if is_code_file(file))
    
    processed_size = 0
    progress_bar = st.progress(0)
    status_text = st.empty()
    file_info = st.empty()
    
    color_map = {
        'cm': HexColor('#7CC379'),  # Comments - Green
        'st': HexColor('#E6DB74'),  # Strings - Yellow
        'kw': HexColor('#FF6188'),  # Keywords - Pink
        'nm': HexColor('#78DCE8'),  # Names - Light blue
        'op': HexColor('#FF9D00'),  # Operators - Orange
        'pu': HexColor('#F8F8F2'),  # Punctuation - Light gray
        'bi': HexColor('#A9DC76'),  # Built-in - Lime
        'fn': HexColor('#A9DC76'),  # Functions - Lime
        'cl': HexColor('#FFB86C'),  # Classes - Orange
        'ns': HexColor('#78DCE8'),  # Namespace - Light blue
        'ow': HexColor('#FF6188'),  # Operator words - Pink
    }
    
    for idx, (file_path, relative_path) in enumerate(files_to_process):
        processed_size += os.path.getsize(file_path)
        progress = processed_size / total_size
        progress_bar.progress(progress)
        
        status_text.text(f"Processing: {relative_path}")
        file_info.text(f"Processed: {format_size(processed_size)} / {format_size(total_size)} ({int(progress * 100)}%)")
        
        try:
            with open(file_path, 'rb') as f:
                code = extract_code_content(f.read())
            
            try:
                lexer = get_lexer_for_filename(file_path)
            except:
                try:
                    lexer = guess_lexer(code)
                except:
                    lexer = PythonLexer()
            
            formatter = HtmlFormatter(
                style='material',  # Changed to material theme
                linenos=True,
                cssclass="source",
                noclasses=True,    # Include CSS directly in output
                wrapcode=True
            )
            highlighted_code = highlight(code, lexer, formatter)
            
            code_text = add_color_markers(highlighted_code)
            
            # Split code into smaller chunks if too large
            lines = [line for line in code_text.splitlines() if line.strip()]
            max_lines_per_page = 50  # Adjust this value based on font size and page size
            
            for i in range(0, len(lines), max_lines_per_page):
                chunk = lines[i:i + max_lines_per_page]
                code_chunk = '\n'.join(chunk)
                
                if i == 0:  # First chunk includes the header
                    content.append(Paragraph(
                        f'<font color="blue"><b>📄 {relative_path}</b></font>',
                        heading_style
                    ))
                    content.append(Spacer(1, 5))
                
                colored_code = ColoredPreformattedText(code_chunk, code_style, color_map)
                content.append(colored_code)
                content.append(Spacer(1, 8))
                
                if i + max_lines_per_page < len(lines):  # Add continuation marker
                    content.append(Paragraph(
                        '<i>(continued on next page...)</i>',
                        styles['Italic']
                    ))
                    content.append(Spacer(1, 8))
            
        except Exception as e:
            error_msg = f"❌ Error processing {relative_path}: {str(e)}"
            content.append(Paragraph(
                f'<font color="red">{error_msg}</font>',
                styles['Normal']
            ))
            content.append(Spacer(1, 8))
    
    progress_bar.empty()
    status_text.empty()
    file_info.empty()
    
    st.success(f"✅ Successfully processed {len(files_to_process)} files ({format_size(total_size)})")
    
    doc.build(content)
    return True

def main():
    st.set_page_config(page_title="Code to PDF Converter", page_icon="📄", layout="wide")
    st.title("Code to PDF Converter")
    st.markdown("""
    Convert your code files or entire folders to a PDF document with syntax highlighting.
    
    **Upload Options:**
    1. Select individual code files
    2. Upload a ZIP file containing code
    3. Upload an entire folder (drag and drop folder)
    
    **Supported file types:** Python, JavaScript, TypeScript, Java, C++, C#, Swift, Ruby, Go, PHP, HTML, CSS, and more...
    """)
    
    uploaded_files = st.file_uploader(
        "Upload your code files, folder, or ZIP archive",
        type=['py', 'js', 'java', 'cpp', 'cs', 'swift', 'html', 'css', 'php', 'rb', 'go', 'ts', 'zip'],
        accept_multiple_files=True,
        help="Drag and drop a folder, multiple files, or a ZIP archive"
    )
    
    if uploaded_files:
        total_upload_size = sum(file.size for file in uploaded_files)
        st.info(f"Total upload size: {format_size(total_upload_size)}")
        
        # Add preview section
        st.subheader("📝 Preview")
        preview_container = st.empty()
        
        with st.spinner("Processing files and generating PDF..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                for uploaded_file in uploaded_files:
                    if uploaded_file.name.endswith('.zip'):
                        zip_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(zip_path, 'wb') as f:
                            f.write(uploaded_file.getvalue())
                        
                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)
                    else:
                        file_path = os.path.join(temp_dir, uploaded_file.name)
                        os.makedirs(os.path.dirname(file_path), exist_ok=True)
                        with open(file_path, 'wb') as f:
                            f.write(uploaded_file.getvalue())
                        
                        # Show preview for non-zip files
                        with open(file_path, 'rb') as f:
                            code = extract_code_content(f.read())
                            with st.expander(f"Preview: {uploaded_file.name}", expanded=False):
                                show_preview(file_path, code)
                
                pdf_path = os.path.join(temp_dir, 'code_documentation.pdf')
                success = create_pdf_from_code(temp_dir, pdf_path)
                
                if success:
                    with open(pdf_path, 'rb') as pdf_file:
                        pdf_bytes = pdf_file.read()
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_bytes,
                            file_name="code_documentation.pdf",
                            mime="application/pdf",
                            key="download_pdf"
                        )
                    st.success("✅ PDF generated successfully!")
                else:
                    st.error("❌ Failed to generate PDF. Please check your files and try again.")

if __name__ == "__main__":
    main()