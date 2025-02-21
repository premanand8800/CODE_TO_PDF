import streamlit as st
import os
import tempfile
from pptx import Presentation
import win32com.client
import pythoncom
import base64
from PIL import Image
import io

def convert_ppsx_to_ppt(input_path, output_path):
    try:
        # Initialize PowerPoint COM object
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = False

        # Open the PPSX file
        presentation = powerpoint.Presentations.Open(input_path)
        
        # Save as PPT
        presentation.SaveAs(output_path, 17)  # 17 is the value for PPT format
        
        # Close
        presentation.Close()
        powerpoint.Quit()
        return True
    except Exception as e:
        st.error(f"Error during conversion: {str(e)}")
        return False
    finally:
        pythoncom.CoUninitialize()

def get_download_link(file_path, file_name):
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    b64 = base64.b64encode(bytes_data).decode()
    return f'<a href="data:application/vnd.ms-powerpoint;base64,{b64}" download="{file_name}">Download PPT File</a>'

def extract_preview_images(input_path):
    """Extract preview images from PPSX file"""
    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = False

        presentation = powerpoint.Presentations.Open(input_path)
        images = []
        
        for i in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides.Item(i)
            temp_image_path = f"temp_slide_{i}.png"
            
            # Export slide as image
            slide.Export(temp_image_path, "PNG")
            
            # Read the image and store in memory
            with open(temp_image_path, 'rb') as img_file:
                img_bytes = img_file.read()
                images.append(img_bytes)
            
            # Clean up temporary image file
            os.remove(temp_image_path)

        presentation.Close()
        powerpoint.Quit()
        return images
    except Exception as e:
        st.error(f"Error extracting previews: {str(e)}")
        return []
    finally:
        pythoncom.CoUninitialize()

def main():
    st.set_page_config(page_title="PPSX Viewer & Converter", page_icon="📊", layout="wide")
    
    st.title("PPSX Viewer & Converter")
    st.markdown("""
    View and convert PowerPoint Show (PPSX) files to PowerPoint Presentation (PPT) format.
    
    **Features:**
    1. Preview PPSX slides
    2. Convert to editable PPT format
    3. Download converted file
    """)

    uploaded_file = st.file_uploader("Upload PPSX file", type=['ppsx'])

    if uploaded_file:
        st.info(f"File uploaded: {uploaded_file.name}")
        
        # Create tabs for viewer and converter
        tab1, tab2 = st.tabs(["📺 Preview", "🔄 Convert"])
        
        # Process the file once and store in session state
        if 'preview_images' not in st.session_state:
            with st.spinner("Loading preview..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(input_path, 'wb') as f:
                        f.write(uploaded_file.getvalue())
                    st.session_state.preview_images = extract_preview_images(input_path)
        
        # Preview tab
        with tab1:
            if st.session_state.preview_images:
                col1, col2 = st.columns([5, 1])
                with col2:
                    current_slide = st.number_input(
                        "Slide Number", 
                        min_value=1, 
                        max_value=len(st.session_state.preview_images),
                        value=1
                    )
                with col1:
                    st.image(
                        st.session_state.preview_images[current_slide-1],
                        caption=f"Slide {current_slide} of {len(st.session_state.preview_images)}",
                        use_column_width=True
                    )
                
                # Add navigation buttons
                cols = st.columns(5)
                with cols[1]:
                    if st.button("⏮️ First"):
                        st.session_state.current_slide = 1
                with cols[2]:
                    if st.button("◀️ Previous") and current_slide > 1:
                        st.session_state.current_slide = current_slide - 1
                with cols[3]:
                    if st.button("Next ▶️") and current_slide < len(st.session_state.preview_images):
                        st.session_state.current_slide = current_slide + 1
                with cols[4]:
                    if st.button("Last ⏭️"):
                        st.session_state.current_slide = len(st.session_state.preview_images)
        
        # Convert tab
        with tab2:
            if st.button("Convert to PPT"):
                with st.spinner("Converting..."):
                    with tempfile.TemporaryDirectory() as temp_dir:
                        input_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(input_path, 'wb') as f:
                            f.write(uploaded_file.getvalue())
                        
                        output_filename = uploaded_file.name.rsplit('.', 1)[0] + '.ppt'
                        output_path = os.path.join(temp_dir, output_filename)
                        
                        if convert_ppsx_to_ppt(input_path, output_path):
                            st.success("Conversion completed!")
                            
                            with open(output_path, "rb") as f:
                                bytes_data = f.read()
                                st.download_button(
                                    label="📥 Download PPT",
                                    data=bytes_data,
                                    file_name=output_filename,
                                    mime="application/vnd.ms-powerpoint"
                                )
                        else:
                            st.error("Conversion failed. Please try again.")

    st.markdown("""
    ---
    ### Supported Features
    - Preview PPSX slides before conversion
    - Navigate through slides
    - Convert PPSX to editable PPT format
    - Preserve all animations and transitions
    - Maintain original formatting
    
    ### Note
    This application requires Microsoft PowerPoint to be installed on the server.
    """)

if __name__ == "__main__":
    main()
