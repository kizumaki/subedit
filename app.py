import streamlit as st
import pandas as pd
import re
import io
import os
import pysrt
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
from docx.enum.text import WD_TAB_ALIGNMENT
import random
from datetime import datetime

# --- GLOBAL CONFIGURATION ---
TARGET_FONT = 'Times New Roman'
TARGET_SIZE_PT = Pt(12)

# --- UTILITIES for WORD FORMATTER (Colors and Formatting) ---

def generate_vibrant_rgb_colors(count=150):
    """Generates a list of highly saturated, distinct RGB colors for speaker highlighting."""
    colors = set()
    while len(colors) < count:
        h = random.random(); s = 0.8; v = 0.9 
        
        if s == 0.0: r = g = b = v
        else:
            i = int(h * 6.0); f = h * 6.0 - i; p = v * (1.0 - s); q = v * (1.0 - s * f); t = v * (1.0 - s * (1.0 - f))
            if i % 6 == 0: r, g, b = v, t, p
            elif i % 6 == 1: r, g, b = q, v, p
            elif i % 6 == 2: r, g, b = p, v, t
            elif i % 6 == 3: r, g, b = p, q, v
            elif i % 6 == 4: r, g, b = t, p, v
            else: r, g, b = v, p, q
        
        r, g, b = int(r * 255), int(g * 255), int(b * 255)
        if (r < 50 and g < 50 and b < 50) or (r > 200 and g > 200 and b > 200): continue 
        colors.add((r, g, b))
    return list(colors)

# Initialize global color pool and mapping
FONT_COLORS_RGB_150 = generate_vibrant_rgb_colors(150)
speaker_color_map = {}
used_colors = []

def get_speaker_color(speaker_name):
    """Assigns and retrieves a unique color for a given speaker name."""
    global used_colors
    global speaker_color_map
    
    if speaker_name not in speaker_color_map:
        if not used_colors:
            # Reinitialize and shuffle if exhausted
            used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB_150]
            random.shuffle(used_colors)
            
        color_object = used_colors.pop()
        speaker_color_map[speaker_name] = color_object
        
    return speaker_color_map[speaker_name]

# Regex to find HTML tags (to split the text)
HTML_SPLIT_REGEX = re.compile(r'(<[ibu]>.*?<\/[ibu]>|<[ibu]/>)', re.IGNORECASE)
# Regex to extract content and tag name from a matched HTML block
HTML_EXTRACT_REGEX = re.compile(r"<([ibu])>(.*?)<\/\1>", re.IGNORECASE | re.DOTALL) 

def set_font_and_size(run, font_name, font_size):
    """Applies Font and Size to a specific run."""
    run.font.name = font_name
    run.font.size = font_size

def set_all_text_formatting(doc):
    """Applies Times New Roman 12pt and specific Spacing (Before: 0pt, After: 6pt, Single Line) to all runs/paragraphs."""
    for paragraph in doc.paragraphs:
        # Apply Font and Size
        for run in paragraph.runs:
            run.font.name = TARGET_FONT
            run.font.size = TARGET_SIZE_PT
        
        # Apply spacing
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(6)

# --- UTILITIES for EXCEL CONVERTER (SRT Parsing) ---

MAX_SPEAKER_NAME_LENGTH = 35 
MAX_SPEAKER_NAME_WORDS = 4 

NON_SPEAKER_PHRASES = [
    "the only problem", "note", "warning", "things", "and on the way we came across this", 
    "this is the highest swing in europe", "and i swear", "which meant", "the only thing is", 
    "and remember", "official distance", "first and foremost", "i said", 
    "here we go", "next up", "step 1", "step 2", "step 3", "and step 3", "first up", 
    "so the question is", "i was growing up", "you might be wondering", "update", 
    "nashville to miami", "all i know is", "unlike judy", "the good news is", 
    "aer lingus seat", "the true test is", "just as i suspected", "like i said", 
    "star review and said", "i told them all", "and best of all", "the point is", 
    "americans", "i was thinking", "and they go", "first of all", "second", 
    "are you like", "as a reminder", "round 2", "round 1", "round 3", "round 4", 
    "round 5", "welcome to round 3", "the question is", "quick reminder", 
    "in 2nd place", "coming up", "first stop", "next step", "and that means", 
    "hashtag", "so to be clear", "your second word", "welcome to round 6", 
    "battle finale time", "number 1", "number 2", "but the truth is", 
    "score to beat", "and your winner", "\"crafty\" and \"betcha\". coming up", 
    "next one", "keep in mind", "and it says", "you could say", "welcome to round 2", 
    "and the best part", "onto round 2", "the ride we chose", "good news is", 
    "bad news", "good news", "he thought", "3 teams remain"
]

COLOR_PALETTE = [
    'background-color: #ADD8E6; color: #000000', 'background-color: #90EE90; color: #000000', 
    'background-color: #FFB6C1; color: #000000', 'background-color: #FFFFE0; color: #000000', 
    'background-color: #DDA0DD; color: #000000', 'background-color: #AFEEEE; color: #000000', 
    'background-color: #F0E68C; color: #000000', 'background-color: #FFA07A; color: #000000', 
    'background-color: #E0FFFF; color: #000000', 'background-color: #F5F5DC; color: #000000', 
    'background-color: #2F4F4F; color: #FFFFFF', 'background-color: #191970; color: #FFFFFF', 
    'background-color: #006400; color: #FFFFFF', 'background-color: #800000; color: #FFFFFF', 
    'background-color: #4B0082; color: #FFFFFF', 'background-color: #556B2F; color: #FFFFFF', 
    'background-color: #8B4513; color: #FFFFFF', 'background-color: #36454F; color: #FFFFFF',
]

def clean_dialogue_text_for_excel(text):
    """
    Converts HTML/XML style formatting tags (i, b, u) to text enclosed in parentheses ().
    Removes any other HTML/XML tags. Used specifically for the Excel output.
    """
    text = re.sub(r'<i[^>]*>(.*?)</i[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<b[^>]*>(.*?)</b[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<u[^>]*>(.*?)</u[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<[^>]*>', '', text, flags=re.DOTALL)
    return re.sub(r'\s+', ' ', text).strip()

def is_valid_speaker_tag(tag):
    """Checks if a tag is likely a speaker name using linguistic heuristics."""
    tag = tag.strip()
    if not tag: return False
    if tag.lower() in NON_SPEAKER_PHRASES: return False
    if len(tag) > MAX_SPEAKER_NAME_LENGTH: return False
    normalized_tag = tag.replace(' and ', ' ').replace(' and', '').replace('&', ' ').strip()
    if not normalized_tag: return False
    word_count = len(normalized_tag.split())
    if word_count > MAX_SPEAKER_NAME_WORDS: return False 
    first_word = normalized_tag.split()[0] if normalized_tag.split() else normalized_tag
    if first_word[0].isalpha() and first_word[0].islower(): return False
    return True

def parse_srt_raw(srt_content):
    """
    Parses SRT content to extract Start, End timecodes, Speaker, and RAW Dialogue.
    (RAW dialogue means HTML tags like <b> and <i> are PRESERVED in the Dialogue column).
    """
    data = []
    blocks = re.split(r'\n\s*\n', srt_content.strip())
    last_known_speaker = "Unknown" 

    def append_row_and_update_state_raw(speaker, raw_dialogue):
        nonlocal last_known_speaker
        # Append RAW dialogue text (no cleaning applied here)
        data.append([time_start, time_end, speaker, raw_dialogue]) 
        last_known_speaker = speaker 

    for block in blocks:
        lines = block.strip().split('\n')
        if len(lines) < 3: continue

        time_line = lines[1].strip()
        time_match = re.match(r'(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})', time_line)
        if not time_match: continue

        time_start = time_match.group(1) 
        time_end = time_match.group(2)   

        dialogue_lines = lines[2:]
        current_dialogue = ""
        block_initial_speaker = last_known_speaker
        
        for line in dialogue_lines:
            line = line.strip()
            if not line: continue

            segments = re.split(r'((?:[\w\s&]+?): )', line)
            
            i = 0
            while i < len(segments):
                segment = segments[i].strip()
                i += 1
                
                if not segment: continue

                if segment.endswith(':') and len(segment) > 1:
                    speaker_tag = segment[:-1].strip()
                    
                    if is_valid_speaker_tag(speaker_tag):
                        
                        if current_dialogue:
                            speaker_to_use = block_initial_speaker if not data or data[-1][0] != time_start else last_known_speaker
                            append_row_and_update_state_raw(speaker_to_use, current_dialogue)
                            current_dialogue = "" 
                            
                        speaker = speaker_tag
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1
                        
                        if dialogue_segment:
                            append_row_and_update_state_raw(speaker, dialogue_segment)
                            
                        if block_initial_speaker == last_known_speaker:
                             block_initial_speaker = speaker
                            
                    else:
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1
                        recombined_text = segment + " " + dialogue_segment
                        
                        if current_dialogue: current_dialogue += " " + recombined_text
                        else: current_dialogue = recombined_text
                        
                else:
                    if current_dialogue: current_dialogue += " " + segment
                    else: current_dialogue = segment

        if current_dialogue:
            speaker_to_use = block_initial_speaker if not data or data[-1][0] != time_start else last_known_speaker
            append_row_and_update_state_raw(speaker_to_use, current_dialogue)

    return pd.DataFrame(data, columns=['Start', 'End', 'Speaker', 'Dialogue'])


# --- CORE FUNCTION: SRT to BASIC WORD (No Change Needed) ---

def process_srt_to_docx_basic(uploaded_file, file_name_without_ext):
    """Reads SRT file and converts it to DOCX with basic formatting."""
    
    srt_content = uploaded_file.getvalue().decode('utf-8')
    subs = pysrt.from_string(srt_content)
    document = Document()
    
    document.add_heading(f"SRT Conversion: {file_name_without_ext}", level=1)

    for sub in subs:
        # Add Index
        p_index = document.add_paragraph(f"{sub.index}")
        set_font_and_size(p_index.runs[0], TARGET_FONT, TARGET_SIZE_PT)
        p_index.paragraph_format.space_after = Pt(0) 

        # Add Timecode
        timecode_str = f"{sub.start} --> {sub.end}"
        p_timecode = document.add_paragraph(timecode_str)
        set_font_and_size(p_timecode.runs[0], TARGET_FONT, TARGET_SIZE_PT)
        p_timecode.paragraph_format.space_after = Pt(0)
        
        # Add Content (cleans up tags using pysrt)
        p_content = document.add_paragraph(sub.text_without_tags)
        if p_content.runs:
            set_font_and_size(p_content.runs[0], TARGET_FONT, TARGET_SIZE_PT)
        p_content.paragraph_format.space_after = Pt(12) 

    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    return modified_file

# --- CORE FUNCTION: WORD SCRIPT FORMATTER (FIXED LOGIC) ---

def build_formatted_docx_from_df(df_raw, file_name_without_ext):
    """
    Builds the final formatted DOCX document from the clean, structured DataFrame 
    (which contains raw dialogue text with HTML tags).
    """
    
    # Reset color mapping and color pool for a new file
    global speaker_color_map
    global used_colors
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB_150]
    random.shuffle(used_colors)
    
    document = Document()
    
    # --- A. Set Main Title (25pt, 2 blank lines after) ---
    title_paragraph = document.add_paragraph(file_name_without_ext.upper())
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_before = Pt(0)
    title_paragraph.paragraph_format.space_after = Pt(0) 
    
    title_run = title_paragraph.runs[0]
    title_run.font.name = TARGET_FONT
    title_run.font.size = Pt(25) 
    title_run.bold = True
    
    # Add two blank paragraphs for spacing
    document.add_paragraph().paragraph_format.space_after = Pt(0)
    document.add_paragraph().paragraph_format.space_after = Pt(0)

    # --- B. Process DataFrame rows and add to new document ---
    
    # Group by Timecode to ensure all dialogues under the same time block are together
    grouped_rows = df_raw.groupby(['Start', 'End'])
    
    for (start_time, end_time), rows in grouped_rows:
        
        # B.1 Add Timecode (Bold)
        timecode_str = f"{start_time} --> {end_time}"
        time_paragraph = document.add_paragraph(timecode_str)
        time_paragraph.paragraph_format.space_after = Pt(0) # Minimal space after timecode

        for run in time_paragraph.runs:
            run.font.bold = True
            
        # B.2 Add Dialogue rows
        for index, row in rows.iterrows():
            speaker_name = row['Speaker']
            dialogue_text = row['Dialogue']
            
            new_paragraph = document.add_paragraph()
            new_paragraph.style = document.styles['Normal']
            new_paragraph.paragraph_format.space_before = Pt(0) 
            new_paragraph.paragraph_format.space_after = Pt(6) 
            
            # Speaker and Indent Formatting
            if speaker_name not in ["Unknown", ""]:
                # Apply Hanging Indent (1 inch) and Tab Stop
                new_paragraph.paragraph_format.left_indent = Inches(1.0)
                new_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
                new_paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.0), WD_TAB_ALIGNMENT.LEFT)
                
                # 1. Run for the speaker name (Bold and Font Color)
                speaker_tag = f"{speaker_name}:"
                font_color_object = get_speaker_color(speaker_name) 
                
                run_speaker = new_paragraph.add_run(speaker_tag)
                run_speaker.font.bold = True
                run_speaker.font.color.rgb = font_color_object 
                
                # 2. Insert Tab character
                new_paragraph.add_run('\t') 
                
            else:
                # No speaker -> No indent
                new_paragraph.paragraph_format.left_indent = None
                new_paragraph.paragraph_format.first_line_indent = None
                
            # B.3 Process Dialogue Content for HTML Tags (i, b, u)
            
            # Use the fixed logic: Split the RAW dialogue text by the HTML tags
            parts = re.split(HTML_SPLIT_REGEX, dialogue_text)
            
            for part in parts:
                if not part:
                    continue
                
                # Check if the part is a standalone HTML tag block
                html_match = HTML_EXTRACT_REGEX.match(part)
                
                if html_match:
                    tag = html_match.group(1).lower()
                    content = html_match.group(2)
                    
                    run = new_paragraph.add_run(content)
                    run.font.bold = True  # Always Bold if coming from a tag (b or i)
                    run.font.italic = (tag == 'i') # Italic only if the tag is 'i'
                else:
                    # This is regular text (unstyled)
                    new_paragraph.add_run(part)

    # C. Apply General Font/Size and Spacing (Global settings)
    set_all_text_formatting(document)
    
    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    return modified_file


# --------------------------------------------------------------------------------
# --- STREAMLIT PAGES ---
# --------------------------------------------------------------------------------

def srt_to_docx_page():
    st.markdown("## 📄 SRT to Word (.docx) Converter - Basic")
    st.markdown("This function converts SRT files to DOCX, preserving line structure (index, timecode, content) with **Times New Roman, 12pt** format.")
    st.markdown("---")

    uploaded_file = st.file_uploader(
        "1. Upload your SRT file (.srt)",
        type=['srt'],
        key="srt_docx_uploader",
        help="Only .srt format is accepted."
    )

    if uploaded_file is not None:
        original_filename = uploaded_file.name
        file_name_without_ext = os.path.splitext(original_filename)[0]
        
        st.info(f"File received: **{original_filename}**.")
        
        if st.button("2. RUN WORD CONVERSION", key="run_srt_docx"):
            with st.spinner('Processing and creating Word file...'):
                try:
                    modified_file_io = process_srt_to_docx_basic(uploaded_file, file_name_without_ext)
                    
                    new_filename = f"BASIC_CONVERTED_{file_name_without_ext}.docx"

                    st.success("✅ Conversion complete! You can download the file.")
                    
                    st.download_button(
                        label="3. Download Converted Word File",
                        data=modified_file_io,
                        file_name=new_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    st.markdown("---")
                    st.balloons()

                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
                    st.warning("Please check the format of the input SRT file.")


def srt_to_excel_page():
    st.markdown("## 📊 Analyze & Convert SRT to Excel (.xlsx)")
    st.markdown("This function analyzes the SRT file to extract detailed dialogue and corresponding speaker, then exports to an Excel file. Highly useful for data control or content analysis.")
    st.markdown("---")
    
    # NOTE FOR USER: This section explains why coloring is not in the download (FIXED)
    st.warning("⚠️ **IMPORTANT NOTE ON EXCEL FORMATTING:** The colorful highlighting you see in the preview is for the web display only and CANNOT be included in the downloaded Excel (.xlsx) file due to file format limitations. The downloaded file will contain clean, organized data.")
    st.markdown("---")

    uploaded_file = st.file_uploader("1. Upload your SRT file (.srt)", type="srt", key="srt_excel_uploader")

    if uploaded_file is not None:
        try:
            try:
                # Try UTF-8 first
                srt_content = uploaded_file.read().decode("utf-8")
            except UnicodeDecodeError:
                # Fallback to Latin-1
                srt_content = uploaded_file.read().decode("latin-1")
                
        except Exception:
            st.error("File encoding error. Please ensure your SRT file is correctly encoded (UTF-8 recommended).")
            return

        with st.spinner('Analyzing SRT data...'):
            # Use the RAW parser (returns data structure with HTML tags)
            df_raw = parse_srt_raw(srt_content)
        
        if df_raw.empty:
            st.error("Could not parse any subtitles.")
            return
            
        # Create a CLEANED DataFrame for download and display (Excel format)
        df_cleaned = df_raw.copy()
        df_cleaned['Dialogue'] = df_cleaned['Dialogue'].apply(clean_dialogue_text_for_excel) # Clean dialogue text for Excel/Display


        st.subheader("📊 Speaker Statistics")
        
        unique_speakers = df_cleaned['Speaker'].unique()
        actual_speakers = [s for s in unique_speakers if s not in ["Unknown", ""]]
        speaker_count = len(actual_speakers)

        st.success(f"**Total Recognized Speakers:** {speaker_count} people.")
        
        if speaker_count > 0:
            speaker_list_str = ", ".join(actual_speakers)
            st.markdown(f"**List of Speakers:** {speaker_list_str}")
        else:
            st.info("No clear speakers found.")
            
        st.subheader("Converted Data Preview (Web Styling Only)")
        
        # Apply styling ONLY for the web preview
        styled_df_display = apply_styles(df_cleaned)
        st.dataframe(styled_df_display, use_container_width=True)

        st.markdown("---")
        
        output = io.BytesIO()
        # Save the CLEANED (un-styled) DataFrame for accurate Excel output
        df_cleaned.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        original_name_base = uploaded_file.name.rsplit('.', 1)[0]
        file_name = f"{original_name_base}_DATA.xlsx"
        
        st.download_button(
            label="💾 Download Analyzed Excel File (.xlsx)",
            data=output.read(), 
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.document"
        )
        st.success(f"File ready for download: **{file_name}**!")


def word_formatter_page():
    st.markdown("## 📝 Automatic Word Script Formatter Tool (SRT Input)")
    st.markdown("⚠️ **FIXED:** This tool now requires a **SRT file** (.srt) as input for correct formatting. It uses advanced parsing to separate speakers and applies professional layout:")
    st.markdown("- **Title:** Uppercase, 25pt size, centered.")
    st.markdown("- **Timecode:** Bold, minimal line spacing.")
    st.markdown("- **Speaker:** Bold, unique color (per speaker), professional hanging indent and tab stop.")
    st.markdown("- **HTML tags (e.g., `<i>` or `<b>`):** Converted to Bold/Italic formatting in the Word document.")
    st.markdown("---")

    uploaded_file = st.file_uploader(
        "1. Upload your SRT file (.srt)",
        type=['srt'],
        key="word_formatter_uploader",
        help="Please upload an SRT file for formatting."
    )

    if uploaded_file is not None:
        original_filename = uploaded_file.name
        file_name_without_ext = os.path.splitext(original_filename)[0]
        
        st.info(f"File received: **{original_filename}**.")
        
        if st.button("2. RUN AUTOMATIC FORMATTING", key="run_word_formatter"):
            with st.spinner('Processing and formatting file...'):
                try:
                    # FIX: Parse the SRT file to get the clean, structured data
                    df_raw = parse_srt_raw(uploaded_file.getvalue().decode('utf-8'))
                    
                    # FIX: Build the DOCX directly from the clean data structure
                    modified_file_io = build_formatted_docx_from_df(df_raw, file_name_without_ext)
                    
                    new_filename = f"FORMATTED_{file_name_without_ext}.docx"

                    st.success("✅ Formatting complete! You can download the file.")
                    
                    st.download_button(
                        label="3. Download Formatted Word File",
                        data=modified_file_io,
                        file_name=new_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    st.markdown("---")
                    st.balloons()

                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")
                    st.warning("Please check the format of the input SRT file.")


# --------------------------------------------------------------------------------
# --- MAIN APPLICATION ENTRY POINT ---
# --------------------------------------------------------------------------------

def main():
    """Defines the Streamlit application structure using sidebar navigation."""
    
    st.set_page_config(
        page_title="Subtitle & Script Toolkit", 
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.sidebar.title("🛠️ COMPREHENSIVE TOOLKIT")
    st.sidebar.markdown("Select a function to use:")
    
    # Navigation Radio Buttons
    app_mode = st.sidebar.radio(
        "Function",
        (
            "1. SRT to Word (Basic)",
            "2. SRT to Excel (Analysis)",
            "3. Word Script Formatting"
        )
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown(
        """
        **Usage:**
        - Each function operates independently.
        - Upload the file, run the process, and download the result.
        """
    )
    
    # Route to the selected page function
    if app_mode == "1. SRT to Word (Basic)":
        srt_to_docx_page()
    elif app_mode == "2. SRT to Excel (Analysis)":
        srt_to_excel_page()
    elif app_mode == "3. Word Script Formatting":
        word_formatter_page()

if __name__ == "__main__":
    # Ensure environment is ready for color mapping
    if not FONT_COLORS_RGB_150:
         FONT_COLORS_RGB_150 = generate_vibrant_rgb_colors(150)
         
    main()
