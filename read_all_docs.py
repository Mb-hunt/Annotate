import zipfile
import xml.etree.ElementTree as ET
import glob
import os
import sys
import re

def extract_text_from_docx(docx_path):
    print(f"--- Processing DOCX: {os.path.basename(docx_path)} ---")
    try:
        with zipfile.ZipFile(docx_path) as docx:
            xml_content = docx.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            
            text_parts = []
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            for element in tree.iter():
                if element.tag.endswith('}p'): # Paragraph
                    text_parts.append('\n')
                elif element.tag.endswith('}t'): # Text
                    if element.text:
                        text_parts.append(element.text)
                elif element.tag.endswith('}tab'):
                    text_parts.append('\t')
                elif element.tag.endswith('}br'): # Line break
                    text_parts.append('\n')
                    
            content = "".join(text_parts)
            print(content)
            print("\n" + "="*50 + "\n")
            
    except Exception as e:
        print(f"Error reading {docx_path}: {e}")

def extract_text_from_pptx(pptx_path):
    print(f"--- Processing PPTX: {os.path.basename(pptx_path)} ---")
    try:
        with zipfile.ZipFile(pptx_path) as pptx:
            # Get list of slide files, sorted numerically
            slides = [f for f in pptx.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            
            # Sort slides numerically (slide1, slide2, slide10...)
            def slide_number(s):
                match = re.search(r'slide(\d+)\.xml', s)
                return int(match.group(1)) if match else 0
            
            slides.sort(key=slide_number)
            
            full_text = []

            for slide in slides:
                xml_content = pptx.read(slide)
                tree = ET.fromstring(xml_content)
                slide_text = []
                
                # Namespace for DrawingML usually contains 'a' tags like a:t
                # But simple iteration often works without full namespace handling if we look for tags ending in specific names
                
                for element in tree.iter():
                    # Text happens in <a:t>
                    if element.tag.endswith('}t'):
                        if element.text:
                            slide_text.append(element.text)
                    elif element.tag.endswith('}p'): # Paragraph break often implied
                        slide_text.append('\n')
                
                if slide_text:
                    full_text.append(f"[Slide {slide_number(slide)}]\n" + "".join(slide_text))
            
            print("\n\n".join(full_text))
            print("\n" + "="*50 + "\n")

    except Exception as e:
        print(f"Error reading {pptx_path}: {e}")

def main():
    docx_files = glob.glob("*.docx")
    pptx_files = glob.glob("*.pptx")
    
    with open("updated_docs_content.txt", "w", encoding="utf-8") as f:
        original_stdout = sys.stdout
        sys.stdout = f
        
        # Process specific priority files first if present
        priorities = ["FAQs Computer Use Eval (1).docx", "Weekly Training_ Factual Accuracy.pptx", "Weekly Training 2_ Best Response Guide & Overall Justification.pptx"]
        
        # Helper to process if exists
        processed = set()
        
        for p in priorities:
            if os.path.exists(p):
                if p.endswith('.docx'):
                    extract_text_from_docx(p)
                elif p.endswith('.pptx'):
                    extract_text_from_pptx(p)
                processed.add(p)
        
        # Process remaining
        for file in docx_files:
            if file not in processed:
                extract_text_from_docx(file)
                
        for file in pptx_files:
            if file not in processed:
                extract_text_from_pptx(file)

        sys.stdout = original_stdout

if __name__ == "__main__":
    main()
