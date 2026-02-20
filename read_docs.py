import zipfile
import xml.etree.ElementTree as ET
import os
import glob
import sys

def extract_text_from_docx(docx_path):
    print(f"--- Reading {os.path.basename(docx_path)} ---")
    try:
        with zipfile.ZipFile(docx_path) as docx:
            xml_content = docx.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            
            text_parts = []
            
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

def main():
    docx_files = glob.glob("*.docx")
    with open("all_docs_content.txt", "w", encoding="utf-8") as f:
        original_stdout = sys.stdout
        sys.stdout = f
        
        for file in docx_files:
            extract_text_from_docx(file)
            
        sys.stdout = original_stdout

if __name__ == "__main__":
    main()
