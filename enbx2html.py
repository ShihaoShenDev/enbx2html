import os
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path
import re
import argparse
import zipfile
import tempfile

class EnbxConverter:
    def __init__(self, source_dir, output_dir, file_title=None):
        self.source_dir = Path(source_dir)
        self.output_dir = Path(output_dir)
        self.file_title = file_title
        self.resources_dir = self.source_dir / "Resources"
        self.slides_dir = self.source_dir / "Slides"
        
        self.board_info = {}
        self.slide_order = []
        self.resource_map = {}
        self.slide_file_map = {}
        self.metadata = {}
        
        # XML Namespaces (if any, usually strictly required for ElementTree if defined in root)
        # In the XMLs we saw, Board.xml didn't have xmlns.
        # [Content_Types].xml had one.
        # We'll handle namespaces dynamically if needed, or ignore them.

    def parse_metadata(self):
        doc_xml = self.source_dir / "Document.xml"
        if not doc_xml.exists():
            print("Document.xml not found!")
            return

        try:
            tree = ET.parse(doc_xml)
            root = tree.getroot()
            self.metadata['Name'] = root.find('Name').text if root.find('Name') is not None else "Unknown"
            self.metadata['Creator'] = root.find('Creator').text if root.find('Creator') is not None else "Unknown"
            self.metadata['CreatedDateTime'] = root.find('CreatedDateTime').text if root.find('CreatedDateTime') is not None else "Unknown"
            self.metadata['ModifiedDateTime'] = root.find('ModifiedDateTime').text if root.find('ModifiedDateTime') is not None else "Unknown"
            print(f"Document Metadata: {self.metadata}")
        except Exception as e:
            print(f"Error parsing Document.xml: {e}")

    def parse_board(self):
        board_xml = self.source_dir / "Board.xml"
        if not board_xml.exists():
            print("Board.xml not found!")
            return

        tree = ET.parse(board_xml)
        root = tree.getroot()
        
        self.board_info['width'] = float(root.find('SlideWidth').text)
        self.board_info['height'] = float(root.find('SlideHeight').text)
        
        slides_node = root.find('Slides')
        if slides_node is not None:
            for item in slides_node.findall('Item'):
                self.slide_order.append(item.text)
        
        print(f"Board parsed: {self.board_info}, {len(self.slide_order)} slides found.")

    def parse_references(self):
        ref_xml = self.source_dir / "Reference.xml"
        if not ref_xml.exists():
            print("Reference.xml not found!")
            return

        tree = ET.parse(ref_xml)
        root = tree.getroot()
        
        rels = root.find('Relationships')
        if rels is not None:
            for rel in rels.findall('Relationship'):
                res_id = rel.find('Id').text
                target = rel.find('Target').text
                self.resource_map[res_id] = target.replace('\\', '/')
        
        print(f"References parsed: {len(self.resource_map)} resources found.")

    def map_slides(self):
        if not self.slides_dir.exists():
            print("Slides directory not found!")
            return

        for xml_file in self.slides_dir.glob("*.xml"):
            try:
                tree = ET.parse(xml_file)
                root = tree.getroot()
                slide_id = root.find('Id').text
                self.slide_file_map[slide_id] = xml_file
            except Exception as e:
                print(f"Error parsing {xml_file}: {e}")
        
        print(f"Mapped {len(self.slide_file_map)} slide files.")

    def copy_resources(self):
        if self.source_dir.resolve() == self.output_dir.resolve():
            print("Source and Output are the same directory. Skipping resource copy.")
            return

        dest_res = self.output_dir / "Resources"
        if dest_res.exists():
            shutil.rmtree(dest_res)
        
        if self.resources_dir.exists():
            shutil.copytree(self.resources_dir, dest_res)
            print("Resources copied.")
        else:
            print("No Resources folder to copy.")

    def generate_html(self):
        if not self.output_dir.exists():
            self.output_dir.mkdir(parents=True)
        
        # HTML Header
        html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{self.file_title if self.file_title else self.metadata.get('Name', 'EasiNote Export')}</title>
    <style>
        body {{
            margin: 0;
            padding: 0;
            background-color: #333;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            overflow: hidden;
            font-family: "Microsoft YaHei", sans-serif;
        }}
        #container {{
            position: relative;
            width: {self.board_info.get('width', 1280)}px;
            height: {self.board_info.get('height', 720)}px;
            background-color: white;
            overflow: hidden;
            box-shadow: 0 0 20px rgba(0,0,0,0.5);
        }}
        .slide {{
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            display: none;
            background-size: 100% 100%;
        }}
        .slide.active {{
            display: block;
        }}
        .element {{
            position: absolute;
            transform-origin: 50% 50%;
            white-space: pre-wrap; /* Preserve formatting */
            display: flex; /* For alignment */
            flex-direction: column;
        }}
        .element img {{
            width: 100%;
            height: 100%;
            display: block;
        }}
        .nav-buttons {{
            position: fixed;
            bottom: 20px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 1000;
        }}
        .info-button {{
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 1000;
        }}
        button {{
            padding: 10px 20px;
            font-size: 16px;
            cursor: pointer;
            background: rgba(255, 255, 255, 0.8);
            border: none;
            border-radius: 5px;
            margin: 0 5px;
        }}
        button:hover {{
            background: white;
        }}
        
        /* Modal Styles */
        .modal {{
            display: none; 
            position: fixed; 
            z-index: 2000; 
            left: 0;
            top: 0;
            width: 100%; 
            height: 100%; 
            overflow: auto; 
            background-color: rgba(0,0,0,0.4); 
        }}
        .modal-content {{
            background-color: #fefefe;
            margin: 15% auto; 
            padding: 20px;
            border: 1px solid #888;
            width: 50%; 
            border-radius: 10px;
            box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
        }}
        .close {{
            color: #aaa;
            float: right;
            font-size: 28px;
            font-weight: bold;
        }}
        .close:hover,
        .close:focus {{
            color: black;
            text-decoration: none;
            cursor: pointer;
        }}
        .info-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }}
        .info-table td, .info-table th {{
            border: 1px solid #ddd;
            padding: 8px;
        }}
        .info-table tr:nth-child(even){{background-color: #f2f2f2;}}
        .info-table th {{
            padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: #4CAF50;
            color: white;
        }}
    </style>
</head>
<body>
    <div id="container">
"""
        
        # Generate Slides
        for index, slide_id in enumerate(self.slide_order):
            if slide_id not in self.slide_file_map:
                print(f"Warning: Slide ID {slide_id} not found in mapped files.")
                continue
            
            slide_file = self.slide_file_map[slide_id]
            slide_html = self.render_slide(slide_file, index == 0)
            html_content += slide_html
        
        # HTML Footer & Scripts
        metadata_rows = ""
        # Localization and formatting
        key_map = {
            'Name': '文档名称',
            'Creator': '作者',
            'CreatedDateTime': '创建时间',
            'ModifiedDateTime': '上次修改时间'
        }
        
        for key, label in key_map.items():
            val = self.metadata.get(key, "")
            if val:
                if key == 'Creator':
                    val = f'<a href="https://k.seewo.com/personalPage/{val}" target="_blank">{val}</a>'
                metadata_rows += f"<tr><td>{label}</td><td>{val}</td></tr>"

        html_content += f"""
    </div>
    <div class="nav-buttons">
        <button onclick="prevSlide()">上一页</button>
        <button onclick="nextSlide()">下一页</button>
    </div>
    
    <div class="info-button">
        <button onclick="showInfo()">关于文档</button>
    </div>

    <div id="infoModal" class="modal">
        <div class="modal-content">
            <span class="close" onclick="closeInfo()">&times;</span>
            <h2>文档信息</h2>
            <table class="info-table">
                {metadata_rows}
            </table>
        </div>
    </div>

    <script>
        let currentSlide = 0;
        const slides = document.querySelectorAll('.slide');
        const modal = document.getElementById("infoModal");
        
        function showSlide(n) {{
            slides[currentSlide].classList.remove('active');
            currentSlide = (n + slides.length) % slides.length;
            slides[currentSlide].classList.add('active');
        }}
        
        function nextSlide() {{
            if (currentSlide < slides.length - 1) {{
                showSlide(currentSlide + 1);
            }}
        }}
        
        function prevSlide() {{
            if (currentSlide > 0) {{
                showSlide(currentSlide - 1);
            }}
        }}
        
        function showInfo() {{
            modal.style.display = "block";
        }}
        
        function closeInfo() {{
            modal.style.display = "none";
        }}
        
        window.onclick = function(event) {{
            if (event.target == modal) {{
                modal.style.display = "none";
            }}
        }}
        
        // Keyboard navigation
        document.addEventListener('keydown', (e) => {{
            if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === ' ') {{
                nextSlide();
            }} else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {{
                prevSlide();
            }}
        }});
    </script>
</body>
</html>
"""
        
        with open(self.output_dir / "index.html", "w", encoding="utf-8") as f:
            f.write(html_content)
        print("HTML generation complete.")

    def resolve_image_source(self, source_text):
        if not source_text or not source_text.startswith("id://"):
            return None
        res_id = source_text.replace("id://", "")
        if res_id in self.resource_map:
            return self.resource_map[res_id]
        return None

    def render_slide(self, xml_file, is_active):
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        active_class = " active" if is_active else ""
        style = ""
        
        # Slide Background
        bg_node = root.find("Background")
        if bg_node is not None:
            img_brush = bg_node.find("ImageBrush")
            if img_brush is not None:
                source = img_brush.find("Source")
                if source is not None:
                    img_path = self.resolve_image_source(source.text)
                    if img_path:
                        style = f'style="background-image: url(\'{img_path}\');"'
        
        html = f'<div class="slide{active_class}" {style}>\n'
        
        elements_node = root.find("Elements")
        if elements_node is not None:
            for elem in elements_node:
                html += self.render_element(elem)
                
        html += '</div>\n'
        return html

    def render_element(self, elem):
        # Common properties
        x = float(elem.find('X').text) if elem.find('X') is not None else 0
        y = float(elem.find('Y').text) if elem.find('Y') is not None else 0
        w = float(elem.find('Width').text) if elem.find('Width') is not None else 0
        h = float(elem.find('Height').text) if elem.find('Height') is not None else 0
        rot = float(elem.find('Rotation').text) if elem.find('Rotation') is not None else 0
        
        style = f"left: {x}px; top: {y}px; width: {w}px; height: {h}px;"
        if rot != 0:
            style += f" transform: rotate({rot}deg);"
            
        content = ""
        
        tag = elem.tag
        
        # Handle Text
        if tag == "Text" or (tag == "ActivityItem" and elem.find("Text") is not None):
            # For ActivityItem, the Text node is a child
            text_root = elem if tag == "Text" else elem.find("Text")
            rich_text = text_root.find("RichText")
            if rich_text is not None:
                content = self.render_rich_text(rich_text)
                
                # Activity Item Background
                if tag == "ActivityItem":
                    bg_node = elem.find("Background")
                    if bg_node is not None:
                        img_brush = bg_node.find("ImageBrush")
                        if img_brush is not None:
                            src_node = img_brush.find("Source")
                            if src_node is not None:
                                img_path = self.resolve_image_source(src_node.text)
                                if img_path:
                                    # Add background image to style
                                    style += f" background-image: url('{img_path}'); background-size: 100% 100%;"

        # Handle Image/Picture
        # If there is a direct Source tag in the element (like ActivityItem had), use it if no other content
        if not content:
            source_node = elem.find("Source")
            if source_node is not None:
                img_path = self.resolve_image_source(source_node.text)
                if img_path:
                    content = f'<img src="{img_path}" draggable="false">'
        
        return f'<div class="element" style="{style}">{content}</div>\n'

    def render_rich_text(self, rich_text_node):
        html = ""
        
        # Vertical Alignment
        vert_align = rich_text_node.find("VerticalTextAlignment")
        justify_content = "flex-start"
        if vert_align is not None:
            if vert_align.text == "Center":
                justify_content = "center"
            elif vert_align.text == "Bottom":
                justify_content = "flex-end"
        
        # Need to wrap inner content in a div that handles alignment
        inner_html = ""
        
        text_lines = rich_text_node.find("TextLines")
        if text_lines is not None:
            for line in text_lines.findall("TextLine"):
                line_html = '<div style="display: block; width: 100%;">' # Line container
                
                # Horizontal Alignment
                text_align = line.find("TextAlignment")
                text_align_style = "left"
                if text_align is not None:
                    if text_align.text == "Center":
                        text_align_style = "center"
                    elif text_align.text == "Right":
                        text_align_style = "right"
                
                line_html = f'<div style="text-align: {text_align_style}; line-height: 1.2;">'

                text_runs = line.find("TextRuns")
                if text_runs is not None:
                    for run in text_runs.findall("TextRun"):
                        text_content = run.find("Text").text
                        if not text_content:
                            continue
                            
                        # Style mapping
                        font_size = run.find("FontSize")
                        font_family_node = run.find("FontFamily/Source")
                        foreground_node = run.find("Foreground/ColorBrush")
                        font_weight = run.find("FontWeight")
                        
                        style = ""
                        if font_size is not None:
                            style += f"font-size: {font_size.text}px; "
                        if font_family_node is not None:
                            style += f"font-family: '{font_family_node.text}', sans-serif; "
                        if foreground_node is not None:
                            # Hex ARGB to CSS RGBA? 
                            # EasiNote uses #AARRGGBB usually. CSS wants #RRGGBB or rgba().
                            color_hex = foreground_node.text
                            if color_hex.startswith("#") and len(color_hex) == 9:
                                a = int(color_hex[1:3], 16) / 255.0
                                r = int(color_hex[3:5], 16)
                                g = int(color_hex[5:7], 16)
                                b = int(color_hex[7:9], 16)
                                style += f"color: rgba({r},{g},{b},{a}); "
                            else:
                                style += f"color: {color_hex}; "
                        if font_weight is not None and font_weight.text == "Bold":
                            style += "font-weight: bold; "
                            
                        line_html += f'<span style="{style}">{text_content}</span>'
                
                line_html += "</div>"
                inner_html += line_html
        
        return f'<div style="display: flex; flex-direction: column; justify-content: {justify_content}; height: 100%;">{inner_html}</div>'

def process_enbx(input_path, output_dir=None, show_info=False):
    input_path = Path(input_path).resolve()
    
    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        return

    # Determine extraction directory
    is_zip = False
    if input_path.suffix.lower() == '.enbx' and zipfile.is_zipfile(input_path):
        is_zip = True
        default_output_name = input_path.stem + "_html"
    else:
        # Assume it's a directory
        default_output_name = input_path.name + "_html"

    if output_dir:
        final_output_dir = Path(output_dir).resolve()
    else:
        final_output_dir = input_path.parent / default_output_name

    source_dir = input_path
    
    if is_zip:
        print(f"Unzipping {input_path}...")
        # If we extract directly to output_dir, we don't need a temp dir unless we want to separate raw from html
        # But enbx structure is root-level. 
        # Strategy: Extract to output_dir. The HTML will be generated there too.
        if not final_output_dir.exists():
            final_output_dir.mkdir(parents=True)
        
        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(final_output_dir)
        
        source_dir = final_output_dir
    
    print(f"Converting from {source_dir} to {final_output_dir}...")
    
    file_title = input_path.stem
    converter = EnbxConverter(source_dir, final_output_dir, file_title=file_title)
    converter.parse_metadata()
    
    if show_info:
        # Already printed in parse_metadata if available
        pass
        
    converter.parse_board()
    converter.parse_references()
    converter.map_slides()
    converter.copy_resources()
    converter.generate_html()
    
    print(f"Done! Output at: {final_output_dir / 'index.html'}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert ENBX (EasiNote) files to HTML5.")
    parser.add_argument("input_file", help="Path to .enbx file or extracted directory")
    parser.add_argument("-o", "--output", help="Output directory (optional)")
    parser.add_argument("--info", action="store_true", help="Show document metadata")
    
    args = parser.parse_args()
    
    process_enbx(args.input_file, args.output, args.info)
