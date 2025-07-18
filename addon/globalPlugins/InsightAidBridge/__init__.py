# __init__.py - InsightAid NVDA addon

import globalPluginHandler
import tones
import os
from datetime import datetime
import api
import json
import globalVars

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """
    InsightAid NVDA addon for on-demand element and document capture.
    """

    def __init__(self):
        super(GlobalPlugin, self).__init__()
        
        # Configuration - easily customizable
        self.APP_NAME = "InsightAid"
        self.DATA_FOLDER = os.path.join(os.path.expandvars(r'%appdata%'), self.APP_NAME, "input")
        self.ELEMENT_FILE = "current_element.json"
        self.DOCUMENT_FILE = "document_content.json"
        self.VERSION_FILE = "nvda_version.json"
        
        # Ensure data folder exists
        os.makedirs(self.DATA_FOLDER, exist_ok=True)
        
        # Full file paths
        self.element_file_path = os.path.join(self.DATA_FOLDER, self.ELEMENT_FILE)
        self.document_file_path = os.path.join(self.DATA_FOLDER, self.DOCUMENT_FILE)
        self.version_file_path = os.path.join(self.DATA_FOLDER, self.VERSION_FILE)
        
        # Create/update version ping file on startup
        self._create_version_file()

    def _create_version_file(self):
        """Create/update the version ping file with NVDA and addon info."""
        try:
            # Get NVDA version
            nvda_version = getattr(globalVars, 'appVersion', 'unknown')
            
            # Create version info
            version_info = {
                "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                "nvda_version": nvda_version,
                "addon_version": "0.1.0",  # TODO: Get from manifest.ini
                "addon_name": "InsightAid Bridge"
                # TODO: Add NVDA language info
            }
            
            # Write to file
            with open(self.version_file_path, 'w', encoding='utf-8') as f:
                json.dump(version_info, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            # Fail silently - don't break addon if version file fails
            pass

    def _capture_current_element(self):
        """Capture detailed info about the current navigator object."""
        try:
            current_nav = api.getNavigatorObject()
            
            # Get object details
            name = current_nav.name or ""
            role = current_nav.role.name if current_nav.role else "(no role)"
            value = current_nav.value or ""
            
            # Try to get content for elements with no name
            content = name
            if not content or content == "(no name)":
                try:
                    if hasattr(current_nav, 'makeTextInfo'):
                        ti = current_nav.makeTextInfo('all')
                        text_content = ti.text.strip()
                        if text_content:
                            content = text_content
                    elif hasattr(current_nav, 'displayText'):
                        content = current_nav.displayText
                    elif value:
                        content = value
                except:
                    pass
            
            if not content:
                content = "(no name)"
            
            # Add value info if different from name
            if value and value != name and value != content:
                content = f"{content} (value: {value})"
            
            # Create capture data
            capture_data = {
                "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                "content": content,
                "role": role
            }
            
            # For images, add location info
            if "GRAPHIC" in role.upper() or "IMAGE" in role.upper():
                try:
                    location = current_nav.location
                    if location:
                        x, y, width, height = location
                        capture_data["location"] = {"x": x, "y": y, "width": width, "height": height}
                except:
                    capture_data["location"] = "could not retrieve"
            
            return capture_data
            
        except Exception as e:
            return {"error": f"Failed to capture element: {str(e)}"}

    # ... rest of your existing methods stay exactly the same ...

    def _get_structured_document_content(self, start_obj):
        """Traverse document and collect structured content with roles in document order."""
        structured_content = []
        max_nodes = 1000  # Prevent infinite loops
        max_depth = 20   # Don't go too deep
        node_count = 0
        
        def traverse_node(node, depth, skip_children=False):
            nonlocal node_count, structured_content
            
            if node_count >= max_nodes or depth > max_depth:
                return
            
            node_count += 1
            
            try:
                # Get node details
                name = node.name or ""
                role = node.role.name if node.role else ""
                value = node.value or ""
                
                # Skip junk roles early to avoid processing entire sections
                if role.upper() in ['BANNER', 'NAVIGATION', 'CONTENTINFO', 'COMPLEMENTARY', 'GENERIC', 'TOOLBAR', 'MENUBAR', 'MENU']:
                    return
                
                # Determine content to use
                content = name
                if not content and value:
                    content = value
                
                # Track if we captured this element completely
                captured_complete_element = False
                
                # Process current node if it has meaningful content
                if content and role:
                    # Map roles to simpler names and only include meaningful ones
                    role_mapping = {
                        'HEADING': 'HEADING',
                        'LINK': 'LINK', 
                        'BUTTON': 'BUTTON',
                        'PARAGRAPH': 'PARAGRAPH',
                        'STATICTEXT': 'TEXT',
                        'GRAPHIC': 'GRAPHIC',
                        'IMAGE': 'IMAGE',
                        'LIST': 'LIST',
                        'LISTITEM': 'LISTITEM',
                        'TABLE': 'TABLE',
                        'CELL': 'CELL',
                        'EDITABLETEXT': 'INPUT',
                        'COMBOBOX': 'DROPDOWN',
                        'CHECKBOX': 'CHECKBOX'
                    }
                    
                    if role.upper() in role_mapping:
                        role_label = role_mapping[role.upper()]
                        
                        # Add level for headings
                        if role.upper() == 'HEADING' and hasattr(node, 'level'):
                            try:
                                level = node.level
                                role_label = f"HEADING{level}"
                            except:
                                pass
                        
                        # Clean up content
                        content = content.strip()
                        if content:
                            structured_content.append(f"[{role_label}] {content}")
                            
                            # Mark elements that contain complete content - skip their text children
                            if role.upper() in ['HEADING', 'LINK', 'LISTITEM', 'BUTTON', 'PARAGRAPH']:
                                captured_complete_element = True
                
                # Recursively process children in order (unless we captured complete content)
                if not captured_complete_element and not skip_children:
                    try:
                        children = node.children
                        if children:
                            for child in children:
                                traverse_node(child, depth + 1)
                    except:
                        pass
                    
            except:
                # Skip problematic nodes
                pass
        
        # Start traversal
        traverse_node(start_obj, 0)
        return structured_content

    def _capture_full_document(self):
        """Capture structured content of the entire focused document."""
        try:
            # Get the focused document
            focus = api.getFocusObject()
            
            # Find the document object
            document_obj = focus
            while document_obj and document_obj.role.name != 'DOCUMENT':
                try:
                    document_obj = document_obj.parent
                except:
                    break
            
            if not document_obj or document_obj.role.name != 'DOCUMENT':
                return {"error": "No document found"}
            
            # Get structured content
            structured_lines = self._get_structured_document_content(document_obj)
            
            return {
                "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                "document_name": document_obj.name or "Untitled Document",
                "structured_content": structured_lines,
                "total_elements": len(structured_lines)
            }
            
        except Exception as e:
            return {"error": f"Failed to capture document: {str(e)}"}

    def script_captureElement(self, gesture):
        """NVDA+i: Capture current element and save to file."""
        tones.beep(800, 200)  # Confirmation beep
        
        capture_data = self._capture_current_element()
        
        try:
            # Write to file (overwrite each time for now)
            with open(self.element_file_path, 'w', encoding='utf-8') as f:
                json.dump(capture_data, f, indent=2, ensure_ascii=False)
            
            tones.beep(1000, 200)  # Success beep
        except Exception as e:
            tones.beep(400, 500)  # Error beep

    def script_captureFullDocument(self, gesture):
        """NVDA+shift+i: Capture full document structure and save to file."""
        tones.beep(600, 300)  # Different confirmation beep
        
        document_data = self._capture_full_document()
        
        try:
            # Write to separate file for full document
            with open(self.document_file_path, 'w', encoding='utf-8') as f:
                json.dump(document_data, f, indent=2, ensure_ascii=False)
            
            tones.beep(1200, 300)  # Higher success beep
        except Exception as e:
            tones.beep(300, 800)  # Lower error beep

    __gestures = {
        "kb:NVDA+i": "captureElement",
        "kb:NVDA+shift+i": "captureFullDocument"
    }