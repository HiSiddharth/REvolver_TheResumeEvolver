import re
import os

def minify():
    # Resolve paths dynamically relative to the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    
    js_path = os.path.join(project_root, "bookmarklet", "bookmarklet.js")
    min_path = os.path.join(project_root, "bookmarklet", "bookmarklet.min.js")
    
    if not os.path.exists(js_path):
        print(f"Error: {js_path} does not exist.")
        return
        
    with open(js_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # 1. Remove block comments /* ... */
    content_no_block = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
    
    # 2. Remove single line comments // safely by respecting string literals
    def strip_single_line_comment(line):
        in_single = False
        in_double = False
        in_backtick = False
        i = 0
        while i < len(line):
            char = line[i]
            if char == '\\' and (in_single or in_double or in_backtick):
                i += 2
                continue
            if char == "'" and not in_double and not in_backtick:
                in_single = not in_single
            elif char == '"' and not in_single and not in_backtick:
                in_double = not in_double
            elif char == '`' and not in_single and not in_double:
                in_backtick = not in_backtick
            elif char == '/' and i + 1 < len(line) and line[i+1] == '/':
                if not in_single and not in_double and not in_backtick:
                    return line[:i]
            i += 1
        return line

    lines = content_no_block.split('\n')
    cleaned_lines = [strip_single_line_comment(line) for line in lines]
    content_no_comments = '\n'.join(cleaned_lines)
    
    # 3. Collapse multiple whitespaces/newlines (excluding template literals)
    # Split by backticks to preserve styling blocks
    parts = content_no_comments.split('`')
    for i in range(len(parts)):
        if i % 2 == 0:
            # Outside template literal: remove newlines and excess whitespace
            parts[i] = re.sub(r'\s+', ' ', parts[i])
            # Remove spaces around operators/delimiters
            parts[i] = re.sub(r'\s*([=\{\}\(\),;:><!\+\-\*/])\s*', r'\1', parts[i])
    
    minified_js = '`'.join(parts).strip()
    
    # Wrap in javascript: prefix and write
    bookmarklet_url = "javascript:" + minified_js
    
    # Ensure directory exists
    os.makedirs(os.path.dirname(min_path), exist_ok=True)
    
    with open(min_path, "w", encoding="utf-8") as f:
        f.write(bookmarklet_url)
        
    print(f"Minified bookmarklet written to {min_path}")

if __name__ == "__main__":
    minify()
