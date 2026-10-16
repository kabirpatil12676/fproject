import os
import re

search_patterns = [
    (re.compile(r'mahdinavaei1367@gmail\.com', re.IGNORECASE), 'author@example.com'),
    (re.compile(r'linkedin\.com/in/mahdinavaei', re.IGNORECASE), 'linkedin.com/in/author'),
    (re.compile(r'github\.com/MahdiNavaei', re.IGNORECASE), 'github.com/author'),
    (re.compile(r'mahdinavaei\.github\.io', re.IGNORECASE), 'author.github.io'),
    (re.compile(r'Mahdi Navaei', re.IGNORECASE), 'Project Author'),
    (re.compile(r'MahdiNavaei', re.IGNORECASE), 'ProjectAuthor'),
    (re.compile(r'mahdinavaei', re.IGNORECASE), 'projectauthor')
]

def process_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            
        modified = False
        new_content = content
        for pattern, replacement in search_patterns:
            if pattern.search(new_content):
                new_content = pattern.sub(replacement, new_content)
                modified = True
                
        if modified:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print(f"Modified: {filepath}")
    except Exception as e:
        pass # Ignore binary files or unreadable files

def walk_directory(directory):
    exclude_dirs = {'node_modules', 'venv', '__pycache__', '.git', '.next'}
    for root, dirs, files in os.walk(directory):
        dirs[:] = [d for d in dirs if d not in exclude_dirs]
        for file in files:
            if file.endswith(('.md', '.py', '.ts', '.tsx', '.json', '.js', '.jsx', '.html', '.css', '.txt')):
                process_file(os.path.join(root, file))

if __name__ == '__main__':
    walk_directory(r'e:\pharmaceutical-supply-chain-agentic-ai')
    print("Done")
