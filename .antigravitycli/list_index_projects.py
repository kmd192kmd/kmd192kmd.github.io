with open(r"C:\Users\a\Desktop\coding\forGithub\kmd192kmd.github.io\index.html", 'r', encoding='utf-8') as f:
    content = f.read()

import re
matches = re.findall(r'<h3 class="text-3xl font-black[^>]*>(.*?)</h3>', content)
for i, m in enumerate(matches):
    print(f"Project {i+1}: {m}")
