import os

extract_dir = r"C:\Users\a\Desktop\coding\forGithub\kmd192kmd.github.io\images\extracted"
files = os.listdir(extract_dir)

# Sort by size to make it easier to view
files_with_size = []
for file in files:
    path = os.path.join(extract_dir, file)
    size = os.path.getsize(path)
    files_with_size.append((file, size))

files_with_size.sort(key=lambda x: x[1], reverse=True)

print(f"Total extracted files: {len(files_with_size)}")
for name, size in files_with_size[:20]:
    print(f"File: {name}, Size: {size} bytes")
