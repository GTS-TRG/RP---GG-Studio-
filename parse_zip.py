import zipfile
import io
import re
import xml.etree.ElementTree as ET

# We can search for the zip signature PK\x03\x04
def try_parse():
    import sys
    # Read raw stdin or string if available
    with open('/tmp/user_prompt.raw', 'rb') as f:
        data = f.read()
    
    start = data.find(b'PK\x03\x04')
    if start != -1:
        zip_data = data[start:]
        z = zipfile.ZipFile(io.BytesIO(zip_data))
        print("Files in zip:", z.namelist())
        for name in z.namelist():
            if 'sharedStrings' in name or 'sheet' in name or 'table' in name:
                print(f"=== {name} ===")
                content = z.read(name)
                print(content[:500])

if __name__ == '__main__':
    try_parse()
