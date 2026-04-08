import zipfile, os, shutil, re

def fix_months(template_path, output_path, target_m, prev_m):
    temp_dir = 'temp_pptx_fix'
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    with zipfile.ZipFile(template_path, 'r') as zip_ref: zip_ref.extractall(temp_dir)

    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    for filename in os.listdir(slides_dir):
        if not filename.endswith('.xml') or '_' in filename: continue
        path = os.path.join(slides_dir, filename)
        with open(path, 'r', encoding='utf-8') as f: content = f.read()
        orig = content
        
        if filename == 'slide3.xml':
             # The template has "January-2026" on the left, "February-2026" on the right. 
             # Wait, the template original text was:
             # Let's replace whatever month string is in the specific label.
             # It's probably easier to just replace 'January-2026' -> prev_m and 'February-2026' -> target_m if we were starting from scratch.
             # Actually, we replaced 'February-2026' with "January and followed by February". 
             # This means the graph labels (if they were separate text boxes) might have been caught in that.
             # Let's look at the original template's slide 3.
             pass

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, ds, fs in os.walk(temp_dir):
            for file in fs:
                fpath = os.path.join(root, file); zipf.write(fpath, os.path.relpath(fpath, temp_dir))
    shutil.rmtree(temp_dir)

