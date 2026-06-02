import pandas as pd, os, json

base = r"C:\Users\Hyper_hui\PycharmProjects\PythonProject1\题库"
img_dir = os.path.join(base, "试题图片")

print(f"试题图片 dir exists: {os.path.isdir(img_dir)}")
print(f"试题图片 files: {len(os.listdir(img_dir))}")
print()

found_any = False
for f in sorted(os.listdir(base)):
    if not f.endswith('.xlsx'):
        continue
    found_any = True
    path = os.path.join(base, f)
    df = pd.read_excel(path)
    img_count = 0
    if "图片" in df.columns:
        imgs = df["图片"].dropna().astype(str).str.strip()
        imgs = imgs[imgs != ""]
        img_count = len(imgs)
    if img_count > 0:
        print(f"{f}: {img_count} image references")
        for _, v in imgs.items():
            for pair in v.split(";"):
                if "=" in pair:
                    ref, fname = pair.split("=", 1)
                    exists = os.path.exists(os.path.join(img_dir, fname))
                    print(f"  {ref} -> {fname} [{'OK' if exists else 'MISSING'}]")

if not found_any:
    print("NO xlsx files found!")
else:
    print("\nDone checking.")
