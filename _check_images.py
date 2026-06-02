import pandas as pd
import os

base = r"C:\Users\Hyper_hui\PycharmProjects\PythonProject1\题库"
files = [
    "C语言-88.xlsx",
    "C语言-80.xlsx",
    "C语言-12.xlsx",
    "C语言-79.xlsx",
    "C语言-8.xlsx",
    "C语言-81.xlsx",
]
img_dir = os.path.join(base, "试题图片")

print("=== Images in 试题图片 ===")
for fn in sorted(os.listdir(img_dir)):
    fp = os.path.join(img_dir, fn)
    print(f"  {fn} ({os.path.getsize(fp)} bytes)")

print()

for f in files:
    path = os.path.join(base, f)
    if not os.path.exists(path):
        print(f"MISSING: {f}")
        continue
    df = pd.read_excel(path)
    if "图片" in df.columns:
        imgs = df["图片"].dropna().astype(str).str.strip()
        imgs = imgs[imgs != ""]
        if len(imgs) > 0:
            print(f"=== {f} ===")
            for _, v in imgs.items():
                print(f"  {v}")
                for pair in v.split(";"):
                    if "=" in pair:
                        ref, fname = pair.split("=", 1)
                        fpath = os.path.join(img_dir, fname)
                        if os.path.exists(fpath):
                            print(f"    -> OK ({fname})")
                        else:
                            print(f"    -> MISSING ({fname})")
        else:
            print(f"{f}: no image references")
    else:
        print(f"{f}: no 图片 column")
