import pandas as pd, os, json

base = r"C:\Users\Hyper_hui\PycharmProjects\PythonProject1\题库"
img_dir = os.path.join(base, "试题图片")

report = {}
for f in sorted(os.listdir(base)):
    if not f.endswith('.xlsx') or 'C语言' not in f:
        continue
    path = os.path.join(base, f)
    df = pd.read_excel(path)
    missing = []
    if "图片" in df.columns:
        imgs = df["图片"].dropna().astype(str).str.strip()
        imgs = imgs[imgs != ""]
        for _, v in imgs.items():
            for pair in v.split(";"):
                if "=" in pair:
                    ref, fname = pair.split("=", 1)
                    if not os.path.exists(os.path.join(img_dir, fname)):
                        missing.append(fname)
    if missing:
        report[f] = missing

print(json.dumps(report, ensure_ascii=False, indent=2))
if not report:
    print("ALL OK - no missing images")
