import fitz
import pandas as pd


doc = fitz.open("CBSA_01-99_2025.pdf")
lines = []
for page in doc:
    lines.extend(page.get_text().split("\n"))
doc.close()


raw_rows = []
valid_count = 0
for line in lines:
    parts = line.strip().split(" ", 1)
    if parts and parts[0].replace(".", "").isdigit():
        code = parts[0].replace(".", "")
        if 2 <= len(code) <= 10:
            valid_count += 1
            if valid_count <= 92:
                continue
            desc = parts[1].strip() if len(parts) > 1 else ""
            raw_rows.append((code, desc))

formatted = []
i = 0
while i < len(raw_rows):
    code, desc = raw_rows[i]
    code_len = len(code)

    if code_len == 2 and i + 1 < len(raw_rows):
        next_code = raw_rows[i + 1][0]
        # Only treat as Chapter if the next code starts with this
        if len(next_code) == 4 and next_code.startswith(code):
            current_chapter = code.zfill(2)
            formatted.append({
                "Chapter": current_chapter,
                "Heading": "",
                "Subheading": "",
                "Description": desc
            })
            i += 1
            continue

    if code_len == 4:
        formatted.append({
            "Chapter": "",
            "Heading": code,
            "Subheading": "",
            "Description": desc
        })
    elif code_len == 6:
        formatted.append({
            "Chapter": "",
            "Heading": "",
            "Subheading": code,
            "Description": desc
        })
    elif code_len > 6:
        formatted.append({
            "Chapter": "",
            "Heading": "",
            "Subheading": code,
            "Description": desc
        })
    i += 1


df = pd.DataFrame(formatted)
df = df[["Chapter", "Heading", "Subheading", "Description"]]
df.to_excel("CBSA_HS_Tree_CORRECT_FINAL.xlsx", index=False)
print("✅ Output ready: CBSA_HS_Tree_CORRECT_FINAL.xlsx")