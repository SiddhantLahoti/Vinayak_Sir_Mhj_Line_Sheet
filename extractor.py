import json
import os
import openpyxl

INPUT_FILE = "Lume LAB - tot wts - with IG styles.xlsx"
IMAGE_DIR = "extracted_images"
OUTPUT_JSON = "extracted_data.json"


def extract_data():
    os.makedirs(IMAGE_DIR, exist_ok=True)

    print(f"Loading input workbook: {INPUT_FILE}...")
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    ws = wb.active

    # Map images anchored near Column O (openpyxl col index 14 is Column O)
    images_by_row = {}
    for img in ws._images:
        anchor = img.anchor
        if hasattr(anchor, "_from"):
            col_idx = anchor._from.col
            row_idx = anchor._from.row  # 0-indexed row
            # Col 14 corresponds to Column O
            if col_idx == 14:
                images_by_row[row_idx] = img

    extracted_styles = []

    # Find rows where Column A has "Style No."
    for r in range(1, ws.max_row + 1):
        cell_a = ws.cell(row=r, column=1).value
        if cell_a and "Style No." in str(cell_a):
            style_no = str(ws.cell(row=r, column=3).value).strip()
            print(f"\nProcessing Style: {style_no} at row {r}")

            # Extract sketch image from Column O near this style block
            # In openpyxl, 0-indexed row (r-1) up to (r+10)
            image_path = None
            for row_offset in range(-1, 12):
                target_0_indexed_row = (r - 1) + row_offset
                if target_0_indexed_row in images_by_row:
                    matched_img = images_by_row[target_0_indexed_row]
                    image_filename = f"{style_no}.png"
                    image_path = os.path.join(IMAGE_DIR, image_filename)
                    with open(image_path, "wb") as f:
                        f.write(matched_img._data())
                    print(f"  Extracted image saved to {image_path}")
                    break

            # Read diamond details starting at (style_row + 2)
            first_diamond_row = r + 2
            curr_row = first_diamond_row
            diamond_rows = []

            while curr_row <= ws.max_row:
                qty_val = ws.cell(row=curr_row, column=8).value  # Column H is Qty
                if qty_val is None or str(qty_val).strip() == "":
                    break

                val_d = ws.cell(row=curr_row, column=4).value  # Shape
                val_e = ws.cell(row=curr_row, column=5).value  # Sieve (min-max)
                val_f = ws.cell(row=curr_row, column=6).value  # MM
                val_h = qty_val  # Qty
                val_i = ws.cell(row=curr_row, column=9).value  # Total Wt
                val_j = ws.cell(row=curr_row, column=10).value  # Setting Description

                # Parse Sieve (Column E)
                min_sieve, max_sieve = None, None
                if val_e is not None and str(val_e).strip() != "":
                    sieve_str = str(val_e).strip()
                    if "-" in sieve_str:
                        parts = sieve_str.split("-")
                        min_sieve = parts[0].strip()
                        max_sieve = parts[1].strip()
                    else:
                        min_sieve = sieve_str

                diamond_rows.append(
                    {
                        "shape": str(val_d).strip() if val_d is not None else None,
                        "min_sieve": min_sieve,
                        "max_sieve": max_sieve,
                        "mm": val_f,
                        "qty": val_h,
                        "weight": val_i,
                        "setting": (
                            str(val_j).strip() if val_j is not None else None
                        ),
                    }
                )
                curr_row += 1

            print(f"  Found {len(diamond_rows)} diamond details rows.")

            extracted_styles.append(
                {
                    "style_no": style_no,
                    "image_path": image_path,
                    "diamond_rows": diamond_rows,
                }
            )

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(extracted_styles, f, indent=4)

    print(
        f"\nExtraction complete! Saved data for {len(extracted_styles)} styles to {OUTPUT_JSON}."
    )


if __name__ == "__main__":
    extract_data()