import fitz
import os

pdf_path = "D:/OneDrive - HSCC CHINA I E LTD/MARKETING/DIGITAL + ECOM + PHOTOS/SAHI LONDON/1.2.1 SAHI 4 WORLD/2.2 Exhibitions/MODE TORONTO/2026/SAHI_Order_1645_Lookbook_Images.pdf"
out_dir = "D:/sahi_recovery/public/images/lookbook"

doc = fitz.open(pdf_path)

image_pages = []
collections = []

for i in range(len(doc)):
    page = doc[i]
    imgs = page.get_images()
    text = page.get_text().strip()
    has_sku = "SKU" in text and "Description" in text
    
    # Collection title pages
    if "ORDER DETAIL" in text and not has_sku:
        # Extract collection name from the first line
        lines = text.split("\n")
        collection_name = lines[0].strip() if lines else f"Collection {i+1}"
        collections.append({"page": i+1, "name": collection_name})
    
    if len(imgs) >= 1 and not has_sku:
        image_pages.append(i)

print(f"Pages with product images: {len(image_pages)}")
print(f"Collections found: {len(collections)}")
for c in collections:
    print(f"  Page {c['page']}: {c['name']}")

for pg_num in image_pages:
    page = doc[pg_num]
    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat)
    out_path = os.path.join(out_dir, f"look_{pg_num+1:02d}.png")
    pix.save(out_path)
    print(f"  Saved: look_{pg_num+1:02d}.png")

doc.close()
print("Done!")
