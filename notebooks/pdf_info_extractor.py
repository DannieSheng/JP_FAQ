# Databricks notebook source
# %pip install PyMuPDF
# %pip install pdf2image

# dbutils.library.restartPython()

# COMMAND ----------

# MAGIC %md
# MAGIC ```
# MAGIC sudo apt-get install poppler-utils tesseract-ocr
# MAGIC
# MAGIC sudo apt update && sudo apt upgrade
# MAGIC sudo apt install tesseract-ocr
# MAGIC sudo apt install libtesseract-dev
# MAGIC
# MAGIC ```
# MAGIC

# COMMAND ----------

# dbutils.library.restartPython()

# COMMAND ----------

import os
import pymupdf
from pdf2image import convert_from_path

# COMMAND ----------

# path_pdf = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/data/JP Label/stelara iv Japanese PI.pdf"
path_pdf = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/data/JP Label/stelara sc Japanese PI.pdf"
filename = path_pdf.split("/")[-1].split(".pdf")[0]
dir_images = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/JP Label images/"
if not os.path.exists(dir_images):
    os.makedirs(dir_images)

# COMMAND ----------

texts = []
doc = pymupdf.open(path_pdf) # open a document
for page in doc: # iterate the document pages
    text = page.get_text() # get plain text encoded as UTF-8
    texts.append(text)

# COMMAND ----------

tbs = {}
bboxes = {}
n = 1
for page in doc:
    tbs[f"page {n}"] = []
    bboxes[f"page {n}"] = []
    print(f"Page {n}:")
    n_tb = 0
    for tb in page.find_tables(vertical_strategy="lines_strict", horizontal_strategy="lines_strict", snap_x_tolerance=1): 
        # print(len(tb.cells))
        n_tb += 1
        tbs[f"page {n}"].append(tb.extract())
        bboxes[f"page {n}"].append(tb.bbox) # x0, y0, x1, y1
    n += 1
    print(f"{n_tb} tables")

# COMMAND ----------

tbs['page 1'][2]

# COMMAND ----------

bboxes['page 3']

# COMMAND ----------

tb.to_markdown()

# COMMAND ----------

tb.to_pandas()

# COMMAND ----------

# i = 0
# tbs_i = []
# page = doc[i]
# for tb in page.find_tables(vertical_strategy="lines_strict", horizontal_strategy="lines_strict", snap_x_tolerance=1): #, snap_y_tolerance=10):#):
#     tbs_i.append(tb.extract())
# print(f"page {i+1}: {len(tbs_i)} tables")
# tbs_i

# COMMAND ----------

tbs['page 1'][2]

# COMMAND ----------

tbs["page 3"]

# COMMAND ----------

"\n".join(texts)

# COMMAND ----------

import re

text = "\n".join(texts)
result = {}
current_key = ''
lines = text.splitlines()
for line in lines:
    section_match = re.match(r'^(\d+\.\s+.+)', line)
    if section_match:
        # current_key = line.split(' ', 1)[1]
        current_key = section_match.group(1).strip()
        result[current_key] = []
    else:
        if current_key in result:
            result[current_key].append(line.strip())

# COMMAND ----------

result.keys()

# COMMAND ----------

result['4. 効能又は効果']

# COMMAND ----------

result['5. 効能又は効果に関連する注意']

# COMMAND ----------

result['6. 用法及び用量']

# COMMAND ----------

texts[0]

# COMMAND ----------

result['3. 組成・性状']

# COMMAND ----------

re.match(r'^(\d+\.\s+.+)', line)

# COMMAND ----------

section_match

# COMMAND ----------

section_match = re.match(r'^(\d+\.\s+.+)', line)
section_match

# COMMAND ----------

import re

text = "\n".join(texts)

sections = {}
current_section = None
current_subsection = None
current_paragraphs = []

lines = text.splitlines()
for line in lines:
    # Check for section headings like 1. 警告
    section_match = re.match(r'^(\d+\.\s+.+)', line)
    if section_match:
        # If we have existing data, store it
        if current_section and current_paragraphs:
            sections[current_section][current_subsection] = "\n".join(current_paragraphs)
            current_paragraphs = []

        current_section = section_match.group(1).strip()
        sections[current_section] = {}
        current_subsection = None

    # # Check for subsection headings like 1.1, 1.2, etc.
    # elif re.match(r'^(\d+\.\d+\s+.+)', line):
    #     if current_subsection and current_paragraphs:
    #         sections[current_section][current_subsection] = "\n".join(current_paragraphs)
    #         current_paragraphs = []

    #     current_subsection = line.strip()
    #     sections[current_section][current_subsection] = {}

    # # Check for sub-paragraphs with indented content
    # elif line.startswith("    "):
    #     if current_paragraphs:
    #         current_paragraphs.append(line.strip())
    #     else:
    #         current_paragraphs = [line.strip()]

    # Treat lines as paragraphs under current section and subsection
    else:
        if current_paragraphs:
            sections[current_section][current_subsection] = "\n".join(current_paragraphs)
            current_paragraphs = []

        current_paragraphs = [line.strip()]

# Final section cleanup
if current_section and current_paragraphs:
    sections[current_section][current_subsection] = "\n".join(current_paragraphs)

# COMMAND ----------

sections.keys()


# COMMAND ----------

sections['17. 臨床成績']

# COMMAND ----------

sections['1. 警告']

# COMMAND ----------

sections['2. 禁忌']

# COMMAND ----------

sections['3. 組成・性状']

# COMMAND ----------

sections['4. 効能又は効果']

# COMMAND ----------

sections['5. 効能又は効果に関連する注意']

# COMMAND ----------

sections['6. 用法及び用量']

# COMMAND ----------

sections['7. 用法及び用量に関連する注意']

# COMMAND ----------

sections['8. 重要な基本的注意']

# COMMAND ----------

sections.keys()

# COMMAND ----------



# COMMAND ----------

re.match(r'^\d+\.', line)

# COMMAND ----------

current_section

# COMMAND ----------

sections[current_section]

# COMMAND ----------

sections

# COMMAND ----------

print(texts[0])

# COMMAND ----------

images = convert_from_path(path_pdf)

# COMMAND ----------

# Store Pdf with convert_from_path function
images = convert_from_path(path_pdf)
im_path = os.path.join(dir_images, filename)
if not os.path.exists(im_path):
    os.makedirs(im_path)
for i in range(len(images)):
   
      # Save pages as images in the pdf
    images[i].save(os.path.join(im_path, f'page{i}.jpg'), 'JPEG')

# COMMAND ----------

im_path

# COMMAND ----------

import cv2
import pytesseract

# COMMAND ----------

im_path

# COMMAND ----------


