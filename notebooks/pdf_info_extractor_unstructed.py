# Databricks notebook source
# MAGIC %pip install -r "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/requirements.txt"
# MAGIC %pip install "unstructured[all-docs]"

# COMMAND ----------

# MAGIC %md
# MAGIC In terminal:  
# MAGIC ```
# MAGIC sudo apt-get update
# MAGIC sudo apt-get install poppler-utils tesseract-ocr
# MAGIC ```

# COMMAND ----------

dbutils.library.restartPython() 

# COMMAND ----------

import os
from unstructured.partition.pdf import partition_pdf

# COMMAND ----------

from unstructured_inference.models.tables import cells_to_html

# COMMAND ----------

path_pdf = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/data/JP Label/stelara iv Japanese PI.pdf"
# path_pdf = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/data/JP Label/stelara sc Japanese PI.pdf"
filename = path_pdf.split("/")[-1].split(".pdf")[0]
dir_partitions = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/JP Label images partition/"
if not os.path.exists(dir_partitions):
    os.makedirs(dir_partitions)

# COMMAND ----------

raw_pdf_elements = partition_pdf(
    filename=path_pdf,                  # mandatory
    strategy="hi_res",                                     # mandatory to use ``hi_res`` strategy
    extract_images_in_pdf=True,                            # mandatory to set as ``True``
    extract_image_block_types=["Image", "Table"],          # optional
    extract_image_block_to_payload=False,                  # optional
    # extract_image_block_output_dir=path + "images/" + file_name.split(".")[0] + "/",  # optional - only works when 
    chunking_strategy="by_title",
    # extract_image_block_to_payload=True
    multipage_sections = "False",
    include_page_breaks = "False",
    )

# COMMAND ----------

print(f"In total {len(raw_pdf_elements)} elements")
print(f"Categories of all elements are: {set([ele.category for ele in raw_pdf_elements])}")

tables = []
for ele in raw_pdf_elements:
    if ele.category == "Table":
        tables.append(ele)
print(f"{len(tables)} tables")

# COMMAND ----------

print(tables[-1].metadata.text_as_html)

# COMMAND ----------

dir(tables[0].metadata)

# COMMAND ----------

tables[0].metadata.page_number

# COMMAND ----------

tables[0].metadata.fields

# COMMAND ----------

dir(tables[0].metadata)

# COMMAND ----------

dir(tables[0])

# COMMAND ----------

dir(tables[0])

# COMMAND ----------

tables[0].to_dict()

# COMMAND ----------

print("\n\n".join([str(el) for el in raw_pdf_elements]))

# COMMAND ----------


