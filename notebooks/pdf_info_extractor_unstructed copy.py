# ---
# jupyter:
#   jupytext:
#     cell_metadata_filter: -all
#     custom_cell_magics: kql
#     text_representation:
#       extension: .py
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.11.2
#   kernelspec:
#     display_name: jpfaq
#     language: python
#     name: python3
# ---

# %%
import os
from unstructured.partition.pdf import partition_pdf


# Get the current working directory
current_working_directory = os.getcwd()
# Print the current working directory
print(current_working_directory)

# %% [markdown]
# ```
# conda install conda-forge::poppler  
# conda install conda-forge::tesseract  
# ```

# %% [markdown]
# On Mac:  
# `brew install tesseract`  
# `brew install poppler-utils`

# %%
# from unstructured_inference.models.tables import cells_to_html

# %%
path_pdf = "../data/JP Label/stelara iv Japanese PI.pdf"
# path_pdf = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/data/JP Label/stelara sc Japanese PI.pdf"
filename = path_pdf.split("/")[-1].split(".pdf")[0]
dir_partitions = "../output/JP Label images partition/"
if not os.path.exists(dir_partitions):
    os.makedirs(dir_partitions)

# %%
raw_pdf_elements = partition_pdf(
    filename=path_pdf,                  # mandatory
    strategy="hi_res",                                     # mandatory to use ``hi_res`` strategy
    extract_images_in_pdf=True,                            # mandatory to set as ``True``
    extract_image_block_types=["Image", "Table"],          # optional
    extract_image_block_to_payload=False,                  # optional
    extract_image_block_output_dir=dir_partitions + filename, #.split(".")[0] + "/",  # optional - only works when 
    chunking_strategy="by_title",
    # extract_image_block_to_payload=True
    multipage_sections = "False",
    include_page_breaks = "False",
    )

# %%
print(f"In total {len(raw_pdf_elements)} elements")
ele_types = set([ele.category for ele in raw_pdf_elements])
print(f"Categories of all elements are: {ele_types}")

# %%
ele_groups = dict((key, []) for key in ele_types)

for ele in raw_pdf_elements:
    ele_groups[ele.category].append(ele)
print(f"{len(ele_groups['Table'])} tables")

# %%
txt_tables = [x.to_dict() for x in ele_groups['Table']]
txt_images = [x.to_dict() for x in ele_groups['CompositeElement']]

# %%
