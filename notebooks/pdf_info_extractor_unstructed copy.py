# %%
import os
from unstructured.partition.pdf import partition_pdf


# Get the current working directory
current_working_directory = os.getcwd()
# Print the current working directory
print(current_working_directory)

# %%
from unstructured_inference.models.tables import cells_to_html

# %%
path_pdf = "data/JP Label/stelara iv Japanese PI.pdf"
# path_pdf = "/Workspace/Users/dsheng3@its.jnj.com/github/JP_FAQ/data/JP Label/stelara sc Japanese PI.pdf"
filename = path_pdf.split("/")[-1].split(".pdf")[0]
dir_partitions = "output/JP Label images partition/"
if not os.path.exists(dir_partitions):
    os.makedirs(dir_partitions)

# %%
import nltk

# %%
nltk.download('punkt')

# %%
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

# %%
print(f"In total {len(raw_pdf_elements)} elements")
print(f"Categories of all elements are: {set([ele.category for ele in raw_pdf_elements])}")

# %%
tables = []
for ele in raw_pdf_elements:
    if ele.category == "Table":
        tables.append(ele)
print(f"{len(tables)} tables")

# %%
print(tables[-1].metadata.text_as_html)

# %%
dir(tables[0].metadata)

# %%
tables[0].metadata.page_number

# %%
tables[0].metadata.fields

# %% [markdown]
# COMMAND ----------

# %%
dir(tables[0].metadata)

# %% [markdown]
# COMMAND ----------

# %%
dir(tables[0])

# %% [markdown]
# COMMAND ----------

# %%
dir(tables[0])

# %% [markdown]
# COMMAND ----------

# %%
tables[0].to_dict()

# %% [markdown]
# COMMAND ----------

# %%
print("\n\n".join([str(el) for el in raw_pdf_elements]))

# %% [markdown]
# COMMAND ----------

# %%
