# Databricks notebook source
# MAGIC %md
# MAGIC # Unstructured

# COMMAND ----------

# MAGIC %md
# MAGIC 安装库

# COMMAND ----------



! pip3 install langchain unstructured[all-docs] pydantic lxml poppler-utils unstructured[azure] PyMuPDF pdf2docx langchain_openai weasyprint pdfkit Spire.Doc pillow PyPDF2

# !cp /Workspace/Users/jfeng51@its.jnj.com/Transformed/builder.py databricks/python/lib/python3.10/site-packages/google/protobuf/internal/

#### apt package
#System Dependencies: Ensure the subsequent system dependencies are installed. Your requirements might vary based on the document types you’re handling:
# libmagic-dev : Essential for filetype detection.
# poppler-utils : Needed for images and PDFs.
# tesseract-ocr : Essential for images and PDFs.
# libreoffice : For MS Office documents.
# pandoc : For EPUBs, RTFs, and Open Office documents. Please note that to handle RTF files, you need version 2.14.2 or newer. Running this script will install the correct version for you.

# sudo apt-get install poppler-utils tesseract-ocr

# COMMAND ----------

# MAGIC %md
# MAGIC 分割Text和Table

# COMMAND ----------

from typing import Any
from pydantic import BaseModel
from unstructured.partition.pdf import partition_pdf
import google.protobuf

path = "/Workspace/Users/jfeng51@its.jnj.com/Transformed/"
file_name = "output.pdf"


"""  Text Table Image """
raw_pdf_elements = partition_pdf(
    filename= path + "PDF/" + file_name,                  # mandatory
    strategy="hi_res",                                     # mandatory to use ``hi_res`` strategy
    extract_images_in_pdf=True,                            # mandatory to set as ``True``
    extract_image_block_types=["Image", "Table"],          # optional
    extract_image_block_to_payload=False,                  # optional
    extract_image_block_output_dir=path + "images/" + file_name.split(".")[0] + "/",  # optional - only works when 
    chunking_strategy="by_title",
    # extract_image_block_to_payload=True
    multipage_sections = "False",
    include_page_breaks = "False",
    )

# category_counts = {}
# for element in raw_pdf_elements:
#     category = str(type(element))
#     if category in category_counts:
#         category_counts[category] += 1
#     else:
#         category_counts[category] = 1
# unique_categories = set(category_counts.keys())
# print(category_counts)


# path = "/Workspace/Users/jfeng51@its.jnj.com/Transformed/"
# file_name = "885383.pdf"
# raw_pdf_elements = partition_pdf(path + "PDF/" + file_name,
#                                  strategy="auto")

for ele in raw_pdf_elements:
    print(ele.category)

# COMMAND ----------

# MAGIC %md
# MAGIC PDF去除页脚

# COMMAND ----------

from PyPDF2 import PdfReader, PdfWriter
 
path = "/Workspace/Users/jfeng51@its.jnj.com/Transformed/"
file_name = "885383.pdf"


def remove_footer(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        # 获取页面的原始大小
        page_width = page.mediabox.width
        page_height = page.mediabox.height
        print(page_width)
        print(page_height)

        # 设置裁剪区域
        # left = 10
        # right = page_width - 10
        # top = 50
        # bottom = page_height - 50

        left = 0
        right = page_width
        top = 50
        bottom = page_height

        # 裁剪页面
        page.cropbox.lower_left = (left, bottom)
        page.cropbox.upper_right = (right, top)

        writer.add_page(page)

    with open(output_pdf, 'wb') as out_file:
        writer.write(out_file)

# 指定输入和输出 PDF 文件路径
input_file = '/Workspace/Users/jfeng51@its.jnj.com/Transformed/PDF/885383.pdf'
output_file = '/Workspace/Users/jfeng51@its.jnj.com/Transformed/output.pdf'

remove_footer(input_file, output_file)

# COMMAND ----------

# MAGIC %md
# MAGIC PDF转化为Image

# COMMAND ----------

from pdf2image import convert_from_path
 
# PDF文件路径
pdf_path = path + "PDF/" + "output.pdf"

image_path = "/Workspace/Users/jfeng51@its.jnj.com/Transformed/Global_Image/"
 
# 将PDF文件的每一页转换为图像
images = convert_from_path(pdf_path)
 
# 可以遍历images列表来保存或处理每一张图像
for i, image in enumerate(images):
    # 图像保存路径
    image_name = f'page_{i + 1}.png'
    # 保存图像
    image.save(image_path + image_name, 'PNG')

# COMMAND ----------

# MAGIC %md
# MAGIC Image去裁剪页脚

# COMMAND ----------

from PIL import Image

def crop_image(image_path, save_path):
    image = Image.open(image_path)
    print(image.width)
    print(image.height)
    resized_image = image.crop((0, 0, 1700, 2100))
    resized_image.save(save_path)

# COMMAND ----------

# MAGIC %md
# MAGIC 多页Image合并为一页

# COMMAND ----------

from PIL import Image

def merge_images(files, output_file):
    # 打开第一张图片
    base_img = Image.open(files[0])
    width, height = base_img.size

    # 创建一个新的空白图片，用于存储合并后的结果
    merge_img = Image.new('RGB', (width, height * len(files)), 0xffffff)

    # 依次将每张图片粘贴到新图片中
    for i, file in enumerate(files):
        img = Image.open(file)
        merge_img.paste(img, (0, i * height))

    # 保存合并后的图片
    merge_img.save(output_file)

# 要合并的图片文件列表
files = ['/Workspace/Users/jfeng51@its.jnj.com/Transformed/Global_Image/page_1.png',
         '/Workspace/Users/jfeng51@its.jnj.com/Transformed/Global_Image/page_2.png',
         '/Workspace/Users/jfeng51@its.jnj.com/Transformed/Global_Image/page_3.png']

crop_image_list = []
i = 1
for address in files:
    crop_image(address, '/Workspace/Users/jfeng51@its.jnj.com/Transformed/' + f'page_{i}.png')
    crop_image_list.append('/Workspace/Users/jfeng51@its.jnj.com/Transformed/' + f'page_{i}.png')
    i += 1


""" 合并Image """
# # 输出文件路径
output_file ='/Workspace/Users/jfeng51@its.jnj.com/Transformed/merged_image.jpg'

# # 调用函数进行图片合并
merge_images(crop_image_list, output_file)

# COMMAND ----------



# COMMAND ----------



# COMMAND ----------



# COMMAND ----------



# COMMAND ----------


