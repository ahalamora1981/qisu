import streamlit as st
import os
import shutil
import zipfile
from zipfile import ZipFile

from doc_process import doc_process


extract_folder = 'documents'
compressed_file_name = "output"
input_folder = None

# 压缩目录到zip文件
def compress_directory(directory_path, output_zip):
    try:
        shutil.make_archive(output_zip, 'zip', directory_path)
        return True
    except Exception as e:
        return str(e)

def recode(raw: str) -> str:
    '''
    编码修正
    '''
    
    try:
        return raw.encode('cp437').decode('utf-8')
    
    except:
        return raw.encode('utf-8').decode('utf-8')

def zip_extract_all(src_zip_file: ZipFile, target_path: str) -> None:

    # 遍历压缩包内所有内容，创建所有目录
    for file_or_path in file.namelist():
        
        # 若当前节点是文件夹
        if file_or_path.endswith('/'):
            try:
                # 基于当前文件夹节点创建多层文件夹
                os.makedirs(os.path.join(target_path, recode(file_or_path)))
            except FileExistsError:
                # 若已存在则跳过创建过程
                pass
        
        # 否则视作文件进行写出
        else:
            pass

    # 遍历压缩包内所有内容，解压文件
    for file_or_path in file.namelist():
        
        # 若当前节点是文件夹
        if file_or_path.endswith('/'):
            pass
        
        # 否则视作文件进行写出
        else:
            # 利用shutil.copyfileobj，从压缩包io流中提取目标文件内容写出到目标路径
            with open(os.path.join(target_path, recode(file_or_path)), 'wb') as z:
                # 这里基于Zipfile.open()提取文件内容时需要使用原始的乱码文件名
                shutil.copyfileobj(src_zip_file.open(file_or_path), z)

st.header('起诉书 & 委托书 - 自动处理程序')

with open('README.md', 'r') as f:
    readme = f.read()

if st.toggle('显示说明文档'):
    st.markdown(readme.split('[img]')[0])
    st.image('docs_format.jpg')
    st.markdown(readme.split('[img]')[1])

# 添加一个文件上传组件
uploaded_file = st.file_uploader("选择要上传的文件", type=["zip"])

# 如果有文件上传
if uploaded_file:
    # 保存上传的ZIP文件到本地临时目录
    with open("temp.zip", "wb") as f:
        f.write(uploaded_file.read())

    # 创建文档目录
    os.makedirs(extract_folder, exist_ok=True)

    # 解压ZIP文件中的文件并处理文件名和内容
    with zipfile.ZipFile("temp.zip", "r") as file:
        # for file_or_path in file.namelist():
        #     print(file_or_path, ' -------> ' , recode(file_or_path))
        zip_extract_all(file, extract_folder)

    # 显示解压缩完成的消息
    st.success(f"ZIP文件已成功解压缩到目录 {extract_folder}")
    input_folder = os.listdir(extract_folder)[0] if os.listdir(extract_folder)[0] != 'output' else os.listdir(extract_folder)[1]

    # 删除临时文件
    os.remove("temp.zip")

if st.button('自动处理并生成ZIP文件'):
    if input_folder:
        input_path = os.path.join(extract_folder, input_folder)
        output_path = os.path.join(extract_folder, 'output')
        result = doc_process(input_path=input_path, output_path=output_path)
    else:
        result = '请先上传ZIP文件'
    st.markdown(result)

    if input_folder:
        # 在Streamlit中压缩目录
        if compress_directory(os.path.join(extract_folder, 'output'), compressed_file_name):
            st.success("导出目录已成功压缩为ZIP文件")

            # 创建下载链接
            with open(f"{compressed_file_name}.zip", "rb") as file:
                st.download_button("点击此处下载ZIP文件", file.read(), f"{compressed_file_name}.zip")
        else:
            st.error("目录压缩失败。")

if st.button('清空输出文档', type='primary'):
    try:
        shutil.rmtree(extract_folder)
    except Exception as e:
        # st.write(f"删除文件夹 {extract_folder} 时发生错误：{str(e)}")
        pass
    
    try:
        os.remove(f"{compressed_file_name}.zip")
    except Exception as e:
        # st.write(f"删除文件 f'{compressed_file_name}.zip' 时发生错误：{str(e)}")
        pass

    if extract_folder not in os.listdir() and f"{compressed_file_name}.zip" not in os.listdir():
        st.markdown('所有输出文档已清空')

if extract_folder in os.listdir() or f"{compressed_file_name}.zip" in os.listdir():
    st.markdown(':red[完成任务后请点击“清空输出文档”]')   