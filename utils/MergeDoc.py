# pip install python-docx docxcompose
import os
import re
from docx import Document
from docxcompose.composer import Composer
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .logger_util import logger
import subprocess
import traceback
from copy import deepcopy
def insert_page_break(doc):
    """在文档末尾插入分页符"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    # 添加分页符
    run._element.addnext(OxmlElement("w:br"))
    run._element.getnext().set(qn("w:type"), "page")
    

def merge_docx_in_folder(folder_path:str, output_path:str):
    '''
    folder_path:待合并的文件夹路径
    output_path:输出文件路径
    '''
    try:
        # 收集所有 docx 文件
        if folder_path[0].lower().endswith(".docx"):
            files=folder_path
        else:
            files = []
            for f in os.listdir(folder_path):
                if not f.lower().endswith(".docx"):
                    continue
                full = os.path.join(folder_path, f)
                if os.path.isfile(full):
                    files.append(full)
                else:
                    logger.warning(f"不是文件，跳过: {full}")
        # 以-前面的数字排序
        # files.sort(key=lambda x: int(re.match(r"(\d+)-", os.path.basename(x)).group(1)))
            # 排序：如果有以数字开头的，也保守处理无数字前缀的
        def sort_key(path):
            base = os.path.basename(path)
            m = re.match(r"(\d+)-", base)
            if m:
                return int(m.group(1))
            else:
                # 文件没有数字前缀，就排在后面，按名字排序
                return float('inf'), base.lower()
        files.sort(key=sort_key)

        if not files:
            logger.warning("没有找到 docx 文件")
            return

        # 第一个文档作为 master
        master = Document(files[0])
        composer = Composer(master)
        logger.info(f"已合并 {files[0]}")
        errorinfo=[]
        # 合并其余文档
        for filepath in files[1:]:
            # 在 master 中插入分页符
            insert_page_break(master)
            try:
                # 再追加新文档
                doc = Document(filepath)
                composer.append(doc)
                logger.info(f"已合并 {filepath}")
            except Exception as e:
                ## 源码修改：
                ##   File "/home/ljc/project/uv_lushan/.venv/lib/python3.10/site-packages/docxcompose/composer.py", line 213, in add_shapes
                logger.warning("合并 %s 时出错，已跳过。错误信息: %s", filepath, repr(e))
                errorinfo.append(f"{os.path.basename(filepath)[:2]}")
                # 还可以打印子文档的 rels keys / rid 列表
                # doc = Document(filepath)
                # rels = doc.part.rels
                # for rid, rel in doc.part.rels.items():
                #     print("rid:", rid, "-> target:", rel.target_ref or rel.target_part)
                # logger.debug(f"该文档的 part.rels keys: {rels.keys()}")
                logger.debug("详细堆栈：\n%s", traceback.format_exc())
                continue
        # for f in files[1:]:
        #     try:
        #         insert_page_break(master)
        #         doc = Document(f)
        #         composer.append(doc)
        #         logger.info(f"已合并 {f}")
        #     except Exception as e:
        #         logger.warning(f"composer.append 合并 {f} 出错，尝试 fallback: {e}")
        #         try:
        #             # fallback: 直接拼 body 元素
        #             sub_doc = Document(f)
        #             insert_page_break(master)
        #             for element in sub_doc.element.body:
        #                 master.element.body.append(element)
        #             logger.info(f"fallback 合并成功: {f}")
        #         except Exception as e2:
        #             logger.warning(f"fallback 也失败: {f}, 错误: {e2}")
        #             continue        
        # # 保存结果
        composer.save(output_path)
        # errorinfo=['10','11','12']
        logger.info(f"合并完成，结果保存在：{output_path}")
        return '合并文件过程中，文件'+','.join(errorinfo)+'出现错误，已跳过。'
    except Exception as e:
        logger.info(f"合并失败：{e}")
        pass



if __name__=="__main__":
    docx_files = [
        # "data/test_data/验收测试/四川省雅砻江木罗水电站/四川省雅砻江木罗水电站.docx",
        # "data/test_data/验收测试/四川省雅砻江新龙水电站/四川省雅砻江新龙水电站.docx",
        # "data/test_data/验收测试/四川省雅砻江牙根二级水电站/四川省雅砻江牙根二级水电站.docx",

        # "data/test_data/合并文档测试/01-理塘-预可--综合说明--（汇总稿）.docx",
        # "data/test_data/合并文档测试/02-理塘-预可-工程任务和建设必要性（院审稿）.docx",
        # "data/test_data/合并文档测试/03-理塘-预可-水文泥沙20250617.docx",
        # "data/test_data/合并文档测试/05-理塘-预可-工程规模（院审稿）.docx",
        # "data/test_data/合并文档测试/06-理塘-预可 建设征地和移民安置-2025.6.18.docx",
        # "data/test_data/合并文档测试/07-理塘-预可-环境保护与水土保持-送审稿.docx",
        "data/test_data/合并文档测试/08-理塘-预可-工程布置及建筑物-20250618.docx",
        "data/test_data/合并文档测试/09-理塘-预可--机电及金属结构.docx",
        "data/test_data/合并文档测试/10-理塘-预可-施工组织设计2025.6.18.docx",
        # "data/test_data/合并文档测试/12-理塘-预可-经济评价（院审稿）.docx",
        # "data/test_data/合并文档测试/merged (5).docx"
    ]
    out_dir="/home/ljc/project/Demo2.9/zbb_output/ljc_merge"
    file_name = os.path.basename(docx_files[0])
    out_path=os.path.join(out_dir,file_name)
    print(merge_docx_in_folder(docx_files,out_path))