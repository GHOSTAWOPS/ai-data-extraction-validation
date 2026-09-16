# emerge.py  —— 带“删除目录及其之前内容”预处理的 COM 合并脚本
import os, glob, re, zipfile, time, shutil
from win32com.client import Dispatch, constants

# ====== 配置输入/输出（写死路径，便于先测试）======
input_folder = r"M:\Users\zty\Desktop\理塘"
output_file  = r"M:\Users\zty\Desktop\理塘\merged.docx"
work_temp    = r"M:\Users\zty\Desktop\理塘\~merge_working.docx"  # 临时工作文件
# ===================================

# 预处理：从你放在 docx_utils.py 里的函数调用（不修改其实现）
from docx_utils import remove_toc_and_before

DEFAULT_STYLE_IDS = {"Heading1", "Heading2", "Heading3",
                     "TOC1", "TOC2", "TOC3", "Normal"}

def natural_sort_key(s):
    basename = os.path.basename(s)
    parts = re.findall(r"\d+|\D+", basename)
    key = []
    for p in parts:
        if p.isdigit():
            key.append((0, int(p)))
        else:
            key.append((1, p.lower()))
    return key

def check_default_styles(docx_path):
    """是否含默认 Heading/TOC 样式（视为高风险文档）"""
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            if "word/styles.xml" not in zf.namelist():
                return False
            xml = zf.read("word/styles.xml").decode("utf-8", errors="ignore")
            return any(f'w:styleId="{sid}"' in xml for sid in DEFAULT_STYLE_IDS)
    except Exception:
        return False

def rebind_doc_by_fullname(word, fullname, retries=3, sleep_sec=0.2):
    """通过 FullName 获取（或重开）目标文档，尽量规避 COM 断开"""
    fullname_l = fullname.lower()
    for _ in range(retries):
        try:
            for d in word.Documents:
                if d.FullName.lower() == fullname_l:
                    return d
            # 若没在打开列表里，尝试打开
            return word.Documents.Open(fullname, ReadOnly=False, AddToRecentFiles=False)
        except Exception:
            time.sleep(sleep_sec)
    # 最后一次抛出，让上层感知
    return word.Documents.Open(fullname, ReadOnly=False, AddToRecentFiles=False)

def collapse_to_end(doc):
    """返回一个已折叠到文末的 Range"""
    rng = doc.Content
    rng.Collapse(Direction=constants.wdCollapseEnd)
    return rng

def merge_docs(folder, output, temp_path):
    # 收集原始文件
    raw_files = glob.glob(os.path.join(folder, "*.docx")) + glob.glob(os.path.join(folder, "*.doc"))
    raw_files = [f for f in raw_files if not os.path.basename(f).startswith("~$")]
    raw_files.sort(key=natural_sort_key)
    if not raw_files:
        print("⚠️ 未找到要合并的文件"); return

    # ===== 预处理：删除“最后一个 TOC 及其之前内容”，统一产出为 .docx =====
    preclean_dir = os.path.join(folder, "~toc_clean")
    os.makedirs(preclean_dir, exist_ok=True)
    cleaned_files = []
    print(f"预处理输出目录：{preclean_dir}")
    for src in raw_files:
        # 统一命名为 .docx（remove_toc_and_before 内部已兼容 .doc -> .docx）
        out_name = os.path.splitext(os.path.basename(src))[0] + ".docx"
        dst = os.path.join(preclean_dir, out_name)
        try:
            remove_toc_and_before(src, dst)
            cleaned_files.append(dst)
            print(f"🧹 清理目录完成：{os.path.basename(src)} -> {out_name}")
        except Exception as e:
            print(f"⚠️ 预处理失败，已跳过：{os.path.basename(src)} | {e}")

    if not cleaned_files:
        print("⚠️ 所有文件预处理失败，没有可合并的文档"); return

    # ====== 合并（沿用你原来的逻辑：高/低风险分流、分页符等）======
    word = Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        try:
            word.ScreenUpdating = False
            word.Options.UpdateLinksAtOpen = False
        except Exception:
            pass

        # 先建目标文档并保存为“工作文件”，后续一律通过 FullName 重绑
        target = word.Documents.Add()
        os.makedirs(os.path.dirname(temp_path), exist_ok=True)
        target.SaveAs2(temp_path)            # 保存到工作路径
        target_fullname = target.FullName    # 记录稳定路径
        print("工作文档：", target_fullname)

        # 注意：此处基于“清理后的文件”进行风险判断与合并
        for i, fp in enumerate(cleaned_files):
            risky = check_default_styles(fp)
            print(f"正在处理：{os.path.basename(fp)} → {'高风险(FormattedText)' if risky else '低风险(InsertFile)'}")

            # 每次操作前，通过 FullName 重新获取目标 doc 句柄
            target = rebind_doc_by_fullname(word, target_fullname)

            # 插入分节符（下一页）
            if i > 0:
                rng_break = collapse_to_end(target)
                rng_break.InsertBreak(constants.wdSectionBreakNextPage)

            if risky:
                # 先准备目标插入点(已折叠到文末)
                dest = collapse_to_end(target)
                # 打开源文档，取整文 Range
                src = word.Documents.Open(fp, ReadOnly=True, AddToRecentFiles=False)
                src_rng = src.Range()
                src_rng.WholeStory()
                # 关键：带格式插入
                dest.FormattedText = src_rng.FormattedText
                src.Close(SaveChanges=False)
            else:
                dest = collapse_to_end(target)
                # 直接在文末插入整份文件
                dest.InsertFile(fp, "", False, False, False)

            # 每次合并后保存一次工作文档，确保可重绑
            target.Save()
            print(f"✅ 已合并：{os.path.basename(fp)}")

        # 最终保存到输出路径（另存为）
        os.makedirs(os.path.dirname(output), exist_ok=True)
        target = rebind_doc_by_fullname(word, target_fullname)
        target.SaveAs2(output)
        target.Close(SaveChanges=False)

        print(f"\n🎉 合并完成：{output}")
    finally:
        try:
            word.ScreenUpdating = True
        except Exception:
            pass
        word.Quit()
        # 清理预处理临时目录
        try:
            shutil.rmtree(preclean_dir, ignore_errors=True)
        except Exception:
            pass

if __name__ == "__main__":
    merge_docs(input_folder, output_file, work_temp)
