'''
一键拆表工具 V4（tkinter 版，替代 WPS 智能工具箱会员功能）
海上钢琴师出品（不是上海钢琴师，不是上海~不是上海）

用法：选择 Excel 文件 -> 选工作表 -> 选拆分依据列 -> 可选数据起始行 -> 一键拆分
支持两种模式：
  1. 拆到当前工作簿：在原文件内按取值生成多个工作表（原表保留，自动备份）
  2. 拆成多个文件：每个取值生成一个独立 xlsx 到指定文件夹

V3 -> V4 修复清单：
  1. 【丢行】不再信任文件里的 <dimension>。V3 在只读模式下读取范围会回落到
     dimension 声明的行数，dimension 少报时超出的行被静默丢弃（预览也统计不到）。
  2. 【丢列】列数口径统一为「文件里真实存在的单元格列数」：直接扫工作表 XML 的
     <c r=..> 标签统计，界面下拉框 / 预览 / 实际拷贝三处口径完全一致，
     不再出现"界面显示 200 列、拆出来只有 150 列"。
  3. 【数据安全】拆到当前工作簿前自动备份原文件为 原名.backup.xlsx
  4. 【性能】拆成多个文件时工作簿只加载一次；拷贝只遍历需要的行，不再每个分组全表扫描
  5. 【文件污染】拆成多个文件不再把原工作簿的其它工作表带进每个输出文件
  6. 【覆盖】输出文件名去重 + 补全 Windows 非法字符 + 长度限制
  7. 【格式】条件格式的行号随数据重排，不再指向错误的行
  8. 【格式】保留全部合并单元格（V3 只保留第 1 行的）
  9. 【格式】冻结窗格按数据起始行校正
 10. 【易用】列下拉框不再被 512 列上限截断
 11. 【易用】结果消息明确告知有多少行因拆分列为空被跳过
 12. 【易用】界面提前提示 dimension 与实际数据不一致，不再"闷声少列少行"

V4 -> V4.1 修复清单：
 13. 【崩溃】拆"样式很多的大表格"时保存阶段报 IndexError: list index out of range。
     原因：V4.0 直接拷单元格的 _style 索引数组，而这些索引指向的是**源工作簿**的
     样式表；「拆成多个文件」每个输出文件都是全新的空工作簿，样式表只有默认
     1~2 项，保存时 openpyxl 用索引去查就越界了。表格越大、样式越多越容易踩到。
     修法：改为按"样式组件"逐个搬运（字体/填充/边框/对齐/保护/数字格式…），
     让 openpyxl 把组件登记进**目标工作簿**并写入目标表的索引；按源样式数组
     缓存，百万单元格也不重复转换。详见 make_style_applier()。
 14. 【排错】报错弹窗不再只显示最外面 5 帧（那会把真正的出错位置切掉，用户截图
     根本看不出原因）。现在显示"外层调用位置 + 内层出错点"的首尾摘要，并把
     完整 traceback 写到源文件旁边的「一键拆表-错误日志.txt」。

V4.1 -> V4.2 修复清单：
 15. 【卡死】大表格（台账类）预览和拆分时窗口卡死。四个独立原因，全部修掉：
     (a) pick_file() 用 openpyxl.load_workbook(path).sheetnames 读工作表名，
         341MB 的台账要 8.8 秒，全卡在主线程上 -> 改成只读 workbook.xml + rels；
     (b) _populate_cols() 在主线程上扫真实行列数，要 64 秒 -> 改成后台线程，
         并把整表正则扫描改成"逐行只看最后一个单元格"，降到 8 秒，再加缓存；
     (c) 取值统计走 openpyxl 只读迭代，每行都要解析全部 413 个单元格
         （单行 1.7ms，16 万行要 278 秒），而且纯 Python 解析长时间占住 GIL，
         界面照样不响应 -> 改成定向 XML 引擎（只解析目标列），8 秒，放后台线程；
     (d) 【最致命】拆成多个文件时 openpyxl.load_workbook() 普通模式要给 6650 万个
         单元格各建一个 Cell 对象，实测约 4.7GB，而可用内存只有 3.6GB ——
         程序疯狂换页，看起来就是"卡死" -> 拆分成多个文件改走**流式 XML 通道**
         （见 split_to_files_stream），内存占用与文件大小无关。
 16. 【拆分】「拆到当前工作簿」模式做内存体检：装不下就立刻给出明确提示并建议
     改用「拆成多个文件」，不再让你等一个永远不来的结果（见 TooBigForSheets）。
 17. 【正确性】流式通道把共享公式（<f t="shared" si=..>）就地展开成普通公式。
     源表里 master 的 ref 是连续行区间，拆行后 master 和依赖单元格会被分到
     不同文件，照搬会让 Excel 报"文件已损坏"。展开后与 openpyxl 路径完全一致。
 18. 【体验】拆出的文件不再继承源表的行隐藏状态。台账常常整表处于筛选状态
     （每一行都带 hidden="1"），拆出来的文件里筛选条件已经没了，保留 hidden
     会让用户看到一片折叠的行、以为又丢数据了。
 19. 【格式】流式通道额外保留：自动筛选范围（重排到新数据区）、数据有效性
     （重排 sqref）、超链接（重排 ref，外部链接重建 sheet rels）。
     openpyxl 路径这些是直接丢掉的。
 20. 【正确性】修掉一个坐标系搞混的 BUG：openpyxl 的 coordinate_to_tuple 返回
     (行, 列)，而 range_boundaries 返回 (列, 行, 列, 行)。V4.1 处理超链接 ref
     时按 (列, 行) 解包，导致 A3 被写成 C1。
 21. 【体积】流式通道对共享字符串表做「子集化」：**保持条目总数与索引位置完全不变**，
     只把本组没引用到的 <si> 置空成 <si/>，所以不用改写任何一个 <v> 的索引值。
     台账源表 33MB/111 万条，实测输出从 421MB 降到约 57MB，取值不受影响。

性能要点（沿袭 V2/V3，V4.2 大幅加强）：
  - 取值扫描全部走单遍流式，禁止在 read_only 表上逐行 cell() 随机访问
  - 拆分按行号映射只拷需要的行：单元格值 + 样式 + 列宽/行高/冻结窗格/合并
  - 【V4.2】大文件（> 约 300 万格）拆分走流式 XML 通道：
      解压一次工作表 -> 按 <row> 切块 -> 需要的行落 spool -> 逐组拼装输出
      样式表/共享字符串/主题从源文件原样搬运（格式零失真）
      内存占用与文件大小无关；所有耗时步骤都在后台线程，界面全程可响应

V4.2 -> V4.3 修复清单：
 22. 【致命 · 拆出来的表格在 WPS/Excel 里一片空白】
     根因：xl/_rels/workbook.xml.rels 里的 Target 必须**相对 xl/ 目录**写，
     而 V4.2 把 zip 内的完整路径直接写进去了（Target="xl/styles.xml"），
     解析出来变成 xl/xl/styles.xml —— 这个部件根本不存在，整个包按 OPC
     规范是**非法**的。
     WPS / Excel 严格走关系表，找不到样式表、字符串表、主题，于是
     所有 t="s" 的文字单元格**全部显示为空白**（数字还能显示），
     用户看到的就是"拆出来的台账表格是空白的"。
     修法：新增 _rel_target() 剥掉 xl/ 前缀；并新增 check_package_rels()
     在每写出一个文件后自检"关系表里每个 Target 都必须落到真实存在的部件上"，
     不合法立刻报错中止，绝不再把 WPS 打不开的文件交出去。
     【为什么原来测试全绿】openpyxl 读共享字符串是**直接按固定路径**
     xl/sharedStrings.xml 去找的、不查关系表，所以它照样能读出文字；
     这个 BUG 只有真正的 Office 才会暴露。新增 verify_package.py 专门盯这类问题。
 23. 【兼容】_subset_sst 生成共享字符串表时补回 XML 声明（源文件里有，
     切片时会丢掉；规范上可选，但补上更保险）。
 24. 【外观】把工具图标换成用户指定的那只边牧。
     - exe 图标资源：app.ico 含 16/20/24/32/40/48/64/128/256 共 9 档
       （资源管理器、任务栏、Alt+Tab 都清晰）；
     - **标题栏必须程序自己设**：Tk 在 Windows 上不会自动用 exe 的图标资源，
       不设就永远显示 Tk 自带的羽毛图标（看着像个蓝色笔尖）。
       新增 set_window_icon() + resource_path()，并用 --add-data 把 app.ico
       打进包，自检里也会报告 window_icon 是否生效。
     - 去白底用"四角泛洪 + 小阈值(16) + 只保留最大连通块"：
       全局去白会把边牧自己的白毛挖掉；阈值给大(40~72)会让毛边出碎白点；
       不清理连通块会留下脚下那块白色残影。

V4.3 -> V4.4 修复清单（本轮由代码审查发现，全部有回归用例钉住）：
 25. 【正确性 · 公式引用错行】共享公式平移的锚点用错了行号。
     现象：源第 5 行的共享公式依赖单元格展开后得到 =B5*2，但它在新文件里
     被排到了第 3 行 —— 引用的是**旧行号**；而输出里还保留着旧的缓存值，
     所以打开时数值看着是对的，一旦编辑触发重算就变了（"看着对、一改就错"）。
     根因：_expand_shared_formulas 只拿到旧行号，把共享公式展开成
     旧坐标系下的普通公式就直接写出了。
     修法：展开时同时传入新行号（_rewrite_row 本来就知道），master 和
     依赖单元格的公式一律按「旧行号 -> 新行号」平移后再写出。
 26. 【正确性 · 公式引用错行】普通公式（非共享）在两种模式下都不平移的问题
     一并修掉：流式通道和 openpyxl 通道现在都把公式按旧行号 -> 新行号平移
     （复用 openpyxl 的 Translator，与共享公式同一套机制），两条路径行为一致。
     另外 master 行本身被跳过（拆分列空白等）时，现在也会先学习它的公式正文
     （_learn_shared_formulas），它的依赖单元格不再因查不到 master 而丢公式。
 27. 【数据丢失 · 边缘】取值恰好是"(空白)"时会静默丢数据。
     勾选「空白行也拆出」后，空白行会以"(空白)"为名分组；如果拆分列里恰好
     有一行的真实取值就是"(空白)"，原来 groups['(空白)'] = empties 直接赋值
     会把这一组**覆盖掉**，这些行从所有输出里消失，统计上也看不出来。
     修法：改为合并进同一组（数据行在前，空白行跟在后面）。
 28. 【防呆】拆分进行中切换工作表、等拆分结束后按钮会被无条件恢复，
     此时列还没选，点下去会把整列判成空白。现在 do_split 先校验已选中列，
     _on_done 只在确实有选中列时才恢复「一键拆分」按钮。
 29. 【体验】状态栏的详细进度（"已处理 N 行"）会被"已用 X 秒"计时器每 80ms
     覆盖一次，基本看不见。现在进度消息会同步进计时器文本。
 30. 【健壮】输出文件名把「基础名 + 取值」总长一起限制（V4.3 只在重名分支
     截断，基础名很长时首个文件名仍可能超 MAX_PATH 写不出去）；
     check_package_rels 对"输出包本身读不开"不再静默放行。
 31. 【清理】删除死代码：split_to_files_stream 里从未使用的 notes、
     MainWindow 里从未读取的 _col_cache。

V4.4 -> V4.5 修复清单：
 32. 【功能 · 打不开 xlsb】原来文件对话框只认 *.xlsx / *.xlsm，.xlsb 就算选中，
     也会被当成损坏的 xlsx —— zipfile 能打开这个包，但里面根本没有
     xl/workbook.xml（xlsb 的部件是 workbook.bin / sheet1.bin /
     sharedStrings.bin），于是直接弹「无法打开文件」。
     xlsb 是 BIFF12 **二进制**格式，而本工具整套引擎都建立在「ZIP + XML」之上，
     自己解析是不可能的。
     修法：检测到 .xlsb 时，先调用本机 WPS 表格 / Microsoft Excel 的 COM 自动化
     另存为 xlsx（FileFormat=51），再走原有流程 —— 格式、公式、合并、列宽全部
     由 Office 自己搬运，保真度最高。
     - 转换结果按「绝对路径 + mtime + 大小」缓存到本机用户数据目录下的
       「一键拆表 / xlsb缓存」文件夹，同一个文件第二次拆不用再等；
       源文件一改，缓存自动失效。
     - 转换在后台线程执行，界面不卡，状态栏显示已用时间
       （180MB 的文件约需 2~4 分钟，取决于机器）。
     - 机器上既没有 WPS 也没有 Excel（或 WPS 是绿色版/精简版，没注册 COM）时，
       给出明确提示，引导用户手动另存为 xlsx 后再用本工具。
     - 新增依赖 comtypes（打包时已加入 hiddenimports）。

 33. 【数据安全 · 只拆选定的表，其他子表不动】「拆到当前工作簿」原来是 openpyxl
     直接重写源文件（虽然会先备份成 .backup.xlsx）。openpyxl 保存时会重新生成
     整个工作簿，源文件里其他工作表的图表、图片、数据透视表、切片器等部件会
     丢失 —— 用户看到的就是"我只想拆台账那一张，结果别的表被搞坏了"。
     修法：这个模式改为**源文件只读、结果另存**：
     - 源文件一个字节都不改（连备份都不再需要，make_backup 已删除）；
     - 在源文件所在目录生成「<原名>-拆分.xlsx」，内容 = 源工作簿的全部原有
       工作表（原样搬过去）+ 新拆出来的若干工作表；
     - 界面文案同步改成「拆到新工作簿（源文件不动…）」。
     注意：新模式对超大表（几千万单元格）依然走内存体检，装不下就提示改用
     「拆成多个文件」—— 那条通道是流式的，多大都不怕。

 34. 【细节】xlsb 源文件的输出命名统一用**原始**文件名（不再用临时转换出来的
     xlsx 名），错误日志也写到源文件旁边。

 35. 【致命 · 大表拆到一半失败】工作表 XML 现在**始终**按 zip64 写。
     现象：拆 16 万行 × 412 列的台账时，小分组能写出来，紧接着的大分组直接抛
     RuntimeError: File size too large, try using force_zip64，整个拆分中止
     —— 实测必踩，而且 V4.4 就有这个毛病。
     根因：_assemble_one_file 靠 est > 3.5GB 的估算来决定要不要启用 zip64，
     而 zipfile 的硬限制是 2GB（ZIP64_LIMIT）。est 是拿源表 XML 的**未压缩**
     大小按行数比例外推的，误差一大就会"以为不用 zip64、结果写出超过 2GB"。
     修法：工作表 XML 一律 force_zip64=True —— zip64 对小于 4GB 的文件没有任何
     副作用，Excel / WPS 都正常识别；同时删掉那个不可靠的 est 估算。

V4.5 -> V4.6 修复清单：
 36. 【BUG · 筛选按钮跑到数据行上去了】用户三行表头、数据从第 4 行开始，
     拆出来的文件里筛选按钮却出现在**第 4 行**（第一行数据），本该在第 3 行
     （表头最后一行）。
     根因：_prepare_suffix 里的 _af 直接用 start_row 当 autoFilter 的起始行：
         <autoFilter ref="A{start_row}:{col}{last_row}"/>
     autoFilter 的 ref 起始行就是**筛选按钮所在行**，所以按钮跟着数据行跑了。
     而且它把整个标签**重建**成自闭合标签，源表带的筛选条件
     （<filterColumn>，比如"专案二室"）和 WPS 的扩展属性
     （etc:filterBottomFollowUsedRange）全被丢掉。
     修法：起始行取 max(1, start_row - 1)（表头最后一行）；源表 ref 的起始行
     如果本来就在表头区（表头是原样复制的，行号不变），优先沿用它，
     多行表头也不会错位。只替换 ref 属性，标签里的其它内容原样保留。

 37. 【BUG · 拆出来的文件里看不到源文件的其它工作表】用户拆表一时，
     输出文件里只有拆出来的数据表，源文件里那张 Sheet1 不见了。
     根因：「拆成多个文件」通道是**手工组装 ZIP 包**的，只写了一张 sheet1.xml，
     workbook.xml / rels / content-types 也都只登记了一张表。
     修法：新增「保留源文件的其他工作表」选项（界面复选框，**默认勾选**），
     开启后每个输出文件里 = 拆出的数据表（第一张）+ 源文件里其它工作表
     （按原顺序跟在后面，部件原样搬运，含它自己的 rels）。
     三个连带处理：
       (a) workbook.xml 的每个 <sheet> 都要有对应的 r:id，
           所以 rels 里的编号由 _wb_rels_xml 统一分配并回传；
       (b) 拆出的数据固定写 xl/worksheets/sheet1.xml，若源表里那个
           sheet1.xml 恰好是要保留的表（用户拆的不是第一张表），给它换编号避让；
       (c) **共享字符串子集化必须把其它表用到的 <si> 一并保留**
           （_collect_si_refs 分块扫描收集下标）—— 否则那些表的文字会被置空，
           打开是一片空白而数字还在，症状极难排查。
     注：「拆到新工作簿」模式本来就保留全部工作表，不受这个选项影响。
'''
import copy
import hashlib
import os
import posixpath
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
import traceback

import xml.etree.ElementTree as ET
import zipfile
from array import array

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.styles.stylesheet import Stylesheet
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel, from_ISO8601
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.xml.functions import fromstring

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

BASE_FONT = ('Microsoft YaHei UI', 12)
TITLE_FONT = ('Microsoft YaHei UI', 12, 'bold')
MAX_ROW = 1048576
MAX_COL_LIMIT = 16384
VERSION = 'V4.6'


def resource_path(name):
    '''取打包后内置资源的真实路径。

    PyInstaller onefile 会把 --add-data 带进来的文件解到临时目录 sys._MEIPASS，
    源码直接运行时资源就在脚本旁边。
    '''
    base = getattr(sys, '_MEIPASS', None)
    if not base:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, name)


def set_window_icon(root):
    '''把窗口图标设成 app.ico（那只边牧）。

    【为什么必须显式设置】
    Tk 在 Windows 上**不会**自动使用 exe 里的图标资源 —— 不设的话标题栏永远是
    Tk 自带的羽毛图标（看起来像个蓝色笔尖）。exe 的图标资源只影响资源管理器里
    的文件图标和任务栏，标题栏必须靠 wm iconbitmap 单独指定。

    app.ico 里有 16/20/24/32/40/48/64/128/256 共 9 档，Windows 会按场景自己挑，
    所以标题栏（16px）、任务栏（32/48px）、Alt+Tab（大图标）都能拿到清晰的版本。
    '''
    for cand in ('app.ico', os.path.join('assets', 'app.ico')):
        p = resource_path(cand)
        if not os.path.exists(p):
            continue
        try:
            root.iconbitmap(p)
        except Exception:
            continue
        return True
    return False


def setup_fonts(root: tk.Tk):
    '''统一放大全局字体（默认 9 号在高分屏上太小）。'''
    root.option_add('*Font', BASE_FONT)
    style = ttk.Style(root)
    try:
        style.theme_use('vista')
    except tk.TclError:
        pass
    style.configure('.', font=BASE_FONT)
    style.configure('TLabel', font=BASE_FONT)
    style.configure('TButton', font=BASE_FONT, padding=(10, 6))
    style.configure('TRadiobutton', font=BASE_FONT)
    style.configure('TCheckbutton', font=BASE_FONT)
    style.configure('TCombobox', font=BASE_FONT)
    style.configure('TSpinbox', font=BASE_FONT)
    style.configure('Go.TButton', font=TITLE_FONT, padding=(10, 8))
    root.option_add('*TCombobox*Listbox.font', BASE_FONT)
    root.option_add('*TCombobox*Listbox.height', 15)


INVALID_SHEET_CHARS = set(':\\/?*[]')
INVALID_FILE_CHARS = set(':\\/?*[]<>"|')


def safe_sheet_name(name, used):
    if not str(name).strip():
        name = '(空白)'
    name = ''.join('_' if c in INVALID_SHEET_CHARS else c for c in name)[:28]
    i = 1
    base = name
    while name.lower() in used:
        i += 1
        name = f'{base[:26]}_{i}'
    used.add(name.lower())
    return name


def safe_file_stem(val: str, limit: int = 120) -> str:
    '''把取值变成合法的文件名片段（V4 新增）。'''
    s = ''.join('_' if c in INVALID_FILE_CHARS else c for c in str(val))
    s = s.strip().strip('.')
    s = re.sub('\\s+', ' ', s)
    return s[:limit] or '(空白)'


def _sheet_xml_map(path):
    '''返回 {工作表名: xl/worksheets/sheetN.xml}'''
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        wbxml = z.read('xl/workbook.xml').decode('utf-8', 'ignore')
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'ignore')
    attr_re = re.compile('([\\w:]+)\\s*=\\s*"([^"]*)"')
    rid2target = {}
    for tag in re.findall('<Relationship\\b[^>]*/?>', rels):
        a = dict(attr_re.findall(tag))
        if 'Id' in a and 'Target' in a:
            rid2target[a['Id']] = a['Target']
    out = {}
    for tag in re.findall('<sheet\\b[^>]*/?>', wbxml):
        a = dict(attr_re.findall(tag))
        nm = a.get('name')
        rid = a.get('r:id') or a.get('id')
        t = rid2target.get(rid, '')
        if t.startswith('/'):
            t = t[1:]
        elif t and not t.startswith('xl/'):
            t = 'xl/' + t
        if not nm or t not in names:
            continue
        out[_xml_unescape(nm)] = t
    return out


def _xml_unescape(s):
    return s.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')


def declared_dimension(path, sheet):
    '''读取该工作表在文件里声明的 <dimension>，返回 (max_row, max_col)；读不到返回 None。

    注意：很多导出工具会写一个与实际数据不符的 dimension，
    V3 正是因为信任它才会丢行丢列。这里只用来做"提示"，不再作为唯一依据。
    '''
    try:
        target = _sheet_xml_map(path).get(sheet)
        if not target:
            return None
        with zipfile.ZipFile(path) as z:
            with z.open(target) as f:
                head = f.read(8192).decode('utf-8', 'ignore')
        m = re.search('<dimension[^>]*\\bref="([^"]+)"', head)
        if not m:
            return None
        _c1, _r1, c2, r2 = range_boundaries(m.group(1))
        return (r2, c2)
    except Exception:
        return None


_ROW_R_ATTR_RE = re.compile(b'(?<![A-Za-z0-9_])r="(\\d+)"')
_COL_R_ATTR_RE = re.compile(b'(?<![A-Za-z0-9_])r="([A-Za-z]{1,3})')
_XML_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def _col_letters_to_num(letters):
    if isinstance(letters, (bytes, bytearray)):
        letters = letters.decode('ascii', 'ignore')
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_to_col_letters(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _cast_number_bytes(v):
    '''与 openpyxl.worksheet._reader._cast_number 等价（入参是 bytes）。'''
    if b'.' in v or b'E' in v or b'e' in v:
        return float(v)
    return int(v)


def sheet_extent(path, sheet):
    '''返回该工作表「真实存在」的 (最大行号, 最大列号)。

    V4.2：不再对全部 <c> 标签做正则匹配。单元格在同一行内按列号升序排列，
    所以「每行最后一个单元格的列号」的最大值就是整表最大列号 ——
    每行只需要一次 rfind。2.2GB 工作表从 64 秒降到 ~8 秒。
    分块流式处理，块之间只保留不完整的行，避免标签被切断。
    '''
    target = _sheet_xml_map(path).get(sheet)
    if not target:
        return (0, 0)
    max_row = 0
    max_col = 0
    with zipfile.ZipFile(path) as z:
        with z.open(target) as f:
            tail = b''
            while True:
                chunk = f.read(1048576)
                if not chunk:
                    break
                buf = tail + chunk
                consumed = 0
                pos = 0
                while True:
                    i = buf.find(b'<row', pos)
                    if i < 0:
                        consumed = len(buf)
                        break
                    gt = buf.find(b'>', i)
                    if gt < 0:
                        break
                    tag = buf[i:gt + 1]
                    m = _ROW_R_ATTR_RE.search(tag)
                    if m:
                        v = int(m.group(1))
                        if v > max_row:
                            max_row = v
                    if tag.endswith(b'/>'):
                        pos = consumed = gt + 1
                        continue
                    end = buf.find(b'</row>', gt)
                    if end < 0:
                        break
                    k = buf.rfind(b'<c ', gt, end)
                    if k >= 0:
                        cm = _COL_R_ATTR_RE.search(buf, k, end)
                        if cm:
                            c = _col_letters_to_num(cm.group(1))
                            if c > max_col:
                                max_col = c
                    pos = consumed = end + 6
                tail = buf[consumed:]
                if not tail and consumed == 0:
                    break
    if tail:
        m = _ROW_R_ATTR_RE.search(tail)
        if m:
            v = int(m.group(1))
            if v > max_row:
                max_row = v
        k = tail.rfind(b'<c ')
        if k >= 0:
            cm = _COL_R_ATTR_RE.search(tail, k)
            if cm:
                c = _col_letters_to_num(cm.group(1))
                if c > max_col:
                    max_col = c
    return (max_row, max_col)


_EXTENT_CACHE = {}
_EXTENT_CACHE_MAX = 8


def _file_stamp(path):
    '''(绝对路径, mtime_ns, 大小)，文件一变缓存就自动失效。'''
    try:
        st = os.stat(path)
        return (os.path.abspath(path), st.st_mtime_ns, st.st_size)
    except OSError:
        return (os.path.abspath(path), 0, 0)


def sheet_extent_cached(path, sheet):
    '''带缓存的 sheet_extent。

    用户每动一下「数据起始行」都会重新读表头，但真实范围没变 ——
    341MB 的台账每次重扫要 8 秒，必须缓存，否则连点几下箭头就要等半天。
    '''
    key = (_file_stamp(path), sheet)
    hit = _EXTENT_CACHE.get(key)
    if hit is None:
        hit = sheet_extent(path, sheet)
        if len(_EXTENT_CACHE) >= _EXTENT_CACHE_MAX:
            _EXTENT_CACHE.clear()
        _EXTENT_CACHE[key] = hit
    return hit


_META_CACHE = {}
_META_CACHE_MAX = 3


def _wb_meta(path):
    '''取（并缓存）工作簿元信息。'''
    try:
        st = os.stat(path)
        key = (os.path.abspath(path), st.st_mtime_ns, st.st_size)
    except OSError:
        key = (os.path.abspath(path), 0, 0)
    while True:
        hit = _META_CACHE.get(key)
        if hit is not None:
            return hit
        while len(_META_CACHE) >= _META_CACHE_MAX:
            _META_CACHE.pop(next(iter(_META_CACHE)))
        meta = {'shared': [], 'shared_loaded': False, 'date_formats': set(),
                'timedelta_formats': set(), 'epoch': WINDOWS_EPOCH}
        _META_CACHE[key] = meta
        _load_styles_meta(path, meta)
        return meta


def _load_styles_meta(path, meta):
    '''读 <workbookPr date1904> 和 styles.xml 里的数字格式（判定日期列用）。

    date_formats 直接取 openpyxl Stylesheet 算好的结果
    （Stylesheet.__init__ 里 _normalise_numbers() 已经建好 xf 下标 -> 是否日期格式）。
    注意不要用 Stylesheet.number_formats —— 那是"自定义格式表"，不是按 xf 索引的。
    '''
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            wbxml = z.read('xl/workbook.xml') if 'xl/workbook.xml' in names else b''
            sx = z.read('xl/styles.xml') if 'xl/styles.xml' in names else None
        if re.search(b'<workbookPr\\b[^>]*\\bdate1904="(?:1|true)"', wbxml):
            meta['epoch'] = MAC_EPOCH
        if sx:
            sheet_styles = Stylesheet.from_tree(fromstring(sx))
            meta['date_formats'] = set(getattr(sheet_styles, 'date_formats', ()) or ())
            meta['timedelta_formats'] = set(getattr(sheet_styles, 'timedelta_formats', ()) or ())
    except Exception:
        pass


def _local_name(tag):
    '''标签的本地名（openpyxl 的 fromstring 会剥掉命名空间，这里两种都兼容）。'''
    if not isinstance(tag, str):
        return ''
    return tag.rsplit('}', 1)[-1]


def _si_text(node):
    '''一个 <si>/<is> 节点的纯文本（优先直接子 <t>，否则拼所有 <r>/<t>）。'''
    direct = [c.text or '' for c in node if _local_name(c.tag) == 't']
    if direct:
        return ''.join(direct)
    parts = []
    for r in node:
        if _local_name(r.tag) != 'r':
            continue
        for t in r:
            if _local_name(t.tag) == 't':
                parts.append(t.text or '')
    return ''.join(parts)


def _ensure_shared_strings(path, meta):
    '''加载共享字符串表（只加载一次）。1.1M 条字符串约 4 秒、60MB。'''
    if meta['shared_loaded']:
        return meta['shared']
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    for _evt, node in ET.iterparse(f, events=('end',)):
                        if _local_name(node.tag) != 'si':
                            continue
                        text = _si_text(node)
                        out.append(text.replace('x005F_', ''))
                        node.clear()
    except Exception:
        out = []
    meta['shared'] = out
    meta['shared_loaded'] = True
    return out


def _cell_type_style(tag):
    '''从 <c ...> 标签里取 t（类型）和 s（样式下标）。'''
    t = b'n'
    s = 0
    i = tag.find(b't="')
    if i > 0:
        j = tag.find(b'"', i + 3)
        if j > 0:
            t = tag[i + 3:j]
    i = tag.find(b's="')
    if i > 0:
        j = tag.find(b'"', i + 3)
        if j > 0:
            try:
                s = int(tag[i + 3:j])
            except ValueError:
                s = 0
    return (t, s)


def _inline_text(inner):
    '''<is> 里的文本（inlineStr 类型）。'''
    try:
        node = fromstring(inner)
        return _si_text(node)
    except Exception:
        return ''


def _cell_payload(t, inner):
    '''取 <c> 的内容：返回 (原始 <v> 字节 或 None, inline 文本 或 None)。'''
    if t == b'inlineStr':
        return (None, _inline_text(inner))
    i = inner.find(b'<v>')
    if i < 0:
        i = inner.find(b'<v ')
        if i < 0:
            return (None, None)
    i = inner.find(b'>', i)
    if i < 0:
        return (None, None)
    j = inner.find(b'</v>', i)
    if j < 0:
        return (None, None)
    return (inner[i + 1:j], None)


def _formula_value(inner, coordinate, shared_formulae):
    """复刻 openpyxl.worksheet._reader.WorkSheetParser.parse_formula。

    openpyxl 默认 data_only=False，公式单元格返回的是**公式串**（如 '=SUM(B2:C2)'），
    而不是缓存值。共享公式（Excel 大量使用）还要按单元格坐标做相对引用平移 ——
    这里直接复用 openpyxl 自己的 Translator，保证结果逐字节一致。
    """
    try:
        i = inner.find(b'<f')
        if i < 0:
            return None
        gt = inner.find(b'>', i)
        if gt < 0:
            return None
        if inner[i:gt + 1].endswith(b'/>'):
            raw = inner[i:gt + 1]
        else:
            end = inner.find(b'</f>', gt)
            if end < 0:
                return None
            raw = inner[i:end + 4]
        fnode = fromstring(raw)
    except Exception:
        return None
    try:
        ftype = fnode.get('t')
        value = '=' + (fnode.text or '')
        if ftype == 'array':
            value = ArrayFormula(fnode.get('ref'), text=value)
            return value
        if ftype == 'shared':
            idx = fnode.get('si')
            if idx in shared_formulae:
                value = shared_formulae[idx].translate_formula(coordinate)
                return value
            if value != '=':
                shared_formulae[idx] = Translator(value, coordinate)
            return value
        if ftype == 'dataTable':
            value = DataTableFormula(**fnode.attrib)
        return value
    except Exception:
        return value


def _resolve_value(t, s, v, inline, meta):
    '''把原始 XML 片段转成与 openpyxl 完全一致的 Python 值。'''
    if t == b'f':
        return v
    if t == b'inlineStr':
        return inline
    if v is None:
        return None
    if t == b's':
        try:
            return meta['shared'][int(v)]
        except (ValueError, IndexError, TypeError):
            return None
    if t == b'b':
        try:
            return bool(int(v))
        except ValueError:
            return None
    if t == b'str':
        return v.decode('utf-8', 'replace')
    if t == b'd':
        try:
            return from_ISO8601(v.decode('utf-8', 'replace'))
        except Exception:
            return v.decode('utf-8', 'replace')
    try:
        num = _cast_number_bytes(v)
    except (ValueError, TypeError):
        return v.decode('utf-8', 'replace')
    try:
        if s in meta['date_formats']:
            return from_excel(num, meta['epoch'], timedelta=s in meta['timedelta_formats'])
        return num
    except (OverflowError, ValueError):
        return '#VALUE!'


def _iter_rows_column(path, sheet, col_letters):
    '''按行号顺序产出 (行号, t, s, v, inline)。

    只解析目标列：先按 </row> 切出行，再在行体内 find 目标单元格，
    同一行的其它 412 列完全不解析。
    '''
    target = _sheet_xml_map(path).get(sheet)
    if not target:
        return
    needle = b'<c r="' + col_letters
    shared_formulae = {}
    with zipfile.ZipFile(path) as z:
        with z.open(target) as f:
            tail = b''
            while True:
                chunk = f.read(1048576)
                if not chunk:
                    break
                buf = tail + chunk
                cut = buf.rfind(b'</row>')
                if cut < 0:
                    tail = buf
                    continue
                limit = cut + 6
                pos = 0
                while True:
                    i = buf.find(b'<row', pos, limit)
                    if i < 0:
                        break
                    gt = buf.find(b'>', i, limit)
                    if gt < 0:
                        break
                    tag = buf[i:gt + 1]
                    m = _ROW_R_ATTR_RE.search(tag)
                    if not m:
                        pos = gt + 1
                        continue
                    row_num = int(m.group(1))
                    if tag.endswith(b'/>'):
                        yield (row_num, b'n', 0, None, None)
                        pos = gt + 1
                        continue
                    end = buf.find(b'</row>', gt, limit)
                    if end < 0:
                        break
                    j = buf.find(needle, gt, end)
                    if j < 0:
                        yield (row_num, b'n', 0, None, None)
                        pos = end + 6
                        continue
                    k = j + len(needle)
                    e = k
                    while e < end and 48 <= buf[e] <= 57:
                        e += 1
                    if e == k:
                        # 列字母比目标长（比如扫 A 列却先遇到 AB1）-> 本行没有目标列
                        yield (row_num, b'n', 0, None, None)
                        pos = end + 6
                        continue
                    cgt = buf.find(b'>', e, end)
                    if cgt < 0:
                        yield (row_num, b'n', 0, None, None)
                        pos = end + 6
                        continue
                    ctag = buf[j:cgt + 1]
                    if ctag.endswith(b'/>'):
                        yield (row_num, b'n', 0, None, None)
                        pos = end + 6
                        continue
                    cend = buf.find(b'</c>', cgt, end)
                    if cend < 0:
                        yield (row_num, b'n', 0, None, None)
                        pos = end + 6
                        continue
                    t, s = _cell_type_style(ctag)
                    inner = buf[cgt + 1:cend]
                    if b'<f' in inner:
                        coord = col_letters.decode('ascii', 'ignore') + str(row_num)
                        yield (row_num, b'f', s, _formula_value(inner, coord, shared_formulae), None)
                        pos = end + 6
                        continue
                    v, inline = _cell_payload(t, inner)
                    yield (row_num, t, s, v, inline)
                    pos = end + 6
                tail = buf[limit:]
            if tail:
                i = tail.find(b'<row')
                if i >= 0:
                    gt = tail.find(b'>', i)
                    if gt >= 0:
                        tag = tail[i:gt + 1]
                        m = _ROW_R_ATTR_RE.search(tag)
                        if m:
                            yield (int(m.group(1)), b'n', 0, None, None)


def _iter_column_values(path, sheet, col_idx, start_row, meta):
    '''产出 (行号, 值)。缺失的行按 openpyxl 只读模式的口径补 None。

    openpyxl 的只读迭代器会把「XML 里不存在的行号」也当成空行吐出来，
    这里必须照做，否则"因拆分列为空被跳过多少行"的统计会和旧版对不上。
    '''
    letters = _num_to_col_letters(col_idx).encode('ascii')
    expected = start_row
    for row_num, t, s, v, inline in _iter_rows_column(path, sheet, letters):
        if row_num < start_row:
            continue
        while expected < row_num:
            yield (expected, None)
            expected += 1
        yield (row_num, _resolve_value(t, s, v, inline, meta))
        expected = row_num + 1


def scan_column(path, sheet, col_idx, start_row=2):
    '''单遍流式扫描某列：返回 (取值顺序, 计数, 空白行数)。'''
    meta = _wb_meta(path)
    _ensure_shared_strings(path, meta)
    counts = {}
    order = []
    empty = 0
    for _r, val in _iter_column_values(path, sheet, col_idx, start_row, meta):
        if val is None or str(val).strip() == '':
            empty += 1
            continue
        k = str(val).strip()
        if k not in counts:
            order.append(k)
        counts[k] = counts.get(k, 0) + 1
    return (order, counts, empty)


def group_rows(path, sheet, col_idx, start_row=2, keep_blank=True):
    '''按取值分组行号（只记行号，省内存）。返回 (groups, empties)。

    【V4.4 修复】空白行并入 "(空白)" 组时用 setdefault + extend：
    如果拆分列里恰好有一行的真实取值就是 "(空白)"，旧版的直接赋值
    groups['(空白)'] = empties 会把这一组数据**覆盖掉**（静默丢数据），
    现在改为合并进同一组，两种来源的行都保留。
    '''
    meta = _wb_meta(path)
    _ensure_shared_strings(path, meta)
    empties = []
    groups = {}
    for r, val in _iter_column_values(path, sheet, col_idx, start_row, meta):
        if val is None or str(val).strip() == '':
            empties.append(r)
            continue
        groups.setdefault(str(val).strip(), []).append(r)
    if keep_blank and empties:
        groups.setdefault('(空白)', []).extend(empties)
    return (groups, empties)


def read_row_values(path, sheet, row_num, max_col):
    '''读某一行（1 ~ max_col 列）的取值列表，缺失列补 None。

    只解析这一行，不走全表。表头行（通常第 1~3 行）用得上。
    '''
    if row_num < 1 or max_col < 1:
        return []
    target = _sheet_xml_map(path).get(sheet)
    if not target:
        return []
    meta = _wb_meta(path)
    _ensure_shared_strings(path, meta)
    vals = {}
    shared_formulae = {}
    with zipfile.ZipFile(path) as z:
        with z.open(target) as f:
            tail = b''
            done = False
            while not done:
                chunk = f.read(1048576)
                if not chunk:
                    break
                buf = tail + chunk
                cut = buf.rfind(b'</row>')
                if cut < 0:
                    tail = buf
                    continue
                limit = cut + 6
                pos = 0
                while True:
                    i = buf.find(b'<row', pos, limit)
                    if i < 0:
                        break
                    gt = buf.find(b'>', i, limit)
                    if gt < 0:
                        break
                    tag = buf[i:gt + 1]
                    m = _ROW_R_ATTR_RE.search(tag)
                    if not m:
                        pos = gt + 1
                        continue
                    if int(m.group(1)) > row_num:
                        done = True
                        break
                    elif int(m.group(1)) != row_num or tag.endswith(b'/>'):
                        pos = gt + 1
                        continue
                    end = buf.find(b'</row>', gt, limit)
                    if end < 0:
                        break
                    p = gt + 1
                    while True:
                        j = buf.find(b'<c ', p, end)
                        if j < 0:
                            break
                        cgt = buf.find(b'>', j, end)
                        if cgt < 0:
                            break
                        ctag = buf[j:cgt + 1]
                        cm = _COL_R_ATTR_RE.search(buf, j, cgt)
                        if cm:
                            col = _col_letters_to_num(cm.group(1))
                            if 1 <= col <= max_col:
                                if ctag.endswith(b'/>'):
                                    vals[col] = None
                                else:
                                    cend = buf.find(b'</c>', cgt, end)
                                    if cend >= 0:
                                        t, s = _cell_type_style(ctag)
                                        inner = buf[cgt + 1:cend]
                                        if b'<f' in inner:
                                            coord = _num_to_col_letters(col) + str(row_num)
                                            vals[col] = _formula_value(inner, coord, shared_formulae)
                                        else:
                                            v, inline = _cell_payload(t, inner)
                                            vals[col] = _resolve_value(t, s, v, inline, meta)
                        p = cgt + 1
                    done = True
                    break
                tail = buf[limit:]
    return [vals.get(c) for c in range(1, max_col + 1)]


def probe_sheet(path, sheet, header_row):
    '''返回 (真实最大列数, 表头行取值列表, 真实最大行数)。

    列数口径 = 文件里真实存在的单元格列数（由 sheet_extent 从 XML 统计），
    与拆分时拷贝的范围完全一致，不会再出现"界面显示 200 列、拆出来只有 150 列"。
    表头只解析这一行，不做全表遍历。
    '''
    max_row, max_col = sheet_extent_cached(path, sheet)
    header = []
    if header_row >= 1 and max_col >= 1:
        header = read_row_values(path, sheet, header_row, max_col)
    return (max_col, header, max_row)


def build_row_map(rows, start_row):
    '''表头区（1 ~ start_row-1）原位保留，数据行依次压缩排列。
    映射方向：旧行号 -> 新行号。
    '''
    m = {r: r for r in range(1, start_row)}
    for new_r, old_r in enumerate(sorted(rows), start=start_row):
        m[old_r] = new_r
    return m


def _remap_sqref(rng, row_map):
    '''把条件格式的范围按行号映射重排（V4 新增）。'''
    lo = get_column_letter(rng.min_col)
    hi = get_column_letter(rng.max_col)
    parts = []
    for r in range(rng.min_row, rng.max_row + 1):
        nr = row_map.get(r)
        if nr is None:
            continue
        parts.append(f'{lo}{nr}:{hi}{nr}')
    return ' '.join(parts)


_STYLE_ATTRS = ('font', 'fill', 'border', 'alignment', 'protection', 'number_format', 'quotePrefix', 'pivotButton')


def make_style_applier():
    '''返回 apply(src_cell, dst_cell)，把源单元格样式搬到**目标工作簿**。

    【为什么不能直接 nc._style = copy.copy(sc._style)】

    StyleArray 里存的是一串**索引**（fontId / fillId / borderId / numFmtId /
    protectionId / alignmentId / xfId ...），这些索引指向的是**源工作簿**的
    _fonts / _fills / _borders / _alignments / _protections / _number_formats 表。

    而「拆成多个文件」的每个输出文件用的是全新的 openpyxl.Workbook()，
    它这些表都只有默认的 1~2 项。源表里任何带对齐或保护的单元格，
    保存时就会在这里越界：

        openpyxl/styles/stylesheet.py :: write_stylesheet()
            if style.alignmentId:
                xf.alignment = wb._alignments[style.alignmentId]      # IndexError!
            if style.protectionId:
                xf.protection = wb._protections[style.protectionId]   # IndexError!

    表现为「拆大表格时报 IndexError: list index out of range」——
    表格越大、样式越多，越容易踩到。

    【正确做法】

    按"样式组件"逐个赋值。openpyxl 的 StyleDescriptor.__set__ 会把组件登记进
    目标工作簿的样式表（相同组件自动去重），并写入属于**目标表**的索引：

        nc.font = copy(src_cell.font)      # 注意要 copy：
                                           # 取到的是只读 StyleProxy，
                                           # copy() 之后才是可入库的真实对象

    为了让上百万个单元格不重复做这套转换，按「源样式数组」缓存第一次算出来的
    目标样式数组，后续单元格直接套用。
    '''
    cache = {}

    def apply(src_cell, dst_cell):
        st = src_cell._style
        if st is None:
            return
        key = st.tobytes()
        new_style = cache.get(key)
        if new_style is None:
            for attr in _STYLE_ATTRS:
                setattr(dst_cell, attr, copy.copy(getattr(src_cell, attr)))
            new_style = copy.copy(dst_cell._style)
            cache[key] = new_style
        else:
            dst_cell._style = copy.copy(new_style)
    return apply


def copy_sheet_rows(src_ws, dst_ws, row_map, max_col, start_row):
    '''按行号映射拷贝需要的行（值 + 样式 + 超链接 + 列宽/行高/冻结/合并/条件格式）。

    V4 改动：
      - 只遍历 row_map 里真正需要的行，不再每拷一个分组就把整张表扫一遍
      - max_col 由调用方传入，与界面显示的口径完全一致
      - 合并单元格全部按行号映射保留（V3 只保留第 1 行）
      - 条件格式 sqref 随行号重排
      - 冻结窗格按 start_row 校正

    V4.1 改动：
      - 样式改为逐组件搬运（见 make_style_applier）。原先直接拷 _style 索引数组，
        遇到带对齐/保护的表格会在保存时抛 IndexError

    V4.4 改动：
      - 普通公式（字符串值、以 = 开头）按「旧行号 -> 新行号」平移，
        与流式通道行为一致。V4.3 直接照抄公式串，行号压缩后引用会指错行。
    '''
    if max_col < 1:
        max_col = 1
    cells = src_ws._cells
    apply_style = make_style_applier()
    touched = set()
    for old_r, new_r in sorted(row_map.items()):
        row_has = False
        for col in range(1, max_col + 1):
            sc = cells.get((old_r, col))
            if sc is None:
                continue
            val = sc.value
            styled = sc.has_style
            link = sc.hyperlink
            if val is None and not styled and link is None:
                continue
            if old_r != new_r and isinstance(val, str) and val.startswith('='):
                try:
                    val = Translator(val, get_column_letter(col) + str(old_r)).translate_formula(
                        get_column_letter(col) + str(new_r))
                except Exception:
                    pass
            nc = dst_ws.cell(row=new_r, column=col, value=val)
            if styled:
                apply_style(sc, nc)
            if link is not None:
                nc.hyperlink = copy.copy(link)
            row_has = True
        if not row_has:
            continue
        touched.add(new_r)
    for nr in row_map.values():
        if nr not in touched:
            dst_ws.cell(row=nr, column=1)
    for letter, dim in src_ws.column_dimensions.items():
        try:
            d = dst_ws.column_dimensions[letter]
            d.width = dim.width
            d.hidden = dim.hidden
            d.bestFit = dim.bestFit
        except Exception:
            continue
    for old_r, new_r in row_map.items():
        if old_r not in src_ws.row_dimensions:
            continue
        sd = src_ws.row_dimensions[old_r]
        dd = dst_ws.row_dimensions[new_r]
        dd.height = sd.height
        dd.hidden = sd.hidden
    if src_ws.freeze_panes:
        try:
            frow, fcol = coordinate_to_tuple(str(src_ws.freeze_panes))
            frow = max(1, min(frow, max(1, start_row)))
            fcol = max(1, min(fcol, max_col))
            if frow > 1 or fcol > 1:
                dst_ws.freeze_panes = dst_ws.cell(row=frow, column=fcol).coordinate
        except Exception:
            pass
    for rng in list(src_ws.merged_cells.ranges):
        try:
            src_rows = list(range(rng.min_row, rng.max_row + 1))
            if not all(r in row_map for r in src_rows):
                continue
            new_rows = [row_map[r] for r in src_rows]
            if new_rows != list(range(new_rows[0], new_rows[0] + len(new_rows))):
                continue
            lo = get_column_letter(rng.min_col)
            hi = get_column_letter(rng.max_col)
            dst_ws.merge_cells(f'{lo}{new_rows[0]}:{hi}{new_rows[-1]}')
        except Exception:
            continue
    for cf in src_ws.conditional_formatting:
        try:
            for rng in cf.sqref.ranges:
                new_ref = _remap_sqref(rng, row_map)
                if not new_ref:
                    continue
                for rule in cf.rules:
                    dst_ws.conditional_formatting.add(new_ref, copy.copy(rule))
        except Exception:
            continue


# ---------------- xlsb 支持（V4.5 新增） ----------------
#
# xlsb 是 BIFF12 二进制格式：包里的部件是 workbook.bin / sheet1.bin /
# sharedStrings.bin，本工具整套引擎（ZIP + XML）没法解析它。
# 唯一保真的办法是让 Office 自己把它另存为 xlsx，再走原有流程。

XLSB_EXT = '.xlsb'


def is_xlsb_path(path):
    '''是不是 .xlsb（二进制工作簿）。'''
    return os.path.splitext(str(path))[1].lower() == XLSB_EXT


def looks_like_xlsb(path):
    '''看包里的部件名判断是不是 xlsb。

    有人会把 .xlsb 改名成 .xlsx 发出来，光看扩展名会漏判；xlsb 的部件叫
    xl/workbook.bin，xlsx 的部件叫 xl/workbook.xml，一眼能分辨。
    （只读 ZIP 中央目录，几百 MB 的文件也是毫秒级。）
    '''
    try:
        with zipfile.ZipFile(path) as z:
            return 'xl/workbook.bin' in set(z.namelist())
    except Exception:
        return False


def xlsb_cache_dir():
    '''xlsb 转换结果的缓存目录（放本机用户数据目录，不污染源文件所在文件夹）。'''
    base = os.environ.get('LOCALAPPDATA') or os.environ.get('TEMP') or tempfile.gettempdir()
    d = os.path.join(base, '一键拆表', 'xlsb缓存')
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        d = tempfile.gettempdir()
    return d


def xlsb_cache_path(src):
    '''转换结果的缓存路径。

    键 = 绝对路径 + mtime + 大小 —— 源文件一改，缓存自动失效，不会拿旧数据糊弄人。
    '''
    st = os.stat(src)
    key = '%s|%d|%d' % (os.path.abspath(src).lower(), st.st_mtime_ns, st.st_size)
    tag = hashlib.md5(key.encode('utf-8')).hexdigest()[:16]
    stem = safe_file_stem(os.path.splitext(os.path.basename(src))[0], 48)
    return os.path.join(xlsb_cache_dir(), '%s_%s.xlsx' % (stem, tag))


def _new_com_app():
    '''创建 WPS 表格 / Microsoft Excel 的 COM 应用对象（优先 WPS，再退回 Excel）。

    本机只装 WPS 时，KET.Application 是 WPS 表格自己的 ProgID；
    Excel.Application 常常也被 WPS 注册成别名，所以两个都试。
    '''
    try:
        import comtypes.client as cc
    except Exception as e:
        raise RuntimeError('本机缺少 comtypes 组件，无法自动转换 xlsb（%s）。' % e)
    errs = []
    for prog in ('KET.Application', 'Excel.Application'):
        try:
            app = cc.CreateObject(prog)
        except Exception as e:
            errs.append('%s：%s' % (prog, e))
            continue
        for attr, val in (('Visible', False), ('DisplayAlerts', False)):
            try:
                setattr(app, attr, val)
            except Exception:
                pass
        return app
    raise RuntimeError('本机没有可用的 WPS 表格 / Microsoft Excel：\n' + '\n'.join(errs))


def convert_xlsb_to_xlsx(src, dst, on_progress=None):
    '''用本机 Office（WPS / Excel）把 .xlsb 另存为 .xlsx，返回 dst 路径。

    转换动作是 Office 自己做的「另存一份」，所以格式、公式、合并单元格、
    列宽行高、条件格式、数据验证全部原样保留 —— 比任何纯 Python 方案都保真。
    '''
    if on_progress:
        on_progress('正在启动 WPS / Excel…')
    app = _new_com_app()
    wb = None
    tmp = dst + '.tmp.xlsx'
    try:
        if on_progress:
            on_progress('正在打开 xlsb（大文件需要几分钟，请勿关闭 WPS 窗口）…')
        wb = app.Workbooks.Open(os.path.abspath(src), UpdateLinks=0, ReadOnly=True)
        for p in (tmp, dst):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass
        if on_progress:
            on_progress('正在另存为 xlsx…')
        wb.SaveAs(tmp, FileFormat=51)   # 51 = xlOpenXMLWorkbook
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
    if not os.path.exists(tmp):
        raise RuntimeError('转换没有产出文件，可能被 Office 的弹窗拦住了。')
    os.replace(tmp, dst)
    return dst


def unique_workbook_path(src):
    '''「拆到新工作簿」的输出路径：源文件所在目录下的 <原名>-拆分.xlsx（重名加序号）。'''
    d = os.path.dirname(os.path.abspath(src))
    base = os.path.splitext(os.path.basename(src))[0]
    path = os.path.join(d, '%s-拆分.xlsx' % base)
    i = 1
    while os.path.exists(path):
        i += 1
        path = os.path.join(d, '%s-拆分_%d.xlsx' % (base, i))
    return path


def unique_output_path(out_dir, base, val, used):
    '''生成不冲突的输出文件名（V4 新增去重）。

    【V4.4】首个名字也限制「基础名 + 取值」总长。V4.3 只在重名分支截断取值，
    基础名很长时首个文件名仍可能把完整路径顶过 MAX_PATH，写出直接失败。
    '''
    stem = safe_file_stem(val)
    budget = max(10, 150 - len(base))
    if len(stem) > budget:
        stem = stem[:budget]
    name = f'{base}-{stem}.xlsx'
    i = 1
    while name.lower() in used:
        i += 1
        name = f'{base}-{stem[:110]}_{i}.xlsx'
    used.add(name.lower())
    return os.path.join(out_dir, name)


def format_error(exc_text, keep_head=4, keep_tail=18):
    '''把完整 traceback 裁成「外层调用位置 + 内层真正出错点」。

    V4.0 用的是 traceback.format_exc(limit=5)，只保留最外面 5 帧，
    真正出错的帧被切掉了 —— 用户截图里的报错完全看不出原因。
    '''
    lines = [ln for ln in exc_text.rstrip().splitlines() if ln.strip()]
    if len(lines) <= keep_head + keep_tail + 1:
        return '\n'.join(lines)
    dropped = len(lines) - keep_head - keep_tail
    return '\n'.join(lines[:keep_head] + ['    ...（中间省略 %d 行）...' % dropped] + lines[-keep_tail:])


def write_error_log(src_path, exc_text):
    '''把完整 traceback 写到源文件旁边，方便用户发回来定位。返回路径或 None。'''
    try:
        d = os.path.dirname(os.path.abspath(src_path)) if src_path else os.getcwd()
        p = os.path.join(d, '一键拆表-错误日志.txt')
        with open(p, 'w', encoding='utf-8') as f:
            f.write('时间：%s\n' % time.strftime('%Y-%m-%d %H:%M:%S'))
            f.write('版本：%s\n' % VERSION)
            f.write('源文件：%s\n\n' % src_path)
            f.write(exc_text)
        return p
    except Exception:
        return None


_XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
_PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
_OD_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

_ROW_START_RE = re.compile(b'<row(?=[\\s/>])')
_F_TAG_RE = re.compile(b'<f\\b([^>]*?)(/?)>')
_ATTR_RE = re.compile(b'([\\w:]+)\\s*=\\s*"([^"]*)"')
_CELL_REF_SUB_RE = re.compile(b'r="([A-Za-z]{0,3})\\d+"')
_ROW_HIDDEN_RE = re.compile(b'\\s+hidden="[^"]*"')
_DIMENSION_RE = re.compile(b'<dimension\\b[^>]*/?>')
_FILTERMODE_RE = re.compile(b'(<sheetPr\\b[^>]*?)\\s+filterMode="[^"]*"')
_PANE_RE = re.compile(b'<pane\\b[^>]*?/>|<pane\\b[^>]*?>.*?</pane>', re.S)
_SELECTION_RE = re.compile(b'<selection\\b[^>]*?/>|<selection\\b[^>]*?>.*?</selection>', re.S)
_AUTOFILTER_RE = re.compile(b'<autoFilter\\b[^>]*?/>|<autoFilter\\b[^>]*?>.*?</autoFilter>', re.S)
_MERGECELLS_RE = re.compile(b'<mergeCells\\b[^>]*?>.*?</mergeCells>|<mergeCells\\b[^>]*?/>', re.S)
_MERGECELL_RE = re.compile(b'<mergeCell\\b[^>]*?/>')
_CF_SQREF_RE = re.compile(b'(<conditionalFormatting\\b[^>]*?\\bsqref=")([^"]*)(")')
_DV_SQREF_RE = re.compile(b'(<dataValidation\\b[^>]*?\\bsqref=")([^"]*)(")')
_HYPERLINKS_RE = re.compile(b'<hyperlinks\\b[^>]*?>.*?</hyperlinks>', re.S)
_HYPERLINK_RE = re.compile(b'<hyperlink\\b[^>]*?/>')
_SST_SPLIT_RE = re.compile(b'(?s)(<sst\\b[^>]*>)(.*)</sst>\\s*$')
_SST_SI_RE = re.compile(b'<si\\b.*?</si>', re.S)
_SI_REF_RE = re.compile(b't="s"[^>]*><v>(\\d+)')
_REL_TARGET_RE = re.compile(b'Target="([^"]*)"')
_SST_EMPTY = b'<si/>'
_SST_SUBSET_LIMIT = 60000000

_DROP_ELEMS = (b'drawing', b'legacyDrawing', b'legacyDrawingHF', b'picture', b'oleObjects', b'controls', b'tableParts', b'rowBreaks', b'colBreaks', b'extLst', b'ignoredErrors', b'smartTags', b'webPublishItems')
_DROP_RE = re.compile(b'<(?:drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|tableParts|rowBreaks|colBreaks|extLst|ignoredErrors|smartTags|webPublishItems)\\b(?:[^>]*?/>|[^>]*?>.*?</(?:drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|tableParts|rowBreaks|colBreaks|extLst|ignoredErrors|smartTags|webPublishItems)>)', re.S)

_CONTENT_TYPES_HEAD = (_XML_DECL + ('<Types xmlns="%s">' % _CT_NS).encode()
                       + b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                       + b'<Default Extension="xml" ContentType="application/xml"/>'
                       + b'<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                       + b'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                       + b'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                       + b'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')

_CT_OPTIONAL = ((b'styles', b'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'),
                (b'sharedStrings', b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'),
                (b'theme', b'<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'))

_ROOT_RELS_XML = (_XML_DECL + ('<Relationships xmlns="%s">' % _PKG_REL_NS).encode()
                  + ('<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>' % _OD_REL_NS).encode()
                  + ('<Relationship Id="rId2" Type="%s/extended-properties" Target="docProps/app.xml"/>' % _OD_REL_NS).encode()
                  + ('<Relationship Id="rId3" Type="%s/metadata/core-properties" Target="docProps/core.xml"/>' % _PKG_REL_NS).encode()
                  + b'</Relationships>')

_COPYABLE_PARTS = ((b'styles', 'xl/styles.xml', 'styles'),
                   (b'sharedStrings', 'xl/sharedStrings.xml', 'sharedStrings'),
                   (b'theme', 'xl/theme/theme1.xml', 'theme'))


_CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'


def _content_types_xml(present, extra_entries=(), extra_ct=b''):
    '''[Content_Types].xml。

    【V4.6】extra_entries 是随拆出文件一起保留的**其它工作表**的部件路径，
    每一个都要在这里声明，否则 WPS / Excel 会认为包里有个没登记的表。
    extra_ct 是从源包搬过来的声明（图片、绘图等依赖部件用），
    调用方需保证里面不含 extra_entries 自己的 Override（重复声明是非法包）。
    '''
    out = [_CONTENT_TYPES_HEAD]
    for key, frag in _CT_OPTIONAL:
        if key in present:
            out.append(frag)
    for entry in extra_entries:
        out.append(('<Override PartName="/%s" ContentType="%s"/>'
                    % (entry, _CT_WORKSHEET)).encode())
    if extra_ct:
        out.append(extra_ct)
    out.append(b'</Types>')
    return b''.join(out)


def _rel_target(zip_path):
    '''把 zip 内的部件路径转成「相对 xl/ 的」关系目标。

    【这是 V4.2 的一个致命 BUG，务必别改回去】
    xl/_rels/workbook.xml.rels 这个文件本身位于 xl/ 目录下，按 OPC 规范，
    它的 Target 必须**相对 xl/ 目录**来写：
        styles.xml / sharedStrings.xml / theme/theme1.xml
    V4.2 之前直接把 zip 内的完整路径写进去了（xl/styles.xml），解析出来就变成
    xl/xl/styles.xml —— 这个部件根本不存在，整个包是**非法**的。

    为什么测试没发现：openpyxl 读共享字符串是**直接按固定路径**
    xl/sharedStrings.xml 去找的，不查关系表，所以它照样能读出文字；
    而 WPS / Excel 老老实实走关系表 → 找不到字符串表和样式表 →
    所有 t="s" 的文字单元格**全部显示为空白**（数字还能显示），
    用户看到的就是"拆出来的表格是空白的"。
    '''
    if zip_path.startswith('xl/'):
        return zip_path[3:]
    return zip_path


def _wb_rels_xml(present, extra_entries=()):
    '''xl/_rels/workbook.xml.rels。返回 (xml, extra_rids)。

    【V4.6】extra_rids 与 extra_entries 一一对应 —— workbook.xml 里每个
    <sheet> 的 r:id 必须和这里的 Id 对上，所以由本函数统一分配编号。
    '''
    out = [_XML_DECL + ('<Relationships xmlns="%s">' % _PKG_REL_NS).encode(),
           ('<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>' % _OD_REL_NS).encode()]
    rid = 1
    for key, target, typ in _COPYABLE_PARTS:
        if key not in present:
            continue
        rid += 1
        out.append(('<Relationship Id="rId%d" Type="%s/%s" Target="%s"/>' % (rid, _OD_REL_NS, typ, _rel_target(target))).encode())
    extra_rids = []
    for entry in extra_entries:
        rid += 1
        extra_rids.append('rId%d' % rid)
        out.append(('<Relationship Id="rId%d" Type="%s/worksheet" Target="%s"/>'
                    % (rid, _OD_REL_NS, _rel_target(entry))).encode())
    out.append(b'</Relationships>')
    return b''.join(out), extra_rids


def _xml_escape(s) -> str:
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def _workbook_xml(sheet_name, date1904, extra_sheets=()):
    '''workbook.xml。extra_sheets 是 [(工作表名, rId), ...]。

    【V4.6】拆出来的表排第一（打开就能看到数据），源文件里保留的其它工作表
    按原顺序跟在后面。
    '''
    pr = b'<workbookPr date1904="1"/>' if date1904 else b'<workbookPr/>'
    sheets = [b'<sheet name="' + _xml_escape(sheet_name).encode('utf-8')
              + b'" sheetId="1" r:id="rId1"/>']
    for i, (nm, rid) in enumerate(extra_sheets, start=2):
        sheets.append(b'<sheet name="' + _xml_escape(nm).encode('utf-8')
                      + ('" sheetId="%d" r:id="%s"/>' % (i, rid)).encode())
    return (_XML_DECL
            + ('<workbook xmlns="%s" xmlns:r="%s">' % (_MAIN_NS, _OD_REL_NS)).encode()
            + b'<fileVersion appName="xl" lastEdited="3" lowestEdited="5" rupBuild="9302"/>'
            + pr
            + b'<bookViews><workbookView/></bookViews><sheets>'
            + b''.join(sheets)
            + b'</sheets>'
            + b'<calcPr calcId="191029"/></workbook>')


def _core_xml():
    now = time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())
    return (_XML_DECL
            + b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            + b'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            + b'xmlns:dcterms="http://purl.org/dc/terms/" '
            + b'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            + b'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            + b'<dc:creator>\xe4\xb8\x80\xe9\x94\xae\xe6\x8b\x86\xe8\xa1\xa8</dc:creator>'
            + b'<dcterms:created xsi:type="dcterms:W3CDTF">'
            + now.encode()
            + b'</dcterms:created></cp:coreProperties>')


def _app_xml(sheet_name, extra_names=()):
    '''docProps/app.xml。

    【V4.6】工作表数量与名称必须和实际一致 —— 文档属性里写着"1 个工作表"
    而包里有两张，虽然 Excel 一般容错，但有些查看器/工具会据此显示错乱。
    '''
    names = [sheet_name] + [str(n) for n in extra_names]
    parts = b''.join(b'<vt:lpstr>' + _xml_escape(n).encode('utf-8') + b'</vt:lpstr>'
                     for n in names)
    return (_XML_DECL
            + b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            + b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            + b'<Application>\xe4\xb8\x80\xe9\x94\xae\xe6\x8b\x86\xe8\xa1\xa8</Application>'
            + b'<HeadingPairs><vt:vector size="2" baseType="variant">'
            + b'<vt:variant><vt:lpstr>\xe5\xb7\xa5\xe4\xbd\x9c\xe8\xa1\xa8</vt:lpstr></vt:variant>'
            + ('<vt:variant><vt:i4>%d</vt:i4></vt:variant></vt:vector></HeadingPairs>'
               % len(names)).encode()
            + ('<TitlesOfParts><vt:vector size="%d" baseType="lpstr">' % len(names)).encode()
            + parts
            + b'</vt:vector></TitlesOfParts></Properties>')


def _avail_ram_bytes():
    '''当前可用物理内存（字节）。取不到返回 -1。'''
    try:
        import ctypes

        class _MEMSTATUS(ctypes.Structure):
            _fields_ = [('dwLength', ctypes.c_ulong),
                        ('dwMemoryLoad', ctypes.c_ulong),
                        ('ullTotalPhys', ctypes.c_ulonglong),
                        ('ullAvailPhys', ctypes.c_ulonglong),
                        ('ullTotalPageFile', ctypes.c_ulonglong),
                        ('ullAvailPageFile', ctypes.c_ulonglong),
                        ('ullTotalVirtual', ctypes.c_ulonglong),
                        ('ullAvailVirtual', ctypes.c_ulonglong),
                        ('ullAvailExtendedVirtual', ctypes.c_ulonglong)]

        st = _MEMSTATUS()
        st.dwLength = ctypes.sizeof(_MEMSTATUS)
        if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(st)):
            return -1
        return int(st.ullAvailPhys)
    except Exception:
        return -1


_CELL_BYTES = 71


def ram_enough_for_sheets(cells):
    '''判断 openpyxl 普通模式能否安全装下这么多单元格。'''
    need = cells * _CELL_BYTES
    avail = _avail_ram_bytes()
    if avail <= 0:
        return cells <= 6000000
    return need <= avail * 0.45


def _pick_spool_dir(out_dir, need):
    '''挑一个放得下 need 字节的临时目录（优先输出目录所在盘）。'''
    cands = []
    if out_dir and os.path.isdir(out_dir):
        cands.append(out_dir)
    cands.append(tempfile.gettempdir())
    for d in cands:
        try:
            if shutil.disk_usage(d).free >= need:
                return d
        except Exception:
            continue
    return cands[-1]


class TooBigForSheets(Exception):
    '''工作表太大，openpyxl 装不下，不适合「拆到当前工作簿」。'''


def _iter_sheet_parts(fobj, chunk_size=4194304):
    '''流式产出工作表 XML 的片段：(kind, data)。

    kind == 'prefix' : <sheetData> 之前（含 <sheetData> 标签本身）
    kind == 'row'    : 一个完整的 <row ...>...</row>（或自闭合的 <row .../>）
    kind == 'suffix' : </sheetData> 之后（含 </sheetData>）
    '''
    buf = b''
    while True:
        k = buf.find(b'<sheetData>')
        k2 = buf.find(b'<sheetData/>')
        if k2 >= 0 and (k < 0 or k2 < k):
            yield ('prefix', buf[:k2 + len(b'<sheetData/>')])
            buf = buf[k2 + len(b'<sheetData/>'):]
            break
        if k >= 0:
            yield ('prefix', buf[:k + len(b'<sheetData>')])
            buf = buf[k + len(b'<sheetData>'):]
            break
        if len(buf) > 16 * chunk_size:
            raise RuntimeError('工作表 XML 结构异常：找不到 <sheetData> 标签。')
        more = fobj.read(chunk_size)
        if not more:
            yield ('prefix', buf)
            return
        buf += more
    while True:
        e = buf.find(b'</sheetData>')
        search_end = len(buf) if e < 0 else e
        pos = consumed = 0
        while True:
            m = _ROW_START_RE.search(buf, pos, search_end)
            if m is None:
                break
            i = m.start()
            gt = buf.find(b'>', i)
            if gt < 0 or gt >= search_end:
                break
            if buf[gt - 1:gt] == b'/':
                yield ('row', buf[i:gt + 1])
                pos = consumed = gt + 1
                continue
            j = buf.find(b'</row>', gt)
            if j < 0 or j + 6 > search_end:
                break
            yield ('row', buf[i:j + 6])
            pos = consumed = j + 6
        if e >= 0:
            yield ('suffix', buf[e:])
            return
        buf = buf[consumed:]
        more = fobj.read(chunk_size)
        if not more:
            if buf:
                yield ('suffix', buf)
            return
        buf += more


def _expand_shared_formulas(row_bytes, old_row, new_row, shared):
    '''把行里的共享公式就地展开成普通公式（并按新行号平移），返回新的行字节。

    【为什么必须展开】
    共享公式的写法是：一个 master 单元格带 ref + si 和公式正文，
    同一 si 的其它单元格只写 <f t="shared" si="0"/>，正文靠 master 提供。
    源表里 master 的 ref 是 "OC5:OC68" 这种**连续行区间**，而拆表会把行号
    重排、把不同行分到不同文件 —— 依赖单元格和它的 master 会被拆散，
    输出文件里只剩一个没有正文的 <f t="shared" si="0"/>，
    Excel 打开会报"文件已损坏，需要修复"。
    所以这里按 openpyxl 的口径统一展开成普通公式：
        master    -> <f>公式正文</f>
        dependent -> <f>把正文按相对偏移翻译到本单元格的公式</f>

    【V4.4 修复：平移锚点】
    V4.3 展开时只拿到**旧行号**，写出的公式引用的还是旧坐标 —— 源第 5 行的
    =B5*2 挪到新表第 3 行后仍是 =B5*2，而输出里还带着旧缓存值，打开时数值
    看着对，一编辑重算就错。现在展开时同时传入新行号：
        master    -> 以旧坐标为锚建 Translator（供依赖单元格复用），
                     自身正文再按 旧坐标 -> 新坐标 平移后写出；
        dependent -> 先用 master 的翻译器得到**旧坐标**下的公式，
                     再以本单元格旧坐标为锚，平移到**新坐标**后写出。

    优化：Translator 在构造时就要分词，所以每个 si 只建一次、复用给该 si 的
    所有依赖单元格（源表里 master 少、依赖多，省掉绝大部分分词开销）。
    '''
    pos = 0
    while True:
        i = row_bytes.find(b'<f', pos)
        if i < 0:
            return row_bytes
        cstart = row_bytes.rfind(b'<c', 0, i)
        if cstart < 0:
            pos = i + 2
            continue
        ctag_end = row_bytes.find(b'>', cstart)
        cm = _COL_R_ATTR_RE.search(row_bytes, cstart, ctag_end + 1)
        if cm is None:
            pos = i + 2
            continue
        mf = _F_TAG_RE.match(row_bytes, i)
        if mf is None:
            pos = i + 2
            continue
        attrs = dict(_ATTR_RE.findall(mf.group(1)))
        ftype = attrs.get(b't', b'')
        if ftype not in (b'shared', b''):
            # 数组公式 / 数据表公式等保持原样（ref 是区域属性，正文平移反而不对）
            pos = mf.end()
            continue
        if ftype == b'shared':
            if mf.group(2) == b'/':
                text = b''
                end = mf.end()
            else:
                j = row_bytes.find(b'</f>', mf.end())
                if j < 0:
                    pos = mf.end()
                    continue
                text = row_bytes[mf.end():j]
                end = j + 4
            si = attrs.get(b'si', b'0')
            col = cm.group(1).decode('ascii')
            old_coord = col + str(old_row)
            new_coord = col + str(new_row)
            if b'ref' in attrs:
                # master 单元格：以旧坐标为锚注册翻译器（供同 si 依赖复用），
                # 自身正文按 旧 -> 新 平移后写出
                try:
                    if text:
                        tr = Translator('=' + text.decode('utf-8', 'replace'), old_coord)
                        shared[si] = tr
                        txt = tr.translate_formula(new_coord)
                        if txt.startswith('='):
                            txt = txt[1:]
                        repl = b'<f>' + txt.encode('utf-8') + b'</f>'
                    else:
                        repl = b'<f>' + text + b'</f>'
                except Exception:
                    shared[si] = None
                    repl = b'<f>' + text + b'</f>'
            else:
                # 依赖单元格：先平移到旧坐标，再以自身旧坐标为锚平移到新坐标
                try:
                    tr = shared.get(si)
                    if tr is None:
                        repl = b''
                    else:
                        txt = tr.translate_formula(old_coord)
                        if txt.startswith('='):
                            txt = txt[1:]
                        if old_row != new_row:
                            txt = Translator('=' + txt, old_coord).translate_formula(new_coord)
                        if txt.startswith('='):
                            txt = txt[1:]
                        repl = b'<f>' + txt.encode('utf-8') + b'</f>'
                except Exception:
                    txt = ''
                    repl = b'<f>' + txt.encode('utf-8') + b'</f>'
            row_bytes = row_bytes[:i] + repl + row_bytes[end:]
            pos = i + len(repl)
        else:
            # 普通公式（无 t 属性）：【V4.4】随行号平移（与 openpyxl 路径统一）
            if old_row == new_row or mf.group(2) == b'/':
                pos = mf.end()
                continue
            j = row_bytes.find(b'</f>', mf.end())
            if j < 0:
                pos = mf.end()
                continue
            text = row_bytes[mf.end():j]
            new_text = _translate_formula_bytes(text, cm.group(1), old_row, new_row)
            if new_text != text:
                row_bytes = row_bytes[:mf.end()] + new_text + row_bytes[j:]
                pos = mf.end() + len(new_text)
            else:
                pos = j + 4


def _learn_shared_formulas(row_bytes, old_row, shared):
    '''【V4.4】只学习共享公式的 master（不改写行内容）。

    master 行可能因为拆分列空白而整行被跳过；不学习的话，它后面的
    依赖单元格会因为查不到 si 而把公式丢掉。这里只登记 Translator，
    不产出任何字节。
    '''
    if b'<f' not in row_bytes:
        return
    pos = 0
    while True:
        i = row_bytes.find(b'<f', pos)
        if i < 0:
            return
        cstart = row_bytes.rfind(b'<c', 0, i)
        if cstart < 0:
            pos = i + 2
            continue
        ctag_end = row_bytes.find(b'>', cstart)
        cm = _COL_R_ATTR_RE.search(row_bytes, cstart, ctag_end + 1)
        if cm is None:
            pos = i + 2
            continue
        mf = _F_TAG_RE.match(row_bytes, i)
        if mf is None:
            pos = i + 2
            continue
        attrs = dict(_ATTR_RE.findall(mf.group(1)))
        if attrs.get(b't') != b'shared' or b'ref' not in attrs:
            pos = mf.end()
            continue
        if mf.group(2) == b'/':
            text = b''
            end = mf.end()
        else:
            j = row_bytes.find(b'</f>', mf.end())
            if j < 0:
                pos = mf.end()
                continue
            text = row_bytes[mf.end():j]
            end = j + 4
        si = attrs.get(b'si', b'0')
        if text and si not in shared:
            try:
                coord = cm.group(1).decode('ascii') + str(old_row)
                shared[si] = Translator('=' + text.decode('utf-8', 'replace'), coord)
            except Exception:
                pass
        pos = end


def _translate_formula_bytes(text, col, old_row, new_row):
    '''【V4.4】把一个普通公式的正文（不含 <f> 标签）按行号平移。失败返回原文。'''
    try:
        txt = Translator('=' + text.decode('utf-8', 'replace'),
                         col.decode('ascii') + str(old_row)).translate_formula(
            col.decode('ascii') + str(new_row))
        if txt.startswith('='):
            txt = txt[1:]
        return txt.encode('utf-8')
    except Exception:
        return text


def _rewrite_row(row_bytes, old_row, new_row, shared):
    '''把一行重写成目标行号：改行号、改单元格引用、展开/平移公式、去掉隐藏标记。

    「去掉隐藏标记」是有意的：台账这类表常常整表处于筛选状态，
    每一行都带 hidden="1"。拆出来的文件里筛选条件已经没了，
    要是还保留 hidden，用户打开会看到一片折叠的行，以为又丢数据了。

    【性能】整行 413 个引用要一起改，所以用一条 C 层正则扫一遍整行，
    而不是逐个单元格处理（6650 万个单元格逐个处理要几分钟）。
    实测：161598 行 / 6650 万个引用，这一遍约 30 秒。

    【V4.4】_expand_shared_formulas 现在同时拿到新行号：共享公式展开后一律
    平移到新行号，普通公式（非共享/非数组）也在同一遍扫描里平移。
    '''
    out = row_bytes
    if b'<f' in out:
        out = _expand_shared_formulas(out, old_row, new_row, shared)
    if old_row == new_row and b'hidden="' not in out[:out.find(b'>')]:
        return out
    nb = str(new_row).encode()
    out = _CELL_REF_SUB_RE.sub(b'r="\\g<1>' + nb + b'"', out)
    gt = out.find(b'>')
    if gt > 0:
        out = _ROW_HIDDEN_RE.sub(b'', out[:gt]) + out[gt:]
    return out


def _fix_pane(pane_bytes, max_col, start_row):
    '''冻结窗格按新的表头行数 / 列数校正。'''
    a = dict(_ATTR_RE.findall(pane_bytes))

    def _num(key):
        try:
            return int(float(a.get(key, b'0')))
        except Exception:
            return 0

    xs = _num(b'xSplit')
    ys = _num(b'ySplit')
    if xs <= 0 and ys <= 0:
        return pane_bytes
    ys = max(0, min(ys, max(1, start_row) - 1))
    xs = max(0, min(xs, max(1, max_col)))
    if xs <= 0 and ys <= 0:
        return b''
    tl = (_num_to_col_letters(xs + 1) + str(ys + 1)).encode()
    if xs and ys:
        ap = b'bottomRight'
    elif ys:
        ap = b'bottomLeft'
    else:
        ap = b'topRight'
    return (b'<pane'
            + (b' xSplit="%d"' % xs if xs else b'')
            + (b' ySplit="%d"' % ys if ys else b'')
            + b' topLeftCell="' + tl + b'" activePane="' + ap + b'" state="frozen"/>')


def _prepare_prefix(prefix, max_col, last_row, start_row):
    '''改写 <sheetData> 之前的部分（dimension / 冻结窗格 / 选中区域）。'''
    out = prefix
    if last_row >= 1:
        ref = b'A1:' + _num_to_col_letters(max(1, max_col)).encode() + str(last_row).encode()
    else:
        ref = b'A1'
    out = _DIMENSION_RE.sub(b'<dimension ref="' + ref + b'"/>', out, count=1)
    out = _FILTERMODE_RE.sub(b'\\1', out, count=1)
    out = _PANE_RE.sub(lambda m: _fix_pane(m.group(0), max_col, start_row), out, count=1)
    out = _SELECTION_RE.sub(b'', out)
    return out


def _ref_bounds(ref):
    '''把 "A3" / "A3:H9" 解析成 (min_col, min_row, max_col, max_row)。

    注意：openpyxl 的 coordinate_to_tuple 返回的是 (行, 列)，
    而 range_boundaries 返回的是 (列, 行, 列, 行) —— 顺序不一样，别搞混。
    '''
    if ':' in ref:
        return range_boundaries(ref)
    row, col = coordinate_to_tuple(ref)
    return (col, row, col, row)


def _remap_sqref_str(refs, row_map):
    '''把 "A3:OW161598 B5:D9" 这类 sqref 按行号映射重排（连续段会合并）。'''
    out = []
    for tok in refs.decode('ascii', 'ignore').split():
        try:
            c1, r1, c2, r2 = _ref_bounds(tok)
        except Exception:
            continue
        lo = _num_to_col_letters(min(c1, c2))
        hi = _num_to_col_letters(max(c1, c2))
        mapped = [row_map[r] for r in range(r1, r2 + 1) if r in row_map]
        i = 0
        while i < len(mapped):
            j = i
            while j + 1 < len(mapped) and mapped[j + 1] == mapped[j] + 1:
                j += 1
            if j > i:
                out.append('%s%d:%s%d' % (lo, mapped[i], hi, mapped[j]))
            else:
                out.append('%s%d' % (lo, mapped[i]))
            i = j + 1
    return ' '.join(out)


def _fix_merges(block, row_map):
    '''合并区域：只保留行全在本次输出里的，并把行号映射过去。'''
    kept = []
    for m in _MERGECELL_RE.finditer(block):
        a = dict(_ATTR_RE.findall(m.group(0)))
        ref = a.get(b'ref', b'').decode('ascii', 'ignore')
        try:
            c1, r1, c2, r2 = _ref_bounds(ref)
        except Exception:
            continue
        rows = list(range(r1, r2 + 1))
        if not all(r in row_map for r in rows):
            continue
        new_rows = [row_map[r] for r in rows]
        if new_rows != list(range(new_rows[0], new_rows[0] + len(new_rows))):
            continue
        lo = _num_to_col_letters(min(c1, c2))
        hi = _num_to_col_letters(max(c1, c2))
        kept.append(b'<mergeCell ref="' + ('%s%d:%s%d' % (lo, new_rows[0], hi, new_rows[-1])).encode() + b'"/>')
    if not kept:
        return b''
    return b'<mergeCells count="' + str(len(kept)).encode() + b'">' + b''.join(kept) + b'</mergeCells>'


def _lookup_rel(rels_bytes, rid):
    '''在 sheetN.xml.rels 里查 rId 对应的 (Type, Target, 是否 External)。'''
    if not rels_bytes:
        return None
    for tag in re.finditer(b'<Relationship\\b[^>]*?/>', rels_bytes):
        a = dict(_ATTR_RE.findall(tag.group(0)))
        if a.get(b'Id') != rid:
            continue
        ext = a.get(b'TargetMode', b'') == b'External'
        return (a.get(b'Type', b''), a.get(b'Target', b''), ext)
    return None


def _fix_hyperlinks(suffix, row_map, src_rels):
    '''超链接：行号跟着重排；外部链接会重建一份 sheet rels。'''
    m = _HYPERLINKS_RE.search(suffix)
    if not m:
        return (suffix, None)
    rid_out = []
    tags = []
    for hm in _HYPERLINK_RE.finditer(m.group(0)):
        try:
            tag = hm.group(0)
            a = dict(_ATTR_RE.findall(tag))
            ref = a.get(b'ref', b'').decode('ascii', 'ignore')
            c1, r1, c2, r2 = _ref_bounds(ref)
            rows = [row_map[r] for r in range(r1, r2 + 1) if r in row_map]
        except Exception:
            continue
        if not rows:
            continue
        new_ref = ('%s%d:%s%d' % (_num_to_col_letters(min(c1, c2)), rows[0],
                                  _num_to_col_letters(max(c1, c2)), rows[-1])).encode()
        rid = a.get(b'r:id')
        if rid is not None:
            info = _lookup_rel(src_rels, rid)
            if info is None:
                continue
            new_rid = b'rId%d' % (len(rid_out) + 1)
            rid_out.append((new_rid, info[0], info[1], info[2]))
            tag = re.sub(b'(?<=\\s)r:id="[^"]*"', b'r:id="' + new_rid + b'"', tag, count=1)
        tag = re.sub(b'(?<=\\s)ref="[^"]*"', b'ref="' + new_ref + b'"', tag, count=1)
        tags.append(tag)
    new_part = b''
    if tags:
        new_part = b'<hyperlinks>' + b''.join(tags) + b'</hyperlinks>'
    rels = None
    if rid_out:
        rels = (_XML_DECL + ('<Relationships xmlns="%s">' % _PKG_REL_NS).encode()
                + b''.join(b'<Relationship Id="%s" Type="%s" Target="%s"%s/>' % (r[0], r[1], r[2], b' TargetMode="External"' if r[3] else b'') for r in rid_out)
                + b'</Relationships>')
    return (suffix.replace(m.group(0), new_part), rels)


def _prepare_suffix(suffix, row_map, last_row, max_col, start_row, src_rels):
    '''改写 </sheetData> 之后的部分。返回 (新后缀, 工作表 rels 或 None)。'''
    out = _DROP_RE.sub(b'', suffix)

    def _af(m):
        if last_row < start_row:
            return b''
        tag = m.group(0)
        # 【V4.6 修正】筛选按钮要留在**表头最后一行**（start_row - 1）。
        # 原来直接写 start_row，用户看到的就是"筛选从第 4 行开始"——
        # 按钮跑到第一行数据上去了（三行表头、数据从第 4 行起时尤其明显）。
        # 源表 ref 的起始行如果本来就在表头区，它在输出里行号不变
        # （表头是原样复制的），优先沿用它，这样多行表头也不会错位。
        first = max(1, start_row - 1)
        ref_b = dict(_ATTR_RE.findall(tag)).get(b'ref')
        if ref_b:
            try:
                r1 = _ref_bounds(ref_b.decode('ascii'))[1]
                if r1 < start_row:
                    first = r1
            except Exception:
                pass
        new_ref = ('A%d:%s%d' % (first, _num_to_col_letters(max(1, max_col)),
                                 last_row)).encode()
        if b'ref="' in tag:
            # 只换 ref，源表带的筛选条件（filterColumn）、WPS 扩展属性
            # （etc:filterBottomFollowUsedRange 之类）原样保留
            return re.sub(b'(?<=\\s)ref="[^"]*"', b'ref="' + new_ref + b'"', tag, count=1)
        return tag.replace(b'<autoFilter', b'<autoFilter ref="' + new_ref + b'"', 1)

    out = _AUTOFILTER_RE.sub(_af, out)
    out = _MERGECELLS_RE.sub(lambda m: _fix_merges(m.group(0), row_map), out)
    out = _CF_SQREF_RE.sub(lambda m: m.group(1) + _remap_sqref_str(m.group(2), row_map).encode() + m.group(3), out)
    out = _DV_SQREF_RE.sub(lambda m: m.group(1) + _remap_sqref_str(m.group(2), row_map).encode() + m.group(3), out)
    return _fix_hyperlinks(out, row_map, src_rels)


def _read_sst(path, entry_names):
    '''把 sharedStrings.xml 拆成 (头部标签, 条目列表)；结构不对就返回 None。

    返回 None 时调用方会退回"整份照搬"，绝不会产生坏文件。
    '''
    if 'xl/sharedStrings.xml' not in entry_names:
        return None
    try:
        with zipfile.ZipFile(path) as z:
            raw = z.read('xl/sharedStrings.xml')
        m = _SST_SPLIT_RE.search(raw)
        if not m:
            return None
        head, body = m.group(1), m.group(2)
        entries = _SST_SI_RE.findall(body)
        if not entries or body.count(b'<si') != len(entries):
            return None
        return (head, entries)
    except Exception:
        return None


def _collect_si_refs(zsrc, entry, chunk=4194304):
    '''扫一个工作表部件，收集它引用到的全部共享字符串下标。

    【V4.6】为什么需要它：共享字符串表会做「子集化」（把本组没用到的 <si>
    置空来缩体积）。如果随拆出文件一起保留了源文件的其它工作表，就必须把这些
    表引用到的 <si> 也标记为保留 —— 否则那些表的文字会被一并清空，
    打开就是一片空白（而数字还在，症状极具迷惑性）。

    分块扫描并保留尾巴，避免把跨块的单元格切断；重复计数由 set 天然去重。
    '''
    refs = set()
    tail = b''
    try:
        with zsrc.open(entry) as f:
            while True:
                buf = f.read(chunk)
                if not buf:
                    break
                data = tail + buf
                if b't="s"' in data:
                    refs.update(int(x) for x in _SI_REF_RE.findall(data))
                tail = data[-256:]
    except Exception:
        pass
    return refs


def _subset_sst(head, entries, bitmap):
    '''按 bitmap 生成共享字符串表：用到的条目原样保留，没用到的置成 <si/>。

    【为什么不重排索引】
    重排索引就得把工作表里每一处 <v>433</v> 都改掉（1300 万处），又慢又容易错。
    这里改成"保留条数、把没用到的条目清空"—— 索引完全不动，
    没被引用的条目内容是什么都无所谓，而 <si/> 高度重复、压缩后几乎不占体积。
    实测 43 个分组的输出从 421MB 降到约 57MB。
    '''
    EMPTY = _SST_EMPTY
    body = b''.join(e if u else EMPTY for e, u in zip(entries, bitmap))
    return _XML_DECL + head + body + b'</sst>'


def _copy_entry(zsrc, name, zo, chunk=1048576):
    # 【V4.6】force_zip64 一律打开：共享字符串表这类部件在超大台账里未压缩
    # 也可能顶到 zipfile 的 2GB 硬限制（ZIP64_LIMIT），
    # 报错形式还是那句 File size too large, try using force_zip64。
    # 对小于 4GB 的部件没有任何副作用。
    with zsrc.open(name) as f:
        with zo.open(name, 'w', force_zip64=True) as d:
            while True:
                b = f.read(chunk)
                if not b:
                    break
                d.write(b)


def _sheet_rels_name(entry):
    d, base = entry.rsplit('/', 1)
    return f'{d}/_rels/{base}.rels'


def _part_closure(zsrc, names, src_entry, dst_entry, out, seen, depth=3):
    '''收集「部件 + 它的 .rels + rels 里引用的包内部件」的闭包。

    【V4.6】为什么需要：随拆出文件一起保留的其它工作表，可能带图片、图表、
    批注（对应 xl/drawings/*、xl/media/*、xl/comments*.xml 等部件，通过
    xl/worksheets/_rels/sheetN.xml.rels 关联）。只搬工作表本身、不搬这些依赖，
    包里的关系表就指向不存在的部件 —— Excel / WPS 直接判文件损坏，
    check_package_rels 也会拦下来报「结构不合法」。

    依赖部件一律保持**原名**搬运：改名的只有工作表本身，而 rels 里的 Target
    是相对路径，不随工作表改名而变。
    '''
    if src_entry in seen or depth < 0:
        return
    seen.add(src_entry)
    out.append((src_entry, dst_entry))
    rel_src = _sheet_rels_name(src_entry)
    if rel_src not in names:
        return
    out.append((rel_src, _sheet_rels_name(dst_entry)))
    try:
        data = zsrc.read(rel_src)
    except Exception:
        return
    base = posixpath.dirname(src_entry)
    for tag in re.findall(rb'<Relationship\b[^>]*/?>', data):
        a = dict(_ATTR_RE.findall(tag))
        if a.get(b'TargetMode') == b'External':
            continue
        t = a.get(b'Target', b'').decode('utf-8', 'ignore')
        if not t or t.startswith(('http://', 'https://', 'mailto:', 'file:', '#')):
            continue
        resolved = (posixpath.normpath(t.lstrip('/')) if t.startswith('/')
                    else posixpath.normpath(posixpath.join(base, t)))
        if resolved in names and resolved not in seen:
            _part_closure(zsrc, names, resolved, resolved, out, seen, depth - 1)


def _extra_ct_fragments(zsrc, entries, skip=()):
    '''从源包的 [Content_Types].xml 里搬出依赖部件需要的声明。

    Default 按扩展名生效（png / jpeg / vml 之类），全部搬过来；
    Override 按部件路径生效，只搬 entries 里那几个、且不在 skip 里的
    （skip = 我们自己已经声明过的部件，重复 Override 是非法包）。
    rels / xml 两个 Default 也已经在 _CONTENT_TYPES_HEAD 里声明过，跳过。
    '''
    try:
        ct = zsrc.read('[Content_Types].xml').decode('utf-8', 'ignore')
    except Exception:
        return b''
    out = []
    for tag in re.findall(r'<Default\b[^>]*/>', ct):
        if re.search(r'Extension="(rels|xml)"', tag):
            continue
        out.append(tag)
    want = {e.lstrip('/') for e in entries}
    skip_set = {s.lstrip('/') for s in skip}
    for tag in re.findall(r'<Override\b[^>]*/>', ct):
        m = re.search(r'PartName="([^"]*)"', tag)
        if not m:
            continue
        name = m.group(1).lstrip('/')
        if name in want and name not in skip_set:
            out.append(tag)
    return ''.join(out).encode('utf-8')


def check_package_rels(path):
    '''校验输出包里的每个关系目标都能落到真实存在的部件上，返回问题列表。

    【为什么必须有这一步 —— V4.2 的血泪教训】
    xl/_rels/workbook.xml.rels 里的 Target 必须相对 xl/ 目录写。一旦多写一层
    （写成 xl/styles.xml），解析出来就是 xl/xl/styles.xml —— 不存在的部件，
    整个包按 OPC 规范是**非法**的。
    但 openpyxl 读共享字符串是**直接按固定路径**取的、不查关系表，所以测试全绿；
    而 WPS/Excel 严格走关系表，找不到字符串表就把所有 t="s" 单元格显示成空白，
    用户看到的是"拆出来的表格是空白的"。
    这个校验极便宜（只读几个 .rels），却能拦住整类"文件看起来合法、用户却打不开"的问题。

    【V4.4】输出包本身读不开（zip 损坏等）也算问题，不再静默放行。
    '''
    problems = []
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return ['输出包不是合法的 zip（%s）' % e]
    with z:
        try:
            names = set(z.namelist())
        except Exception as e:
            return ['输出包无法读取文件列表（%s）' % e]
        for n in sorted(names):
            if not n.endswith('.rels'):
                continue
            owner = os.path.dirname(os.path.dirname(n))
            try:
                targets = _REL_TARGET_RE.findall(z.read(n))
            except Exception as e:
                problems.append('%s 读取失败（%s）' % (n, e))
                continue
            for raw in targets:
                t = raw.decode('utf-8')
                if t.startswith(('http://', 'https://', 'mailto:', 'file:', '#')):
                    continue
                if t.startswith('/'):
                    resolved = posixpath.normpath(t.lstrip('/'))
                else:
                    resolved = posixpath.normpath(posixpath.join(owner, t))
                if resolved in names:
                    continue
                problems.append(f'{n} 里的 Target={t!r} 解析成 {resolved!r}，包里没有这个部件')
    return problems


def _read_entry_or_none(zsrc, name):
    try:
        return zsrc.read(name)
    except Exception:
        return None


def _assemble_one_file(zsrc, out_path, sheet_name, prefix, suffix, sheet_rels, date1904,
                       header_spans, offsets, lengths, spool, present, sst_data=None,
                       extra_sheets=()):
    '''写出一个 xlsx：小部件自己生成，样式/共享字符串从源文件原样搬，工作表流式写。

    【V4.6】extra_sheets = [(名称, 源部件路径, 输出部件路径), ...] ——
    源文件里除选定表之外的其它工作表，原样搬进每个输出文件
    （用户要的是"只拆表一，表2 照样在"）。
    '''
    extra_entries = [e[2] for e in extra_sheets]
    names = set(zsrc.namelist())
    # 【V4.6】先算出要搬运的部件闭包（工作表 + 它的 .rels + 依赖的图片/绘图/批注…），
    # 才能在写 [Content_Types].xml 时把这些部件都声明上
    extra_parts = []
    if extra_sheets:
        _seen = set()
        for _nm, src_entry, dst_entry in extra_sheets:
            _part_closure(zsrc, names, src_entry, dst_entry, extra_parts, _seen)
    _declared = {'docProps/app.xml', 'docProps/core.xml', 'xl/workbook.xml',
                 'xl/worksheets/sheet1.xml', 'xl/styles.xml',
                 'xl/sharedStrings.xml', 'xl/theme/theme1.xml'}
    _declared |= set(extra_entries)
    _deps = [src for src, _dst in extra_parts if src not in _declared]
    _extra_ct = _extra_ct_fragments(zsrc, _deps, _declared) if _deps else b''
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=1, allowZip64=True) as zo:
        zo.writestr('[Content_Types].xml',
                    _content_types_xml(present, extra_entries, _extra_ct))
        zo.writestr('_rels/.rels', _ROOT_RELS_XML)
        zo.writestr('docProps/core.xml', _core_xml())
        zo.writestr('docProps/app.xml', _app_xml(sheet_name, [e[0] for e in extra_sheets]))
        rels_xml, extra_rids = _wb_rels_xml(present, extra_entries)
        zo.writestr('xl/workbook.xml', _workbook_xml(
            sheet_name, date1904,
            [(e[0], extra_rids[i]) for i, e in enumerate(extra_sheets)]))
        zo.writestr('xl/_rels/workbook.xml.rels', rels_xml)
        for key, target, _typ in _COPYABLE_PARTS:
            if key not in present:
                continue
            if key == b'sharedStrings' and sst_data is not None:
                zo.writestr(target, sst_data)
            else:
                _copy_entry(zsrc, target, zo)
        # 【V4.6】保留的其它工作表及其依赖部件（图片/绘图/批注…），按闭包原样搬
        for src_entry, dst_entry in extra_parts:
            if dst_entry == src_entry:
                _copy_entry(zsrc, src_entry, zo)
            else:
                zo.writestr(dst_entry, zsrc.read(src_entry))
        if sheet_rels:
            zo.writestr('xl/worksheets/_rels/sheet1.xml.rels', sheet_rels)
        # 【V4.5】工作表 XML 必须**始终**按 zip64 写。
        # 原来靠 est > 3.5GB 估算来开关 zip64，而 zipfile 的硬限制是 2GB
        # （ZIP64_LIMIT）—— 估算偏一点就会出现"以为不用 zip64、结果写出超 2GB"，
        # 直接抛 RuntimeError: File size too large, try using force_zip64，
        # 拆到一半整个任务失败（实测 16 万行 × 412 列的台账必踩）。
        # 估算本身也不可靠：它是拿源表 XML 的未压缩大小按行数比例外推的。
        # zip64 对小于 4GB 的文件没有任何副作用，Excel / WPS 都正常识别，索性常开。
        with zo.open('xl/worksheets/sheet1.xml', 'w', force_zip64=True) as dst:
            dst.write(prefix)
            for off, ln in header_spans:
                spool.seek(off)
                dst.write(spool.read(ln))
            for off, ln in zip(offsets, lengths):
                spool.seek(off)
                dst.write(spool.read(ln))
            dst.write(suffix)
    problems = check_package_rels(out_path)
    if problems:
        raise RuntimeError('生成的文件结构不合法（关系表指向了不存在的部件），已中止：\n  '
                           + '\n  '.join(problems[:6])
                           + '\n\n这是程序 BUG，请把这条信息连同源文件反馈。')


def split_to_files_stream(path, sheet, groups, start_row, out_dir, base, on_progress=None,
                          keep_others=False):
    '''超大文件流式拆分（拆成多个文件）。返回写出文件数。

    【V4.4】master 行不在任何分组里（拆分列空白被跳过）时也会学习它的
    共享公式正文（_learn_shared_formulas），依赖单元格不再丢公式；
    删掉了从未使用的 notes 返回值。

    【V4.6】keep_others=True 时，把源文件里除选定表之外的其它工作表也原样
    搬进每个输出文件 —— 用户要的是"只拆表一，表2 照样在"。
    '''
    zmap = _sheet_xml_map(path)
    entry = zmap.get(sheet)
    if not entry:
        raise RuntimeError('找不到工作表「%s」对应的 XML 部件。' % sheet)
    # 【V4.6】要一起保留的其它工作表。
    # 拆出的数据总是写到 xl/worksheets/sheet1.xml，所以源表里那个 sheet1.xml
    # 如果不在本次拆分范围内（用户拆的不是第一张表），得换个编号避开冲突。
    extra_sheets = []
    if keep_others:
        taken = {'xl/worksheets/sheet1.xml'}
        all_entries = set(zmap.values())
        k = len(all_entries) + 1
        for nm, e in zmap.items():
            if nm == sheet:
                continue
            dst = e
            if dst in taken:
                while True:
                    cand = 'xl/worksheets/sheet%d.xml' % k
                    k += 1
                    if cand not in taken and cand not in all_entries:
                        dst = cand
                        break
            taken.add(dst)
            extra_sheets.append((nm, e, dst))
    nrow, ncol = sheet_extent_cached(path, sheet)
    ncol = max(1, ncol)
    items = list(groups.items())
    ng = len(items)
    row_maps = []
    new_of = {}
    row2g = {}
    for gi, (_val, rows) in enumerate(items):
        rm = build_row_map(rows, start_row)
        row_maps.append(rm)
        for old_r, new_r in rm.items():
            if old_r < start_row:
                continue
            row2g[old_r] = gi
            new_of[old_r] = new_r
    if on_progress:
        on_progress(2, '准备中…')
    with zipfile.ZipFile(path) as zsrc:
        zi = zsrc.getinfo(entry)
        names = set(zsrc.namelist())
        sst = _read_sst(path, names)
        subset_ok = bool(sst) and len(sst[1]) * ng <= _SST_SUBSET_LIMIT
        if not subset_ok:
            sst = None
        # 【V4.6】保留的其它工作表引用到的共享字符串必须一并留在子集里，
        # 否则那些表的文字会被置空（数字还在，看着就像"表空了"）
        other_used = set()
        if extra_sheets:
            for _nm, src_entry, _dst in extra_sheets:
                other_used |= _collect_si_refs(zsrc, src_entry)
        need = int(zi.file_size * 1.2) + 314572800
        spool_dir = _pick_spool_dir(out_dir, need)
        fd, spool_path = tempfile.mkstemp(prefix='一键拆表-', suffix='.spool', dir=spool_dir)
        os.close(fd)
        spool = open(spool_path, 'w+b', buffering=4194304)
        offsets = [array('q') for _ in range(ng)]
        lengths = [array('q') for _ in range(ng)]
        shared = {}
        header_spans = []
        used_ids = [set() for _ in range(ng)] if subset_ok else None
        hdr_used = set()
        prefix = suffix = b''
        seen = last_r = 0
        written = 0
        try:
            with zsrc.open(entry) as f:
                for kind, data in _iter_sheet_parts(f):
                    if kind == 'prefix':
                        prefix = data
                        continue
                    if kind == 'suffix':
                        suffix = data
                        continue
                    m = _ROW_R_ATTR_RE.search(data, 0, 64)
                    old_r = int(m.group(1)) if m else last_r + 1
                    last_r = old_r
                    if old_r < start_row:
                        nb = _rewrite_row(data, old_r, old_r, shared)
                        off = spool.tell()
                        spool.write(nb)
                        header_spans.append((off, len(nb)))
                        if used_ids is not None and b't="s"' in nb:
                            hdr_used.update(int(s) for s in _SI_REF_RE.findall(nb))
                        continue
                    gi = row2g.get(old_r)
                    if gi is None:
                        # 【V4.4】master 行可能整行不在任何分组里（拆分列空白等），
                        # 也要学习它的共享公式正文，依赖单元格才不会丢公式
                        if b'<f' in data:
                            _learn_shared_formulas(data, old_r, shared)
                        continue
                    nb = _rewrite_row(data, old_r, new_of[old_r], shared)
                    off = spool.tell()
                    spool.write(nb)
                    offsets[gi].append(off)
                    lengths[gi].append(len(nb))
                    if used_ids is not None and b't="s"' in nb:
                        used_ids[gi].update(map(int, _SI_REF_RE.findall(nb)))
                    seen += 1
                    if on_progress and nrow and seen % 8192 == 0:
                        on_progress(3 + int(52 * min(1.0, seen / float(nrow))),
                                    '读取源表数据…（已处理 %d 行）' % seen)
            spool.flush()
            meta = _wb_meta(path)
            date1904 = meta.get('epoch') == MAC_EPOCH
            present = set()
            for key, target, _typ in _COPYABLE_PARTS:
                if target in names:
                    present.add(key)
            sheet_rels = _read_entry_or_none(zsrc, _sheet_rels_name(entry))
            if sheet_rels is None:
                sheet_rels = b''
            used_names = set()
            for gi, (val, _rows) in enumerate(items):
                if not offsets[gi]:
                    continue
                if on_progress:
                    on_progress(56 + int(42 * gi / max(1, ng)), '写出文件：%s' % val)
                last_new = max(row_maps[gi].values())
                pfx = _prepare_prefix(prefix, ncol, last_new, start_row)
                sfx, rels = _prepare_suffix(suffix, row_maps[gi], last_new, ncol, start_row, sheet_rels)
                out_path = unique_output_path(out_dir, base, val, used_names)
                sst_data = None
                if subset_ok:
                    bm = bytearray(len(sst[1]))
                    for i in used_ids[gi]:
                        bm[i] = 1
                    for i in hdr_used:
                        bm[i] = 1
                    for i in other_used:
                        if i < len(bm):
                            bm[i] = 1
                    sst_data = _subset_sst(sst[0], sst[1], bm)
                # 【V4.6】拆出的表名要避开保留下来的其它表名，否则一个包里出现
                # 两张同名工作表，Excel / WPS 直接判为损坏
                sheet_used = {nm.lower() for nm, _s, _d in extra_sheets}
                _assemble_one_file(zsrc, out_path, safe_sheet_name(val, sheet_used),
                                   pfx, sfx, rels,
                                   date1904, header_spans, offsets[gi], lengths[gi], spool,
                                   present, sst_data=sst_data, extra_sheets=extra_sheets)
                written += 1
            if on_progress:
                on_progress(99, '整理输出…')
        finally:
            spool.close()
            try:
                os.remove(spool_path)
            except OSError:
                pass
    return written


class SplitWorker:
    def __init__(self, path, sheet, col_idx, mode, out_dir, keep_blank, start_row,
                 src_path=None, keep_others=False):
        self.path = path            # 真正被解析的文件（xlsb 已转成 xlsx）
        self.src_path = src_path or path   # 用户选中的原始文件（只用于命名/报错定位）
        self.sheet = sheet
        self.col_idx = col_idx
        self.mode = mode
        self.out_dir = out_dir
        self.keep_blank = keep_blank
        self.keep_others = keep_others   # 【V4.6】拆成多个文件时保留其它工作表
        self.start_row = max(1, int(start_row))
        self.on_progress = None

    def run(self):
        '''返回 (ok: bool, msg: str)。'''
        try:
            if not isinstance(self.col_idx, int) or self.col_idx < 1:
                # 【V4.4】引擎级兜底：没选列时明确提示，不再把整列当空白处理
                return (True, '请先选择拆分依据列。')
            if self.mode == 'sheets':
                return (True, self._split_to_sheets())
            return (True, self._split_to_files())
        except TooBigForSheets as e:
            return (False, str(e))
        except Exception:
            full = traceback.format_exc()
            log = write_error_log(self.src_path, full)
            msg = format_error(full)
            if log:
                msg += '\n\n完整错误日志已写入：\n%s' % log
            return (False, msg)

    def _split_to_sheets(self):
        try:
            nrow, ncol = sheet_extent_cached(self.path, self.sheet)
        except Exception:
            nrow = ncol = 0
        cells = nrow * ncol
        if cells and not ram_enough_for_sheets(cells):
            avail = _avail_ram_bytes()
            avail_txt = ('%.1f GB' % (avail / 1073741824.0)) if avail > 0 else '未知'
            raise TooBigForSheets('这个工作表太大了，不适合「拆到新工作簿」。\n\n'
                                  '工作表：%s 列 × %s 行，约 %s 万个单元格。\n'
                                  '「拆到新工作簿」需要把整个工作簿读进内存再写出，预计需要约 %.1f GB，'
                                  '而当前可用内存只有 %s —— 继续下去程序会被拖死（看起来就是卡死）。\n\n'
                                  '请改用「拆成多个文件」（那个模式是流式处理，多大都不怕），'
                                  '或者先在源表里删掉不需要的列再拆。'
                                  % (ncol, nrow, '%.0f' % (cells / 10000.0),
                                     cells * _CELL_BYTES / 1073741824.0, avail_txt))
        groups, empties = group_rows(self.path, self.sheet, self.col_idx,
                                     self.start_row, self.keep_blank)
        if not groups:
            return '该列没有可拆分的数据。'
        # 【V4.5】源文件只读：load_workbook 不写源文件，结果另存到源目录下的
        # 「<原名>-拆分.xlsx」。源工作簿里的其他工作表原样搬进新文件，
        # 源文件一个字节都不动 —— 彻底避免"只拆一张表，别的表被搞坏"。
        out_path = unique_workbook_path(self.src_path)
        wb = openpyxl.load_workbook(self.path)
        src = wb[self.sheet]
        # 【V4.4】与界面/流式通道的列数口径对齐：取 XML 扫描值与 openpyxl 的较大者，
        # 两个口径谁多都不会漏列（此前只用 src.max_column）
        max_col = max(src.max_column, ncol)
        used = {s.lower() for s in wb.sheetnames}
        kept = len(wb.sheetnames)
        total = len(groups)
        for i, (val, rows) in enumerate(groups.items(), 1):
            if self.on_progress:
                self.on_progress(int((i - 1) / total * 100), f'生成工作表：{val}')
            new = wb.create_sheet(title=safe_sheet_name(val, used))
            copy_sheet_rows(src, new, build_row_map(rows, self.start_row), max_col, self.start_row)
        if self.on_progress:
            self.on_progress(99, '保存新工作簿…')
        wb.save(out_path)
        wb.close()
        summary = '\n'.join(f'  {k}: {len(v)} 行' for k, v in groups.items())
        tail = f'\n\n源文件未做任何改动。原有 {kept} 个工作表已原样保留在新文件里。'
        if os.path.splitext(self.src_path)[1].lower() in ('.xlsm', '.xltm', '.xlsb'):
            tail += '\n\n注意：新文件是 .xlsx 格式，宏（VBA）不会保留。'
        if empties:
            tail += f'\n\n另有 {len(empties)} 行因拆分列为空未拆出（如需保留请勾选「空白行也拆出」）。'
        return f'已生成新工作簿（源文件未改动）：\n{out_path}\n\n拆出 {total} 个工作表：\n{summary}{tail}'

    def _split_to_files(self):
        if not self.out_dir:
            return '请先选择输出文件夹。'
        groups, empties = group_rows(self.path, self.sheet, self.col_idx,
                                     self.start_row, self.keep_blank)
        if not groups:
            return '该列没有可拆分的数据。'
        base = os.path.splitext(os.path.basename(self.src_path))[0]
        written = split_to_files_stream(self.path, self.sheet, groups, self.start_row,
                                        self.out_dir, base, self.on_progress,
                                        keep_others=self.keep_others)
        summary = '\n'.join(f'  {k}: {len(v)} 行' for k, v in groups.items())
        tail = ''
        if self.keep_others:
            # 文件都已经写出来了，这里只是补一句说明；读表名失败也绝不能
            # 把整个拆分判成失败（否则用户看到"失败"却发现文件好端端在）
            try:
                others = [n for n in _sheet_xml_map(self.path) if n != self.sheet]
            except Exception:
                others = []
            if others:
                tail += ('\n\n源文件的其它工作表已一并保留在每个文件里：'
                         + '、'.join(others))
        if os.path.splitext(self.src_path)[1].lower() in ('.xlsm', '.xltm', '.xlsb'):
            tail += '\n\n注意：拆出来的文件是 .xlsx 格式，宏（VBA）不会保留。'
        if empties:
            tail += f'\n\n另有 {len(empties)} 行因拆分列为空未拆出（如需保留请勾选「空白行也拆出」）。'
        return f'已拆出 {written} 个文件到：\n{self.out_dir}\n{summary}{tail}'


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.file_path = ''    # 实际被解析的文件（xlsb 已转成 xlsx）
        self.src_path = ''     # 用户选中的原始文件（用于命名与报错定位）
        self.worker = None
        self.q = queue.Queue()
        self._real_rows = 0
        self._dim_warn = ''
        self._gen = 0
        self._probe_pending = False
        self._scan_pending = False
        self._scan_token = 0
        self._busy_t0 = 0.0
        self._busy_text = ''
        root.title('一键拆表工具 ' + VERSION)
        root.minsize(880, 820)
        root.geometry('1020x950')
        setup_fonts(root)
        pad = dict(padx=12, pady=6)
        frm = ttk.Frame(root)
        frm.pack(fill='both', expand=True, **pad)
        row1 = ttk.Frame(frm)
        row1.pack(fill='x', **pad)
        self.ed_file = ttk.Entry(row1, font=BASE_FONT)
        self.ed_file.pack(side='left', fill='x', expand=True)
        ttk.Button(row1, text='选择 Excel 文件…', command=self.pick_file).pack(side='left', padx=(6, 0))
        grid = ttk.Frame(frm)
        grid.pack(fill='x', **pad)
        grid.columnconfigure(1, weight=1)
        ttk.Label(grid, text='工作表：').grid(row=0, column=0, sticky='w', pady=6)
        self.cb_sheet = ttk.Combobox(grid, state='readonly', font=BASE_FONT)
        self.cb_sheet.grid(row=0, column=1, columnspan=2, sticky='ew', pady=6)
        self.cb_sheet.bind('<<ComboboxSelected>>', self.on_sheet_changed)
        ttk.Label(grid, text='拆分依据列：').grid(row=1, column=0, sticky='w', pady=6)
        self.cb_col = ttk.Combobox(grid, state='readonly', font=BASE_FONT)
        self.cb_col.grid(row=1, column=1, columnspan=2, sticky='ew', pady=6)
        self.cb_col.bind('<<ComboboxSelected>>', self.on_col_changed)
        ttk.Label(grid, text='数据起始行：').grid(row=2, column=0, sticky='w')
        row_start = ttk.Frame(grid)
        row_start.grid(row=2, column=1, columnspan=2, sticky='w', pady=6)
        self.sp_start = ttk.Spinbox(row_start, from_=1, to=1048576, width=8,
                                    font=BASE_FONT, command=self.on_start_changed)
        self.sp_start.set(2)
        self.sp_start.pack(side='left')
        self.sp_start.bind('<Return>', self.on_start_changed)
        self.sp_start.bind('<FocusOut>', self.on_start_changed)
        ttk.Label(row_start, text='（第 1 行 ~ 起始行-1 作为表头原样保留）').pack(side='left', padx=(8, 0))
        ttk.Label(frm, text='取值预览（拆分后每个取值一张表/一个文件）：').pack(fill='x', **pad)
        pv_wrap = ttk.Frame(frm)
        pv_wrap.pack(fill='both', expand=True, **pad)
        self.preview = tk.Text(pv_wrap, height=8, wrap='none', font=BASE_FONT,
                               state='disabled', relief='groove')
        sb = ttk.Scrollbar(pv_wrap, command=self.preview.yview)
        self.preview.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.preview.pack(side='left', fill='both', expand=True)
        box = ttk.LabelFrame(frm, text='拆分方式')
        box.pack(fill='x', **pad)
        self.mode_var = tk.StringVar(value='sheets')
        ttk.Radiobutton(box, text='拆到新工作簿（源文件不动，另存为「原名-拆分.xlsx」，原有工作表一并保留）',
                        variable=self.mode_var, value='sheets',
                        command=self.on_mode_changed).pack(anchor='w', padx=10, pady=(6, 0))
        row_mode = ttk.Frame(box)
        row_mode.pack(fill='x', padx=10, pady=(4, 6))
        ttk.Radiobutton(row_mode, text='拆成多个文件（每个取值一个 xlsx）',
                        variable=self.mode_var, value='files',
                        command=self.on_mode_changed).pack(side='left')
        self.ed_out = ttk.Entry(row_mode, font=BASE_FONT)
        self.btn_out = ttk.Button(row_mode, text='输出文件夹…', command=self.pick_out_dir)
        self.ed_out.pack_forget()
        self.btn_out.pack_forget()
        self.blank_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(box, text='空白行也拆出（生成「(空白)」表/文件）',
                        variable=self.blank_var).pack(anchor='w', padx=10, pady=(0, 4))
        # 【V4.6】拆成多个文件时，把源文件里其它工作表一起带进每个输出文件
        self.keep_others_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(box, text='保留源文件的其他工作表（仅「拆成多个文件」；如「说明」「参数」页）',
                        variable=self.keep_others_var).pack(anchor='w', padx=10, pady=(0, 8))
        self.btn_go = ttk.Button(frm, text='一键拆分', command=self.do_split, style='Go.TButton')
        self.btn_go.pack(fill='x', **pad)
        self.btn_go.configure(state='disabled')
        self.status = ttk.Label(frm, text='就绪。')
        self.status.pack(fill='x', **pad)
        self.root.after(80, self._poll)

    def pick_file(self):
        path = filedialog.askopenfilename(
            title='选择 Excel 文件',
            filetypes=[('Excel 文件（xlsx / xlsm / xlsb）', '*.xlsx *.xlsm *.xlsb'),
                       ('Excel 工作簿', '*.xlsx *.xlsm'),
                       ('Excel 二进制工作簿（会自动转换）', '*.xlsb'),
                       ('所有文件', '*.*')])
        if not path:
            return
        self.src_path = path
        # 旧版 .xls 是 OLE 复合文档（BIFF8），连 ZIP 都不是，给个明确出路，
        # 别让用户对着"无法打开文件"发呆
        if os.path.splitext(path)[1].lower() in ('.xls', '.xlt', '.xla', '.xlm'):
            messagebox.showinfo(
                '提示',
                '这是旧版 .xls 格式（二进制 BIFF），程序读不了。\n\n'
                '请先在 WPS / Excel 里把它「另存为」xlsx 或 xlsb，再用本工具处理。',
                parent=self.root)
            return None
        # 扩展名是 .xlsb，或者虽然挂着别的扩展名、内容其实是 xlsb，都走转换流程
        if is_xlsb_path(path) or looks_like_xlsb(path):
            self._open_xlsb(path)
            return None
        self._show_path(path)
        self._load_sheets(path)
        return None

    def _show_path(self, text):
        '''把文件路径写进只读输入框。'''
        self.ed_file.config(state='normal')
        self.ed_file.delete(0, 'end')
        self.ed_file.insert(0, text)
        self.ed_file.config(state='readonly')

    def _load_sheets(self, path):
        '''读工作表名并填充界面。path 必须是能解析的 xlsx / xlsm。'''
        self.file_path = path
        try:
            names = list(_sheet_xml_map(path).keys())
        except Exception as e:
            messagebox.showerror('错误', '无法打开文件：\n%s' % e, parent=self.root)
            return None
        if not names:
            messagebox.showerror('错误', '这个文件里没有工作表。', parent=self.root)
            return None
        self._real_rows = 0
        self.cb_sheet['values'] = names
        self.cb_sheet.set(names[0])
        self.on_sheet_changed(None)
        return None

    def _open_xlsb(self, src):
        '''xlsb（二进制工作簿）先转成 xlsx 再处理。

        【V4.5】xlsb 的部件是 workbook.bin / sheet1.bin 这类二进制文件，
        本工具整套引擎建立在「ZIP + XML」上，自己解析不了，只能让 Office
        自己另存一份 xlsx。转换结果按「路径 + mtime + 大小」缓存，第二次不用等。
        '''
        self._gen += 1
        gen = self._gen
        self.file_path = ''
        self.cb_sheet['values'] = []
        self.cb_sheet.set('')
        self.cb_col['values'] = []
        self.cb_col.set('')
        self.btn_go.config(state='disabled')
        self._show_path(src)
        try:
            cache = xlsb_cache_path(src)
        except OSError as e:
            self._clear_busy('读取失败。')
            messagebox.showerror('错误', '读取文件失败：\n%s' % e, parent=self.root)
            return None
        if os.path.exists(cache) and os.path.getsize(cache) > 0:
            self._after_convert(src, cache, cached=True)
            return None
        self._set_preview('源文件是 .xlsb（二进制工作簿），程序要先调用本机 WPS / Excel '
                          '把它转成 xlsx 才能拆。\n\n'
                          '· 转换由 Office 自己完成，格式、公式、合并单元格完全保真；\n'
                          '· 大文件（一两百 MB）大约要几分钟，期间请不要关闭 WPS 窗口；\n'
                          '· 转换结果会缓存，同一个文件第二次打开就不用再等了。')
        self._set_busy('正在用 WPS / Excel 转换 xlsb…')

        def work():
            try:
                convert_xlsb_to_xlsx(src, cache)
                self.q.put(('converted', gen, src, cache, None))
            except Exception as e:
                self.q.put(('converted', gen, src, cache, e))
        threading.Thread(target=work, daemon=True).start()
        return None

    def _after_convert(self, src, xlsx, cached):
        '''xlsb 转换完成（或命中缓存）后，按普通 xlsx 继续。'''
        self._load_sheets(xlsx)
        # 让用户一眼看出"我选的是 xlsb，程序用的是转换出来的 xlsx"
        self._show_path('%s    （xlsb 已自动转换为 xlsx%s）'
                        % (src, '，用的是上次的缓存' if cached else ''))
        return None

    def _render_converted(self, gen, src, cache, err):
        '''转换线程的回调（在主线程里跑）。'''
        if gen != self._gen:
            return None
        if err is not None:
            self._clear_busy('转换失败。')
            self._set_preview('')
            messagebox.showerror(
                '无法转换 xlsb',
                '这个 .xlsb 文件转换失败：\n\n%s\n\n'
                '转换依赖本机的 WPS 表格或 Microsoft Excel（需要 COM 自动化支持）。\n'
                '如果这台机器上两者都没有，或者 WPS 是绿色版 / 精简版，\n'
                '请先手动在 WPS 里把文件「另存为」xlsx，再用本工具处理。' % err,
                parent=self.root)
            return None
        self._clear_busy('xlsb 转换完成。')
        self._after_convert(src, cache, cached=False)
        return None

    def _set_preview(self, text):
        self.preview.config(state='normal')
        self.preview.delete('1.0', 'end')
        self.preview.insert('1.0', text)
        self.preview.config(state='disabled')

    def _set_busy(self, text):
        '''进入"忙"状态。状态栏会由 _poll 持续刷新已用时间。'''
        self._busy_t0 = time.time()
        self._busy_text = text
        self.status.config(text=text)
        self.btn_go.config(state='disabled')

    def _clear_busy(self, text=''):
        self._busy_t0 = 0.0
        self._busy_text = ''
        self.status.config(text=text)

    def on_sheet_changed(self, _evt=None):
        self._gen += 1
        self._probe_pending = False
        self._scan_pending = False
        self.cb_col.set('')
        self.cb_col['values'] = []
        self._set_preview('')
        self.btn_go.config(state='disabled')
        self._clear_busy()
        if self.file_path and self.cb_sheet.get():
            self._populate_cols()

    def _real_max_col(self):
        '''真实列数（与拆分时拷贝的口径一致）。'''
        return sheet_extent_cached(self.file_path, self.cb_sheet.get())[1]

    def _populate_cols(self):
        '''后台读取表头行填充「拆分依据列」。

        大表（341MB 台账 / 6650 万单元格）扫真实范围要 8 秒左右，
        以前是在主线程做的 —— 窗口会直接冻结十几秒，用户以为程序死了。
        现在放到后台线程，期间界面照常响应，状态栏显示已用时间。
        '''
        self.cb_col.set('')
        self.cb_col['values'] = []
        self._set_preview('')
        self.btn_go.config(state='disabled')
        if not self.file_path or not self.cb_sheet.get():
            return
        self._gen += 1
        gen = self._gen
        path = self.file_path
        sheet = self.cb_sheet.get()
        start = self._start_row()
        hdr_row = max(1, start - 1)
        self._probe_pending = True
        self._set_preview('正在读取表头与真实行列数，请稍候…\n\n（源文件很大时这一步需要几秒到几十秒，界面不会卡住）')
        self._set_busy('正在读取表头…')

        def work():
            try:
                res = probe_sheet(path, sheet, hdr_row)
            except Exception as e:
                res = e
            self.q.put(('probe', gen, path, sheet, start, res))
        threading.Thread(target=work, daemon=True).start()

    def _render_probe(self, gen, path, sheet, start, res):
        if gen != self._gen:
            return
        self._probe_pending = False
        if isinstance(res, Exception):
            self._clear_busy('读取失败。')
            self._set_preview('读取失败：%s' % res)
            messagebox.showerror('错误', '读取表头失败：\n%s' % res, parent=self.root)
            return
        maxc, header, nrows = res
        self._real_rows = nrows
        try:
            declared = declared_dimension(path, sheet)
        except Exception:
            declared = None
        limit = min(max(maxc, 0), MAX_COL_LIMIT)
        cols = []
        for i in range(limit):
            h = header[i] if i < len(header) else None
            text = str(h).strip() if h is not None else ''
            cols.append(f'{get_column_letter(i + 1)} - {text}' if text else get_column_letter(i + 1))
        self.cb_col['values'] = cols
        warn = []
        if declared:
            drow, dcol = declared
            if dcol > maxc:
                warn.append(f'注意：文件声明 {dcol} 列，实际只有 {maxc} 列数据，将按 {maxc} 列拆分。')
            if maxc > dcol:
                warn.append(f'注意：文件声明 {dcol} 列，实际有 {maxc} 列，已按 {maxc} 列处理。')
            if drow < nrows:
                warn.append(f'注意：文件声明 {drow} 行，实际有 {nrows} 行，已按 {nrows} 行处理（旧版会丢掉多出来的行）。')
        if start == 1:
            warn.append('注意：起始行为 1，第 1 行会作为数据参与拆分，不会作为表头保留。')
        self._dim_warn = '\n'.join(warn)
        if cols:
            self.cb_col.current(0)
        self._clear_busy()
        self.on_col_changed(None)

    def on_start_changed(self, _evt=None):
        '''起始行变化 -> 表头行随之变化 -> 刷新列列表（真实范围有缓存，很快）。'''
        if self.file_path and self.cb_sheet.get():
            self._populate_cols()
            return None
        return None

    def on_col_changed(self, _evt=None):
        '''选列后统计取值（后台线程，界面不会卡）。'''
        self._gen += 1
        gen = self._gen
        self._set_preview('')
        if not self.cb_col.get() or not self.file_path:
            self.btn_go.config(state='disabled')
            self._clear_busy()
            self._scan_pending = False
            return None
        idx = self.cb_col.current() + 1
        path = self.file_path
        sheet = self.cb_sheet.get()
        start = self._start_row()
        self._scan_pending = True
        self._set_busy('正在统计取值…')
        self._set_preview('正在统计取值，请稍候…\n\n（源文件很大时这一步需要一点时间，期间界面仍可正常操作）')

        def work():
            try:
                res = scan_column(path, sheet, idx, start_row=start)
            except Exception as e:
                res = e
            self.q.put(('scan', gen, path, sheet, start, res))
        threading.Thread(target=work, daemon=True).start()

    def _render_scan(self, gen, path, sheet, start, res):
        if gen != self._gen:
            return
        self._scan_pending = False
        if isinstance(res, Exception):
            self._clear_busy('读取失败。')
            self._set_preview('读取失败：%s' % res)
            self.btn_go.config(state='disabled')
            return None
        order, counts, empty = res
        lines = [f'{k}: {counts[k]} 行' for k in order]
        if empty:
            lines.append(f'(空白): {empty} 行（默认不拆出，可勾选保留后拆出）')
        head = ''
        try:
            ncol = self._real_max_col()
            nrow = self._real_rows or '?'
            head = f'源表：{ncol} 列 × {nrow} 行（第 {start} 行起为数据）\n'
        except Exception:
            pass
        if self._dim_warn:
            head += self._dim_warn + '\n'
        self._set_preview(head + ('\n'.join(lines) or '（无数据行）'))
        self._clear_busy('共 %d 个取值%s。可以点「一键拆分」了。'
                         % (len(order), ('，另有 %d 个空白行' % empty) if empty else ''))
        self.btn_go.config(state='normal' if (order or empty) else 'disabled')
        return None

    def on_mode_changed(self):
        if self.mode_var.get() == 'files':
            self.ed_out.pack(side='left', fill='x', expand=True, padx=(10, 4))
            self.btn_out.pack(side='left')
            return None
        self.ed_out.pack_forget()
        self.btn_out.pack_forget()
        return None

    def pick_out_dir(self):
        d = filedialog.askdirectory(title='选择输出文件夹')
        if d:
            self.ed_out.delete(0, 'end')
            self.ed_out.insert(0, d)
        return None

    def _start_row(self):
        try:
            v = int(str(self.sp_start.get()).strip())
            return max(1, v)
        except (TypeError, ValueError):
            return 2

    def do_split(self):
        if not self.file_path:
            return
        # 【V4.6 修正】两道防呆，避免"闷头拆旧数据 / 拆到一半说文件不存在"：
        # (1) 转换出来的中间文件被清掉了（用户手动清了缓存目录，或系统清了临时文件）；
        # (2) 源文件在程序打开之后又被外部改动过（在 WPS / Excel 里保存、或者重新
        #     下载了一份）—— 缓存键是「路径 + mtime + 大小」，此时早已失效，
        #     程序手里还是上次那份转换结果，继续拆就会拆到**旧数据**。
        if not os.path.exists(self.file_path):
            messagebox.showwarning(
                '文件不存在',
                '源文件或它的转换缓存已经不在了：\n%s\n\n请重新「选择 Excel 文件」。'
                % self.file_path, parent=self.root)
            return None
        if self.src_path and self.src_path != self.file_path:
            try:
                stale = xlsb_cache_path(self.src_path) != self.file_path
            except OSError:
                stale = False
            if stale:
                messagebox.showwarning(
                    '源文件已变化',
                    '源文件在打开之后又被修改过（在 WPS / Excel 里保存过，或者重新下载了）。\n\n'
                    '程序手里还是上次的转换结果，继续拆会拆到**旧数据**。\n'
                    '请重新「选择 Excel 文件」再拆。', parent=self.root)
                return None
        mode = self.mode_var.get()
        out_dir = self.ed_out.get().strip() if mode == 'files' else ''
        if mode == 'files' and not out_dir:
            messagebox.showwarning('提示', '请先选择输出文件夹。', parent=self.root)
            return None
        path = self.file_path
        sheet = self.cb_sheet.get()
        col_idx = self.cb_col.current() + 1
        if col_idx < 1:
            # 【V4.4】没选中列时直接不拆（此前会静默把整列判成空白）
            return None
        keep_blank = bool(self.blank_var.get())
        keep_others = bool(self.keep_others_var.get())
        start_row = self._start_row()
        self._set_busy('正在拆分…')

        def report(pct, msg):
            self.q.put(('progress', pct, msg))

        def run():
            w = SplitWorker(path, sheet, col_idx, mode, out_dir, keep_blank, start_row,
                            src_path=self.src_path or path, keep_others=keep_others)
            w.on_progress = report
            ok, msg = w.run()
            self.q.put(('done', ok, msg))
        self.worker = run
        threading.Thread(target=run, daemon=True).start()
        return None

    def _poll(self):
        try:
            while True:
                kind, *args = self.q.get_nowait()
                if kind == 'progress':
                    pct, msg = args
                    # 【V4.4】进度消息同步进计时器文本，否则"已用 X 秒"
                    # 每 80ms 就把它覆盖掉，详细进度根本看不见
                    self._busy_text = msg
                    self.status.config(text=f'[{pct}%] {msg}')
                elif kind == 'probe':
                    self._render_probe(*args)
                elif kind == 'scan':
                    self._render_scan(*args)
                elif kind == 'converted':
                    self._render_converted(*args)
                elif kind == 'done':
                    self._on_done(*args)
        except queue.Empty:
            pass
        if self._busy_t0:
            self.status.config(text='%s（已用 %.0f 秒）' % (self._busy_text, time.time() - self._busy_t0))
        self.root.after(80, self._poll)

    def _on_done(self, ok, msg):
        self._busy_t0 = 0.0
        self._busy_text = ''
        if ok:
            self.status.config(text='完成。')
            messagebox.showinfo('拆分完成', msg, parent=self.root)
            self._egg(msg)
        else:
            self.status.config(state='normal')
            self.status.config(text='失败。')
            messagebox.showerror('拆分失败', msg, parent=self.root)
        # 【V4.4】只在列仍处于选中状态时恢复按钮：
        # 拆分期间切换工作表会把列清空，此时不该让用户直接再点「一键拆分」
        if self.cb_col.get():
            self.btn_go.config(state='normal')
        return None

    def _egg(self, msg):
        '''彩蛋弹窗（保留 V2 的那句感谢）。'''
        win = tk.Toplevel(self.root)
        win.title('彩蛋')
        win.resizable(False, False)
        win.grab_set()
        ttk.Label(win, text='🎉 免费拆表成功，感谢铭哥 🎉', padding=(24, 16)).pack()
        ttk.Label(win, text='感谢邹哥测试').pack()
        ttk.Button(win, text='感谢铭哥', command=win.destroy, padding=(16, 2)).pack(pady=(0, 14))
        win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - win.winfo_height()) // 2
        win.geometry(f'+{x}+{y}')


def run_selftest(report_path):
    '''打包后自检：验证 openpyxl / tkinter(tcl-tk) 已正确打进 exe，核心流程可用。

    用法：一键拆表-V4.exe --selftest 报告文件路径
    （打包为窗口程序后没有控制台，所以把结果写到文件里）
    '''
    lines = []

    def w(s):
        lines.append(str(s))

    try:
        w('VERSION=' + VERSION)
        w('frozen=%s' % getattr(sys, 'frozen', False))
        w('executable=%s' % sys.executable)
        w('python=%s' % sys.version.replace('\n', ' '))
        w('openpyxl=%s' % openpyxl.__version__)
        import et_xmlfile
        w('et_xmlfile=%s' % getattr(et_xmlfile, '__version__', '?'))
        r = tk.Tk()
        r.withdraw()
        w('tcl/tk=%s' % r.tk.call('info', 'patchlevel'))
        w('tcl_library=%s' % r.tk.call('info', 'library'))
        w('themes=%s' % ','.join(r.tk.call('ttk::style', 'theme', 'names')))
        r.destroy()
        d = tempfile.mkdtemp(prefix='v4_exe_selftest_')
        raw = os.path.join(d, 'raw.xlsx')
        src = os.path.join(d, 't.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '数据'
        ws.append(['名称', '值', '备注'])
        for i in range(12):
            ws.append(['组%d' % (i % 3), i, 'x%d' % i])
        ws.merge_cells('A1:C1')
        wb.save(raw)
        wb.close()
        with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename == 'xl/worksheets/sheet1.xml':
                    data = re.sub(b'<dimension[^>]*/>', b'<dimension ref="A1:ZZ1000" />', data, count=1)
                zout.writestr(it, data)
        w(f'declared={declared_dimension(src, "数据")!s}')
        w(f'extent={sheet_extent(src, "数据")!s}')
        maxc, header, nrows = probe_sheet(src, '数据', 1)
        w(f'probe={(maxc, nrows)!s} header={header!s}')
        if maxc != 3:
            raise AssertionError('真实列数应为 3，实际 %s（dimension 撒谎未被绕过）' % maxc)
        if nrows != 13:
            raise AssertionError('真实行数应为 13，实际 %s' % nrows)
        groups, empties = group_rows(src, '数据', 1, 2, False)
        w('groups=%s' % {k: len(v) for k, v in groups.items()})
        if sum(len(v) for v in groups.values()) != 12:
            raise AssertionError
        outdir = os.path.join(d, 'out')
        os.makedirs(outdir)
        ok, msg = SplitWorker(src, '数据', 1, 'files', outdir, False, 2).run()
        w('split_ok=%s' % ok)
        w('split_msg=%s' % msg.replace('\n', ' | '))
        if not ok:
            raise AssertionError(msg)
        files = sorted(os.listdir(outdir))
        w('files=%s' % files)
        for fn in files:
            wb2 = openpyxl.load_workbook(os.path.join(outdir, fn))
            w2 = wb2.active
            w(f'  {fn!s} -> {w2.max_row!s}行 x {w2.max_column!s}列 合并={sorted(str(x) for x in w2.merged_cells.ranges)!s}')
            if w2.max_column != 3:
                raise AssertionError('%s 列数错' % fn)
            wb2.close()
        root = tk.Tk()
        icon_ok = set_window_icon(root)
        w('window_icon=%s' % ('ok' if icon_ok else 'MISSING（app.ico 没打进包？）'))
        w('icon_path=%s' % resource_path('app.ico'))
        MainWindow(root)
        root.update()
        w('GUI=ok')
        root.destroy()
        # 【V4.5】xlsb 支持自检：转换要靠 comtypes + 本机 Office 的 COM 自动化。
        # 这里只查注册情况，不真的启动 Office（自检不该弹窗、不该占用几分钟）。
        try:
            import comtypes  # noqa: F401
            w('comtypes=%s' % getattr(comtypes, '__version__', '?'))
        except Exception as e:
            w('comtypes=MISSING（xlsb 无法自动转换：%s）' % e)
        try:
            import winreg
            for _prog in ('KET.Application', 'Excel.Application'):
                try:
                    winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, _prog)
                    w('com_%s=已注册' % _prog)
                except OSError:
                    w('com_%s=未注册' % _prog)
        except Exception as e:
            w('com_check=FAIL %s' % e)
        try:
            w('xlsb_cache_dir=%s' % xlsb_cache_dir())
        except Exception as e:
            w('xlsb_cache_dir=FAIL %s' % e)
        w('----------------------------------------')
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment as _Align
        NSC, NSR, NSS = (20, 400, 40)
        st_path = os.path.join(d, 'styled.xlsx')
        swb = openpyxl.Workbook()
        sws = swb.active
        sws.title = '数据'
        _fonts = [Font(name='Calibri', size=9 + i % 6, bold=bool(i % 2),
                       color='FF%06X' % (1250067 * (i + 1) % 16777215)) for i in range(NSS)]
        _fills = [PatternFill('solid', start_color='FF%06X' % (789516 * (i + 1) % 16777215),
                              end_color='FF%06X' % (789516 * (i + 1) % 16777215)) for i in range(NSS)]
        _bords = [Border(left=Side(style='thin'), right=Side(style='thin')) for _ in range(NSS)]
        _aligns = [_Align(horizontal=['left', 'center', 'right'][i % 3],
                          vertical=['top', 'center', 'bottom'][i % 3],
                          wrap_text=bool(i % 2)) for i in range(NSS)]
        _fmts = (['0.00', '#,##0', '0.0%', 'yyyy-mm-dd', '0.0000'] * 8)[:NSS]
        sws.append(['列%d' % c for c in range(1, NSC + 1)])
        for r in range(2, NSR + 2):
            sws.append(['组%d' % (r % 3)] + ['v%d-%d' % (r, c) for c in range(2, NSC + 1)])
            for c in range(1, NSC + 1):
                cell = sws.cell(row=r, column=c)
                k = (r * 7 + c * 3) % NSS
                cell.font = _fonts[k]
                cell.fill = _fills[k]
                cell.border = _bords[k]
                cell.alignment = _aligns[k]
                cell.number_format = _fmts[k]
        swb.save(st_path)
        swb.close()
        chk = openpyxl.load_workbook(st_path)
        src_scale = (len(chk._fonts), len(chk._fills), len(chk._borders),
                     len(chk._alignments), len(chk._number_formats))
        chk.close()
        w(f'styled_source_scale={src_scale!s}')
        sout = os.path.join(d, 'out_styled')
        os.makedirs(sout)
        sok, smsg = SplitWorker(st_path, '数据', 1, 'files', sout, False, 2).run()
        w('styled_split_ok=%s' % sok)
        if not sok:
            w('styled_split_msg=%s' % smsg.replace('\n', ' | '))
        if not sok:
            raise AssertionError('样式繁多的大表拆分失败（IndexError 未修复）：%s' % smsg)
        sfiles = sorted(f for f in os.listdir(sout) if f.endswith('.xlsx'))
        w('styled_files=%s' % sfiles)
        for fn in sfiles:
            owb = openpyxl.load_workbook(os.path.join(sout, fn))
            ows = owb.active
            oscale = (len(owb._fonts), len(owb._fills), len(owb._borders),
                      len(owb._alignments), len(owb._number_formats))
            w(f'  {fn!s} -> {ows.max_row!s}行 x {ows.max_column!s}列 scale={oscale!s}')
            if ows.max_column != NSC:
                raise AssertionError('%s 列数错' % fn)
            if oscale != src_scale:
                raise AssertionError(f'{fn!s} 样式表规模与源表不一致 {oscale!s} != {src_scale!s}')
            owb.close()

        # ---- 【V4.4】新增回归用例：公式平移（共享 + 普通）与 "(空白)" 冲突 ----
        w('----------------------------------------')

        def _write_min_xlsx(dst_path, sheet_xml, sst_items):
            '''手工构造一个最小可用的 xlsx（openpyxl 不会写共享公式，只能手搓）。'''
            _ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   '</Types>')
            _rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                     '</Relationships>')
            _wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                   '<sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets></workbook>')
            _wbr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                    '</Relationships>')
            _sst = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                    'count="%d" uniqueCount="%d">' % (len(sst_items), len(sst_items))
                    + ''.join('<si><t>%s</t></si>' % _xml_escape(s) for s in sst_items)
                    + '</sst>')
            _sty = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
                    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
                    '<fill><patternFill patternType="gray125"/></fill></fills>'
                    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
                    '</styleSheet>')
            with zipfile.ZipFile(dst_path, 'w') as _z:
                _z.writestr('[Content_Types].xml', _ct)
                _z.writestr('_rels/.rels', _rels)
                _z.writestr('xl/workbook.xml', _wb)
                _z.writestr('xl/_rels/workbook.xml.rels', _wbr)
                _z.writestr('xl/worksheets/sheet1.xml', sheet_xml)
                _z.writestr('xl/sharedStrings.xml', _sst)
                _z.writestr('xl/styles.xml', _sty)

        # (a) 共享公式：master 放在「拆分列空白、整行被跳过」的第 2 行，
        #     依赖单元格分布在两组里 —— 同时钉住 _learn_shared_formulas 和平移锚点。
        sf = os.path.join(d, 'shared_src.xlsx')
        _sheet_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                      '<dimension ref="A1:C7"/><sheetData>'
                      '<row r="1" spans="1:3"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>'
                      '<row r="2" spans="1:3"><c r="B2"><v>10</v></c><c r="C2"><f t="shared" ref="C2:C7" si="0">B2*10</f></c></row>'
                      '<row r="3" spans="1:3"><c r="A3" t="s"><v>4</v></c><c r="B3"><v>20</v></c><c r="C3"><f t="shared" si="0"/><v>200</v></c></row>'
                      '<row r="4" spans="1:3"><c r="A4" t="s"><v>3</v></c><c r="B4"><v>30</v></c><c r="C4"><f t="shared" si="0"/><v>300</v></c></row>'
                      '<row r="5" spans="1:3"><c r="A5" t="s"><v>4</v></c><c r="B5"><v>40</v></c><c r="C5"><f t="shared" si="0"/><v>400</v></c></row>'
                      '<row r="6" spans="1:3"><c r="A6" t="s"><v>3</v></c><c r="B6"><v>50</v></c><c r="C6"><f t="shared" si="0"/><v>500</v></c></row>'
                      '<row r="7" spans="1:3"><c r="A7" t="s"><v>4</v></c><c r="B7"><v>60</v></c><c r="C7"><f t="shared" si="0"/><v>600</v></c></row>'
                      '</sheetData></worksheet>')
        _write_min_xlsx(sf, _sheet_xml, ['编号', '数量', '双倍', '甲', '乙'])
        sh_out = os.path.join(d, 'out_shared')
        os.makedirs(sh_out)
        ok, msg = SplitWorker(sf, '数据', 1, 'files', sh_out, False, 2).run()
        w('shared_split_ok=%s' % ok)
        if not ok:
            raise AssertionError('共享公式用例拆分失败：%s' % msg)
        got = {}
        for fn in sorted(f for f in os.listdir(sh_out) if f.endswith('.xlsx')):
            _wbv = openpyxl.load_workbook(os.path.join(sh_out, fn))
            _wsv = _wbv.active
            got[_wsv.title] = [_wsv.cell(row=rr, column=3).value for rr in range(2, _wsv.max_row + 1)]
            _wbv.close()
        w('shared_formulas=%s' % got)
        # 组「甲」= 源行 4/6 -> 新行 2/3；组「乙」= 源行 3/5/7 -> 新行 2/3/4。
        # 每个输出公式都必须引用**本行**的 B 列（master 行被跳过也要能展开）。
        if got.get('甲') != ['=B2*10', '=B3*10']:
            raise AssertionError('共享公式未按新行号平移：%r' % (got.get('甲'),))
        if got.get('乙') != ['=B2*10', '=B3*10', '=B4*10']:
            raise AssertionError('共享公式未按新行号平移：%r' % (got.get('乙'),))

        # (b) 普通公式（非共享）也要随行号平移，两条通道行为一致
        nf = os.path.join(d, 'norm_src.xlsx')
        nwb = openpyxl.Workbook()
        nws = nwb.active
        nws.title = '数据'
        nws.append(['k', 'v', '倍数'])
        for i in range(1, 5):
            nws.append([i, i * 100, '=B%d*2' % (i + 1)])
        nwb.save(nf)
        nwb.close()
        n_out = os.path.join(d, 'out_norm')
        os.makedirs(n_out)
        ok, msg = SplitWorker(nf, '数据', 1, 'files', n_out, False, 2).run()
        w('norm_split_ok=%s' % ok)
        if not ok:
            raise AssertionError('普通公式用例拆分失败：%s' % msg)
        for fn in sorted(f for f in os.listdir(n_out) if f.endswith('.xlsx')):
            _wbv = openpyxl.load_workbook(os.path.join(n_out, fn))
            _wsv = _wbv.active
            w('  norm %s -> C2=%r' % (fn, _wsv['C2'].value))
            if _wsv['C2'].value != '=B2*2':
                raise AssertionError('普通公式未按新行号平移：%r' % (_wsv['C2'].value,))
            _wbv.close()

        # (c) 取值恰好是 "(空白)" + 勾选「空白行也拆出」：两组必须合并、一行都不能丢
        bf = os.path.join(d, 'blank_src.xlsx')
        bwb = openpyxl.Workbook()
        bws = bwb.active
        bws.title = '数据'
        bws.append(['k', 'v'])
        bws.append(['甲', 1])
        bws.append(['(空白)', 2])
        bws.append([None, 3])
        bws.append(['甲', 4])
        bwb.save(bf)
        bwb.close()
        b_out = os.path.join(d, 'out_blank')
        os.makedirs(b_out)
        ok, msg = SplitWorker(bf, '数据', 1, 'files', b_out, True, 2).run()
        w('blank_split_ok=%s' % ok)
        if not ok:
            raise AssertionError('空白冲突用例拆分失败：%s' % msg)
        bfp = os.path.join(b_out, [f for f in os.listdir(b_out) if '(空白)' in f][0])
        _wbv = openpyxl.load_workbook(bfp)
        _wsv = _wbv.active
        w('blank_rows=%s' % (list(_wsv.iter_rows(values_only=True)),))
        if (_wsv.max_row != 3 or _wsv.cell(row=2, column=1).value != '(空白)'
                or _wsv.cell(row=3, column=1).value is not None):
            raise AssertionError('"（空白)"组丢行：max_row=%s' % _wsv.max_row)
        _wbv.close()

        w('RESULT=PASS')
    except Exception:
        w('RESULT=FAIL')
        w(traceback.format_exc())
    try:
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
    except Exception:
        pass


def main():
    for _name in ('stdout', 'stderr'):
        if getattr(sys, _name, None) is None:
            try:
                setattr(sys, _name, open(os.devnull, 'w', encoding='utf-8'))
            except Exception:
                continue
    if len(sys.argv) >= 2 and sys.argv[1] == '--selftest':
        run_selftest(sys.argv[2] if len(sys.argv) > 2 else 'v4_selftest.txt')
        return None
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass
    root = tk.Tk()
    set_window_icon(root)
    try:
        from ctypes import windll
        dpi = windll.user32.GetDpiForWindow(root.winfo_id)
        root.tk.call('tk', 'scaling', dpi / 72.0)
    except Exception:
        pass
    MainWindow(root)
    root.mainloop()
    return None


if __name__ == '__main__':
    main()
