# ============================================================
# 文件名称: RsoftCad.py
# 模块版本: Rsoft2020_Python3.12_1.0
# 模块功能: 用于自动生成 RSoft BPM 的 .ind 文件（建模+配置）
# 模块说明: 实现建模类 RsoftCad 及辅助结构类 vec3, taper 等
# 作者单位: 中南大学机电工程学院 郑煜教授课题组
# 撰写作者: 2022届硕士研究生 万强伟
# 联系方式: 1025459384@qq.com
# ============================================================

import os
import re

# ============================================================
# 辅助结构体定义部分（向量类型、枚举类、材料类型等）
# ============================================================

# === 三维向量类 ===
# 用于表示三维坐标、偏移或参考信息
class vec3:
    def __init__(self, x=None, y=None, z=None):
        self.x, self.y, self.z = x, y, z

# === 宽度变化类型类 ===
# 表示波导的渐变类型
class taper:
    linar = "TAPER_LINEAR"
    quadratic = "TAPER_QUADRATIC"
    exponential = "TAPER_EXPONENTIAL"

# === 弧形段参数类 ===
# 包含半径、起始角度、终止角度
class vec3_arc:
    def __init__(self, radius=None, iangle=None, fangle=None):
        self.radius, self.iangle, self.fangle = radius, iangle, fangle

# === 二维尺寸类（宽、高） ===
class vec2:
    def __init__(self, width, height):
        self.width, self.height = width, height

# === 材料类型枚举 ===
# 分类管理 Dielectrics / Metals / Semiconductors 等
class material_type:
    class Dielectrics:
        Air, ITO, LiNbO3_e, LiNbO3_o, PMMA, Si3N4, SiO2 =\
            (1, "Dielectrics"), (2, "Dielectrics"), (3, "Dielectrics"), (4, "Dielectrics"), (5, "Dielectrics"), (6, "Dielectrics"), (7, "Dielectrics")
    class Metals:
        Ag, Al, Au, Be, Cr, Cu, Ni, Pd, Pt, Ti, W = \
            (1, "Metals"), (2, "Metals"), (3, "Metals"), (4, "Metals"), (5, "Metals"), (6, "Metals"), (7, "Metals"), (8, "Metals"), (9, "Metals"), (10, "Metals"), (11, "Metals")
    class Semiconductors:
        AlAs, AlGaAs, AsPGa, GaAs, GaInAsP, GaN, GaP = \
            (1, "Semiconductors"), (2, "Semiconductors"), (3, "Semiconductors"), (4, "Semiconductors"), (5, "Semiconductors"), (6, "Semiconductors"), (7, "Semiconductors")
        Ge, InAs, InGaAs, InP, Si, SiGe, Si_amorphous = \
            (8, "Semiconductors"), (9, "Semiconductors"), (10, "Semiconductors"), (11, "Semiconductors"), (12, "Semiconductors"), (13, "Semiconductors"), (14, "Semiconductors")
    class Special:
        Graphene, GrapheneCX, GrapheneCY, GrapheneCZ, PEC = \
            (1, "Special"), (2, "Special"), (3, "Special"), (4, "Special"), (5, "Special")
    class TCAD:
        AlAs, AlGaAs, Aluminum, Copper, GaAs, Gas, Germanium, Gold = \
            (1, "TCAD"), (2, "TCAD"), (3, "TCAD"), (4, "TCAD"), (5, "TCAD"), (6, "TCAD"), (7, "TCAD"), (8, "TCAD")
        InAs, InGaAs, Nitride, Oxide, PolySilicon, Silicon, Silver, Tungsten = \
            (9, "TCAD"), (10, "TCAD"), (11, "TCAD"), (12, "TCAD"), (13, "TCAD"), (14, "TCAD"), (15, "TCAD"), (16, "TCAD")

# === 监视器类型枚举 ===
class monitor_type:
    File_Power = "MONITOR_FILE_POWER"
    File_Phase = "MONITOR_FILE_PHASE"
    Fiber_Mode_Power = "MONITOR_WGMODE_POWER"
    Fiber_Mode_Phase = "MONITOR_WGMODE_PHASE"
    Gaussian_Power = "MONITOR_GAUSS_POWER"
    Gaussian_Phase = "MONITOR_GAUSS_PHASE"
    Launch_Power = "MONITOR_LAUNCH_POWER"
    Launch_Phase = "MONITOR_LAUNCH_PHASE"
    Partial_Power = "MONITOR_WG_POWER"
    Total_Power = "MONITOR_TOTAL_POWER"
    Effective_Index = "MONITOR_FIELD_NEFF"
    Field_1_e_Width = "MONITOR_FIELD_WIDTH"
    Field_1_e_Height = "MONITOR_FIELD_HEIGHT"
    Effective_Area = "MONITOR_FIELD_AEFF"

# === 光源类型枚举 ===
class launch_type:
    File = "LAUNCH_FILE"
    Computed_Mode = "LAUNCH_COMPMODE"
    Fiber_Mode = "LAUNCH_WGMODE"
    Gaussian = "LAUNCH_GAUSSIAN"
    Rectangle = "LAUNCH_RECTANGLE"
    MultiMode = "LAUNCH_MULTIMODE"
    Plane_Wave = "LAUNCH_PLANEWAVE"


# ============================================================
# 类名: RsoftCad
# 功能: 构建RSoft BPM仿真所需的.ind结构文件
# 提供接口: 参数定义、材料插入、段结构生成、路径构建、监视器添加、光源设置等
# ============================================================

class RsoftCad:
    # ------------------------------------------------------------
    # 构造函数: __init__
    # 功能: 创建并初始化一个新的 RSoft .ind 文件
    # 参数:
    #   file_path               - 文件输出目录
    #   file_name               - 文件名（不含扩展名）
    #   dimension               - 维度（2 或 3）
    #   free_space_wavelength   - 自由空间波长
    #   background_material     - 背景介质名称（如 SiO2）
    #   Delta                   - 相对折射率差（用于计算绝对折射率差 delta）
    #   width                   - 初始结构宽度（默认高度与宽度相等）
    # ------------------------------------------------------------
    def __init__(self, file_path=str, file_name=str, dimension=int, free_space_wavelength=float, background_material=str, Delta=float, width=float):
        if not os.path.exists(file_path):
            os.makedirs(file_path)  # 若目录不存在则递归创建
        self.file = file_path + "\\" + file_name + ".ind"  # 构造完整路径
        open(self.file, "w").close()  # 创建空文件
        print(f"{self.file}创建成功")
        self.Rsoftfile = open(self.file, "r+")  # 打开文件用于读写

        # 维度合法性检查
        if dimension not in [2, 3]:
            raise ValueError("dimension must be either 2 or 3")

        # 写入全局设置
        self.Rsoftfile.write(f"dimension = {dimension}\n")
        self.Rsoftfile.write(f"wave = {free_space_wavelength}\n")
        self.Rsoftfile.write(f"free_space_wavelength = wave\n")
        self.Rsoftfile.write(f"background_material = {background_material}\n")
        self.Rsoftfile.write(f"background_alpha = nimag($background_material)\n")
        self.Rsoftfile.write(f"background_index = nreal($background_material)\n")
        self.Rsoftfile.write(f"Delta = {Delta}\n")
        self.Rsoftfile.write("delta = (1/(sqrt(1-2*Delta))-1)*background_index\n")
        self.Rsoftfile.write(f"width = {width}\n")
        self.Rsoftfile.write("height = width\n")
        self.Rsoftfile.write("structure = STRUCT_CHANNEL\n")

        # 插入分段 marker 以便后续插入
        self.symbol_marker_line   = self.Rsoftfile.tell(); self.Rsoftfile.write("\n\n")
        self.material_marker_line = self.Rsoftfile.tell(); self.Rsoftfile.write("\n\n")
        self.segment_marker_line  = self.Rsoftfile.tell(); self.Rsoftfile.write("\n\n")
        self.pathway_marker_line  = self.Rsoftfile.tell(); self.Rsoftfile.write("\n\n")
        self.monitor_marker_line  = self.Rsoftfile.tell(); self.Rsoftfile.write("\n\n")
        self.launch_marker_line   = self.Rsoftfile.tell(); self.Rsoftfile.write("\n\n")

        # 初始化计数器
        self.seg_num      = 1
        self.material_num = 1
        self.pathway_num  = 1
        self.monitor_num  = 1
        self.launch_num   = 1

    # ------------------------------------------------------------
    # 方法名: set_symbol
    # 功能: 在 .ind 文件中插入或更新一个符号变量（例如结构长度、间距等）
    # 参数:
    #   symbol - 字符串类型，变量名（如 'Gap'）
    #   value  - 数值或表达式类型，变量值（如 1.5 或 'width+2'）
    # 返回: 无
    # 实现说明:
    #   - 插入内容到 #symbol_marker 区域（实际是文件中的一个偏移位置）
    #   - 为了插入该行，需要先读取其后所有内容（缓存），再写入变量，最后再写回剩余内容
    #   - 每次插入后，更新所有下游 marker 的位置偏移值，保持文件结构正确
    # ------------------------------------------------------------
    def set_symbol(self, symbol, value):
        # === 定位到 symbol_marker 行 ===
        # seek(pos): 设置文件指针位置到 symbol_marker_line
        self.Rsoftfile.seek(self.symbol_marker_line)

        # === 读取 marker 之后的所有内容（为了后续回填）===
        # read(): 从当前指针开始读取文件剩余内容，保存至变量
        remaining_content = self.Rsoftfile.read()

        # === 重新将指针移回 symbol_marker 位置 ===
        self.Rsoftfile.seek(self.symbol_marker_line)

        # === 写入新的变量定义行 ===
        # 写入语法: 变量名 = 变量值，例如: Lin = 500
        # f-string 格式化支持字符串表达式或数值
        self.Rsoftfile.write(f"{symbol} = {value}\n")

        # === 计算刚刚写入的行的长度（偏移）===
        # tell(): 获取当前文件指针位置
        marker_offset = self.Rsoftfile.tell() - self.symbol_marker_line

        # === 更新所有下游 marker 的偏移值 ===
        # 防止原 marker 指针错位，确保写入区域正确
        self.symbol_marker_line   += marker_offset
        self.material_marker_line += marker_offset
        self.segment_marker_line  += marker_offset
        self.pathway_marker_line  += marker_offset
        self.monitor_marker_line  += marker_offset
        self.launch_marker_line   += marker_offset

        # === 恢复写入剩余内容（保持原有结构）===
        # 将原先 marker 后的内容继续写入文件末尾
        self.Rsoftfile.write(remaining_content)


    # ------------------------------------------------------------
    # 方法名: add_material
    # 功能: 从指定材料库文件中提取材料描述内容，并插入至当前 .ind 文件的 material 区域
    # 参数:
    #   material_type - 二元元组类型，如 material_type.Dielectrics.SiO2
    #                 - 格式为 (material_index, material_class)，分别表示材料在该类中的序号 和 类名字符串
    # 返回:
    #   新插入材料的编号（从1开始）
    # 实现说明:
    #   - 自动查找 RsoftMaterial 文件夹下的材料库（.mlb）
    #   - 解析出对应编号的材料描述语句块
    #   - 插入至 #material_marker 标记位置，并更新所有下游 marker 的偏移
    # ------------------------------------------------------------
    def add_material(self, material_type):
        # === 解包参数 material_type 为 序号 和 类别字符串（如 SiO2 → (7, "Dielectrics")）===
        material_sequence, material_class = material_type

        # === 获取当前工作目录 ===
        current_path = os.getcwd()

        # === 构造材料库文件路径，例如: RsoftMaterial\Dielectrics.mlb ===
        material_file = current_path + "\\RsoftMaterial\\" + material_class + ".mlb"

        # === 初始化变量，准备读取目标材料段落 ===
        material_start_lines = []  # 存放所有材料块起始行号（匹配 'material n'）
        material_end_lines = []    # 存放所有材料块终止行号（匹配 'end material'）
        content = []               # 存放最终提取出的材料描述内容

        # === 打开材料库文件，读取内容行列表 ===
        with open(material_file, "r") as file:
            lines = file.readlines()
            for line_number, line in enumerate(lines, start=1):
                # 正则匹配 'material 数字' 行 → 认为是一个材料块起始
                if re.search(r'material \d', line, re.IGNORECASE):
                    material_start_lines.append(line_number)
                # 匹配 'end material' 行 → 表示该材料块结束
                if re.search(r'end material', line, re.IGNORECASE):
                    material_end_lines.append(line_number)

            # === 根据传入的材料序号获取材料内容块（减1是因为 Python 索引从0开始）===
            content = lines[material_start_lines[material_sequence - 1]:material_end_lines[material_sequence - 1] - 1]

        # === 定位到 #material_marker 位置，准备写入 ===
        self.Rsoftfile.seek(self.material_marker_line)
        remaining_content = self.Rsoftfile.read()
        self.Rsoftfile.seek(self.material_marker_line)

        # === 写入新材料内容块 ===
        self.Rsoftfile.write(f"material {self.material_num}\n")  # 写入起始标记
        for line in content:
            self.Rsoftfile.write(line)                          # 写入每行材料描述
        self.Rsoftfile.write(f"end material\n\n")               # 写入结束标记

        # === 更新材料计数器 ===
        self.material_num += 1

        # === 计算插入段落的长度变化，更新 marker 偏移 ===
        marker_offset = self.Rsoftfile.tell() - self.material_marker_line
        self.material_marker_line += marker_offset
        self.segment_marker_line  += marker_offset
        self.pathway_marker_line  += marker_offset
        self.monitor_marker_line  += marker_offset
        self.launch_marker_line   += marker_offset

        # === 写回原始的后续文件内容，保持整体结构 ===
        self.Rsoftfile.write(remaining_content)

        # === 返回刚刚插入的材料编号（从1开始）===
        return self.material_num - 1


    # segment通用方法：按照Rsoft语法规则写入 begin/end.x, begin/end.y, begin/end.z 坐标
    def write_segment(self, prefix, axis, rel_type, pos_offset, rel_vertex, rel_num):
        if getattr(rel_type, axis) == "None":
            self.Rsoftfile.write(f"\t{prefix}.{axis} = {getattr(pos_offset, axis)}\n")
        elif getattr(rel_type, axis) == "Offset":
            self.Rsoftfile.write(f"\t{prefix}.{axis} = {getattr(pos_offset, axis)} rel {getattr(rel_vertex, axis)} segment {getattr(rel_num, axis)}\n")
        elif getattr(rel_type, axis) == "Angle":
            self.Rsoftfile.write(f"\t{prefix}.{axis} = {getattr(pos_offset, axis)} deg rel {getattr(rel_vertex, axis)} segment {getattr(rel_num, axis)}\n")
        else:
            raise ValueError(f"Invalid value for rel_type. rel_type Must be 'None', 'Offset', or 'Angle'.") # 确保 rel_type 的值只能是 "None", "Offset", 或 "Angle"

    # ------------------------------------------------------------
    # 方法名: add_segment
    # 功能: 插入一个直线波导段（segment）结构定义
    # 参数说明:
    #   rel_type        - vec3对象, 起点参考类型 ('None'/'Offset'/'Angle')
    #   pos_offset      - vec3对象, 起点偏移量
    #   rel_num         - vec3对象, 起点参照段编号
    #   rel_vertex      - vec3对象, 起点为参照段的 'begin' 或 'end'
    #   rel_type_end    - vec3对象, 终点参考类型
    #   pos_offset_end  - vec3对象, 终点偏移量
    #   rel_num_end     - vec3对象, 终点参照段编号
    #   rel_vertex_end  - vec3对象, 终点为参照段的 'begin' 或 'end'
    #   dimensions      - vec2对象, 起点尺寸 (width, height)
    #   dimensions_end  - vec2对象, 终点尺寸 (width, height)
    #   width_taper     - 渐变类型 (taper.linar / taper.quadratic / taper.exponential)
    # 返回: 当前段编号（int）
    # ------------------------------------------------------------
    def add_segment(self, rel_type=vec3, pos_offset=vec3, rel_num=vec3, rel_vertex=vec3,
                    rel_type_end=vec3, pos_offset_end=vec3, rel_num_end=vec3, rel_vertex_end=vec3,
                    dimensions=vec2, dimensions_end=vec2, width_taper=taper):

        # 检查 vertex 合法性，必须为 'begin' 或 'end'
        for axis in ['x', 'y', 'z']:
            if getattr(rel_vertex, axis) not in ["begin", "end"]:
                raise ValueError(f"Invalid value for rel_vertex: {rel_vertex}. Must be 'begin' or 'end'.")
            if getattr(rel_vertex_end, axis) not in ["begin", "end"]:
                raise ValueError(f"Invalid value for rel_vertex_end: {rel_vertex_end}. Must be 'begin' or 'end'.")

        # 定位到 segment_marker 行，并读取之后所有内容备用
        self.Rsoftfile.seek(self.segment_marker_line)
        remaining_content = self.Rsoftfile.read()
        self.Rsoftfile.seek(self.segment_marker_line)

        # 写入 segment 块头
        self.Rsoftfile.write(f"segment {self.seg_num}\n")
        self.Rsoftfile.write(f"\twidth_taper = {width_taper}\n")

        # 写入起点位置信息
        for axis in ['x', 'y', 'z']:
            self.write_segment("begin", axis, rel_type, pos_offset, rel_vertex, rel_num)
        self.Rsoftfile.write(f"\tbegin.width = {dimensions.width}\n")
        self.Rsoftfile.write(f"\tbegin.height = {dimensions.height}\n")

        # 写入终点位置信息
        for axis in ['x', 'y', 'z']:
            self.write_segment("end", axis, rel_type_end, pos_offset_end, rel_vertex_end, rel_num_end)
        self.Rsoftfile.write(f"\tend.width = {dimensions_end.width}\n")
        self.Rsoftfile.write(f"\tend.height = {dimensions_end.height}\n")
        self.Rsoftfile.write("end segment\n\n")

        # 更新 marker 偏移
        self.seg_num += 1
        marker_offset = self.Rsoftfile.tell() - self.segment_marker_line
        self.segment_marker_line += marker_offset
        self.pathway_marker_line += marker_offset
        self.monitor_marker_line += marker_offset
        self.launch_marker_line += marker_offset

        self.Rsoftfile.write(remaining_content)
        return self.seg_num - 1


    # ------------------------------------------------------------
    # 方法名: add_arc
    # 功能: 插入一个弧形波导段（segment），常用于分支或弯道
    # 参数说明:
    #   rel_type       - 起点参考类型 (vec3)
    #   pos_offset     - 起点偏移量 (vec3)
    #   rel_num        - 起点参考段编号 (vec3)
    #   rel_vertex     - 起点参考点位 (vec3: begin/end)
    #   arcinfo        - 弧形参数（半径、起始角、终止角）(vec3_arc)
    #   dimensions     - 起点尺寸 (vec2)
    #   dimensions_end - 终点尺寸 (vec2)
    #   width_taper    - 渐变类型 (taper)
    # 返回: 当前弧段编号（int）
    # ------------------------------------------------------------
    def add_arc(self, rel_type=vec3, pos_offset=vec3, rel_num=vec3, rel_vertex=vec3,
                arcinfo=vec3_arc, dimensions=vec2, dimensions_end=vec2, width_taper=taper):
        self.Rsoftfile.seek(self.segment_marker_line)
        remaining_content = self.Rsoftfile.read()
        self.Rsoftfile.seek(self.segment_marker_line)

        self.Rsoftfile.write(f"segment {self.seg_num}\n")
        self.Rsoftfile.write(f"\twidth_taper = {width_taper}\n")
        self.Rsoftfile.write(f"\tposition_taper = TAPER_ARC\n")  # 表示该段为弯曲段
        self.Rsoftfile.write(f"\tarc_type = ARC_FREE\n")          # 使用自由角度
        self.Rsoftfile.write(f"\tarc_radius = {arcinfo.radius}\n")
        self.Rsoftfile.write(f"\tarc_iangle = {arcinfo.iangle}\n")
        self.Rsoftfile.write(f"\tarc_fangle = {arcinfo.fangle}\n")

        for axis in ['x', 'y', 'z']:
            self.write_segment("begin", axis, rel_type, pos_offset, rel_vertex, rel_num)
        self.Rsoftfile.write(f"\tbegin.width = {dimensions.width}\n")
        self.Rsoftfile.write(f"\tbegin.height = {dimensions.height}\n")
        self.Rsoftfile.write(f"\tend.width = {dimensions_end.width}\n")
        self.Rsoftfile.write(f"\tend.height = {dimensions_end.height}\n")
        self.Rsoftfile.write("end segment\n\n")

        self.seg_num += 1
        marker_offset = self.Rsoftfile.tell() - self.segment_marker_line
        self.segment_marker_line += marker_offset
        self.pathway_marker_line += marker_offset
        self.monitor_marker_line += marker_offset
        self.launch_marker_line += marker_offset

        self.Rsoftfile.write(remaining_content)
        return self.seg_num - 1


    # ------------------------------------------------------------
    # 方法名: add_pathway
    # 功能: 添加一条完整的光路，由一系列段号构成
    # 参数:
    #   pathlist - 包含多个段编号（segment index）的列表
    # 返回:
    #   新路径编号
    # ------------------------------------------------------------
    def add_pathway(self, pathlist):
        self.Rsoftfile.seek(self.pathway_marker_line)
        remaining_content = self.Rsoftfile.read()
        self.Rsoftfile.seek(self.pathway_marker_line)

        self.Rsoftfile.write(f"pathway {self.pathway_num}\n")
        for i in pathlist:
            self.Rsoftfile.write(f"\t{i}\n")
        self.Rsoftfile.write("end pathway\n\n")

        self.pathway_num += 1
        marker_offset = self.Rsoftfile.tell() - self.pathway_marker_line
        self.pathway_marker_line += marker_offset
        self.monitor_marker_line += marker_offset
        self.launch_marker_line += marker_offset

        self.Rsoftfile.write(remaining_content)
        return self.pathway_num - 1


    # ------------------------------------------------------------
    # 方法名: add_monitor
    # 功能: 添加一个光路监视器，监控功率、相位等指标
    # 参数:
    #   pathway      - 监视器所属路径编号
    #   monitor_type - 监视器类型（枚举）
    # 返回:
    #   监视器编号
    # ------------------------------------------------------------
    def add_monitor(self, pathway, monitor_type):
        self.Rsoftfile.seek(self.monitor_marker_line)
        remaining_content = self.Rsoftfile.read()
        self.Rsoftfile.seek(self.monitor_marker_line)

        self.Rsoftfile.write(f"monitor {self.monitor_num}\n")
        self.Rsoftfile.write(f"\tpathway = {pathway}\n")
        self.Rsoftfile.write(f"\tmonitor_type = {monitor_type}\n")
        self.Rsoftfile.write(f"\tmonitor_tilt = 1\n")  # 默认为倾斜方式
        self.Rsoftfile.write("end monitor\n\n")

        self.monitor_num += 1
        marker_offset = self.Rsoftfile.tell() - self.monitor_marker_line
        self.monitor_marker_line += marker_offset
        self.launch_marker_line += marker_offset

        self.Rsoftfile.write(remaining_content)
        return self.monitor_num - 1


    # ------------------------------------------------------------
    # 方法名: add_launch
    # 功能: 添加一个光源发射器 (launch_field)，指定路径及发射方式
    # 参数:
    #   pathway     - 发射器路径编号
    #   launch_type - 发射类型，如 Computed_Mode, Gaussian 等
    # 返回:
    #   光源编号
    # ------------------------------------------------------------
    def add_launch(self, pathway, launch_type):
        self.Rsoftfile.seek(self.launch_marker_line)
        remaining_content = self.Rsoftfile.read()
        self.Rsoftfile.seek(self.launch_marker_line)

        self.Rsoftfile.write(f"launch_field {self.launch_num}\n")
        self.Rsoftfile.write(f"\tlaunch_pathway = {pathway}\n")
        self.Rsoftfile.write(f"\tlaunch_type = {launch_type}\n")
        self.Rsoftfile.write("end launch_field\n\n")

        # 第一个 launch 默认写入 symbol 表
        if self.launch_num == 1:
            self.set_symbol('launch_type', launch_type)

        marker_offset = self.Rsoftfile.tell() - self.launch_marker_line
        self.launch_marker_line += marker_offset
        self.Rsoftfile.write(remaining_content)

        self.launch_num += 1
        return self.launch_num - 1