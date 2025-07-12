# ============================================================
# 文件名称: RsoftData.py
# 模块功能: 用于读取并分析 RSoft 仿真（.mon）结果
# 功能概述:
#   - 自动识别仿真或扫描（Sim / Scan）结果
#   - 计算 IL / EL / UL / WDL 等性能指标矩阵
#   - 输出表格至 txt 文件，生成性能图像 PNG 文件
# 依赖模块: os, glob, re, math, matplotlib, tabulate
# 作者单位: 中南大学机电工程学院 郑煜教授课题组
# 撰写作者: 2022届硕士研究生 万强伟
# 联系方式: 1025459384@qq.com
# ============================================================

import os
import glob
import re
from math import *
from tabulate import tabulate
import matplotlib.pyplot as plt

class RsoftData:
    # ------------------------------------------------------------
    # 构造函数: __init__
    # 功能: 初始化数据对象，识别结果类型、提取性能指标、输出表格与图像
    # 参数:
    #   file_path - 仿真文件（.mon文件）所在路径，通常为仿真目录或 Scan 子目录
    # ------------------------------------------------------------
    def __init__(self, file_path=str):
        self.file_path = file_path

        # === 查找所有 .mon 文件（RSoft仿真结果） ===
        mon_path_list = glob.glob(os.path.join(self.file_path, "*.mon"))

        if len(mon_path_list) == 1:
            # 仅有一个 .mon 文件 → 单次仿真
            self.mon_path_matrix = [mon_path_list]
            self.rows, self.cols = 1, 1
        else:
            # 多个 .mon 文件 → 扫描仿真（双参数Scan）
            value1, value2 = [], []
            for mon_path in mon_path_list:
                mon_name = os.path.basename(mon_path)
                # 匹配形如 Lta(200.0)_wave(1.55).mon 的文件名结构（可以允许有负号）
                pattern = r'(\w+)\((-?[\d.]+)\)_(\w+)\((-?[\d.]+)\)'
                match = re.search(pattern, mon_name)
                if match:
                    self.symbol1 = match.group(1)            # 参数1名称，例如 Lta
                    value1.append(float(match.group(2)))     # 参数1数值
                    self.symbol2 = match.group(3)            # 参数2名称，例如 wave
                    value2.append(float(match.group(4)))     # 参数2数值
                else:
                    print("未找到匹配项")

            # 去重保序（保留原扫描顺序）
            self.unique_value1 = list(dict.fromkeys(value1))
            self.unique_value2 = list(dict.fromkeys(value2))

            # 划分为矩阵：行是 unique_value1，列是 unique_value2
            self.rows, self.cols = len(self.unique_value1), len(self.unique_value2)
            self.mon_path_matrix = [
                mon_path_list[i * self.cols:(i + 1) * self.cols]
                for i in range(self.rows)
            ]

        # === 创建输出文件：file_path_result.txt ===
        self.result_path = self.file_path + "_result.txt"
        open(self.result_path, "w").close()
        print(f"{self.result_path} 创建成功")
        self.resultfile = open(self.result_path, "r+")

        # === 性能矩阵["output", "IL", "EL", "UL"]计算并写入 ===
        matrix_names = ["output", "IL", "EL", "UL"]
        for name in matrix_names:
            getattr(self, name)()  # 调用 self.output(), self.IL(), ...
            self.print_matrix(getattr(self, f"{name}_matrix"), name)

        # === 性能矩阵"WDL"计算并写入 ===
        if len(mon_path_list) > 1:
            self.WDL()
            self.print_matrix(self.WDL_matrix, "WDL")

        # === 性能矩阵"ILmax_n"计算并写入 ===
        self.ILmax_n()
        self.print_matrix(self.ILmax_n_matrix, "ILmax_n")

        # === 性能矩阵["ILmax", "ELmax", "ULmax", "WDLmax", "mean"]计算并写入 ===
        if len(mon_path_list) > 1:
            matrix_names = ["ILmax", "ELmax", "ULmax", "WDLmax", "mean"]
            for name in matrix_names:
                getattr(self, name)()
                self.print_matrix(getattr(self, f"{name}_matrix"), name)

            # 找出最优点并写入
            self.resultfile.write(f"{self.symbol1}={self.min_symbol},min_mean={self.min_mean[0]}\n")
            # 作图并保存
            self.plot_all()

    # === 返回最小值供optimize输出 ===
    def get_min_symbol(self):
        return self.min_symbol

    # === 打印性能矩阵至 txt 文件（格式化输出）===
    # 函数名: print_matrix
    # 功能:
    #   - 将计算得到的性能矩阵输出至 result 文件
    #   - 根据矩阵种类自动生成表头
    # 参数:
    #   matrix       : 性能矩阵（二维列表）
    #   matrix_name  : 矩阵名称（如 'IL', 'EL', 'UL', 'WDL', 'mean' 等）
    # 返回: 无
    def print_matrix(self, matrix, matrix_name):
        # 写入矩阵名
        self.resultfile.write(f"{matrix_name}:\n")
        rows, cols = len(matrix), len(matrix[0])

        # 单点仿真：直接写入一个数
        if rows == 1 and cols == 1:
            self.resultfile.write(f"{matrix[0][0]}\n\n")
        else:
            table_data = []

            # 构建表头（根据矩阵类型判断）
            if matrix_name == "WDL":
                header = [f"{self.symbol1}/n_out"] + list(range(1, cols + 1))
            elif matrix_name in {"ILmax", "ELmax", "WDLmax", "ULmax", "mean"}:
                header = [f"{self.symbol1}"] + [f"{matrix_name}"]
            else:
                header = [f"{self.symbol1}/{self.symbol2}"] + self.unique_value2
            table_data.append(header)

            # 写入每一行数据（首列为 symbol1 参数值）
            for i in range(rows):
                row = [self.unique_value1[i]] + matrix[i]
                table_data.append(row)

            # 使用 tabulate 美化输出为纯文本表格
            formatted_table = tabulate(table_data, tablefmt="plain")
            self.resultfile.write(formatted_table + "\n\n")


    # === 提取每个 .mon 文件的输出功率值 ===
    # 函数名: output
    # 功能: 提取每个仿真结果文件的最后一行输出值（输出功率），组成二维矩阵
    # 参数: 无
    # 返回: 无（结果存储在 self.output_matrix）
    def output(self):
        # 初始化 output_matrix 为 rows×cols 空矩阵
        self.output_matrix = [[None for _ in range(self.cols)] for _ in range(self.rows)]

        for i in range(self.rows):
            for j in range(self.cols):
                mon_path = self.mon_path_matrix[i][j]
                with open(mon_path, "r") as file:
                    # 读取文件最后一行
                    last_line = None
                    for line in file:
                        last_line = line.strip()

                # 拆分成数字，忽略第一个数字（仿真位置）
                values = last_line.split()
                output = [float(value) for value in values[1:]]

                # 记录输出端口数
                self.n_out = len(output)
                self.output_matrix[i][j] = output


    # === 计算插入损耗 IL（-10log(P)) ===
    # 函数名: IL
    # 功能: 计算每个输出端口的插入损耗，单位 dB，按端口展开
    # 参数: 无
    # 返回: 无（结果存储在 self.IL_matrix）
    def IL(self):
        self.IL_matrix = [[None for _ in range(self.cols)] for _ in range(self.rows)]

        for i in range(self.rows):
            for j in range(self.cols):
                IL = [None for _ in range(self.n_out)]
                for n in range(self.n_out):
                    # IL = -10 * log10(Pout)
                    IL[n] = round(-10 * log10(self.output_matrix[i][j][n]), 4)
                self.IL_matrix[i][j] = IL


    # === 计算每点仿真的最大插入损耗 ILmax_n（每个仿真点多个输出端口中的最大值）===
    # 函数名: ILmax_n
    # 功能: 对于每个仿真点，计算多个输出端口中插入损耗 IL 的最大值
    # 参数: 无
    # 返回: 无（结果存入 self.ILmax_n_matrix）
    def ILmax_n(self):
        self.ILmax_n_matrix = [[None for _ in range(self.cols)] for _ in range(self.rows)]
        for i in range(self.rows):
            for j in range(self.cols):
                # 提取当前点所有输出端口的 IL，计算其最大值
                ILmax_n = round(max(self.IL_matrix[i][j]), 4)
                self.ILmax_n_matrix[i][j] = ILmax_n


    # === 计算每行（wave）下的最大 ILmax 值 ===
    # 函数名: ILmax
    # 功能: 在每一行（固定参数1值）中，取 ILmax 的最大值
    # 参数: 无
    # 返回: 无（结果存入 self.ILmax_matrix）
    def ILmax(self):
        self.ILmax_matrix = [[None for _ in range(1)] for _ in range(self.rows)]
        for i in range(self.rows):
            # 每行中最大 ILmax_n 值
            ILmax = round(max(self.ILmax_n_matrix[i]), 4)
            self.ILmax_matrix[i][0] = ILmax

    # === 计算总损耗 EL（-10log(ΣP)) ===
    # 函数名: EL
    # 功能: 将每个输出端口的功率相加，计算总能量损耗
    # 参数: 无
    # 返回: 无（结果存储在 self.EL_matrix）
    def EL(self):
        self.EL_matrix = [[None for _ in range(self.cols)] for _ in range(self.rows)]

        for i in range(self.rows):
            for j in range(self.cols):
                # EL = -10 * log10(ΣPout)
                self.EL_matrix[i][j] = round(-10 * log10(sum(self.output_matrix[i][j])), 4)


    # === 计算每行的最大总损耗 ELmax ===
    # 函数名: ELmax
    # 功能:
    #   - 对于每一个 symbol1 参数（即每一行），提取所有波长下的 EL 值
    #   - 找出该参数下所有波长中的最大 Excess Loss 值
    # 参数: 无
    # 返回: 无（结果存入 self.ELmax_matrix，尺寸为 rows × 1）
    def ELmax(self):
        # 初始化 ELmax_matrix：行为参数 symbol1，列为 1
        self.ELmax_matrix = [[None for _ in range(1)] for _ in range(self.rows)]

        for i in range(self.rows):
            for j in range(1):
                # 提取该行所有波长下的 EL 值（如 EL[i][0], EL[i][1], ...）
                ELmax = round(max(self.EL_matrix[i]), 4)  # 保留四位有效数字
                self.ELmax_matrix[i][j] = ELmax



    # === 计算不均匀性损耗 UL（-10log(min/max)) ===
    # 函数名: UL
    # 功能: 评估各输出端功率分布是否均匀
    # 参数: 无
    # 返回: 无（结果存储在 self.UL_matrix）
    def UL(self):
        self.UL_matrix = [[None for _ in range(self.cols)] for _ in range(self.rows)]

        for i in range(self.rows):
            for j in range(self.cols):
                # UL = -10 * log10(min(P) / max(P))
                Pmin = min(self.output_matrix[i][j])
                Pmax = max(self.output_matrix[i][j])
                self.UL_matrix[i][j] = round(-10 * log10(Pmin / Pmax), 4)


    # === 计算每行（一个参数设置）下的最大 UL 值 ===
    # 函数名: ULmax
    # 功能: 统计每个参数设置下（每行）UL 指标的最大值
    # 参数: 无
    # 返回: 无（结果存入 self.ULmax_matrix）
    def ULmax(self):
        self.ULmax_matrix = [[None for _ in range(1)] for _ in range(self.rows)]
        for i in range(self.rows):
            ULmax = round(max(self.UL_matrix[i]), 4)
            self.ULmax_matrix[i][0] = ULmax


    # === 计算波长依赖损耗 WDL（每个输出端口）===
    # 函数名: WDL
    # 功能: 对每个输出端口 n_out，计算其在不同波长下的损耗变化
    # 参数: 无
    # 返回: 无（结果存入 self.WDL_matrix）
    def WDL(self):
        self.WDL_matrix = [[None for _ in range(self.n_out)] for _ in range(self.rows)]
        n_output = [None for _ in range(self.cols)]

        for i in range(self.rows):
            for n in range(self.n_out):
                # 提取每列的相同端口输出值（不同波长）
                for j in range(self.cols):
                    n_output[j] = self.output_matrix[i][j][n]
                # WDL = -10 * log10(Pmin / Pmax)
                self.WDL_matrix[i][n] = round(-10 * log10(min(n_output) / max(n_output)), 4)


    # === 计算每行的最大波长相关损耗 WDLmax ===
    # 函数名: WDLmax
    # 功能: 提取每个参数设置下所有输出端口的最大 WDL 值
    # 参数: 无
    # 返回: 无（结果存入 self.WDLmax_matrix）
    def WDLmax(self):
        self.WDLmax_matrix = [[None for _ in range(1)] for _ in range(self.rows)]
        for i in range(self.rows):
            WDLmax = round(max(self.WDL_matrix[i]), 4)
            self.WDLmax_matrix[i][0] = WDLmax


    # === 计算平均性能指标 mean = (ELmax + ULmax + WDLmax) / 3 ===
    # 函数名: mean
    # 功能:
    #   - 综合评估某参数设置下的性能
    #   - 用于确定最优参数（如最小 mean）
    # 参数: 无
    # 返回: 无（结果存入 self.mean_matrix）
    def mean(self):
        self.mean_matrix = [[None for _ in range(1)] for _ in range(self.rows)]
        for i in range(self.rows):
            # 取三项性能指标平均值
            mean = round(
                (self.ELmax_matrix[i][0] + self.WDLmax_matrix[i][0] + self.ULmax_matrix[i][0]) / 3, 4
            )
            self.mean_matrix[i][0] = mean

        # 找出最小 mean 值
        self.min_mean = min(self.mean_matrix)
        min_index = self.mean_matrix.index(self.min_mean)
        # 记录对应 symbol1 参数值（用于优化结果输出）
        self.min_symbol = self.unique_value1[min_index]



    # === 绘图：性能 vs 扫描参数 ===
    # 函数名: plot_symbol_vs_wave
    # 功能: 对 EL / UL / ILmax_n / WDL 等性能在不同参数或波长下的趋势进行可视化
    # 参数:
    #   ax          : matplotlib 的 subplot 轴对象
    #   matrixname  : 指标名（如 "EL", "WDL", "ILmax_n"）
    #   matrix      : 对应矩阵（二维）
    # 返回: 无
    def plot_symbol_vs_wave(self, ax, matrixname, matrix):
        if matrixname == "WDL":
            # WDL 特例：纵轴是不同输出端口
            for i in range(len(matrix[0])):
                y_values = [row[i] for row in matrix]
                ax.plot(self.unique_value1, y_values, marker='o', linestyle='-', label=f"n_out {i+1}")
            ax.set_title(f"{matrixname} vs {self.symbol1} for Different n_out")
            ax.legend(title="n_out")
        else:
            # 普通指标：纵轴是不同波长下的指标值
            for i, wave in enumerate(self.unique_value2):
                y_values = [row[i] for row in matrix]
                ax.plot(self.unique_value1, y_values, marker='o', linestyle='-', label=f"{wave}")
            ax.set_title(f"{matrixname} vs {self.symbol1} for Different Waves")
            ax.legend(title="Wavelength (µm)")

        ax.set_xlabel(self.symbol1)
        ax.set_ylabel(f"{matrixname} (dB)")
        ax.legend(loc="upper right")
        ax.grid(True)


    # === 绘图：绘制 ILmax 相对于 symbol1 的变化曲线 ===
    # 函数名: plot_ILmax
    # 功能:
    #   - 在指定子图轴 ax 上绘制 ILmax 曲线
    #   - 横轴为参数 symbol1（如 Lta），纵轴为 ILmax(dB)
    # 参数:
    #   ax - matplotlib 子图对象（axes[i,j]）
    # 返回:
    #   无（函数仅作图，不返回数据）
    def plot_ILmax(self, ax):
        # 绘制线图：x 为参数值，y 为 ILmax（蓝色线，圆点）
        ax.plot(self.unique_value1, self.ILmax_matrix,
                marker='o', linestyle='-', color='b', label="ILmax")

        # 图标题、标签、网格与图例
        ax.set_title("ILmax")
        ax.set_xlabel(self.symbol1)        # 横轴：扫描参数（如 Lta）
        ax.set_ylabel("ILmax (dB)")        # 纵轴：单位为 dB
        ax.legend(loc="upper right")
        ax.grid(True)



    # === 绘图：ELmax、ULmax、WDLmax、mean 曲线 ===
    # 函数名: plot_maxmatrix
    # 功能: 展示多个评价指标随参数变化的趋势（用于筛选最优设计）
    # 参数:
    #   ax - matplotlib 的子图对象
    # 返回: 无
    def plot_maxmatrix(self, ax):
        metrics = {
            "ELmax": (self.ELmax_matrix, 'o', 2),
            "ULmax": (self.ULmax_matrix, 's', 2),
            "WDLmax": (self.WDLmax_matrix, 'd', 2),
            "mean": (self.mean_matrix, '*', 3),
        }
        for metric_name, (values, marker, linewidth) in metrics.items():
            if metric_name == "mean":
                ax.plot(self.unique_value1, values, marker=marker, markersize=10, linestyle='-', linewidth=linewidth,
                        label=metric_name, markeredgewidth=2)
            else:
                ax.plot(self.unique_value1, values, marker=marker, markersize=6, linestyle='-', linewidth=linewidth,
                        label=metric_name)
        ax.set_title("Performance Metrics")
        ax.set_xlabel(self.symbol1)
        ax.set_ylabel("Performance (dB)")
        ax.legend(loc="upper right")
        ax.grid(True)


    # === 总览绘图 ===
    # 函数名: plot_all
    # 功能: 输出所有重要指标的变化趋势图，共6幅子图（2×3）生成 2x3 大图，按照 1 2 5 | 3 4 6 的顺序
    # 参数: 无
    # 返回: 无
    def plot_all(self):
        fig, axes = plt.subplots(2, 3, figsize=(24, 12))
        self.plot_symbol_vs_wave(axes[0, 0], "ILmax_n", self.ILmax_n_matrix)
        self.plot_symbol_vs_wave(axes[0, 1], "EL", self.EL_matrix)
        self.plot_symbol_vs_wave(axes[1, 0], "UL", self.UL_matrix)
        self.plot_symbol_vs_wave(axes[1, 1], "WDL", self.WDL_matrix)
        self.plot_ILmax(axes[0, 2])       # 单独的 ILmax 曲线
        self.plot_maxmatrix(axes[1, 2])   # 所有 max 指标对比

        plt.tight_layout()
        save_path = self.file_path + "_result.png"
        plt.savefig(save_path, dpi=600, bbox_inches="tight")
        print(f"性能图像已保存到: {save_path}")
