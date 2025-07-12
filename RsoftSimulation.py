# ============================================================
# 文件名称: RsoftSimulation.py
# 模块功能: 控制 RSoft BPM 仿真、参数扫描、自动优化等任务
# 支持功能:
#   - 单次仿真 Sim
#   - 参数扫描 Scan（支持双参数）
#   - 多参数级联优化 Optimize
#   - 自动窗口最小化、许可证弹窗处理、并发仿真调度
# 作者单位: 中南大学机电工程学院 郑煜教授课题组
# 撰写作者: 2022届硕士研究生 万强伟
# ============================================================

import subprocess
import time
import win32gui
import win32con
import threading
from concurrent.futures import ThreadPoolExecutor
import pyautogui
import shutil
from RsoftData import *
from OAT import *


class RsoftSimulation:
    # === 构造函数: 初始化仿真类，设置文件路径、最大并发数、窗口控制等 ===
    # 参数:
    #   file_path        : ind 文件所在路径
    #   file_name        : ind 文件名称（不含后缀）
    #   max_workers      : 最大并发仿真数量
    #   window_minimize  : 是否最小化仿真窗口（"on"/"off"）
    def __init__(self, file_path=str, file_name=str, max_workers=int, window_minimize="on"):
        self.file_name = file_name
        self.file_path = file_path
        self.file = file_path + "\\" + self.file_name + ".ind"  # 拼接完整文件路径
        self.max_workers = max_workers
        self.window_minimize = window_minimize

        # 创建并发线程池（最多 max_workers 个任务同时进行）
        self.command_pool = ThreadPoolExecutor(max_workers=self.max_workers)

        # 最小化窗口的控制标志，只在第一次运行时多次尝试
        self.first_minimize = True
        self.mailnum = 0


    # === 启动 RSoft 仿真命令，并自动处理窗口与许可证 ===
    # 函数名: run_command
    # 功能:
    #   - 启动系统命令运行 RSoft 仿真
    #   - 最小化仿真窗口（可选）
    #   - 启动后台线程监控许可证弹窗
    # 参数:
    #   command  : 系统命令字符串（如 bsimw32 xxx.ind prefix=xxx ...）
    #   work_dir : 命令执行的工作目录（仿真路径）
    # 返回: 无（仿真进程阻塞直至完成）
    def run_command(self, command, work_dir):
        print(f"启动命令: {command}")

        # 启动子进程运行命令
        process = subprocess.Popen(command, shell=True, cwd=work_dir)

        # 启动后台守护线程监测并点击许可证窗口（Query）
        Query_thread = threading.Thread(
            target=self.detect_and_click_query_window,
            args=('Query', 410, 523),
            daemon=True  # 守护线程，主程序退出则自动关闭
        )
        Query_thread.start()


        # 可选：最小化所有包含 "Computation" 的窗口
        if self.window_minimize == "on":
            self.minimize_rsoft_window()

        # 等待仿真进程结束
        process.wait()


    # === 扫描所有窗口，匹配包含特定标题的窗口并最小化 ===
    # 函数名: minimize_rsoft_window
    # 功能:
    #   - 最小化所有窗口标题包含指定字符串的窗口（默认为“Computation”）
    #   - 仅在第一次调用时进行 self.max_workers+1 次尝试，之后仅尝试 2 次
    # 参数:
    #   window_title_part : 匹配窗口标题的关键字（默认 "Computation"）
    # 返回: 无
    def minimize_rsoft_window(self, window_title_part="Computation"):
        def callback(hwnd, _):
            if window_title_part in win32gui.GetWindowText(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        # 控制尝试次数
        if self.first_minimize:
            repeat_times = self.max_workers + 1
            self.first_minimize = False
        else:
            repeat_times = 2
        for _ in range(repeat_times):
            win32gui.EnumWindows(callback, None)
            time.sleep(1.5)


    # === 检测是否弹出“Query”窗口，点击“否”自动关闭 ===
    # 函数名: detect_and_click_query_window
    # 功能:
    #   - 定期扫描是否存在名为 title 的窗口
    #   - 若存在则在相对坐标 (rel_x, rel_y) 模拟鼠标点击
    # 参数:
    #   title    : 窗口标题完整匹配（如 "Query"）
    #   rel_x    : 相对左上角 X 坐标（用于点击“否”按钮）
    #   rel_y    : 相对左上角 Y 坐标
    #   interval : 检查间隔时间（秒）
    def detect_and_click_query_window(self, title, rel_x, rel_y, interval=10):
        while True:
            hwnd = win32gui.FindWindow(None, title)
            if hwnd:
                # 获取窗口左上角坐标
                left, top, _, _ = win32gui.GetWindowRect(hwnd)
                abs_x, abs_y = left + rel_x, top + rel_y
                pyautogui.moveTo(abs_x, abs_y)
                pyautogui.click()
                print(f"点击窗口 '{title}' 内相对坐标 ({rel_x}, {rel_y})的‘否’，屏幕坐标 ({abs_x}, {abs_y})")
                break
            time.sleep(interval)


    # === 等待当前线程池中所有任务完成 ===
    # 函数名: wait_completion
    # 功能: 阻塞直到所有仿真命令结束
    # 返回: 无
    def wait_completion(self):
        self.command_pool.shutdown(wait=True)
        print("所有命令执行完毕")


    # === 优化过程中专用的等待函数（自动重建线程池）===
    # 函数名: wait_Scan
    # 功能:
    #   - 等待线程池中所有优化仿真任务完成
    #   - 重建线程池，为下一轮仿真做准备
    # 返回: 无
    def wait_Scan(self):
        self.command_pool.shutdown(wait=True)
        print("Scan命令执行完毕")
        self.command_pool = ThreadPoolExecutor(self.max_workers)
        self.first_minimize = True


    # === 生成格式化字符串，确保 valuelist 中所有值输出宽度一致 ===
    # 函数名: determine_format
    # 功能:
    #   - 将数值列表中的元素转化为统一对齐的字符串（用于仿真目录命名、文件解析等）
    #   - 自动判断是否为浮点数
    #   - 计算整数和小数部分所需宽度，生成格式字符串
    # 参数:
    #   valuelist : list[float or int]
    # 返回:
    #   format_str : str，Python 格式化模板（如 "{:05.2f}"）
    def determine_format(self, valuelist):
        # 将所有输入值强制转换为 float，确保后续处理统一
        all_values = [float(v) for v in valuelist]

        # 检查是否包含浮点数（存在非整数即为 True）
        has_floats = any(v % 1 != 0 for v in all_values)

        # 获取整数部分所需最大字符长度（例如 1, 100, 1000 -> 长度为 4）
        max_integer_length = max(len(str(int(v))) for v in all_values)

        if has_floats:
            # 获取小数部分最长精度（例如 1.2, 1.23, 1.234 -> 小数位数为 3）
            decimal_places = max(
                len(f"{v}".split(".")[-1]) if "." in f"{v}" else 0 for v in all_values
            )

            # 生成浮点数格式字符串，如 "{:07.3f}"，7 = 整数 + 小数点 + 小数位
            format_str = f"{{:0{max_integer_length + decimal_places + 1}.{decimal_places}f}}"
        else:
            # 若为纯整数，生成整数格式字符串，如 "{:04d}"
            format_str = f"{{:0{max_integer_length}d}}"

        return format_str


    # === 执行单次 BPM 仿真（Sim）任务 ===
    # 函数名: Sim
    # 功能:
    #   - 使用默认参数或指定参数执行单次 RSoft 仿真
    #   - 支持指定 symbol=value 对参数进行覆盖仿真
    # 参数:
    #   symbollist : 'default' 或 参数名列表（如 ['Lta', 'Ln']）
    #   valuelist  : 'default' 或 参数值列表（如 [200, 500]）
    # 返回:
    #   run_path   : 仿真结果目录（用于后续分析）
    def Sim(self, symbollist, valuelist):
        # 创建仿真路径（如 D:\work\test_Sim）
        Sim_path = f"{self.file_path}\\{self.file_name}_Sim"
        if not os.path.exists(Sim_path):
            os.makedirs(Sim_path)

        # === 情况一：使用默认参数 ===
        if symbollist == 'default' and valuelist == 'default':
            run_path = f"{Sim_path}\\default"  # 仿真结果子目录
            if not os.path.exists(run_path):
                os.makedirs(run_path)

            run_prefix = "default"
            commend = "bsimw32 " + self.file + " prefix=" + run_prefix
            self.command_pool.submit(self.run_command, commend, run_path)

        # === 情况二：使用自定义参数 ===
        else:
            # 文件夹命名：Lta(300)_Ln(500)
            symbol_value_bracket = [f"{symbollist[i]}({valuelist[i]})" for i in range(len(symbollist))]
            symbol_value_path = "_".join(symbol_value_bracket)
            run_path = f"{Sim_path}\\{symbol_value_path}"
            if not os.path.exists(run_path):
                os.makedirs(run_path)

            run_prefix = symbol_value_path

            # 构造参数字符串：Lta=300 Ln=500
            symbol_value_equal = [f"{symbollist[i]}={valuelist[i]}" for i in range(len(symbollist))]
            symbol_value_space = " ".join(symbol_value_equal)

            # 构造仿真命令并提交
            commend = "bsimw32 " + self.file + " prefix=" + run_prefix + " " + symbol_value_space
            self.command_pool.submit(self.run_command, commend, run_path)

        return run_path


    # === 双参数扫描仿真（Scan）任务 ===
    # 函数名: Scan
    # 功能:
    #   - 执行两个参数组合下的所有仿真任务（全排列）
    #   - 通常用于某一结构下，对 wave 和某参数进行性能扫描
    # 参数:
    #   symbollist : 参数名列表（必须为两个参数，如 ['Lta', 'wave']）
    #   valuelist  : 对应值列表（如 [[100,200],[1.55,1.65]]）
    #   optimize   : 优化模式标志（"off" 或 "on"）
    # 返回:
    #   run_path   : 仿真结果路径（供后续数据处理）
    def Scan(self, symbollist, valuelist, optimize="off"):
        # === 创建扫描结果根目录 ===
        if optimize == "off":
            Scan_path = f"{self.file_path}\\{self.file_name}_Scan"
            if not os.path.exists(Scan_path):
                os.makedirs(Scan_path)
        elif optimize == "on":
            Scan_path = self.optimize_path  # 优化模式下路径特殊处理

        # 构建文件夹命名（如 Lta(100_800)_wave(1.55_1.65)）
        symbol_value_sta_end_bracket = [f"{symbollist[i]}({valuelist[i][0]}_{valuelist[i][-1]})" for i in range(len(symbollist))]
        symbol_value_path = "_".join(symbol_value_sta_end_bracket)

        # 创建扫描结果路径
        if optimize == "off":
            run_path = f"{Scan_path}\\{symbol_value_path}"
        elif optimize == "on":
            run_path = f"{Scan_path}\\{self.optimize_index}_{symbol_value_path}"
            self.optimize_index += 1
        if not os.path.exists(run_path):
            os.makedirs(run_path)

        # === 将 valuelist 格式化为等宽字符串，防止路径混乱 ===
        valuelist_format = [[], []]
        for i in range(len(valuelist)):
            format_str = self.determine_format(valuelist[i])
            valuelist_format[i] = [format_str.format(v) for v in valuelist[i]]

        # === 构造所有参数组合并提交仿真任务 ===
        for i in range(len(valuelist[0])):
            for j in range(len(valuelist[1])):
                # 构建路径：Lta(100)_wave(1.55)
                symbol_value_bracket = [f"{symbollist[0]}({valuelist_format[0][i]})", f"{symbollist[1]}({valuelist_format[1][j]})"]
                symbol_value_path = "_".join(symbol_value_bracket)
                run_prefix = symbol_value_path

                # 构建仿真参数：Lta=100 wave=1.55
                symbol_value_equal = [f"{symbollist[0]}={valuelist[0][i]}", f"{symbollist[1]}={valuelist[1][j]}"]
                symbol_value_space = " ".join(symbol_value_equal)

                # 构建命令（优化模式路径不同）
                if optimize == "on":
                    commend = "bsimw32 " + self.Optimize_Rsoft + " prefix=" + run_prefix + " " + symbol_value_space
                else:
                    commend = "bsimw32 " + self.file + " prefix=" + run_prefix + " " + symbol_value_space

                self.command_pool.submit(self.run_command, commend, run_path)
        if optimize == "on":
            return run_path
        else:
            self.wait_Scan()
            RsoftData(run_path)


    # === 多参数级联优化仿真（Optimize） ===
    # 函数名: Optimize
    # 功能:
    #   - 多轮双参数扫描，每轮选择最优值替换 ind 文件中的参数
    #   - 每轮都将当前优化参数与 wave 组合进行 Scan
    #   - 最终保留每轮优化结果至 Optimize_result.txt
    # 参数:
    #   symbolList : 参数名列表（如 ['Lta', 'Ln', 'Wn', ..., 'wave']）
    #   valueList  : 与 symbolList 一一对应的值列表（每个是数组）
    # 返回: 无（中间输出包括数据、图、结果文件）
    def Optimize(self, symbolList, valueList):
        # === Step 1: 创建干净优化目录 OptimizeN ===
        self.Optimize_path = self.create_clean_optimize_path(self.file_path, self.file_name)
        self.optimize_index = 1
        self.optimize_path = self.Optimize_path

        # === Step 2: 创建并打开结果记录文件 ===
        Optimize_result_file = f"{self.Optimize_path}\\Optimize_result.txt"
        open(Optimize_result_file, "w").close()  # 清空旧文件
        Optimize_result = open(Optimize_result_file, "r+")

        # === Step 3: 复制 .ind 文件，用于迭代修改 ===
        self.Optimize_Rsoft = f"{self.Optimize_path}\\{self.file_name}_optimize.ind"
        shutil.copyfile(self.file, self.Optimize_Rsoft)

        # === Step 4: 多轮循环优化，每轮只优化一个参数 + wave ===
        for i in range(len(symbolList) - 1):
            # 组合当前优化参数 + wave
            symbollist = [symbolList[i], symbolList[-1]]
            valuelist = [valueList[i], valueList[-1]]

            # 调用 Scan 函数提交所有组合仿真任务
            sacn_path = self.Scan(symbollist, valuelist, optimize="on")
            self.wait_Scan()  # 等待仿真完成

            # 数据分析：提取最优值
            data = RsoftData(sacn_path)
            min_symbol = data.get_min_symbol()

            # 修改 optimize.ind 中当前参数为最优值
            self.change_symbol(self.Optimize_Rsoft, symbolList[i], min_symbol)

            # 记录优化过程到结果文件
            Optimize_result.write(f"{symbolList[i]} {valueList[i]}\n")
            Optimize_result.write(f"{symbolList[i]}={min_symbol}\n")


    # === 自动创建干净的 OptimizeN 文件夹（若存在空文件夹则复用）===
    # 函数名: create_clean_optimize_path
    # 功能:
    #   - 检查当前路径下是否存在 Optimize1、Optimize2 等目录
    #   - 若存在但 result.txt 为空，则删除重建
    #   - 否则新建 OptimizeN
    # 参数:
    #   file_path  : 根路径
    #   file_name  : 文件名（用于构建 OptimizeN 文件夹名）
    # 返回:
    #   new_path   : 最终可用的 OptimizeN 路径
    def create_clean_optimize_path(self, file_path, file_name):
        index = 1
        while True:
            new_path = os.path.join(file_path, f"{file_name}_Optimize{index}")
            result_file = os.path.join(new_path, "Optimize_result.txt")

            if os.path.exists(new_path):
                # 目录存在，判断 result 文件是否为空
                if os.path.isfile(result_file) and os.path.getsize(result_file) == 0:
                    print(f"检测到空文件: {result_file}，清空目录 {new_path}...")
                    shutil.rmtree(new_path)
                    os.makedirs(new_path)
                    break
                else:
                    index += 1
            else:
                os.makedirs(new_path)
                break
        return new_path


    # === 替换 ind 文件中的指定 symbol = value 行 ===
    # 函数名: change_symbol
    # 功能:
    #   - 查找 Optimize.ind 中的 symbol 参数行
    #   - 用新的 min_symbol 替换旧行
    # 参数:
    #   Optimize_Rsoft : ind 文件路径
    #   symbol         : 要替换的 symbol 名（如 "Lta"）
    #   min_symbol     : 优化后最佳值（如 300.0）
    # 返回: 无（文件内容已更新）
    def change_symbol(self, Optimize_Rsoft, symbol, min_symbol):
        new_line = f"{symbol} = {min_symbol}\n"
        # 读入全部行
        with open(Optimize_Rsoft, "r") as f:
            lines = f.readlines()
        # 查找 symbol 所在行
        target_prefix = f"{symbol} ="
        target_index = None
        old_line = None
        for idx, line in enumerate(lines):
            if line.strip().startswith(target_prefix):
                target_index = idx
                old_line = line
                break
        if target_index is None:
            print(f"未找到行: {target_prefix}，不修改文件。")
            return
        # 替换行
        lines.pop(target_index)
        lines.insert(target_index, new_line)
        # 写回文件
        with open(Optimize_Rsoft, "w") as f:
            f.writelines(lines)
        print(f"替换前: {old_line.strip()} 替换后: {new_line.strip()}")


    # === 多参数正交设计优化仿真OEDsim ===
    def OEDsim(self, symbollist, valuelist):
        # === 创建扫描结果根目录 ===
        OEDsim_path = f"{self.file_path}\\{self.file_name}_OEDsim"
        # 构建文件夹命名（如 Lta(100_800)_wave(1.55_1.65)）
        symbol_value_sta_end_bracket = [f"{symbollist[i]}({valuelist[i][0]}_{valuelist[i][-1]})" for i in range(len(symbollist))]
        symbol_value_path = "_".join(symbol_value_sta_end_bracket)
        run_path = f"{OEDsim_path}\\{symbol_value_path}"
        if not os.path.exists(run_path):
            os.makedirs(run_path)

        # 按照格式排列数据，不包含wave
        oat = OAT()
        OED = OrderedDict(zip(symbollist[:-1], valuelist[:-1]))
        # 默认mode=0，宽松模式，只裁剪重复测试集（测试用例参数值可能为None）
        # mode=1，严格模式，除裁剪重复测试集外，还裁剪含None测试集(num为允许None测试集最大数目)
        # 生成正交设计测试用例
        test_OED = oat.genSets(OED, mode=1, num=0)

        # === 构造所有参数组合并提交仿真任务 ===
        for i, case in enumerate(test_OED, 1):
            for wave in valuelist[-1]:
                # 仿真前缀_wave(1.55)格式化为等宽字符串，防止路径混乱
                run_prefix = f"test({i:0{len(str(len(test_OED)))}d})_wave({wave})"
                # 构建仿真参数：Lta=400.0 Ln=400.0 Wn=4.0 Lb=800.0 Lt=80.0
                symbol_value_space = " ".join(f"{key}={value}" for key, value in case.items())
                symbol_value_space_wave =f"{symbol_value_space} wave={wave}"
                # 构建命令（优化模式路径不同）
                commend = "bsimw32 " + self.file + " prefix=" + run_prefix + " " + symbol_value_space_wave
                self.command_pool.submit(self.run_command, commend, run_path)
        # 等待仿真完毕，读取数据并处理
        self.wait_Scan()
        RsoftData(run_path)
        # === 写入正交设计表格 ===
        result_path = run_path + "_result.txt"
        resultfile = open(result_path, "r+")
        # 读取原始内容
        remaining_content = resultfile.read()
        # 回到文件开头
        resultfile.seek(0)
        # 写入正交设计表格
        table_data = []
        # 构建表头（根据矩阵类型判断）
        header = [f"test/symbol"] + symbollist[:-1]
        table_data.append(header)
        # 写入每一行数据
        for i, case in enumerate(test_OED, 1):
            row = [i] + [f"{value}" for _, value in case.items()]
            table_data.append(row)

        # 使用 tabulate 美化输出为纯文本表格
        formatted_table = tabulate(table_data, tablefmt="plain")
        resultfile.write(formatted_table + "\n\n")
        # 保留原有内容
        resultfile.write(remaining_content)
