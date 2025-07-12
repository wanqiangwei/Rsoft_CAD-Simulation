# ---------------------------------------------------------------
# 版本信息: Rsoft2020_Python3.12_1.0
# 作者信息: 中南大学机电工程学院 郑煜教授课题组
# 撰写人  : 2022届硕士研究生 万强伟
# 联系方式: 1025459384@qq.com
# 功能描述: 创建并仿真一个PLC 1×2 Y型分支波导
# ---------------------------------------------------------------

from RsoftCad import *          # 导入波导设计类
from RsoftSimulation import *   # 导入仿真控制类
from RsoftData import *         # 导入数据处理模块
from RsoftMail import *         # 导入邮件通知模块
import numpy as np

# === 初始化波导设计器 ===
# 参数: 路径, 文件名, 维度, 波长, 背景介质, 相对折射率差, 波导宽度
c = RsoftCad(r'D:\work\Python', 'test', 3, 1.55, 'SiO2', 0.0045, 6.5)

# === 设置仿真参数（符号变量）===
c.set_symbol('Lin', 500)
c.set_symbol('Lta', 600)
c.set_symbol('Ln', 550)
c.set_symbol('Wn', 5.5)
c.set_symbol('Lb', 500)
c.set_symbol('Wb', '2*width+Gap')
c.set_symbol('Lt', 100)
c.set_symbol('Wd', 0.3)
c.set_symbol('Gap', 1.5)
c.set_symbol('Offset', 0.4)
c.set_symbol('R', 15000)
c.set_symbol('a0', 0)
c.set_symbol('a1', 'acos(((-25+(width+Gap)/2-Offset/2)/R+1+cos(0))/2)')
c.set_symbol('Lout', 500)

# 网格与仿真设置
c.set_symbol('grid_size', 0.2)              # x方向网格
c.set_symbol('grid_size_y', 0.2)            # y方向网格
c.set_symbol('step_size', 2)                # z方向步进
c.set_symbol('wait', '0')                   # 仿真完成后是否关闭窗口
c.set_symbol('slice_display_mode', 'DISPLAY_CONTOURMAPXZ')  # 输出XZ截面图

# === 材料定义 ===
material1 = c.add_material(material_type.Dielectrics.SiO2)

# === 构建主波导段 ===
# 每段为 add_segment 组成的段体路径
seg1 = c.add_segment(vec3('None','None','None'), vec3(0, 0, 0), vec3(0, 0, 0), vec3('begin','begin','begin'),
                     vec3('Offset','Offset','Offset'), vec3(0, 0, 'Lin'), vec3(1, 1, 1), vec3('begin','begin','begin'),
                     vec2('width','height'), vec2('width','height'), taper.linar)

seg2 = c.add_segment(vec3('Offset','Offset','Offset'), vec3(0, 0, 0), vec3(1, 1, 1), vec3('end','end','end'),
                     vec3('Offset','Offset','Offset'), vec3(0, 0, 'Lta'), vec3(2, 2, 2), vec3('begin','begin','begin'),
                     vec2('width','height'), vec2('Wn','height'), taper.linar)

seg3 = c.add_segment(vec3('Offset','Offset','Offset'), vec3(0, 0, 0), vec3(2, 2, 2), vec3('end','end','end'),
                     vec3('Offset','Offset','Offset'), vec3(0, 0, 'Ln'), vec3(3, 3, 3), vec3('begin','begin','begin'),
                     vec2('Wn','height'), vec2('Wn','height'), taper.linar)

seg4 = c.add_segment(vec3('Offset','Offset','Offset'), vec3(0, 0, 0), vec3(3, 3, 3), vec3('end','end','end'),
                     vec3('Offset','Offset','Offset'), vec3(0, 0, 'Lb'), vec3(4, 4, 4), vec3('begin','begin','begin'),
                     vec2('Wn','height'), vec2('Wb','height'), taper.quadratic)

seg5 = c.add_segment(vec3('Offset','Offset','Offset'), vec3(0, 0, 0), vec3(4, 4, 4), vec3('end','end','end'),
                     vec3('Offset','Offset','Offset'), vec3(0, 0, 'Lt'), vec3(5, 5, 5), vec3('begin','begin','begin'),
                     vec2('Wb','height'), vec2('Wb','height'), taper.linar)

# === 构建分支弯道段（arc）===
seg6 = c.add_arc(vec3('Offset','Offset','Offset'), vec3('(width+Gap)/2', 0, 0), vec3(5, 5, 5), vec3('end','end','end'),
                 vec3_arc('R','a0','a1'), vec2('width','height'), vec2('width','height'), taper.linar)

seg7 = c.add_arc(vec3('Offset','Offset','Offset'), vec3('-(width+Gap)/2', 0, 0), vec3(5, 5, 5), vec3('end','end','end'),
                 vec3_arc('R','a0','-a1'), vec2('width+Wd','height'), vec2('width','height'), taper.linar)

seg8 = c.add_arc(vec3('Offset','Offset','Offset'), vec3('-Offset', 0, 0), vec3(6, 6, 6), vec3('end','end','end'),
                 vec3_arc('R','a1','a0'), vec2('width-Wd','height'), vec2('width','height'), taper.linar)

seg9 = c.add_arc(vec3('Offset','Offset','Offset'), vec3('Offset', 0, 0), vec3(7, 7, 7), vec3('end','end','end'),
                 vec3_arc('R','-a1','a0'), vec2('width','height'), vec2('width','height'), taper.linar)

# === 出口波导段 ===
seg10 = c.add_segment(vec3('Offset','Offset','Offset'), vec3('Offset/2', 0, 0), vec3(8, 8, 8), vec3('end','end','end'),
                      vec3('Offset','Offset','Offset'), vec3(0, 0, 'Lout'), vec3(10, 10, 10), vec3('begin','begin','begin'),
                      vec2('width','height'), vec2('width','height'), taper.linar)

seg11 = c.add_segment(vec3('Offset','Offset','Offset'), vec3('-Offset/2', 0, 0), vec3(9, 9, 9), vec3('end','end','end'),
                      vec3('Offset','Offset','Offset'), vec3(0, 0, 'Lout'), vec3(11, 11, 11), vec3('begin','begin','begin'),
                      vec2('width','height'), vec2('width','height'), taper.linar)

# === 添加路径、监视器、光源 ===
pathway1 = c.add_pathway([1, 2, 3, 4, 5, 6, 8, 10])
pathway2 = c.add_pathway([1, 2, 3, 4, 5, 7, 9, 11])

monitor1 = c.add_monitor(pathway=pathway1, monitor_type=monitor_type.Launch_Power)
monitor2 = c.add_monitor(pathway=pathway2, monitor_type=monitor_type.Launch_Power)

launch1 = c.add_launch(pathway=pathway1, launch_type=launch_type.Computed_Mode)

# === 初始化仿真控制器 ===
s = RsoftSimulation(r'D:\work\Python', 'test', 6, "on")

# === 单次仿真调用Sim ===
Sim1 = s.Sim('default', 'default')
Sim2 = s.Sim(['Lta'], [300])
Sim3 = s.Sim(['Lta', 'Ln', 'Wn'], [300, 400, 5])

# === 双参数扫描Scan（含波长） ===
Lta_list = [0, 50, 100, 200, 160.8, 300.18]
Ln_list = [100, 200, 350, 500.8]
wave_list = [1.27, 1.31, 1.49, 1.55, 1.65]

Scan1 = s.Scan(['Lta', 'wave'], [Lta_list, [1.55]])
Scan2 = s.Scan(['Ln', 'wave'], [Ln_list, [1.27, 1.55]])
Scan3 = s.Scan(['Lta', 'wave'], [Lta_list, wave_list])

# === 多参数优化仿真Optimize ===
Lta_list = [200, 400, 600, 800]
Wn_list = [4, 4.5, 5, 5.5]
Lb_list = [600, 700, 800]
R_list = [12000, 14000, 16000, 18000, 20000]
Offset_list = [-0.4, 0, 0.4, 0.8, 1.2]

o1 = s.Optimize(['Lta', 'Ln', 'Wn', 'Lb', 'wave'], [Lta_list, Ln_list, Wn_list, Lb_list, wave_list])
o2 = s.Optimize(['R', 'Offset', 'wave'], [R_list, Offset_list, wave_list])

# === 多参数正交设计优化仿真OEDsim ===
Lta_list = np.linspace(400, 800, 5)
Ln_list = np.linspace(400, 800, 5)
Wn_list = np.linspace(4, 6, 5)
Lb_list = np.linspace(800, 400, 5)
Lt_list = np.linspace(80, 120, 5)
wave_list = [1.27, 1.31, 1.49, 1.55, 1.65]

OEDsim1 = s.OEDsim(['Lta', 'Ln', 'Wn', 'Lb', 'Lt', 'wave'], [Lta_list, Ln_list, Wn_list, Lb_list, Lt_list, wave_list])


# 等待仿真结束
s.wait_completion()

# === 仿真数据处理 ===
# Scan 与 Optimize 会自动处理数据
RsoftData(Sim1)
RsoftData(Sim2)
RsoftData(Sim3)


# === 仿真结束邮件通知 ===
RsoftMail('1025459384@qq.com')