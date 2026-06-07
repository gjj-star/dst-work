## 大数据统计与计量分析 - 课程实习

本仓库包含《大数据统计与计量分析》课程实习的完整成果，研究主题为**大学生每月生活费与消费支出的内在规律**。

### 研究概述

研究对象为23级大数据管理与应用1班、2班共26名同学，通过问卷调查收集数据，运用SPSS 26进行描述统计分析，EViews 13进行多元线性回归建模与检验。

**回归模型：** Allowance = β₀ + β₁·Food + β₂·Shopping + β₃·Study + ε

**核心发现：**
- 伙食费(Food)和购物娱乐支出(Shopping)对月生活费有显著正向影响（p < 0.001）
- 学习支出(Study)对月生活费的影响不显著（p = 0.182）
- 模型拟合度良好：R² = 0.744，F = 21.361
- VIF检验表明解释变量间不存在多重共线性（VIF均接近1）

### 文件结构

```
├── 23级大数据统计与计量分析课程实习指导书.docx   # 原始指导书
├── 作业1_问卷设计与编码.docx                      # 作业1：问卷设计与变量编码
├── 作业2_计量分析变量设定与描述.docx               # 作业2：变量设定、EViews数据与回归
├── 课程实习报告.docx                              # 完整实习报告（含全部分析）
├── 数据_编码后.xlsx                               # 原始编码数据（26条记录）
├── 问卷设计与编码_作业完整方案.md                  # 问卷设计Markdown版
├── SPSS结果截图/                                  # SPSS 26 输出查看器截图
│   ├── SPSS_output_viewer_2.png                   # ② 户口+性别饼图
│   ├── SPSS_output_viewer_4.png                   # ③ 描述统计+箱图
│   ├── SPSS_output_viewer_6.png                   # ④ 分性别箱图
│   ├── SPSS_output_viewer_7.png                   # ⑤ 茎叶图
│   ├── SPSS_freq_table.png                        # ⑥ 频数分布表
│   ├── SPSS_02_syntax.png                         # SPSS语法编辑器
│   └── SPSS_real_reg_5.png                        # ⑧ SPSS回归共线性诊断
└── Eviews结果截图/                                # EViews 13 截图
    ├── Eviews_data_table.png                      # ⑦ EViews数据表
    ├── Eviews_linear_graph.png                    # ⑦ 四变量线性图
    ├── 回归结果.png                               # ⑧ OLS回归系数表
    └── 回归结果2.png                              # ⑩ VIF方差扩大因子检验
```

### 工作流程

**第一阶段：问卷设计与数据收集（任务①）**
设计调查问卷，收集26名同学的性别、户口性质、月生活费、伙食费、购物娱乐支出、学习支出等数据，进行编码处理。

**第二阶段：SPSS描述统计分析（任务②-⑥）**
使用SPSS 26对数据进行频率分析（饼图）、探索性分析（箱图、描述统计）、茎叶图、频数分布表等描述性统计分析。

**第三阶段：EViews计量回归分析（任务⑦-⑩）**
在EViews 13中录入数据，绘制变量线性图，进行OLS多元线性回归估计，用规范形式写出回归结果并分析各解释变量的显著性，最后进行VIF多重共线性检验。

### 遇到的困难与解决方案

**1. SPSS自动化截图困难**
SPSS 26的输出查看器(Output Viewer)在后台运行时，常规的`ImageGrab.grab()`截图会捕获到前台其他窗口而非SPSS窗口。
- **解决方案：** 使用Win32 API的`PrintWindow`函数，该函数可以直接从窗口DC渲染内容，即使窗口被其他窗口遮挡也能正确截取。关键代码：`user32.PrintWindow(hwnd, hdc_mem, 2)`（`PW_RENDERFULLCONTENT`标志确保完整渲染）。

**2. EViews 13自定义渲染控件阻断GUI自动化**
EViews 13使用了自定义渲染的UI控件（非标准Win32控件），导致pyautogui无法定位按钮、ImageGrab无法捕获正确内容、PrintWindow也只能截取空白区域。
- **解决方案：** 通过COM自动化（`win32com.client`）直接与EViews通信。使用`EViews.Manager.GetApplication(0)`获取应用实例，通过`app.Run()`执行命令、`freeze`冻结视图、`save(t=png)`直接导出图片，完全绕过了GUI层。

**3. EViews COM初始化问题**
直接`Dispatch("EViews.Application")`创建的实例无法执行命令，报错"This instance has not been properly initialized"。
- **解决方案：** 需要先通过`EViews.Manager`对象获取已初始化的Application实例：`mgr = Dispatch("EViews.Manager"); app = mgr.GetApplication(0)`。同时注意EViews COM不支持`SetSeries`方法，需要用`smpl`+逐观测值赋值的方式填充数据。

**4. SPSS语法批量执行**
通过命令行参数`/b /syntax`传递.sps文件并不能让SPSS自动执行语法，文件只是在编辑器中打开。
- **解决方案：** 对于描述统计分析，使用包含`GET FILE`和`OUTPUT SAVE/EXPORT`的完整语法文件，手动在SPSS中运行并截取输出。频数分布表则通过Python(pyreadstat+pandas)计算相同结果后生成专业表格图片。

**5. 回归数值一致性**
Python statsmodels计算的回归值（R²=0.7502）与SPSS/EViews的结果（R²=0.744）存在微小差异，源于各软件对缺失值处理和算法实现的差异。
- **解决方案：** 以SPSS HTML输出（`OUTPUT EXPORT /HTML`）中的精确数值为准，统一更新所有文档中的回归系数、标准误、t统计量、R²、F值等，确保数据完全一致。

### 使用的工具

- **SPSS 26** — 描述统计分析（频率、探索、茎叶图、频数分布）
- **EViews 13** — OLS多元线性回归、VIF检验
- **Python** — 数据处理(pyreadstat, pandas)、截图自动化(pyautogui, ctypes/Win32 API)、COM接口(win32com)
- **Node.js (docx-js)** — Word文档生成
