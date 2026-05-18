# VBA
# Excel VBA 智能量纲推导与单位换算引擎 (`CALC_UNITS`)

这是一个为 Excel 打造的工业级自定义函数（UDF）。它打破了传统 Excel 公式只能计算纯数字的局限，支持直接对带有物理单位的文本进行加、减、乘、除链式代数运算，并具备动态量纲拓扑推导与单位自动对齐功能。
---
## ⚠️ 开源与维护状态声明

> **Module**: CALC_UNITS (Intelligent Dimension Inference & Unit Conversion Engine)  
> **License**: Fully open-source under the **MIT License**. Feel free to fork, modify, and reuse.  
> **Status**: ⚠️ **[DEPRECATED / UNMAINTAINED]** 现作者因精力与环境受限，已无法对该代码提供后续的更新与技术维护。  
> **Credits**: 欢迎各位同行在此基础上继续迭代。若在您的项目中衍生使用，还请保留原作者痕迹或礼貌注明代码出处，共同维护良好的开源社区生态。谢谢！

---

## 🚀 核心特性 (Features)

*   **动态量纲推导 (Dimension Inference)**：系统内部不仅计算数值，还会实时追踪并推导物理维度，支持跨维度的代数拓扑推导。例如：
    *   $\text{质量} \div \text{体积} \rightarrow \text{质量浓度}$（如 `"10 mg" / "2 mL"` $\rightarrow$ `"5 mg/mL"`）
    *   $\text{质量浓度} \times \text{体积} \rightarrow \text{质量}$（如 `"5 mg/mL" * "2 mL"` $\rightarrow$ `"10 mg"`）
    *   $\text{物理量} \div \text{纯数字}$ $\rightarrow$ 保留原物理量纲（如 `"10 mL" / 4` $\rightarrow$ `"2.5 mL"`）
*   **跨数量级自动对齐 (Auto-Alignment)**：内置基于“克 (g)、升 (L)、摩尔 (M)”的系统基准量注册表。当进行加减法时，即便单位数量级不同（如 `500 mg + 1.5 g`），引擎也会自动对齐换算并正确输出 `2 g`。
*   **智能结果正规化 (Format Output)**：计算完成后，系统会根据最终数值的大小，**自动匹配最优雅的单位**进行输出（例如：计算结果为 `0.0005 g` 时，会自动优化显示为 `"500 µg"`）。

---

## 💎 项目两大亮点 (Highlights)

### 1. 无懈可击的“Bug 防御战线”与自愈能力
代码在处理混乱、不规范的用户输入时表现出了极高的鲁棒性（稳健性）：
*   **工作表原生错误拦截**：通过 `IsError(rawInput)` 提前拦截单元格自带的 `#VALUE!` 或 `#DIV/0!` 等原生错误码，防止 VBA 在进行强类型转换时发生致命崩溃。
*   **文本深度标准化流水线**：利用 `StrConv(text, vbNarrow)` 强制将全角字符转为半角，并利用正则和替换引擎，自动将用户乱写的 `ml`、`ul`、`U` 修正为标准大写的 `mL` 或微量符号 `µL`，具备极强的容错和自愈能力。

### 2. 强类型结构与优雅的“计算链”设计
这体现了标准软件工程的编写规范，具有极高的扩展性：
*   **高度封装的拓扑节点**：通过自定义强类型结构 `Private Type ParsedElement`，将每一个输入元素拆解为 `Value（数值）`、`Unit（单位）`、`Dimension（量纲）` 和 `BaseValue（基准值）`。
*   **状态机级联流水线**：主接口函数利用 `ParamArray` 接收变长参数，通过一个 `For...Step 2` 循环配合 `Select Case` 状态机，像一条代数流水线一样层层推进计算。逻辑清晰，未来若需增加“压力、温度”等新量纲，只需在注册表中追加定义，无需重构主逻辑。

---

## 📊 使用示例 (Examples)

| 单元格公式 | 内部量纲推导 | 期望输出结果 | 说明 |
| :--- | :--- | :--- | :--- |
| `=CALC_UNITS("2.5 mg/mL", "*", "4 mL")` | `MASS_CONC * VOLUME -> MASS` | **`10 mg`** | 自动消去体积单位，推导出质量 |
| `=CALC_UNITS("500 mg", "+", "1.5 g")` | `单位自动对齐 (0.5g + 1.5g)` | **`2 g`** | 跨数量级自动换算并相加 |
| `=CALC_UNITS("10 mL", "/", 4)` | `VOLUME / SCALAR -> VOLUME` | **`2.5 mL`** | 支持标量（纯数字）微调 |
| `=CALC_UNITS("  0.05 ml ", "/", 100)` | `文本清洗 + 结果正规化` | **`0.5 µL`** | 空格自愈，小写 `ml` 自动正规化为 `µL` |
| `=CALC_UNITS("10 g", "+", "5 mL")` | `量纲冲突防御触发` | **`【量纲冲突】...`** | 严禁跨维度加减，弹出友好提示 |

---

## 🛠️ 安装与使用 (Installation)

1. 在 Excel 中按下 `Alt + F11` 打开 VBA 编辑器。
2. 点击菜单栏 `插入` $\rightarrow$ `模块`。
3. 将本仓库中的代码完整复制并粘贴到新建的模块中。
4. 返回 Excel 工作表，即可像普通公式一样使用 `=CALC_UNITS()`。
