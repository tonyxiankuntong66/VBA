' ==============================================================================
' 模块名称: CALC_UNITS (智能量纲推导与单位换算引擎)
' 核心功能: 支持 Excel 工作表中带单位文本的链式代数运算，内置量纲拓扑推导与安全自愈机制。
'
' 开源声明: 本代码完全开源，遵循 MIT 开源协议。允许自由复制、修改及商业使用。
' 维护状态: ⚠️ 现作者因精力/环境受限，已无法对该代码提供后续的更新与技术维护。
' 礼貌谢源: 欢迎各位同行在此基础上继续迭代。若在您的项目中衍生使用，还请保留原作者痕迹或
'           礼貌注明代码出处，共同维护良好的开源社区生态。谢谢！
' ==============================================================================
Option Explicit ' 强制显式声明变量，确保代码工业级稳健性

' ==========================================================
' 1. 通用声明区域 (所有 Type 和 Private 变量必须集中在此处)
' ==========================================================
Private Type ParsedElement
    Value As Double       ' 原始数值
    Unit As String        ' 标准化后的单位字符串
    Dimension As String   ' 物理量纲 (MASS, VOLUME, MOLAR_CONC, MASS_CONC, SCALAR)
    BaseValue As Double   ' 换算到系统基准量后的数值
End Type

' 模块级共享变量
Private mUnitRegistry As Object
Private mErrorStorage As String

' ==========================================================
' 2. 主接口函数
' ==========================================================
Public Function CALC_UNITS(ParamArray args() As Variant) As String
    On Error GoTo UnexpectedError
    
    ' 初始化底层量纲注册表
    InitializeRegistry
    
    Dim argc As Long
    argc = UBound(args) - LBound(args) + 1
    
    ' 提前声明变量，规避部分老版本 Excel 的块级作用域编译缺陷
    Dim singleResult As ParsedElement
    Dim accumulator As ParsedElement
    Dim nextToken As ParsedElement
    Dim i As Long
    Dim op As String
    Dim elementLabel As String
    
    ' 边界防错
    If argc = 0 Then
        CALC_UNITS = "【错误】未传入任何参数。"
        Exit Function
    Else
        ' 单个参数直接正规化输出
        If argc = 1 Then
            If Tokenize(args(0), singleResult, "参数1") Then
                CALC_UNITS = FormatOutput(singleResult.BaseValue, singleResult.Dimension)
            Else
                CALC_UNITS = "【解析失败】" & GetLastErrorContext("参数1")
            End If
            Exit Function
        End If
    End If
    
    ' 奇数个参数判定
    If argc Mod 2 = 0 Then
        CALC_UNITS = "【语法错误】参数配对不完整，可能遗漏了运算符或操作数。"
        Exit Function
    End If
    
    ' 解析初始拓扑节点
    If Not Tokenize(args(0), accumulator, "操作数 1 (" & CStr(SafeText(args(0))) & ")") Then
        CALC_UNITS = "【核心阻断】" & GetLastErrorContext("操作数 1")
        Exit Function
    End If
    
    ' 级联代数运算流水线
    For i = 1 To UBound(args) Step 2
        op = Trim(CStr(args(i)))
        elementLabel = "操作数 " & CStr((i \ 2) + 2) & " (" & CStr(SafeText(args(i + 1))) & ")"
        
        ' 解析右侧操作数
        If Not Tokenize(args(i + 1), nextToken, elementLabel) Then
            CALC_UNITS = "【核心阻断】" & GetLastErrorContext(elementLabel)
            Exit Function
        End If
        
        ' 动态量纲拓扑推导
        Select Case op
            Case "+", "-"
                If accumulator.Dimension <> nextToken.Dimension Then
                    CALC_UNITS = "【量纲冲突】无法对不同物理维度进行加减 [" & _
                                 GetFriendlyDimName(accumulator.Dimension) & "] " & op & " [" & _
                                 GetFriendlyDimName(nextToken.Dimension) & "]，请检查数据行。"
                    Exit Function
                End If
                
                If op = "+" Then
                    accumulator.BaseValue = accumulator.BaseValue + nextToken.BaseValue
                Else
                    accumulator.BaseValue = accumulator.BaseValue - nextToken.BaseValue
                End If
                
            Case "*"
                accumulator.BaseValue = accumulator.BaseValue * nextToken.BaseValue
                accumulator.Dimension = InferDimension(accumulator.Dimension, nextToken.Dimension, "*")
                
            Case "/"
                If nextToken.BaseValue = 0 Then
                    CALC_UNITS = "工作表计算链在 [" & elementLabel & "] 试图除以 0 导致溢出。"
                    Exit Function
                End If
                accumulator.BaseValue = accumulator.BaseValue / nextToken.BaseValue
                accumulator.Dimension = InferDimension(accumulator.Dimension, nextToken.Dimension, "/")
                
            Case Else
                CALC_UNITS = "【未知运算符】暂不支持运算符 [" & op & "]。请使用 +, -, *, /"
                Exit Function
        End Select
    Next i
    
    ' 逆向对齐输出
    CALC_UNITS = FormatOutput(accumulator.BaseValue, accumulator.Dimension)
    Exit Function

UnexpectedError:
    CALC_UNITS = "【运行时崩溃】" & Err.Description
End Function

' ==========================================================
' 3. 底层私有辅助子程序与函数
' ==========================================================
Private Function Tokenize(ByVal rawInput As Variant, ByRef outElement As ParsedElement, ByVal label As String) As Boolean
    ' 【Bug防线 1】拦截并自愈 Excel 单元格原生的错误值（如 #VALUE!, #DIV/0!）防止 CStr 转型崩溃
    If IsError(rawInput) Then
        SetLastErrorContext label, "输入源包含 Excel 工作表错误代码（如 #VALUE! 或 #DIV/0!）"
        Tokenize = False
        Exit Function
    End If

    Dim text As String
    ' 读取并安全转型
    If TypeOf rawInput Is Range Then
        text = rawInput.Text
    Else
        text = CStr(rawInput)
    End If
    
    ' 深度清洗流水线 (Normalization Pipeline)
    text = Trim(text)
    text = StrConv(text, vbNarrow) ' 全角转半角
    
    ' 泛化清洗所有微量级前缀变体
    text = Replace(text, "μ", "µ")
    text = Replace(text, "U", "u")
    
    If text = "" Then
        SetLastErrorContext label, "输入内容为空白"
        Tokenize = False
        Exit Function
    End If
    
    ' 拦截纯数字（标量）
    If IsNumeric(text) Then
        outElement.Value = CDbl(text)
        outElement.Unit = ""
        outElement.Dimension = "SCALAR"
        outElement.BaseValue = outElement.Value
        Tokenize = True
        Exit Function
    End If
    
    ' 正则表达式提取结构
    Dim regEx As Object, matches As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^\s*([0-9\.]+(?:[eE][+-]?[0-9]+)?)\s*([a-zA-Zµ\/0-9\.]+)\s*$"
    
    If Not regEx.Test(text) Then
        SetLastErrorContext label, "无法分离数字与单位，格式不合规（示例：'1.5 mg/mL'）"
        Tokenize = False
        Exit Function
    End If
    
    Set matches = regEx.Execute(text)
    outElement.Value = CDbl(matches(0).SubMatches(0))
    outElement.Unit = Trim(matches(0).SubMatches(1))
    
    ' 【Bug防线 2】智能大小写对齐：将单位中的小写 "l" 强制校准为大写 "L"，无视用户输入 ml/ul/l 的习惯
    Dim uKey As String
    uKey = outElement.Unit
    uKey = Replace(uKey, "u", "µ")
    uKey = Replace(uKey, "l", "L")
    
    ' 优先匹配单一标准单位
    If mUnitRegistry.Exists(uKey) Then
        Dim config As Variant
        config = mUnitRegistry(uKey)
        outElement.Dimension = config(0)
        outElement.BaseValue = outElement.Value * config(1)
        Tokenize = True
        Exit Function
    End If
    
    ' 复合单位动态解构
    If InStr(uKey, "/") > 0 Then
        Dim parts() As String
        parts = Split(uKey, "/")
        Dim mPart As String, vPart As String
        mPart = Trim(parts(0)): vPart = Trim(parts(1))
        
        mPart = Replace(mPart, "u", "µ"): mPart = Replace(mPart, "l", "L")
        vPart = Replace(vPart, "u", "µ"): vPart = Replace(vPart, "l", "L")
        
        ' 二级子正则，抓取分母可能携带的特殊体积数字
        Dim denValue As Double
        Dim denUnit As String
        Dim denReg As Object, denMatches As Object
        Set denReg = CreateObject("VBScript.RegExp")
        denReg.Pattern = "^\s*([0-9\.]*)\s*([a-zA-Zµ]+)\s*$"
        
        If denReg.Test(vPart) Then
            Set denMatches = denReg.Execute(vPart)
            If denMatches(0).SubMatches(0) = "" Then
                denValue = 1#
            Else
                denValue = CDbl(denMatches(0).SubMatches(0))
            End If
            denUnit = Trim(denMatches(0).SubMatches(1))
            
            If mUnitRegistry.Exists(mPart) And mUnitRegistry.Exists(denUnit) Then
                Dim mConf As Variant, vConf As Variant
                mConf = mUnitRegistry(mPart): vConf = mUnitRegistry(denUnit)
                
                If mConf(0) = "MASS" And vConf(0) = "VOLUME" Then
                    outElement.Dimension = "MASS_CONC"
                    Dim combinedFactor As Double
                    combinedFactor = mConf(1) / (denValue * vConf(1))
                    outElement.BaseValue = outElement.Value * combinedFactor
                    Tokenize = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' 未注册单位拦截
    SetLastErrorContext label, "未知或不受支持的单位符号 [" & outElement.Unit & "]"
    Tokenize = False
End Function

' ==========================================================
' 【Bug防线 3】无懈可击的物理量纲代数矩阵
' ==========================================================
Private Function InferDimension(ByVal dim1 As String, ByVal dim2 As String, ByVal op As String) As String
    If op = "*" Then
        ' 纯数字（标量）乘法降维保护
        If dim1 = "SCALAR" Then InferDimension = dim2: Exit Function
        If dim2 = "SCALAR" Then InferDimension = dim1: Exit Function
        
        ' 质量浓度 * 体积 = 质量
        If (dim1 = "MASS_CONC" And dim2 = "VOLUME") Or (dim1 = "VOLUME" And dim2 = "MASS_CONC") Then
            InferDimension = "MASS"
        Else
            InferDimension = dim1 & "_MULT_" & dim2
        End If
    ElseIf op = "/" Then
        ' 任何物理量除以纯数字，保留原物理量纲 (例如 5 mL / 2 = 2.5 mL)
        If dim2 = "SCALAR" Then InferDimension = dim1: Exit Function
        
        ' 经典物理代数推导
        If dim1 = "MASS" And dim2 = "VOLUME" Then
            InferDimension = "MASS_CONC"
        ElseIf dim1 = "MASS" And dim2 = "MASS_CONC" Then
            InferDimension = "VOLUME"
        Else
            InferDimension = dim1 & "_DIV_" & dim2
        End If
    End If
End Function

Private Function FormatOutput(ByVal baseValue As Double, ByVal dimension As String) As String
    Dim displayVal As Double
    Dim unitStr As String
    
    Select Case dimension
        Case "MASS"
            If baseValue >= 1# Then
                displayVal = baseValue: unitStr = "g"
            ElseIf baseValue >= 0.001 Then
                displayVal = baseValue * 1000#: unitStr = "mg"
            Else
                displayVal = baseValue * 1000000#: unitStr = "µg"
            End If
        Case "VOLUME"
            If baseValue >= 1# Then
                displayVal = baseValue: unitStr = "L"
            ElseIf baseValue >= 0.001 Then
                displayVal = baseValue * 1000#: unitStr = "mL"
            Else
                displayVal = baseValue * 1000000#: unitStr = "µL"
            End If
        Case "MOLAR_CONC"
            If baseValue >= 1# Then
                displayVal = baseValue: unitStr = "M"
            ElseIf baseValue >= 0.001 Then
                displayVal = baseValue * 1000#: unitStr = "mM"
            Else
                displayVal = baseValue * 1000000#: unitStr = "µM"
            End If
        Case "MASS_CONC"
            displayVal = baseValue
            unitStr = "mg/mL"
        Case "SCALAR"
            displayVal = baseValue
            unitStr = ""
        Case Else
            displayVal = baseValue
            unitStr = "[" & dimension & "]"
    End Select
    
    FormatOutput = Trim(CStr(displayVal) & " " & unitStr)
End Function

Private Sub InitializeRegistry()
    If Not mUnitRegistry Is Nothing Then Exit Sub
    Set mUnitRegistry = CreateObject("Scripting.Dictionary")
    mUnitRegistry.CompareMode = 0
    
    ' MASS 基准: g
    mUnitRegistry.Add "g", Array("MASS", 1#)
    mUnitRegistry.Add "mg", Array("MASS", 0.001)
    mUnitRegistry.Add "µg", Array("MASS", 0.000001)
    
    ' VOLUME 基准: L (统一使用标准大写 L)
    mUnitRegistry.Add "L", Array("VOLUME", 1#)
    mUnitRegistry.Add "mL", Array("VOLUME", 0.001)
    mUnitRegistry.Add "µL", Array("VOLUME", 0.000001)
    
    ' MOLAR_CONC 基准: M
    mUnitRegistry.Add "M", Array("MOLAR_CONC", 1#)
    mUnitRegistry.Add "mM", Array("MOLAR_CONC", 0.001)
    mUnitRegistry.Add "µM", Array("MOLAR_CONC", 0.000001)
End Sub

Private Sub SetLastErrorContext(ByVal label As String, ByVal errMsg As String)
    mErrorStorage = "[" & label & "] 解析受阻原因: " & errMsg
End Sub

Private Function GetLastErrorContext(ByVal label As String) As String
    If mErrorStorage = "" Then
        GetLastErrorContext = "[" & label & "] 遭遇未知格式畸变。"
    Else
        GetLastErrorContext = mErrorStorage
    End If
    mErrorStorage = "" 
End Function

Private Function GetFriendlyDimName(ByVal dimSystem As String) As String
    Select Case dimSystem
        Case "MASS": GetFriendlyDimName = "质量 (Mass)"
        Case "VOLUME": GetFriendlyDimName = "体积 (Volume)"
        Case "MOLAR_CONC": GetFriendlyDimName = "摩尔浓度 (MolarConc)"
        Case "MASS_CONC": GetFriendlyDimName = "质量浓度 (MassConc)"
        Case "SCALAR": GetFriendlyDimName = "无单位数字 (Scalar)"
        Case Else: GetFriendlyDimName = dimSystem
    End Select
End Function

Private Function SafeText(ByVal val As Variant) As String
    If IsError(val) Then
        SafeText = "#ERROR!"
    Else
        SafeText = CStr(val)
    End If
End Function
