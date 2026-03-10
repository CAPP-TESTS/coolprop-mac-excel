Attribute VB_Name = "Module1"
'Information on calling DLL from Excel:
'http://msdn.microsoft.com/en-us/library/office/bb687915.aspx
'
'Information on compiling DLL:
'http://msdn.microsoft.com/en-us/library/office/bb687850.aspx

'Information on 32-bit/64-bit compatibility (64-bit only has one harmonized calling convention)
'http://msdn.microsoft.com/en-us/library/office/ff700513%28v=office.11%29.aspx

' If you use older versions of Excel, you might need to remove the PtrSafe from the function definition

Option Explicit
#If Mac Then
    ' Even though the functions are exported with a leading underscore, Excel 2011 for Mac doesn't want the leading underscore as part of name
    Private Declare PtrSafe Function PropsSI_private Lib "/tmp/libCoolProp.dylib" Alias "PropsSI" (ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String) As Double
    Private Declare PtrSafe Function PhaseSI_private Lib "/tmp/libCoolProp.dylib" Alias "PhaseSI" (ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function Props1SI_private Lib "/tmp/libCoolProp.dylib" Alias "Props1SI" (ByVal output As String, ByVal Ref As String) As Double
    Private Declare PtrSafe Function get_global_param_string_private Lib "/tmp/libCoolProp.dylib" Alias "get_global_param_string" (ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function get_fluid_param_string_private Lib "/tmp/libCoolProp.dylib" Alias "get_fluid_param_string" (ByVal fluid As String, ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function HAPropsSI_private Lib "/tmp/libCoolProp.dylib" Alias "HAPropsSI" (ByVal output As String, ByVal Input1Name As String, ByVal Value1 As Double, ByVal Input2Name As String, ByVal Value2 As Double, ByVal Input3name As String, ByVal Value3 As Double) As Double
    Public Declare PtrSafe Function set_config_string Lib "/tmp/libCoolProp.dylib" (ByVal key As String, ByVal value As String) As Long
    'DEPRECATED
    Private Declare PtrSafe Function Props_private Lib "/tmp/libCoolProp.dylib" Alias "PropsS" (ByVal output As String, ByVal Name1 As Long, ByVal Value1 As Double, ByVal Name2 As Long, ByVal Value2 As Double, ByVal Ref As String) As Double
#ElseIf Win64 Then
    Private Declare PtrSafe Function get_global_param_string_private Lib "CoolProp_x64.dll" Alias "get_global_param_string" (ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function get_fluid_param_string_private Lib "CoolProp_x64.dll" Alias "get_fluid_param_string" (ByVal fluid As String, ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function PropsSI_private Lib "CoolProp_x64.dll" Alias "PropsSI" (ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String) As Double
    Private Declare PtrSafe Function PhaseSI_private Lib "CoolProp_x64.dll" Alias "PhaseSI" (ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function Props1SI_private Lib "CoolProp_x64.dll" Alias "Props1SI" (ByVal output As String, ByVal Ref As String) As Double
    Private Declare PtrSafe Function HAPropsSI_private Lib "CoolProp_x64.dll" Alias "HAPropsSI" (ByVal output As String, ByVal Input1Name As String, ByVal Value1 As Double, ByVal Input2Name As String, ByVal Value2 As Double, ByVal Input3name As String, ByVal Value3 As Double) As Double
    Public Declare PtrSafe Function set_config_string Lib "CoolProp_x64.dll" (ByVal key As String, ByVal value As String) As Long
    'DEPRECATED
    Private Declare PtrSafe Function Props_private Lib "CoolProp_x64.dll" Alias "PropsS" (ByVal output As String, ByVal Name1 As Long, ByVal Value1 As Double, ByVal Name2 As Long, ByVal Value2 As Double, ByVal Ref As String) As Double
#Else
    Private Declare PtrSafe Function get_global_param_string_private Lib "CoolProp_stdcall.dll" Alias "_get_global_param_string@12" (ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function get_fluid_param_string_private Lib "CoolProp_stdcall.dll" Alias "_get_fluid_param_string@16" (ByVal param As String, ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function PropsSI_private Lib "CoolProp_stdcall.dll" Alias "_PropsSI@32" (ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String) As Double
    Private Declare PtrSafe Function PhaseSI_private Lib "CoolProp_stdcall.dll" Alias "_PhaseSI@36" (ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String, ByVal output As String, ByVal n As Integer) As Long
    Private Declare PtrSafe Function Props1SI_private Lib "CoolProp_stdcall.dll" Alias "_Props1SI@8" (ByVal output As String, ByVal Ref As String) As Double
    Private Declare PtrSafe Function HAPropsSI_private Lib "CoolProp_stdcall.dll" Alias "_HAPropsSI@40" (ByVal output As String, ByVal Input1Name As String, ByVal Value1 As Double, ByVal Input2Name As String, ByVal Value2 As Double, ByVal Input3name As String, ByVal Value3 As Double) As Double
    Public Declare PtrSafe Function set_config_string Lib "CoolProp_stdcall.dll" Alias "_set_config_string@8" (ByVal key As String, ByVal value As String) As Long
    'DEPRECATED
    Private Declare PtrSafe Function Props_private Lib "CoolProp_stdcall.dll" Alias "_PropsS@32" (ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String) As Double
#End If

' Flag to avoid re-creating the symlink on every call
Private m_CoolPropInitialized As Boolean

' ---------------------------------------------------------------------------
' Procedures start here
' ---------------------------------------------------------------------------

Private Sub EnsureCoolPropLoaded()
#If Mac Then
    If m_CoolPropInitialized Then Exit Sub

    Dim libName As String
    Dim libPath As String
    Dim addInsFolder As String

    addInsFolder = Environ("Home") _
        & "/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/"

    ' -------------------------------------------------------------------
    ' Step 1: Determine if Excel is running as 64-bit or 32-bit.
    '         MAC_OFFICE_VERSION >= 15 with VBA7 indicates 64-bit
    '         Excel 2016+ for Mac. Older versions are 32-bit.
    ' -------------------------------------------------------------------
    #If MAC_OFFICE_VERSION >= 15 And VBA7 Then
        ' 64-bit Excel for Mac -> need to distinguish x86_64 vs arm64
        Dim archResult As String

        ' Tier 1: Try MacScript (works on Excel 2011 and non-sandboxed
        '         Excel 2016). Fails with error 5 in sandboxed Office 365.
        On Error Resume Next
        archResult = MacScript("do shell script ""uname -m""")
        On Error GoTo 0

        archResult = Trim(archResult)

        If archResult <> "arm64" And archResult <> "x86_64" Then
            ' Tier 2: MacScript was blocked by the sandbox.
            ' The dyld shared cache uses architecture-specific filenames:
            '   ARM Macs  -> dyld_shared_cache_arm64e
            '   Intel Macs -> dyld_shared_cache_x86_64h (or x86_64)
            ' This directory is world-readable and present on macOS 11+.
            Dim dyldArm As String
            On Error Resume Next
            dyldArm = Dir("/System/Library/dyld/dyld_shared_cache_arm64e*")
            On Error GoTo 0
            If dyldArm <> "" Then
                archResult = "arm64"
            Else
                archResult = "x86_64"
            End If
        End If

        If archResult = "arm64" Then
            ' Apple Silicon (M1, M2, M3, M4, ...)
            libName = "libCoolProp_arm_64.dylib"
        Else
            ' Intel 64-bit
            libName = "libCoolProp_x86_64.dylib"
        End If
    #Else
        ' 32-bit Excel for Mac (Excel 2011 or earlier) -> always x86 32-bit
        libName = "libCoolProp_x86_32.dylib"
    #End If

    libPath = addInsFolder & libName

    ' -------------------------------------------------------------------
    ' Step 2: Deploy the correct dylib to /tmp/ so all Declare statements
    '         resolve to /tmp/libCoolProp.dylib.
    '         Uses pure VBA FileCopy — no dependency on MacScript,
    '         libSystem.B.dylib, or any shell command.
    ' -------------------------------------------------------------------
    On Error Resume Next
    Kill "/tmp/libCoolProp.dylib"        ' remove stale file/symlink if present
    On Error GoTo 0
    FileCopy libPath, "/tmp/libCoolProp.dylib"

    m_CoolPropInitialized = True
#End If
End Sub

Public Function get_env_variable(ByVal key As String) As String
    Dim script As String
    Dim output As String
    #If Mac Then
        script = "set pp to system attribute """ & key & """" & vbNewLine & "return pp"
        output = MacScript(script)
    #Else
        output = Environ(key)
    #End If
    get_env_variable = output
End Function

Private Function get_error_message() As String
    Dim errstring As String
    'Make a null-terminated string that is plenty big
    errstring = String(2000, vbNullChar)
    EnsureCoolPropLoaded
    'Get the error string
    Call get_global_param_string_private("errstring", errstring, 2000)
    get_error_message = errstring
End Function

Public Function get_global_param_string(ByVal output As String) As String
    Dim strParam As String
    'Make a null-terminated string that is plenty big
    strParam = String(2000, vbNullChar)
    EnsureCoolPropLoaded
    'Get the version string
    Call get_global_param_string_private(output, strParam, 2000)
    get_global_param_string = strParam
End Function

Public Function PropsSI(ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal fluid As String)
    On Error GoTo ErrorHandler
    Dim PropsSI_temp As Double

    EnsureCoolPropLoaded
    PropsSI_temp = PropsSI_private(output, Name1, Value1, Name2, Value2, fluid)
    
    If Abs(PropsSI_temp) > 1E+30 Then
        'Return error message
        PropsSI = get_error_message()
    Else
        PropsSI = PropsSI_temp
    End If
    Exit Function
    
ErrorHandler:
    If Err = 13 Then
     Exit Function
    End If
    
    MsgBox "The most recent error number is " & Err & ". Its message text is: " & Error(Err)
    Exit Function
End Function

Public Function PhaseSI(ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal fluid As String)
    Dim strPhase As String
    'Make a null-terminated string that is plenty big
    strPhase = String(2000, vbNullChar)
    EnsureCoolPropLoaded
    'Call PhaseSI_private - any errors will return an empty phase string
    Call PhaseSI_private(Name1, Value1, Name2, Value2, fluid, strPhase, 2000)
    If Len(strPhase) = 0 Then
        PhaseSI = get_error_message()
    Else
        PhaseSI = strPhase
    End If
    Exit Function
End Function

Public Function Props1SI(ByVal fluid As String, ByVal output As String)
    On Error GoTo ErrorHandler
    Dim Props1SI_temp As Double

    EnsureCoolPropLoaded
    Props1SI_temp = Props1SI_private(output, fluid)
    
    If Abs(Props1SI_temp) > 1E+30 Then
        'Display the error
        Props1SI = get_error_message()
    Else
        Props1SI = Props1SI_temp
    End If
    Exit Function
    
ErrorHandler:
    If Err = 13 Then
     Exit Function
    End If
    MsgBox "The most recent error number is " & Err & ". Its message text is: " & Error(Err)
    Exit Function
End Function

Public Function Props(ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal fluid As String)
    On Error GoTo ErrorHandler
    Dim Props_temp As Double
    EnsureCoolPropLoaded
    Props_temp = Props_private(output, Name1, Value1, Name2, Value2, fluid)
    If Abs(Props_temp) > 1E+30 Then
        Props = get_error_message()
    Else
        Props = Props_temp
    End If
    Exit Function
ErrorHandler:
    If Err = 13 Then
     Exit Function
    End If
    MsgBox "The most recent error number is " & Err & ". Its message text is: " & Error(Err)
    Exit Function
End Function

Public Function HAPropsSI(ByVal output As String, ByVal Input1Name As String, ByVal Value1 As Double, ByVal Input2Name As String, ByVal Value2 As Double, ByVal Input3name As String, ByVal Value3 As Double) As Double
    On Error GoTo ErrorHandler
    Dim HAPropsSI_temp As Double
    EnsureCoolPropLoaded
    HAPropsSI_temp = HAPropsSI_private(output, Input1Name, Value1, Input2Name, Value2, Input3name, Value3)
    If Abs(HAPropsSI_temp) > 1E+30 Then
        HAPropsSI = get_error_message()
    Else
        HAPropsSI = HAPropsSI_temp
    End If
    Exit Function
ErrorHandler:
    If Err = 13 Then
     Exit Function
    End If
    MsgBox "The most recent error number is " & Err & ". Its message text is: " & Error(Err)
    Exit Function
End Function

Public Function MixtureString(names As Range, fractions As Range) As String
    ' See http://www.functionx.com/vbaexcel/objects/Lesson6.htm
    Dim my_names, my_fractions As Collection
    Dim Cell, i As Variant
    Dim chunk As String
    MixtureString = ""
    Set my_names = New Collection
    Set my_fractions = New Collection
    
    ' Collect all the names
    For Each Cell In names
        my_names.Add Cell.value
    Next
    ' Collect all the fractions
    For Each Cell In fractions
        my_fractions.Add Cell.value
    Next
    ' Zip them back together
    For i = 1 To my_fractions.Count
        chunk = my_names.Item(i) & "[" & LTrim(Str(my_fractions.Item(i))) & "]"
        If i = 1 Then
            MixtureString = MixtureString & chunk
        Else
            MixtureString = MixtureString & "&" & chunk
        End If
    Next

End Function
