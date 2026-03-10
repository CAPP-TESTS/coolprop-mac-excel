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
    ' see http://stackoverflow.com/a/39821415/1360263 for info on interrogating the version of Excel
    #If MAC_OFFICE_VERSION >= 15 And VBA7 Then ' 64-bit Excel 2016 for Mac
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
    #Else ' 32-bit Excel for Mac
        ' Even though the functions are exported with a leading underscore, Excel 2011 for Mac doesn't want the leading underscore as part of name
        Private Declare PtrSafe Function PropsSI_private Lib "/tmp/libCoolProp_32bit.dylib" Alias "PropsSI" (ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String) As Double
        Private Declare PtrSafe Function PhaseSI_private Lib "/tmp/libCoolProp_32bit.dylib" Alias "PhaseSI" (ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal Ref As String, ByVal output As String, ByVal n As Integer) As Long
        Private Declare PtrSafe Function Props1SI_private Lib "/tmp/libCoolProp_32bit.dylib" Alias "Props1SI" (ByVal output As String, ByVal Ref As String) As Double
        Private Declare PtrSafe Function get_global_param_string_private Lib "/tmp/libCoolProp_32bit.dylib" Alias "get_global_param_string" (ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
        Private Declare PtrSafe Function get_fluid_param_string_private Lib "/tmp/libCoolProp_32bit.dylib" Alias "get_fluid_param_string" (ByVal fluid As String, ByVal param As String, ByVal output As String, ByVal n As Integer) As Long
        Private Declare PtrSafe Function HAPropsSI_private Lib "/tmp/libCoolProp_32bit.dylib" Alias "HAPropsSI" (ByVal output As String, ByVal Input1Name As String, ByVal Value1 As Double, ByVal Input2Name As String, ByVal Value2 As Double, ByVal Input3name As String, ByVal Value3 As Double) As Double
        Public Declare PtrSafe Function set_config_string Lib "/tmp/libCoolProp_32bit.dylib" (ByVal key As String, ByVal value As String) As Long
        'DEPRECATED
        Private Declare PtrSafe Function Props_private Lib "/tmp/libCoolProp_32bit.dylib" Alias "PropsS" (ByVal output As String, ByVal Name1 As Long, ByVal Value1 As Double, ByVal Name2 As Long, ByVal Value2 As Double, ByVal Ref As String) As Double
    #End If

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
    'Get the error string
    Call get_global_param_string_private("errstring", errstring, 2000)
    get_error_message = errstring
End Function

Public Function get_global_param_string(ByVal output As String) As String
    Dim strParam As String
    'Make a null-terminated string that is plenty big
    strParam = String(2000, vbNullChar)
    'Get the version string
    Call get_global_param_string_private(output, strParam, 2000)
    get_global_param_string = strParam
End Function

Public Function PropsSI(ByVal output As String, ByVal Name1 As String, ByVal Value1 As Double, ByVal Name2 As String, ByVal Value2 As Double, ByVal fluid As String)
    On Error GoTo ErrorHandler
    Dim PropsSI_temp As Double

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
