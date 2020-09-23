Attribute VB_Name = "modCHeaderGen"
Option Explicit

Public Enum argPassMode_enum
    argError = 0
    argByval = 1
    argByRef = 2
End Enum

Public Type ArgList_entry
    argName As String               'name
    argType As String               'Type
    argPassMode As argPassMode_enum 'Pass mode [byref/byval]
    argSize As Long                 'Size [in bytes] of the variable,4 if pointer [byref]
End Type

Public Type ArgList
    items() As ArgList_entry
    count As Long
End Type

Public Type fctInfo
    retValue As ArgList_entry
    args As ArgList
    name As String
    inmod As String
End Type


'creates the asm cross referenceses and a c .h file needed to call vb functions from asm/c
Public Sub CreateAsmCHeader(ByRef asmh As String, ByRef ch As String)
Dim modules() As String, functions() As String
Dim i As Long, i2 As Long, cheader As String_B
Dim asmheader As String_B, temp As fctInfo

    modules = EnumModuleNames(vbext_ct_StdModule)
    AppendString cheader, "#define vb_bool int" & vbNewLine
    AppendString cheader, "#define vb_true -1" & vbNewLine
    AppendString cheader, "#define vb_false 0" & vbNewLine
    
    For i = 0 To ArrUBound(modules)
        functions = EnumFunctionNames(modules(i))
        For i2 = 0 To ArrUBound(functions)
           temp = ParseVBFunctionDeclatation(GetFunctionCode(modules(i), functions(i2)))
           temp.inmod = modules(i)
           If Len(temp.name) > 0 Then
               AppendString cheader, GetCExternDecl(temp.inmod, temp.name, temp.args, temp.retValue) & vbNewLine
               AppendString asmheader, GetAsmTableRef(temp.inmod, temp.name, temp.args) & vbNewLine & vbNewLine
           End If
        Next i2
    Next i
    
    asmh = GetString(asmheader)
    ch = GetString(cheader)
    
End Sub

Public Function ParseVBFunctionDeclatation(functi As String) As fctInfo
    Dim temp As fctInfo
    Dim t As String
    Dim meth_mode As Long
    Dim funct As String
    
    'Get the function declaration
    funct = Split(Replace$(functi, "_" & vbNewLine, " "), vbNewLine, 2)(0) ', "", "")
    
    t = LCase$(GetFirstWord(funct))
    If t = "public" Then
        RemFisrtWord funct: funct = Trim$(funct)
    ElseIf t = "private" Then
       RemFisrtWord funct: funct = Trim$(funct)
    Else
        'nothing needed
    End If
    
    t = LCase$(GetFirstWord(funct))
    
    If t = "function" Then
        RemFisrtWord funct: funct = Trim$(funct)
        meth_mode = 1
    ElseIf t = "sub" Then
        RemFisrtWord funct: funct = Trim$(funct)
        meth_mode = 2
    End If
    
    
    funct = Replace$(funct, "(", " ( ")
    funct = Replace$(funct, ")", " ) ")
    
    temp.name = GetFirstWord(funct)
    RemFisrtWord funct: funct = Trim$(funct)
    RemFisrtWord funct: funct = Trim$(funct)
    t = GetFirstWord(funct) ' the "("
    
    funct = Replace$(funct, "() as", "_arr as", , , vbTextCompare)
    
    If ((InStr(2, funct, "optional", vbTextCompare) <> 0) Or (InStr(2, funct, "paramarray", vbTextCompare) <> 0)) Then
        Do While t <> ")"
            RemFisrtWord funct: funct = Trim$(funct)
            t = LCase$(GetFirstWord(funct))
        Loop
    Else
        Do While t <> ")"
            Dim targ As ArgList_entry
            t = LCase$(GetFirstWord(funct))
            
            If t = "byval" Then
                targ.argPassMode = argByval
                RemFisrtWord funct: funct = Trim$(funct)
                t = GetFirstWord(funct)
            ElseIf t = "byref" Then
                targ.argPassMode = argByRef
                RemFisrtWord funct: funct = Trim$(funct)
                t = GetFirstWord(funct)
            Else
                targ.argPassMode = argByval
            End If
            
            targ.argName = t
            
            RemFisrtWord funct: funct = Trim$(funct)
            t = LCase$(GetFirstWord(funct))
            If t = "as" Then
                RemFisrtWord funct: funct = Trim$(funct)
                t = LCase$(GetFirstWord(funct))
                targ.argType = t
                RemFisrtWord funct: funct = Trim$(funct)
                t = LCase$(GetFirstWord(funct))
            End If
            
            targ = CVarTypeFromVB(targ)
            ReDim Preserve temp.args.items(temp.args.count)
            temp.args.items(temp.args.count) = targ
            temp.args.count = temp.args.count + 1
        Loop
    End If
    
    RemFisrtWord funct: funct = Trim$(funct)
    
    If meth_mode = 1 Then
        RemFisrtWord funct: funct = Trim$(funct)
        
        Dim t2 As ArgList_entry
        t2.argName = ""
        t2.argPassMode = argByval
        t2.argType = GetFirstWord(funct)
        temp.retValue = CVarTypeFromVB(t2)
    Else
        temp.retValue.argType = "void"
        temp.retValue.argPassMode = argByval
        temp.retValue.argSize = 0
    End If
    
    ParseVBFunctionDeclatation = temp
End Function

Public Function CVarTypeFromVB(var As ArgList_entry) As ArgList_entry
    Dim temp As ArgList_entry
    temp.argName = var.argName
    temp.argPassMode = var.argPassMode
    
    Select Case LCase$(var.argType)
        
        Case "long"
            temp.argType = "int"
            temp.argSize = 4
            
        Case "boolean"
            temp.argType = "vb_bool"
            temp.argSize = 4
        
        Case "integer"
            temp.argType = "short"
            temp.argSize = 2
        
        Case "byte"
            temp.argType = "char"
            temp.argSize = 1
        
        Case "single"
            temp.argType = "float"
            temp.argSize = 4
            
        Case "double"
            temp.argType = "double"
            temp.argSize = 8
        
        Case "currency"
            temp.argType = "double" ' not totaly correct
            temp.argSize = 8        'but both are 8 bytes ;)
        
        Case Else
            
            If temp.argPassMode = argByRef Then         'udt : byref allways -> 4 bytes [void*] or enum : byref/byval allways -> 4 bytes [int/void*] or
                temp.argType = "void"                   'class : byref allways -> 4 bytes [void*]
                temp.argSize = 4
            ElseIf temp.argPassMode = argByval Then
                temp.argType = "int"                    'this can be olny enum : byval allways -> 4 bytes [int] or
                temp.argSize = 4
            Else
                MsgBoxX "Error;variable pass mode not detected"
            End If
            
    End Select
    
    If temp.argPassMode = argByRef Then
        temp.argSize = 4 'pointer...
    End If
    
    CVarTypeFromVB = temp
    
End Function

Public Function GetCExternDecl(inmod As String, infunct As String, from As ArgList, retval As ArgList_entry) As String
    Dim i As Long, temp As String_B
    
    AppendString temp, "//" & inmod & ":" & infunct & vbNewLine
    AppendString temp, "extern " & GetCVarString(retval, False) & " " & inmod & "_" & infunct & "("
    
    For i = 0 To from.count - 1
        AppendString temp, GetCVarString(from.items(i)) & IIf(i = (from.count - 1), "", ",")
    Next i
    
    AppendString temp, ");" & vbNewLine
    
    GetCExternDecl = GetString(temp)
    
End Function

Public Function GetCVarString(from As ArgList_entry, Optional EmitName As Boolean = True) As String
    Dim t As String_B
    
    AppendString t, from.argType & IIf(from.argPassMode = argByRef, "* ", " ")
    
    If EmitName Then
        AppendString t, from.argName
    End If
    
    GetCVarString = GetString(t)
    
End Function

Public Function GetAsmTableRef(inmod As String, infunct As String, from As ArgList) As String
    Dim i As Long, sz As Long, temp As String
    
    For i = 0 To from.count - 1
        sz = sz + from.items(i).argSize
    Next i
    
    'Create the extern decl, and declare the c extern, and write the code for the jump..
    'when c calls the _fff_mmm@y it jumps here , and from here we jump to the VB function ;)
    'Asm code can call both the ?fff@mmm@@AAGXXZ:NEAR or the _fff_mmm@y to archive the same
    ' result altho the ?fff@mmm@@AAGXXZ:NEAR is a bit faster ;)
    temp = "; " & inmod & ":" & infunct & vbNewLine & _
           " EXTRN   ?" & infunct & "@" & inmod & "@@AAGXXZ:NEAR" & vbNewLine & _
           "_TEXT   SEGMENT" & vbNewLine & _
           " PUBLIC  _" & inmod & "_" & infunct & "@" & sz & vbNewLine & _
           " _" & inmod & "_" & infunct & "@" & sz & ": jmp ?" & infunct & "@" & inmod & "@@AAGXXZ" & vbNewLine & _
           "_TEXT   ENDS"
           
    GetAsmTableRef = temp
    
End Function
