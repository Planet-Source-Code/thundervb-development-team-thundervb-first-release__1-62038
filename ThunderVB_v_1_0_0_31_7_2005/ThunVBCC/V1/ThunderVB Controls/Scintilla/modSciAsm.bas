Attribute VB_Name = "modSciAsm"
'Asm-C functions for The sciEdit control ..
'The control compiles with or without inlineAsm/c enabled but if enabled it is much faster
Public Function DebugerExeption() As Long
'#asm'  int 3
End Function

Public Function CallBP(ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long, ByVal p5 As Long) As Long
'#c'int CallBP(int p1,int p2,int p3,int p4,int (__cdecl*fn)(int,int,int,int)){
'#c'if (fn==0) return 0;
'#c'return fn(p1,p2,p3,p4);
'#c'}
End Function

Public Function CCode() As Boolean
'#c'int CCode(){
'#c'return 1 ;
'#c'}
End Function
