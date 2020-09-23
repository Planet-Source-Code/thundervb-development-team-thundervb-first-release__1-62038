Attribute VB_Name = "modEip2Mod"
Option Explicit

Private Type MODULEENTRY32
  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
  szModule As String * 256
  szExePath As String * 260
End Type

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Const TH32CS_SNAPMODULE As Long = &H8

Public Function Eip2Mod(ByVal EIP As Long) As String
Dim tMod As MODULEENTRY32, lProcessID As Long, n As Long
Dim hSnapshot As Long
    
    lProcessID = GetCurrentProcessId
    hSnapshot = CreateToolhelp32Snapshot(8, 0)
    tMod.dwSize = Len(tMod)
    n = Module32First(hSnapshot, tMod)
    
    Do While n
        
        If EIP >= tMod.modBaseAddr And EIP <= (tMod.modBaseAddr + tMod.modBaseSize) Then
            Eip2Mod = Left(tMod.szModule, InStr(tMod.szModule, Chr(0)) - 1)
            GoTo 10
        End If
        
        n = Module32Next(hSnapshot, tMod)
        
    Loop
    
10:
    CloseHandle hSnapshot
    
End Function
