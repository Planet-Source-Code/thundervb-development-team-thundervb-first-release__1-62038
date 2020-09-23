Attribute VB_Name = "ErrorHanlder_asm"
Option Explicit

Private Sub asm_shits()
'#asm' option NOSCOPED
'#asm' exp dd 0
'#asm' exp2 dd 0
'#asm' extern ?ExeptH_bef@ErrorHanlder_vb@@AAGXXZ:NEAR
'#asm' extern ?ExeptH_aft@ErrorHanlder_vb@@AAGXXZ:NEAR
End Sub


Public Function ExeptH() As Long
'#asm'  push ebp
'#asm'  mov ebp,esp
'#asm'  mov eax , [ebp+8+4]
'#asm'  push [ebp+16+4]
'#asm'  push [ebp+12+4]
'#asm'  push [ebp+8+4]
'#asm'  push [ebp+4+4]
'#asm'  push [eax+14h]  ;vb's error handler id :)
'#asm'  call ?ExeptH_bef@ErrorHanlder_vb@@AAGXXZ
'#asm'
'#asm'  mov exp2 ,0
'#asm'  cmp eax,0
'#asm'  je NoVBh
'#asm'  mov exp2 ,1
'#asm'  mov eax ,exp
'#asm'  mov esp,ebp  ;restore epb and esp
'#asm'  pop ebp      ;restore epb and esp
'#asm'  jmp eax
'#asm'  NoVBh:
'#asm'
'#asm'  mov eax ,[exp2]
'#asm'  push [ebp+16+4]
'#asm'  push [ebp+12+4]
'#asm'  push [ebp+8+4]
'#asm'  push [ebp+4+4]
'#asm'  push eax
'#asm'  call ?ExeptH_aft@ErrorHanlder_vb@@AAGXXZ
'#asm'
End Function

Public Sub InitExept(ByVal val As Long)
'#asm' mov eax,[esp+4]
'#asm' mov exp,eax
End Sub

Public Function IsAsmOn() As Boolean
'#asm' mov eax,1
'#asm' ret
End Function


Public Sub Int03()
'#asm' int 3
End Sub

Public Sub Try2Restore(ByVal context As Long)
'#asm' _sEsp$ = -732                       ; size = 4
'#asm' _sEbp$ = -728                       ; size = 4
'#asm' _sEip$ = -724                       ; size = 4
'#asm' _cnt_new$ = -720                    ; size = 716
'#asm' _cnt$ = 8                       ; size = 4
'#asm'
'#asm' ; 3    :     {
'#asm'
'#asm'     push Ebp
'#asm'     mov Ebp, esp
'#asm'     sub esp, 732                ; 000002dcH
'#asm'     push Esi
'#asm'     push Edi
'#asm'
'#asm' ; 4    :     int sEbp;
'#asm' ; 5    :     int sEip;
'#asm' ; 6    :     int sEsp;
'#asm' ; 7    :
'#asm' ; 8    :     CONTEXT cnt_new=*cnt;
'#asm'
'#asm'     mov esi, DWORD PTR _cnt$[ebp]
'#asm'     mov ecx, 179                ; 000000b3H
'#asm'     lea edi, DWORD PTR _cnt_new$[ebp]
'#asm'     rep movsd
'#asm'
'#asm' ; 9    :     cnt_new.Ebp=*(int*)cnt->Ebp;
'#asm'
'#asm'     mov eax, DWORD PTR _cnt$[ebp]
'#asm'     mov ecx, DWORD PTR [eax+180]
'#asm'     mov edx, DWORD PTR [ecx]
'#asm'     mov DWORD PTR _cnt_new$[ebp+180], edx
'#asm'
'#asm' ; 10   :     cnt_new.Eip=*(int*)(((char*)cnt->Ebp)+4);
'#asm'
'#asm'     mov eax, DWORD PTR _cnt$[ebp]
'#asm'     mov ecx, DWORD PTR [eax+180]
'#asm'     mov edx, DWORD PTR [ecx+4]
'#asm'     mov DWORD PTR _cnt_new$[ebp+184], edx
'#asm'
'#asm' ; 11   :     cnt_new.Esp=cnt->Ebp;
'#asm'
'#asm'     mov eax, DWORD PTR _cnt$[ebp]
'#asm'     mov ecx, DWORD PTR [eax+180]
'#asm'     mov DWORD PTR _cnt_new$[ebp+196], ecx
'#asm'
'#asm' ; 12   :
'#asm' ; 13   :     sEbp=cnt_new.Ebp;
'#asm'
'#asm'     mov edx, DWORD PTR _cnt_new$[ebp+180]
'#asm'     mov DWORD PTR _sEbp$[ebp], edx
'#asm'
'#asm' ; 14   :     sEip=cnt_new.Eip;
'#asm'
'#asm'     mov eax, DWORD PTR _cnt_new$[ebp+184]
'#asm'     mov DWORD PTR _sEip$[ebp], eax
'#asm'
'#asm' ; 15   :     sEsp=cnt_new.Esp;
'#asm'
'#asm'     mov ecx, DWORD PTR _cnt_new$[ebp+196]
'#asm'     mov DWORD PTR _sEsp$[ebp], ecx
'#asm'
'#asm' ; 16   :
'#asm' ; 17   :     __asm
'#asm' ; 18   :         {
'#asm' ; 19   :         mov ebp,sEbp
'#asm'
'#asm'     mov ebp, DWORD PTR _sEbp$[ebp]
'#asm'
'#asm' ; 20   :         mov Esp,sEsp
'#asm'
'#asm'     mov esp, DWORD PTR _sEsp$[ebp]
'#asm'
'#asm' ; 21   :         mov eax,sEip
'#asm'
'#asm'     mov eax, DWORD PTR _sEip$[ebp]
'#asm'
'#asm' ; 22   :
'#asm' ; 23   :         jmp eax
'#asm'
'#asm'     jmp Eax
'#asm'
'#asm' ; 24   :         }
'#asm' ; 25   :     }
'#asm'     pop Edi
'#asm'     pop Esi
'#asm'     mov esp, Ebp
'#asm'     pop Ebp
'#asm'     ret 4
End Sub

Public Function GetEip(ByVal context As Long) As Long
'#asm' _cnt$ = 8                       ; size = 4
'#asm'
'#asm'
'#asm' ; 5    :     return cnt->Eip;
'#asm'
'#asm'    mov eax, DWORD PTR _cnt$[esp-4]
'#asm'    mov eax, DWORD PTR [eax+184]
'#asm'
'#asm' ; 6    :     }
'#asm'
'#asm'    ret 4
End Function

'for now this does nothing .. that's why # is missing from 'c'
Public Sub CallStackDump(ByVal context As Long)
'c'lib=kernel32.lib
'c'    #include <windows.h>
'c'    void CallStackDump(CONTEXT *cnt)
'c'    {
'c'    int sEbp;
'c'    int sEip;
'c'    int sEsp;
'c'    //sEbp=*(int*)cnt->Ebp;
'c'    //sEip=*(int*)(((char*)cnt->Ebp)-4);
'c'    //sEsp=cnt->Ebp;
'c'    sEbp=cnt->Ebp;
'c'    sEip=cnt->Eip;
'c'    sEsp=cnt->Esp;
'c'
'c'    ErrorHanlder_vb_CallStackStart();
'c'    ErrorHanlder_vb_CallStackAdd(sEip,sEbp,sEsp);
'c'
'c'
'c'    while ((!IsBadReadPtr((void*)sEbp,4)) && (!IsBadReadPtr( (void*)(sEbp+4) ,4)))
'c'    {
'c'
'c'        sEip=*(int*)(((char*)sEbp)+4);
'c'        sEsp=sEbp;
'c'        sEbp=*(int*)sEbp;
'c'        ErrorHanlder_vb_CallStackAdd(sEip,sEbp,sEsp);
'c'
'c'    }
'c'    ErrorHanlder_vb_CallStackEnd();
'c'
'c'    }
End Sub
        'Dim cnt As CONTEXT
        'Dim newcnt As CONTEXT
        'CopyMemory cnt, ByVal val3, Len(cnt)
        '
        ''unrol the tack up by one ..
        'CopyMemory newcnt, cnt, Len(cnt)     'get the stored ebp
        'CopyMemory newcnt.Ebp, ByVal cnt.Ebp, 4     'get the stored ebp
        'CopyMemory newcnt.Eip, ByVal cnt.Ebp + 4, 4 'get the stored pc
        'newcnt.Esp = cnt.Ebp                 'restore esp
        ''the *bad* thing is that we canot restore all else things back ..
        'SetThreadContext GetCurrentThread(), newcnt
'C code
'c'lib=kernel32.lib
'c'    #include <windows.h>
'c'    void Try2Restore(CONTEXT *cnt)
'c'    {
'c'    int sEbp;
'c'    int sEip;
'c'    int sEsp;
'c'
'c'    CONTEXT cnt_new=*cnt;
'c'    cnt_new.Ebp=*(int*)cnt->Ebp;
'c'    cnt_new.Eip=*(int*)(((char*)cnt->Ebp)+4);
'c'    cnt_new.Esp=cnt->Ebp;
'c'
'c'    sEbp=cnt_new.Ebp;
'c'    sEip=cnt_new.Eip;
'c'    sEsp=cnt_new.Esp;
'c'
'c'    __asm
'c'        {
'c'        mov ebp,sEbp
'c'        mov Esp,sEsp
'c'        mov eax,sEip
'c'
'c'        jmp eax
'c'        }
'c'    }


'c code for Get Eip
'c'lib=kernel32.lib
'c'    #include <windows.h>
'c'    int GetEip(CONTEXT *cnt)
'c'    {
'c'    return cnt->Eip;
'c'    }

