'VA: 4043D4
Private Declare Function DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
'VA: 40438C
Private Declare Function SelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long
'VA: 404344
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
'VA: 4042F4
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'VA: 4042B0
Private Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'VA: 4038F8
Private Declare Sub GetSystemTime Lib "kernel32" Alias "GetSystemTime" (lpSystemTime As SYSTEMTIME)
'VA: 4038B0
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'VA: 403864
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
'VA: 40382C
Private Declare Function FileTimeToSystemTime Lib "kernel32" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'VA: 4037DC
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'VA: 403794
Private Declare Function FindClose Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Long

Private sub Proc_0_0_40D930
  0040D930: push ebp
  0040D931: mov ebp, esp
  0040D933: sub esp, 00000008h
  0040D936: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0040D93B: mov eax, fs:[00h]
  0040D941: push eax
  0040D942: mov fs:[00000000h], esp
  0040D949: sub esp, 000003ACh
  0040D94F: push ebx
  0040D950: push esi
  0040D951: push edi
  0040D952: mov var_8, esp
  0040D955: mov var_4, 00401378h
  0040D95C: mov ecx, 00000094h
  0040D961: xor eax, eax
  0040D963: lea edi, var_264
  0040D969: mov edx, arg_8
  0040D96C: rep stosd 
  0040D96E: mov ecx, 00000050h
  0040D973: lea edi, var_3B8
  0040D979: rep stosd 
  0040D97B: lea ecx, var_268
  0040D981: mov var_268, eax
  0040D987: mov var_270, eax
  0040D98D: call MSVBVM60.DLL.__vbaStrCopy
  0040D993: mov eax, var_268
  0040D999: push 00000001h
  0040D99B: push eax
  0040D99C: call [00401320h] ; Right$(arg_1, arg_2)
  0040D9A2: mov edi, MSVBVM60.DLL.__vbaStrMove
  0040D9A8: mov edx, eax
  0040D9AA: lea ecx, var_270
  0040D9B0: call edi
  0040D9B2: push eax
  0040D9B3: push 0040393Ch
  0040D9B8: call MSVBVM60.DLL.__vbaStrCmp
  0040D9BE: mov ecx, var_268
  0040D9C4: mov ebx, [00401030h] ; Len(arg_1)
  0040D9CA: mov esi, eax
  0040D9CC: push ecx
  0040D9CD: neg esi
  0040D9CF: sbb esi, esi
  0040D9D1: inc esi
  0040D9D2: neg esi
  0040D9D4: call ebx
  0040D9D6: xor edx, edx
  0040D9D8: cmp eax, 00000003h
  0040D9DB: setnle dl
  0040D9DE: neg edx
  0040D9E0: lea ecx, var_270
  0040D9E6: and esi, edx
  0040D9E8: call MSVBVM60.DLL.__vbaFreeStr
  0040D9EE: test si, si
  0040D9F1: jz 40DA1Dh
  0040D9F3: mov eax, var_268
  0040D9F9: push eax
  0040D9FA: call ebx
  0040D9FC: mov ecx, var_268
  0040DA02: sub eax, 00000001h
  0040DA05: jo 0040DAE7h
  0040DA0B: push eax
  0040DA0C: push ecx
  0040DA0D: call [004012FCh] ; Left$(arg_1, arg_2)
  0040DA13: mov edx, eax
  0040DA15: lea ecx, var_268
  0040DA1B: call edi
  0040DA1D: lea edx, var_264
  0040DA23: lea eax, var_3B8
  0040DA29: push edx
  0040DA2A: push eax
  0040DA2B: push 00403734h
  0040DA30: call MSVBVM60.DLL.__vbaRecUniToAnsi
  0040DA36: mov ecx, var_268
  0040DA3C: push eax
  0040DA3D: lea edx, var_270
  0040DA43: push ecx
  0040DA44: push edx
  0040DA45: call MSVBVM60.DLL.__vbaStrToAnsi
  0040DA4B: push eax
  0040DA4C: FindFirstFile(%x1, %x2)
  0040DA51: mov edi, MSVBVM60.DLL.__vbaSetSystemError
  0040DA57: mov esi, eax
  0040DA59: call edi
  0040DA5B: lea eax, var_3B8
  0040DA61: lea ecx, var_264
  0040DA67: push eax
  0040DA68: push ecx
  0040DA69: push 00403734h
  0040DA6E: call MSVBVM60.DLL.__vbaRecAnsiToUni
  0040DA74: mov edx, var_270
  0040DA7A: lea eax, var_268
  0040DA80: push edx
  0040DA81: push eax
  0040DA82: call MSVBVM60.DLL.__vbaStrToUnicode
  0040DA88: lea ecx, var_270
  0040DA8E: call MSVBVM60.DLL.__vbaFreeStr
  0040DA94: xor ecx, ecx
  0040DA96: cmp esi, FFFFFFFFh
  0040DA99: setnz cl
  0040DA9C: neg ecx
  0040DA9E: push esi
  0040DA9F: mov var_26C, ecx
  0040DAA5: FindClose(%x1)
  0040DAAA: call edi
  0040DAAC: push 0040DACDh ; "‹Mðf‹…”ýÿÿ_^d‰'#1"
  0040DAB1: jmp 40DAC0h
  0040DAB3: lea ecx, var_270
  0040DAB9: call MSVBVM60.DLL.__vbaFreeStr
  0040DABF: ret 
End Sub

Private sub Proc_0_1_40DAF0
  0040DAF0: push ebp
  0040DAF1: mov ebp, esp
  0040DAF3: sub esp, 0000000Ch
  0040DAF6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0040DAFB: mov eax, fs:[00h]
  0040DB01: push eax
  0040DB02: mov fs:[00000000h], esp
  0040DB09: sub esp, 00000094h
  0040DB0F: push ebx
  0040DB10: push esi
  0040DB11: push edi
  0040DB12: mov var_C, esp
  0040DB15: mov var_8, 00401388h
  0040DB1C: xor esi, esi
  0040DB1E: lea eax, var_5C
  0040DB21: push esi
  0040DB22: push 004039FCh ; "Scripting.FileSystemObject"
  0040DB27: push eax
  0040DB28: mov var_28, esi
  0040DB2B: mov var_38, esi
  0040DB2E: mov var_48, esi
  0040DB31: mov var_4C, esi
  0040DB34: mov var_5C, esi
  0040DB37: mov var_6C, esi
  0040DB3A: call [00401210h] ; arg_1 = CreateObject(arg_2, arg_3)
  0040DB40: mov ebx, MSVBVM60.DLL.__vbaVarSetVar
  0040DB46: lea ecx, var_5C
  0040DB49: lea edx, var_38
  0040DB4C: push ecx
  0040DB4D: push edx
  0040DB4E: call ebx
  0040DB50: mov edi, arg_10
  0040DB53: lea eax, var_6C
  0040DB56: push esi
  0040DB57: push eax
  0040DB58: mov var_64, edi
  0040DB5B: mov var_6C, 00004008h
  0040DB62: call [0040123Ch] ; Dir
  0040DB68: mov edx, eax
  0040DB6A: lea ecx, var_4C
  0040DB6D: call MSVBVM60.DLL.__vbaStrMove
  0040DB73: push eax
  0040DB74: push 00403558h
  0040DB79: call MSVBVM60.DLL.__vbaStrCmp
  0040DB7F: mov esi, eax
  0040DB81: lea ecx, var_4C
  0040DB84: neg esi
  0040DB86: sbb esi, esi
  0040DB88: neg esi
  0040DB8A: neg esi
  0040DB8C: call MSVBVM60.DLL.__vbaFreeStr
  0040DB92: test si, si
  0040DB95: jz 40DBC2h
  0040DB97: mov ecx, [edi]
  0040DB99: lea edx, var_6C
  0040DB9C: mov var_64, ecx
  0040DB9F: lea ecx, var_5C
  0040DBA2: mov var_6C, 00000008h
  0040DBA9: call MSVBVM60.DLL.__vbaVarDup
  0040DBAF: lea edx, var_5C
  0040DBB2: push edx
  0040DBB3: call [00401154h] ; Kill(arg_1)
  0040DBB9: lea ecx, var_5C
  0040DBBC: call MSVBVM60.DLL.__vbaFreeVar
  0040DBC2: sub esp, 00000010h
  0040DBC5: mov eax, edi
  0040DBC7: mov edi, esp
  0040DBC9: mov ecx, 00004008h
  0040DBCE: mov var_6C, ecx
  0040DBD1: sub esp, 00000010h
  0040DBD4: mov [edi], ecx
  0040DBD6: mov ecx, var_68
  0040DBD9: mov var_64, eax
  0040DBDC: mov esi, 0000000Bh
  0040DBE1: mov [edi+04h], ecx
  0040DBE4: mov ecx, esp
  0040DBE6: or edx, FFFFFFFFh
  0040DBE9: push 00000002h
  0040DBEB: mov [edi+08h], eax
  0040DBEE: mov eax, var_60
  0040DBF1: push 00403A34h ; "CreateTextFile"
  0040DBF6: mov [edi+0Ch], eax
  0040DBF9: mov eax, var_78
  0040DBFC: mov [ecx], esi
  0040DBFE: mov [ecx+04h], eax
  0040DC01: lea eax, var_38
  0040DC04: push eax
  0040DC05: mov [ecx+08h], edx
  0040DC08: mov edx, var_70
  0040DC0B: mov [ecx+0Ch], edx
  0040DC0E: lea ecx, var_5C
  0040DC11: push ecx
  0040DC12: call MSVBVM60.DLL.__vbaVarLateMemCallLd
  0040DC18: add esp, 00000030h
  0040DC1B: lea edx, var_48
  0040DC1E: push eax
  0040DC1F: push edx
  0040DC20: call ebx
  0040DC22: mov ecx, arg_C
  0040DC25: lea eax, var_A8
  0040DC2B: push eax
  0040DC2C: call MSVBVM60.DLL.__vbaVargVar
  0040DC32: mov edx, [eax]
  0040DC34: sub esp, 00000010h
  0040DC37: mov ecx, esp
  0040DC39: mov esi, [00401174h]
  0040DC3F: push 00000001h
  0040DC41: push 00403A54h ; "WriteLine"
  0040DC46: mov [ecx], edx
  0040DC48: mov edx, [eax+04h]
  0040DC4B: mov [ecx+04h], edx
  0040DC4E: mov edx, [eax+08h]
  0040DC51: mov eax, [eax+0Ch]
  0040DC54: mov [ecx+08h], edx
  0040DC57: mov [ecx+0Ch], eax
  0040DC5A: lea ecx, var_48
  0040DC5D: push ecx
  0040DC5E: call MSVBVM60.DLL.__vbaObjVar
  0040DC60: mov edi, MSVBVM60.DLL.__vbaLateMemCall
  0040DC66: push eax
  0040DC67: call edi
  0040DC69: add esp, 0000001Ch
  0040DC6C: lea edx, var_48
  0040DC6F: push 00000000h
  0040DC71: push 00403A68h ; "Close"
  0040DC76: push edx
  0040DC77: call MSVBVM60.DLL.__vbaObjVar
  0040DC79: push eax
  0040DC7A: call edi
  0040DC7C: add esp, 0000000Ch
  0040DC7F: lea edx, var_6C
  0040DC82: lea ecx, var_28
  0040DC85: mov var_64, FFFFFFFFh
  0040DC8C: mov var_6C, 0000000Bh
  0040DC93: call MSVBVM60.DLL.__vbaVarMove
  0040DC99: push 0040DCD3h
  0040DC9E: jmp 40DCC2h
  0040DCA0: test byte ptr var_4, 04h
  0040DCA4: jz 40DCAFh
  0040DCA6: lea ecx, var_28
  0040DCA9: call MSVBVM60.DLL.__vbaFreeVar
  0040DCAF: lea ecx, var_4C
  0040DCB2: call MSVBVM60.DLL.__vbaFreeStr
  0040DCB8: lea ecx, var_5C
  0040DCBB: call MSVBVM60.DLL.__vbaFreeVar
  0040DCC1: ret 
End Sub

Private sub Proc_0_2_40DD10
  0040DD10: push ebp
  0040DD11: mov ebp, esp
  0040DD13: sub esp, 00000008h
  0040DD16: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0040DD1B: mov eax, fs:[00h]
  0040DD21: push eax
  0040DD22: mov fs:[00000000h], esp
  0040DD29: sub esp, 0000002Ch
  0040DD2C: push ebx
  0040DD2D: push esi
  0040DD2E: push edi
  0040DD2F: mov var_8, esp
  0040DD32: mov var_4, 00401398h
  0040DD39: xor eax, eax
  0040DD3B: mov var_18, eax
  0040DD3E: mov var_28, eax
  0040DD41: push eax
  0040DD42: lea eax, var_28
  0040DD45: push 004039FCh ; "Scripting.FileSystemObject"
  0040DD4A: push eax
  0040DD4B: call [00401210h] ; arg_1 = CreateObject(arg_2, arg_3)
  0040DD51: lea ecx, var_28
  0040DD54: push ecx
  0040DD55: call MSVBVM60.DLL.__vbaObjVar
  0040DD5B: lea edx, var_18
  0040DD5E: push eax
  0040DD5F: push edx
  0040DD60: call MSVBVM60.DLL.__vbaObjSetAddref
  0040DD66: mov esi, [00401024h]
  0040DD6C: lea ecx, var_28
  0040DD6F: call MSVBVM60.DLL.__vbaFreeVar
  0040DD71: sub esp, 00000010h
  0040DD74: mov eax, arg_8
  0040DD77: mov edx, esp
  0040DD79: mov ecx, 00004008h
  0040DD7E: push 00000001h
  0040DD80: push 00403A74h ; "FolderExists"
  0040DD85: mov [edx], ecx
  0040DD87: mov ecx, var_34
  0040DD8A: mov [edx+04h], ecx
  0040DD8D: mov ecx, var_18
  0040DD90: push ecx
  0040DD91: mov [edx+08h], eax
  0040DD94: mov eax, var_2C
  0040DD97: mov [edx+0Ch], eax
  0040DD9A: lea edx, var_28
  0040DD9D: push edx
  0040DD9E: call MSVBVM60.DLL.__vbaLateMemCallLd
  0040DDA4: add esp, 00000020h
  0040DDA7: push eax
  0040DDA8: call MSVBVM60.DLL.__vbaBoolVar
  0040DDAE: lea ecx, var_28
  0040DDB1: mov var_14, eax
  0040DDB4: call MSVBVM60.DLL.__vbaFreeVar
  0040DDB6: push 0040DDD1h ; "‹Mðf‹Eì_^d‰'#1"
  0040DDBB: jmp 40DDC7h
  0040DDBD: lea ecx, var_28
  0040DDC0: call MSVBVM60.DLL.__vbaFreeVar
  0040DDC6: ret 
End Sub

Private sub Proc_0_3_40DDF0
  0040DDF0: push ebp
  0040DDF1: mov ebp, esp
  0040DDF3: sub esp, 00000018h
  0040DDF6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0040DDFB: mov eax, fs:[00h]
  0040DE01: push eax
  0040DE02: mov fs:[00000000h], esp
  0040DE09: mov eax, 0000005Ch
  0040DE0E: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  0040DE13: push ebx
  0040DE14: push esi
  0040DE15: push edi
  0040DE16: mov var_18, esp
  0040DE19: mov var_14, 004013A8h
  0040DE20: mov var_10, 00000000h
  0040DE27: mov var_C, 00000000h
  0040DE2E: mov var_4, 00000001h
  0040DE35: mov var_4, 00000002h
  0040DE3C: push FFFFFFFFh
  0040DE3E: call On Error ...
  0040DE44: mov var_4, 00000003h
  0040DE4B: mov eax, arg_8
  0040DE4E: push eax
  0040DE4F: call 0040DD10h
  0040DE54: movsx ecx, ax
  0040DE57: test ecx, ecx
  0040DE59: jz 40DE60h
  0040DE5B: jmp 0040DF1Ah
  0040DE60: mov var_4, 00000006h
  0040DE67: mov edx, arg_8
  0040DE6A: mov var_4C, edx
  0040DE6D: mov var_54, 00004008h
  0040DE74: push 00000001h
  0040DE76: lea eax, var_54
  0040DE79: push eax
  0040DE7A: lea ecx, var_34
  0040DE7D: push ecx
  0040DE7E: call [0040132Ch] ; Right(arg_1, arg_2)
  0040DE84: mov var_5C, 0040393Ch
  0040DE8B: mov var_64, 00008008h
  0040DE92: lea edx, var_34
  0040DE95: push edx
  0040DE96: lea eax, var_64
  0040DE99: push eax
  0040DE9A: call MSVBVM60.DLL.__vbaVarTstEq
  0040DEA0: mov var_68, ax
  0040DEA4: lea ecx, var_34
  0040DEA7: call MSVBVM60.DLL.__vbaFreeVar
  0040DEAD: movsx ecx, word ptr var_68
  0040DEB1: test ecx, ecx
  0040DEB3: jz 40DF07h
  0040DEB5: mov var_4, 00000007h
  0040DEBC: mov edx, arg_8
  0040DEBF: mov var_4C, edx
  0040DEC2: mov var_54, 00004008h
  0040DEC9: mov eax, arg_8
  0040DECC: mov ecx, [eax]
  0040DECE: push ecx
  0040DECF: call [00401030h] ; Len(arg_1)
  0040DED5: sub eax, 00000001h
  0040DED8: jo 40DF4Dh
  0040DEDA: push eax
  0040DEDB: lea edx, var_54
  0040DEDE: push edx
  0040DEDF: lea eax, var_34
  0040DEE2: push eax
  0040DEE3: call [0040130Ch] ; Left(arg_1, arg_2)
  0040DEE9: lea ecx, var_34
  0040DEEC: push ecx
  0040DEED: call MSVBVM60.DLL.__vbaStrVarMove
  0040DEF3: mov edx, eax
  0040DEF5: mov ecx, arg_8
  0040DEF8: call MSVBVM60.DLL.__vbaStrMove
  0040DEFE: lea ecx, var_34
  0040DF01: call MSVBVM60.DLL.__vbaFreeVar
  0040DF07: mov var_4, 00000009h
  0040DF0E: mov edx, arg_8
  0040DF11: mov eax, [edx]
  0040DF13: push eax
  0040DF14: call [00401214h] ; MkDir(arg_1)
  0040DF1A: push 0040DF36h ; "f‹EÜ‹Màd‰'#1"
  0040DF1F: jmp 40DF35h
  0040DF21: lea ecx, var_44
  0040DF24: push ecx
  0040DF25: lea edx, var_34
  0040DF28: push edx
  0040DF29: push 00000002h
  0040DF2B: call MSVBVM60.DLL.__vbaFreeVarList
  0040DF31: add esp, 0000000Ch
  0040DF34: ret 
End Sub

Private sub Proc_0_4_40DF60
  0040DF60: push ebp
  0040DF61: mov ebp, esp
  0040DF63: sub esp, 0000000Ch
  0040DF66: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0040DF6B: mov eax, fs:[00h]
  0040DF71: push eax
  0040DF72: mov fs:[00000000h], esp
  0040DF79: sub esp, 0000022Ch
  0040DF7F: push ebx
  0040DF80: push esi
  0040DF81: push edi
  0040DF82: mov var_C, esp
  0040DF85: mov var_8, 004013F8h
  0040DF8C: mov edi, MSVBVM60.DLL.__vbaVarCopy
  0040DF92: xor esi, esi
  0040DF94: mov ebx, 00000008h
  0040DF99: lea edx, var_148
  0040DF9F: lea ecx, var_74
  0040DFA2: mov var_24, esi
  0040DFA5: mov var_34, esi
  0040DFA8: mov var_44, esi
  0040DFAB: mov var_54, esi
  0040DFAE: mov var_64, esi
  0040DFB1: mov var_74, esi
  0040DFB4: mov var_78, esi
  0040DFB7: mov var_88, esi
  0040DFBD: mov var_98, esi
  0040DFC3: mov var_A8, esi
  0040DFC9: mov var_AC, esi
  0040DFCF: mov var_BC, esi
  0040DFD5: mov var_CC, esi
  0040DFDB: mov var_D0, esi
  0040DFE1: mov var_D4, esi
  0040DFE7: mov var_D8, esi
  0040DFED: mov var_E8, esi
  0040DFF3: mov var_F8, esi
  0040DFF9: mov var_108, esi
  0040DFFF: mov var_118, esi
  0040E005: mov var_128, esi
  0040E00B: mov var_138, esi
  0040E011: mov var_158, esi
  0040E017: mov var_168, esi
  0040E01D: mov var_178, esi
  0040E023: mov var_188, esi
  0040E029: mov var_198, esi
  0040E02F: mov var_1BC, esi
  0040E035: mov var_1DC, esi
  0040E03B: mov var_1EC, esi
  0040E041: mov var_1FC, esi
  0040E047: mov var_20C, esi
  0040E04D: mov var_21C, esi
  0040E053: mov var_22C, esi
  0040E059: mov var_140, 00403558h
  0040E063: mov var_148, ebx
  0040E069: call edi
  0040E06B: lea edx, var_148
  0040E071: lea ecx, var_98
  0040E077: mov var_140, 004040FCh ; "10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,,10,10,10"
  0040E081: mov var_148, ebx
  0040E087: call edi
  0040E089: lea edx, var_148
  0040E08F: lea ecx, var_74
  0040E092: mov var_140, 00404188h ; "NO,NO KK,NIK,NAMA,JK,HUBKEL,STATNIKAH,SUKU,AGAMA,PENDIDIKAN,PENCAHARIAN,IBU,TEMPAT,TGL LAHIR,KEC,DESA,RT,RW,ALAMAT"
  0040E09C: mov var_148, ebx
  0040E0A2: call edi
  0040E0A4: lea edx, var_148
  0040E0AA: lea ecx, var_E8
  0040E0B0: mov var_140, 00403E1Ch
  0040E0BA: mov var_148, ebx
  0040E0C0: call MSVBVM60.DLL.__vbaVarDup
  0040E0C6: push esi
  0040E0C7: lea eax, var_E8
  0040E0CD: push FFFFFFFFh
  0040E0CF: lea ecx, var_74
  0040E0D2: push eax
  0040E0D3: lea edx, var_D4
  0040E0D9: push ecx
  0040E0DA: push edx
  0040E0DB: call MSVBVM60.DLL.__vbaStrVarVal
  0040E0E1: push eax
  0040E0E2: lea eax, var_F8
  0040E0E8: push eax
  0040E0E9: call [004011E8h] ; arg_1 = Split(arg_2, arg_3, arg_4, arg_5)
  0040E0EF: mov edi, MSVBVM60.DLL.__vbaAryVar
  0040E0F5: lea ecx, var_F8
  0040E0FB: push ecx
  0040E0FC: push 00002008h
  0040E101: call edi
  0040E103: mov var_1BC, eax
  0040E109: mov ebx, MSVBVM60.DLL.__vbaAryCopy
  0040E10F: lea edx, var_1BC
  0040E115: lea eax, var_78
  0040E118: push edx
  0040E119: push eax
  0040E11A: call ebx
  0040E11C: lea ecx, var_D4
  0040E122: call MSVBVM60.DLL.__vbaFreeStr
  0040E128: lea ecx, var_F8
  0040E12E: lea edx, var_E8
  0040E134: push ecx
  0040E135: push edx
  0040E136: push 00000002h
  0040E138: call MSVBVM60.DLL.__vbaFreeVarList
  0040E13E: add esp, 0000000Ch
  0040E141: lea edx, var_148
  0040E147: lea ecx, var_E8
  0040E14D: mov var_140, 00403E1Ch
  0040E157: mov var_148, 00000008h
  0040E161: call MSVBVM60.DLL.__vbaVarDup
  0040E167: push esi
  0040E168: lea eax, var_E8
  0040E16E: push FFFFFFFFh
  0040E170: lea ecx, var_98
  0040E176: push eax
  0040E177: lea edx, var_D4
  0040E17D: push ecx
  0040E17E: push edx
  0040E17F: call MSVBVM60.DLL.__vbaStrVarVal
  0040E185: push eax
  0040E186: lea eax, var_F8
  0040E18C: push eax
  0040E18D: call [004011E8h] ; arg_1 = Split(arg_2, arg_3, arg_4, arg_5)
  0040E193: lea ecx, var_F8
  0040E199: push ecx
  0040E19A: push 00002008h
  0040E19F: call edi
  0040E1A1: mov var_1BC, eax
  0040E1A7: lea edx, var_1BC
  0040E1AD: lea eax, var_AC
  0040E1B3: push edx
  0040E1B4: push eax
  0040E1B5: call ebx
  0040E1B7: lea ecx, var_D4
  0040E1BD: call MSVBVM60.DLL.__vbaFreeStr
  0040E1C3: lea ecx, var_F8
  0040E1C9: lea edx, var_E8
  0040E1CF: push ecx
  0040E1D0: push edx
  0040E1D1: push 00000002h
  0040E1D3: call MSVBVM60.DLL.__vbaFreeVarList
  0040E1D9: mov eax, var_78
  0040E1DC: add esp, 0000000Ch
  0040E1DF: push eax
  0040E1E0: push 00000001h
  0040E1E2: call [00401228h] ; UBound
  0040E1E8: lea edx, var_148
  0040E1EE: lea ecx, var_54
  0040E1F1: mov var_140, eax
  0040E1F7: mov var_148, 00000003h
  0040E201: call MSVBVM60.DLL.__vbaVarMove
  0040E207: mov edi, arg_C
  0040E20A: push esi
  0040E20B: push 00000044h
  0040E20D: mov ecx, [edi]
  0040E20F: push ecx
  0040E210: call MSVBVM60.DLL.__vbaLateIdCall
  0040E216: mov ecx, 00000003h
  0040E21B: mov eax, 00000002h
  0040E220: push ecx
  0040E221: mov var_148, ecx
  0040E227: mov edx, esp
  0040E229: mov var_140, eax
  0040E22F: push 00000004h
  0040E231: mov [edx], ecx
  0040E233: mov ecx, var_144
  0040E239: mov [edx+04h], ecx
  0040E23C: mov ecx, [edi]
  0040E23E: mov [edx+08h], eax
  0040E241: mov eax, var_13C
  0040E247: mov [edx+0Ch], eax
  0040E24A: push ecx
  0040E24B: call MSVBVM60.DLL.__vbaLateIdSt
  0040E251: lea edx, var_54
  0040E254: lea eax, var_148
  0040E25A: push edx
  0040E25B: lea ecx, var_E8
  0040E261: push eax
  0040E262: push ecx
  0040E263: mov var_140, 00000001h
  0040E26D: mov var_148, 00000002h
  0040E277: call MSVBVM60.DLL.__vbaVarAdd
  0040E27D: mov ebx, MSVBVM60.DLL.__vbaI4Var
  0040E283: push eax
  0040E284: call ebx
  0040E286: sub esp, 00000010h
  0040E289: mov ecx, 00000003h
  0040E28E: mov edx, esp
  0040E290: mov var_158, ecx
  0040E296: mov var_150, eax
  0040E29C: push 00000005h
  0040E29E: mov [edx], ecx
  0040E2A0: mov ecx, var_154
  0040E2A6: mov [edx+04h], ecx
  0040E2A9: mov ecx, [edi]
  0040E2AB: push ecx
  0040E2AC: mov [edx+08h], eax
  0040E2AF: mov eax, var_14C
  0040E2B5: mov [edx+0Ch], eax
  0040E2B8: call MSVBVM60.DLL.__vbaLateIdSt
  0040E2BE: lea ecx, var_E8
  0040E2C4: call MSVBVM60.DLL.__vbaFreeVar
  0040E2CA: sub esp, 00000010h
  0040E2CD: mov ecx, 00000003h
  0040E2D2: mov edx, esp
  0040E2D4: mov var_148, ecx
  0040E2DA: mov eax, 00000001h
  0040E2DF: push 00000007h
  0040E2E1: mov [edx], ecx
  0040E2E3: mov ecx, var_144
  0040E2E9: mov var_140, eax
  0040E2EF: mov [edx+04h], ecx
  0040E2F2: mov ecx, [edi]
  0040E2F4: push ecx
  0040E2F5: mov [edx+08h], eax
  0040E2F8: mov eax, var_13C
  0040E2FE: mov [edx+0Ch], eax
  0040E301: call MSVBVM60.DLL.__vbaLateIdSt
  0040E307: sub esp, 00000010h
  0040E30A: mov ecx, 00000003h
  0040E30F: mov edx, esp
  0040E311: mov var_148, ecx
  0040E317: mov eax, 00000001h
  0040E31C: push 00000006h
  0040E31E: mov [edx], ecx
  0040E320: mov ecx, var_144
  0040E326: mov var_140, eax
  0040E32C: mov [edx+04h], ecx
  0040E32F: mov ecx, [edi]
  0040E331: push ecx
  0040E332: mov [edx+08h], eax
  0040E335: mov eax, var_13C
  0040E33B: mov [edx+0Ch], eax
  0040E33E: call MSVBVM60.DLL.__vbaLateIdSt
  0040E344: mov eax, 00000002h
  0040E349: lea edx, var_148
  0040E34F: mov var_148, eax
  0040E355: mov var_158, eax
  0040E35B: lea eax, var_54
  0040E35E: push edx
  0040E35F: lea ecx, var_158
  0040E365: push eax
  0040E366: lea edx, var_1EC
  0040E36C: push ecx
  0040E36D: lea eax, var_1DC
  0040E373: push edx
  0040E374: lea ecx, var_A8
  0040E37A: push eax
  0040E37B: push ecx
  0040E37C: mov var_140, 00000001h
  0040E386: mov var_150, esi
  0040E38C: call [004010B4h] ; For
  0040E392: cmp eax, esi
  0040E394: jz 0040E5C9h
  0040E39A: lea edx, var_A8
  0040E3A0: push edx
  0040E3A1: call ebx
  0040E3A3: mov var_140, eax
  0040E3A9: mov eax, var_AC
  0040E3AF: cmp eax, esi
  0040E3B1: mov var_148, 00000003h
  0040E3BB: jz 40E3F5h
  0040E3BD: cmp word ptr [eax], 0001h
  0040E3C1: jnz 40E3F5h
  0040E3C3: lea eax, var_A8
  0040E3C9: push eax
  0040E3CA: call ebx
  0040E3CC: mov ecx, var_AC
  0040E3D2: mov edx, [ecx+14h]
  0040E3D5: sub eax, edx
  0040E3D7: mov edx, [ecx+10h]
  0040E3DA: cmp eax, edx
  0040E3DC: mov var_1C0, eax
  0040E3E2: jb 40E3F0h
  0040E3E4: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0040E3EA: mov eax, var_1C0
  0040E3F0: shl eax, 02h
  0040E3F3: jmp 40E3FBh
  0040E3F5: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0040E3FB: mov ecx, var_AC
  0040E401: mov edx, [ecx+0Ch]
  0040E404: mov eax, [edx+eax]
  0040E407: push eax
  0040E408: call [00401370h] ; Val(arg_1)
  0040E40E: fmul real8 ptr [004013F0h] ; 
  0040E414: fstsw ax
  0040E416: test al, 0Dh
  0040E418: jnz 0040EDE3h
  0040E41E: call MSVBVM60.DLL.__vbaFpI4
  0040E424: mov edx, var_148
  0040E42A: sub esp, 00000010h
  0040E42D: mov ecx, esp
  0040E42F: sub esp, 00000010h
  0040E432: mov var_168, 00000003h
  0040E43C: mov var_160, eax
  0040E442: mov [ecx], edx
  0040E444: mov edx, var_144
  0040E44A: mov [ecx+04h], edx
  0040E44D: mov edx, var_140
  0040E453: mov [ecx+08h], edx
  0040E456: mov edx, var_13C
  0040E45C: mov [ecx+0Ch], edx
  0040E45F: mov edx, var_168
  0040E465: mov ecx, esp
  0040E467: push 00000001h
  0040E469: push 00000039h
  0040E46B: mov [ecx], edx
  0040E46D: mov edx, var_164
  0040E473: mov [ecx+04h], edx
  0040E476: mov [ecx+08h], eax
  0040E479: mov eax, var_15C
  0040E47F: mov [ecx+0Ch], eax
  0040E482: mov ecx, [edi]
  0040E484: push ecx
  0040E485: call MSVBVM60.DLL.__vbaLateIdCallSt
  0040E48B: add esp, 0000002Ch
  0040E48E: lea edx, var_A8
  0040E494: mov var_140, esi
  0040E49A: mov var_148, 00000003h
  0040E4A4: push edx
  0040E4A5: call ebx
  0040E4A7: mov var_160, eax
  0040E4AD: mov eax, var_78
  0040E4B0: lea ecx, var_D0
  0040E4B6: push eax
  0040E4B7: push ecx
  0040E4B8: mov var_168, 00000003h
  0040E4C2: call MSVBVM60.DLL.__vbaAryLock
  0040E4C8: mov eax, var_D0
  0040E4CE: cmp eax, esi
  0040E4D0: jz 40E50Ah
  0040E4D2: cmp word ptr [eax], 0001h
  0040E4D6: jnz 40E50Ah
  0040E4D8: lea edx, var_A8
  0040E4DE: push edx
  0040E4DF: call ebx
  0040E4E1: mov ecx, var_D0
  0040E4E7: mov edx, [ecx+14h]
  0040E4EA: sub eax, edx
  0040E4EC: mov edx, [ecx+10h]
  0040E4EF: cmp eax, edx
  0040E4F1: mov var_1C0, eax
  0040E4F7: jb 40E505h
  0040E4F9: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0040E4FF: mov eax, var_1C0
  0040E505: shl eax, 02h
  0040E508: jmp 40E510h
  0040E50A: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0040E510: mov ecx, var_D0
  0040E516: sub esp, 00000010h
  0040E519: mov edx, esp
  0040E51B: sub esp, 00000010h
  0040E51E: mov ecx, [ecx+0Ch]
  0040E521: add ecx, eax
  0040E523: mov eax, var_148
  0040E529: mov [edx], eax
  0040E52B: mov eax, var_144
  0040E531: mov [edx+04h], eax
  0040E534: mov eax, var_140
  0040E53A: mov [edx+08h], eax
  0040E53D: mov eax, var_13C
  0040E543: mov [edx+0Ch], eax
  0040E546: mov eax, var_168
  0040E54C: mov edx, esp
  0040E54E: sub esp, 00000010h
  0040E551: mov [edx], eax
  0040E553: mov eax, var_164
  0040E559: mov [edx+04h], eax
  0040E55C: mov eax, var_160
  0040E562: mov [edx+08h], eax
  0040E565: mov eax, var_15C
  0040E56B: mov [edx+0Ch], eax
  0040E56E: mov edx, esp
  0040E570: mov eax, 00004008h
  0040E575: push 00000002h
  0040E577: mov [edx], eax
  0040E579: mov eax, var_184
  0040E57F: push 00000041h
  0040E581: mov [edx+04h], eax
  0040E584: mov [edx+08h], ecx
  0040E587: mov ecx, var_17C
  0040E58D: mov [edx+0Ch], ecx
  0040E590: mov edx, [edi]
  0040E592: push edx
  0040E593: call MSVBVM60.DLL.__vbaLateIdCallSt
  0040E599: add esp, 0000003Ch
  0040E59C: lea eax, var_D0
  0040E5A2: push eax
  0040E5A3: call MSVBVM60.DLL.__vbaAryUnlock
  0040E5A9: lea ecx, var_1EC
  0040E5AF: lea edx, var_1DC
  0040E5B5: push ecx
  0040E5B6: lea eax, var_A8
  0040E5BC: push edx
  0040E5BD: push eax
  0040E5BE: call [00401360h] ; Next
  0040E5C4: jmp 0040E392h
  0040E5C9: lea ecx, var_98
  0040E5CF: push ecx
  0040E5D0: push edi
  0040E5D1: call 00416BF0h
  0040E5D6: cmp [00419640h], esi
  0040E5DC: jnz 40E5EEh
  0040E5DE: push 00419640h
  0040E5E3: push 004039B0h
  0040E5E8: call MSVBVM60.DLL.__vbaNew2
  0040E5EE: mov eax, [419640h]
  0040E5F3: lea ecx, var_D8
  0040E5F9: push ecx
  0040E5FA: push eax
  0040E5FB: mov edx, [eax]
  0040E5FD: mov var_1C0, eax
  0040E603: call [edx+1Ch]
  0040E606: cmp eax, esi
  0040E608: fclex 
  0040E60A: jnl 40E621h
  0040E60C: mov edx, var_1C0
  0040E612: push 0000001Ch
  0040E614: push 004039A0h
  0040E619: push edx
  0040E61A: push eax
  0040E61B: call MSVBVM60.DLL.__vbaHresultCheckObj
  0040E621: lea edx, var_D4
  0040E627: mov eax, 0000000Ah
  0040E62C: push edx
  0040E62D: mov ecx, var_D8
  0040E633: sub esp, 00000010h
  0040E636: mov var_148, eax
  0040E63C: mov edx, esp
  0040E63E: mov var_140, 80020004h
  0040E648: mov var_1C8, ecx
  0040E64E: mov ecx, [ecx]
  0040E650: mov [edx], eax
  0040E652: mov eax, var_144
  0040E658: mov [edx+04h], eax
  0040E65B: mov eax, var_140
  0040E661: mov [edx+08h], eax
  0040E664: mov eax, var_13C
  0040E66A: mov [edx+0Ch], eax
  0040E66D: mov edx, var_D8
  0040E673: push edx
  0040E674: call [ecx+5Ch]
  0040E677: cmp eax, esi
  0040E679: fclex 
  0040E67B: jnl 40E692h
  0040E67D: mov ecx, var_1C8
  0040E683: push 0000005Ch
  0040E685: push 00404594h
  0040E68A: push ecx
  0040E68B: push eax
  0040E68C: call MSVBVM60.DLL.__vbaHresultCheckObj
  0040E692: lea edx, var_158
  0040E698: lea ecx, var_E8
  0040E69E: mov var_150, 004045A8h
  0040E6A8: mov var_158, 00000008h
  0040E6B2: call MSVBVM60.DLL.__vbaVarDup
  0040E6B8: mov eax, var_D4
  0040E6BE: push esi
  0040E6BF: lea edx, var_E8
  0040E6C5: push FFFFFFFFh
  0040E6C7: push edx
  0040E6C8: lea ecx, var_F8
  0040E6CE: push eax
  0040E6CF: push ecx
  0040E6D0: call [004011E8h] ; arg_1 = Split(arg_2, arg_3, arg_4, arg_5)
  0040E6D6: lea edx, var_F8
  0040E6DC: lea ecx, var_44
  0040E6DF: call MSVBVM60.DLL.__vbaVarMove
  0040E6E5: lea ecx, var_D4
  0040E6EB: call MSVBVM60.DLL.__vbaFreeStr
  0040E6F1: lea ecx, var_D8
  0040E6F7: call MSVBVM60.DLL.__vbaFreeObj
  0040E6FD: lea ecx, var_E8
  0040E703: call MSVBVM60.DLL.__vbaFreeVar
  0040E709: lea edx, var_44
  0040E70C: mov var_140, 00000001h
  0040E716: push edx
  0040E717: mov var_148, 00000002h
  0040E721: call MSVBVM60.DLL.__vbaRefVarAry
  0040E727: mov eax, [eax]
  0040E729: push eax
  0040E72A: push 00000001h
  0040E72C: call [00401228h] ; UBound
  0040E732: sub eax, 00000001h
  0040E735: lea ecx, var_148
  0040E73B: jo 0040EDE8h
  0040E741: mov var_150, eax
  0040E747: lea edx, var_158
  0040E74D: push ecx
  0040E74E: lea eax, var_168
  0040E754: push edx
  0040E755: lea ecx, var_20C
  0040E75B: push eax
  0040E75C: lea edx, var_1FC
  0040E762: push ecx
  0040E763: lea eax, var_24
  0040E766: push edx
  0040E767: push eax
  0040E768: mov var_158, 00000003h
  0040E772: mov var_160, esi
  0040E778: mov var_168, 00000002h
  0040E782: call [004010B4h] ; For
  0040E788: cmp eax, esi
  0040E78A: jz 0040EC87h
  0040E790: mov ecx, [edi]
  0040E792: push esi
  0040E793: push 00000004h
  0040E795: lea edx, var_E8
  0040E79B: push ecx
  0040E79C: push edx
  0040E79D: call MSVBVM60.DLL.__vbaLateIdCallLd
  0040E7A3: add esp, 00000010h
  0040E7A6: push eax
  0040E7A7: call ebx
  0040E7A9: add eax, 00000001h
  0040E7AC: mov ecx, 00000003h
  0040E7B1: jo 0040EDE8h
  0040E7B7: sub esp, 00000010h
  0040E7BA: mov var_148, ecx
  0040E7C0: mov edx, esp
  0040E7C2: mov var_140, eax
  0040E7C8: push 00000004h
  0040E7CA: mov [edx], ecx
  0040E7CC: mov ecx, var_144
  0040E7D2: mov [edx+04h], ecx
  0040E7D5: mov ecx, [edi]
  0040E7D7: push ecx
  0040E7D8: mov [edx+08h], eax
  0040E7DB: mov eax, var_13C
  0040E7E1: mov [edx+0Ch], eax
  0040E7E4: call MSVBVM60.DLL.__vbaLateIdSt
  0040E7EA: lea ecx, var_E8
  0040E7F0: call MSVBVM60.DLL.__vbaFreeVar
  0040E7F6: sub esp, 00000010h
  0040E7F9: mov ecx, 0000400Ch
  0040E7FE: mov edx, esp
  0040E800: mov var_148, ecx
  0040E806: lea eax, var_24
  0040E809: push 00000001h
  0040E80B: mov [edx], ecx
  0040E80D: mov ecx, var_144
  0040E813: mov var_140, eax
  0040E819: mov [edx+04h], ecx
  0040E81C: lea ecx, var_44
  0040E81F: push ecx
  0040E820: mov [edx+08h], eax
  0040E823: mov eax, var_13C
  0040E829: mov [edx+0Ch], eax
  0040E82C: lea edx, var_E8
  0040E832: push edx
  0040E833: call MSVBVM60.DLL.__vbaVarIndexLoad
  0040E839: add esp, 0000001Ch
  0040E83C: lea edx, var_158
  0040E842: lea ecx, var_F8
  0040E848: mov var_150, 004045B0h
  0040E852: mov var_158, 00000008h
  0040E85C: call MSVBVM60.DLL.__vbaVarDup
  0040E862: push esi
  0040E863: lea eax, var_F8
  0040E869: push FFFFFFFFh
  0040E86B: lea ecx, var_E8
  0040E871: push eax
  0040E872: lea edx, var_D4
  0040E878: push ecx
  0040E879: push edx
  0040E87A: call MSVBVM60.DLL.__vbaStrVarVal
  0040E880: push eax
  0040E881: lea eax, var_108
  0040E887: push eax
  0040E888: call [004011E8h] ; arg_1 = Split(arg_2, arg_3, arg_4, arg_5)
  0040E88E: lea edx, var_108
  0040E894: lea ecx, var_BC
  0040E89A: call MSVBVM60.DLL.__vbaVarMove
  0040E8A0: lea ecx, var_D4
  0040E8A6: call MSVBVM60.DLL.__vbaFreeStr
  0040E8AC: lea ecx, var_F8
  0040E8B2: lea edx, var_E8
  0040E8B8: push ecx
  0040E8B9: push edx
  0040E8BA: push 00000002h
  0040E8BC: call MSVBVM60.DLL.__vbaFreeVarList
  0040E8C2: add esp, 0000000Ch
  0040E8C5: mov var_150, 00000001h
  0040E8CF: mov var_158, 00000002h
  0040E8D9: lea eax, var_24
  0040E8DC: lea ecx, var_158
  0040E8E2: push eax
  0040E8E3: lea edx, var_F8
  0040E8E9: push ecx
  0040E8EA: push edx
  0040E8EB: call MSVBVM60.DLL.__vbaVarAdd
  0040E8F1: push eax
  0040E8F2: call ebx
  0040E8F4: mov var_160, eax
  0040E8FA: lea eax, var_24
  0040E8FD: lea ecx, var_148
  0040E903: push eax
  0040E904: lea edx, var_E8
  0040E90A: push ecx
  0040E90B: push edx
  0040E90C: mov var_168, 00000003h
  0040E916: mov var_140, 00000001h
  0040E920: mov var_148, 00000002h
  0040E92A: call MSVBVM60.DLL.__vbaVarAdd
  0040E930: push eax
  0040E931: call MSVBVM60.DLL.__vbaStrVarMove
  0040E937: mov edx, var_168
  0040E93D: sub esp, 00000010h
  0040E940: mov ecx, esp
  0040E942: sub esp, 00000010h
  0040E945: mov var_108, 00000008h
  0040E94F: mov var_100, eax
  0040E955: mov [ecx], edx
  0040E957: mov edx, var_164
  0040E95D: mov [ecx+04h], edx
  0040E960: mov edx, var_160
  0040E966: mov [ecx+08h], edx
  0040E969: mov edx, var_15C
  0040E96F: mov [ecx+0Ch], edx
  0040E972: mov edx, esp
  0040E974: mov ecx, 00000003h
  0040E979: sub esp, 00000010h
  0040E97C: mov [edx], ecx
  0040E97E: mov ecx, var_184
  0040E984: mov [edx+04h], ecx
  0040E987: xor ecx, ecx
  0040E989: mov [edx+08h], ecx
  0040E98C: mov ecx, var_17C
  0040E992: mov [edx+0Ch], ecx
  0040E995: mov ecx, var_108
  0040E99B: mov edx, esp
  0040E99D: push 00000002h
  0040E99F: push 00000041h
  0040E9A1: mov [edx], ecx
  0040E9A3: mov ecx, var_104
  0040E9A9: mov [edx+04h], ecx
  0040E9AC: mov ecx, [edi]
  0040E9AE: push ecx
  0040E9AF: mov [edx+08h], eax
  0040E9B2: mov eax, var_FC
  0040E9B8: mov [edx+0Ch], eax
  0040E9BB: call MSVBVM60.DLL.__vbaLateIdCallSt
  0040E9C1: lea edx, var_108
  0040E9C7: lea eax, var_E8
  0040E9CD: push edx
  0040E9CE: lea ecx, var_F8
  0040E9D4: push eax
  0040E9D5: push ecx
  0040E9D6: push 00000003h
  0040E9D8: call MSVBVM60.DLL.__vbaFreeVarList
  0040E9DE: add esp, 0000004Ch
  0040E9E1: lea edx, var_BC
  0040E9E7: push edx
  0040E9E8: call MSVBVM60.DLL.__vbaRefVarAry
  0040E9EE: mov eax, [eax]
  0040E9F0: push eax
  0040E9F1: push 00000001h
  0040E9F3: call [00401228h] ; UBound
  0040E9F9: lea ecx, var_E8
  0040E9FF: mov var_E0, eax
  0040EA05: lea edx, var_54
  0040EA08: push ecx
  0040EA09: lea eax, var_F8
  0040EA0F: push edx
  0040EA10: mov var_E8, 00000003h
  0040EA1A: push eax
  0040EA1B: call 00416A10h
  0040EA20: lea edx, var_F8
  0040EA26: lea ecx, var_CC
  0040EA2C: call MSVBVM60.DLL.__vbaVarMove
  0040EA32: lea ecx, var_E8
  0040EA38: call MSVBVM60.DLL.__vbaFreeVar
  0040EA3E: mov eax, 00000002h
  0040EA43: lea ecx, var_148
  0040EA49: mov var_148, eax
  0040EA4F: mov var_158, eax
  0040EA55: lea edx, var_CC
  0040EA5B: push ecx
  0040EA5C: lea eax, var_158
  0040EA62: push edx
  0040EA63: lea ecx, var_22C
  0040EA69: push eax
  0040EA6A: lea edx, var_21C
  0040EA70: push ecx
  0040EA71: lea eax, var_88
  0040EA77: push edx
  0040EA78: push eax
  0040EA79: mov var_140, 00000001h
  0040EA83: mov var_150, esi
  0040EA89: call [004010B4h] ; For
  0040EA8F: cmp eax, esi
  0040EA91: jz 0040EC6Ah
  0040EA97: lea ecx, var_24
  0040EA9A: lea edx, var_158
  0040EAA0: push ecx
  0040EAA1: lea eax, var_118
  0040EAA7: push edx
  0040EAA8: push eax
  0040EAA9: mov var_150, 00000001h
  0040EAB3: mov var_158, 00000002h
  0040EABD: call MSVBVM60.DLL.__vbaVarAdd
  0040EAC3: push eax
  0040EAC4: call ebx
  0040EAC6: lea ecx, var_88
  0040EACC: mov var_170, eax
  0040EAD2: lea edx, var_168
  0040EAD8: push ecx
  0040EAD9: lea eax, var_128
  0040EADF: push edx
  0040EAE0: push eax
  0040EAE1: mov var_160, 00000001h
  0040EAEB: mov var_168, 00000002h
  0040EAF5: call MSVBVM60.DLL.__vbaVarAdd
  0040EAFB: push eax
  0040EAFC: call ebx
  0040EAFE: sub esp, 00000010h
  0040EB01: mov ecx, 0000400Ch
  0040EB06: mov edx, esp
  0040EB08: mov var_148, ecx
  0040EB0E: mov var_190, eax
  0040EB14: lea eax, var_88
  0040EB1A: mov [edx], ecx
  0040EB1C: mov ecx, var_144
  0040EB22: mov var_140, eax
  0040EB28: push 00000001h
  0040EB2A: mov [edx+04h], ecx
  0040EB2D: lea ecx, var_BC
  0040EB33: push ecx
  0040EB34: mov [edx+08h], eax
  0040EB37: mov eax, var_13C
  0040EB3D: mov [edx+0Ch], eax
  0040EB40: lea edx, var_E8
  0040EB46: push edx
  0040EB47: call MSVBVM60.DLL.__vbaVarIndexLoad
  0040EB4D: add esp, 0000001Ch
  0040EB50: push eax
  0040EB51: call MSVBVM60.DLL.__vbaStrErrVarCopy
  0040EB57: mov var_F0, eax
  0040EB5D: lea eax, var_F8
  0040EB63: lea ecx, var_108
  0040EB69: push eax
  0040EB6A: push ecx
  0040EB6B: mov var_F8, 00000008h
  0040EB75: call [004010E8h] ; arg_1 = Trim(arg_2)
  0040EB7B: lea edx, var_108
  0040EB81: push edx
  0040EB82: call MSVBVM60.DLL.__vbaStrVarMove
  0040EB88: sub esp, 00000010h
  0040EB8B: mov ecx, 00000003h
  0040EB90: mov edx, esp
  0040EB92: sub esp, 00000010h
  0040EB95: mov var_138, 00000008h
  0040EB9F: mov var_130, eax
  0040EBA5: mov [edx], ecx
  0040EBA7: mov ecx, var_174
  0040EBAD: mov [edx+04h], ecx
  0040EBB0: mov ecx, var_170
  0040EBB6: mov [edx+08h], ecx
  0040EBB9: mov ecx, var_16C
  0040EBBF: mov [edx+0Ch], ecx
  0040EBC2: mov edx, esp
  0040EBC4: mov ecx, 00000003h
  0040EBC9: sub esp, 00000010h
  0040EBCC: mov [edx], ecx
  0040EBCE: mov ecx, var_194
  0040EBD4: mov [edx+04h], ecx
  0040EBD7: mov ecx, var_190
  0040EBDD: mov [edx+08h], ecx
  0040EBE0: mov ecx, var_18C
  0040EBE6: mov [edx+0Ch], ecx
  0040EBE9: mov ecx, var_138
  0040EBEF: mov edx, esp
  0040EBF1: mov [edx], ecx
  0040EBF3: mov ecx, var_134
  0040EBF9: push 00000002h
  0040EBFB: mov [edx+04h], ecx
  0040EBFE: mov ecx, [edi]
  0040EC00: push 00000041h
  0040EC02: push ecx
  0040EC03: mov [edx+08h], eax
  0040EC06: mov eax, var_12C
  0040EC0C: mov [edx+0Ch], eax
  0040EC0F: call MSVBVM60.DLL.__vbaLateIdCallSt
  0040EC15: lea edx, var_138
  0040EC1B: lea eax, var_108
  0040EC21: push edx
  0040EC22: lea ecx, var_128
  0040EC28: push eax
  0040EC29: lea edx, var_118
  0040EC2F: push ecx
  0040EC30: lea eax, var_F8
  0040EC36: push edx
  0040EC37: lea ecx, var_E8
  0040EC3D: push eax
  0040EC3E: push ecx
  0040EC3F: push 00000006h
  0040EC41: call MSVBVM60.DLL.__vbaFreeVarList
  0040EC47: add esp, 00000058h
  0040EC4A: lea edx, var_22C
  0040EC50: lea eax, var_21C
  0040EC56: lea ecx, var_88
  0040EC5C: push edx
  0040EC5D: push eax
  0040EC5E: push ecx
  0040EC5F: call [00401360h] ; Next
  0040EC65: jmp 0040EA8Fh
  0040EC6A: lea edx, var_20C
  0040EC70: lea eax, var_1FC
  0040EC76: push edx
  0040EC77: lea ecx, var_24
  0040EC7A: push eax
  0040EC7B: push ecx
  0040EC7C: call [00401360h] ; Next
  0040EC82: jmp 0040E788h
  0040EC87: lea edx, var_34
  0040EC8A: lea eax, var_148
  0040EC90: push edx
  0040EC91: push eax
  0040EC92: mov var_140, 00403558h
  0040EC9C: mov var_148, 00008008h
  0040ECA6: call MSVBVM60.DLL.__vbaVarTstEq
  0040ECAC: wait 
  0040ECAD: push 0040EDB4h
  0040ECB2: jmp 40ED1Eh
  0040ECB4: test byte ptr var_4, 04h
  0040ECB8: jz 40ECC3h
  0040ECBA: lea ecx, var_64
  0040ECBD: call MSVBVM60.DLL.__vbaFreeVar
  0040ECC3: lea ecx, var_D0
  0040ECC9: push ecx
  0040ECCA: call MSVBVM60.DLL.__vbaAryUnlock
  0040ECD0: lea ecx, var_D4
  0040ECD6: call MSVBVM60.DLL.__vbaFreeStr
  0040ECDC: lea ecx, var_D8
  0040ECE2: call MSVBVM60.DLL.__vbaFreeObj
  0040ECE8: lea edx, var_138
  0040ECEE: lea eax, var_128
  0040ECF4: push edx
  0040ECF5: lea ecx, var_118
  0040ECFB: push eax
  0040ECFC: lea edx, var_108
  0040ED02: push ecx
  0040ED03: lea eax, var_F8
  0040ED09: push edx
  0040ED0A: lea ecx, var_E8
  0040ED10: push eax
  0040ED11: push ecx
  0040ED12: push 00000006h
  0040ED14: call MSVBVM60.DLL.__vbaFreeVarList
  0040ED1A: add esp, 0000001Ch
  0040ED1D: ret 
End Sub

Private sub Proc_0_5_40EDF0
  0040EDF0: push ebp
  0040EDF1: mov ebp, esp
  0040EDF3: sub esp, 00000008h
  0040EDF6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0040EDFB: mov eax, fs:[00h]
  0040EE01: push eax
  0040EE02: mov fs:[00000000h], esp
  0040EE09: sub esp, 0000016Ch
  0040EE0F: push ebx
  0040EE10: push esi
  0040EE11: push edi
  0040EE12: mov var_8, esp
  0040EE15: mov var_4, 00401408h
  0040EE1C: mov eax, [419640h]
  0040EE21: xor ebx, ebx
  0040EE23: cmp eax, ebx
  0040EE25: mov var_20, ebx
  0040EE28: mov var_30, ebx
  0040EE2B: mov var_40, ebx
  0040EE2E: mov var_44, ebx
  0040EE31: mov var_48, ebx
  0040EE34: mov var_4C, ebx
  0040EE37: mov var_50, ebx
  0040EE3A: mov var_54, ebx
  0040EE3D: mov var_58, ebx
  0040EE40: mov var_5C, ebx
  0040EE43: mov var_60, ebx
  0040EE46: mov var_64, ebx
  0040EE49: mov var_68, ebx
  0040EE4C: mov var_6C, ebx
  0040EE4F: mov var_7C, ebx
  0040EE52: mov var_8C, ebx
  0040EE58: mov var_9C, ebx
  0040EE5E: mov var_AC, ebx
  0040EE64: mov var_BC, ebx
  0040EE6A: mov var_CC, ebx
  0040EE70: mov var_DC, ebx
  0040EE76: mov var_EC, ebx
  0040EE7C: mov var_FC, ebx
  0040EE82: mov var_10C, ebx
  0040EE88: mov var_140, ebx
  0040EE8E: mov var_144, ebx
  0040EE94: mov var_148, ebx
  0040EE9A: jnz 40EEACh
  0040EE9C: push 00419640h
  0040EEA1: push 004039B0h
  0040EEA6: call MSVBVM60.DLL.__vbaNew2
  0040EEAC: mov esi, [00419640h] ; 
  0040EEB2: lea ecx, var_64
  0040EEB5: push ecx
  0040EEB6: push esi
  0040EEB7: mov eax, [esi]
  0040EEB9: call [eax+14h]
  0040EEBC: cmp eax, ebx
  0040EEBE: fclex 
  0040EEC0: jnl 40EED5h
  0040EEC2: mov edi, MSVBVM60.DLL.__vbaHresultCheckObj
  0040EEC8: push 00000014h
  0040EECA: push 004039A0h
  0040EECF: push esi
  0040EED0: push eax
  0040EED1: call edi
  0040EED3: jmp 40EEDBh
  0040EED5: mov edi, MSVBVM60.DLL.__vbaHresultCheckObj
  0040EEDB: mov eax, var_64
  0040EEDE: lea ecx, var_140
  0040EEE4: push ecx
  0040EEE5: push eax
  0040EEE6: mov edx, [eax]
  0040EEE8: mov esi, eax
  0040EEEA: call [edx+000000B8h]
  0040EEF0: cmp eax, ebx
  0040EEF2: fclex 
  0040EEF4: jnl 40EF04h
  0040EEF6: push 000000B8h
  0040EEFB: push 00404788h
  0040EF00: push esi
  0040EF01: push eax
  0040EF02: call edi
  0040EF04: cmp [00419640h], ebx
  0040EF0A: jnz 40EF1Ch
  0040EF0C: push 00419640h
  0040EF11: push 004039B0h
  0040EF16: call MSVBVM60.DLL.__vbaNew2
  0040EF1C: mov esi, [00419640h] ; 
  0040EF22: lea eax, var_68
  0040EF25: push eax
  0040EF26: push esi
  0040EF27: mov edx, [esi]
  0040EF29: call [edx+14h]
  0040EF2C: cmp eax, ebx
  0040EF2E: fclex 
  0040EF30: jnl 40EF3Dh
  0040EF32: push 00000014h
  0040EF34: push 004039A0h
  0040EF39: push esi
  0040EF3A: push eax
  0040EF3B: call edi
  0040EF3D: mov eax, var_68
  0040EF40: lea edx, var_144
  0040EF46: push edx
  0040EF47: push eax
  0040EF48: mov ecx, [eax]
  0040EF4A: mov esi, eax
  0040EF4C: call [ecx+000000C0h]
  0040EF52: cmp eax, ebx
  0040EF54: fclex 
  0040EF56: jnl 40EF66h
  0040EF58: push 000000C0h
  0040EF5D: push 00404788h
  0040EF62: push esi
  0040EF63: push eax
  0040EF64: call edi
  0040EF66: cmp [00419640h], ebx
  0040EF6C: jnz 40EF7Eh
  0040EF6E: push 00419640h
  0040EF73: push 004039B0h
  0040EF78: call MSVBVM60.DLL.__vbaNew2
  0040EF7E: mov esi, [00419640h] ; 
  0040EF84: lea ecx, var_6C
  0040EF87: push ecx
  0040EF88: push esi
  0040EF89: mov eax, [esi]
  0040EF8B: call [eax+14h]
  0040EF8E: cmp eax, ebx
  0040EF90: fclex 
  0040EF92: jnl 40EF9Fh
  0040EF94: push 00000014h
  0040EF96: push 004039A0h
  0040EF9B: push esi
  0040EF9C: push eax
  0040EF9D: call edi
  0040EF9F: mov eax, var_6C
  0040EFA2: lea ecx, var_148
  0040EFA8: push ecx
  0040EFA9: push eax
  0040EFAA: mov edx, [eax]
  0040EFAC: mov esi, eax
  0040EFAE: call [edx+000000C8h]
  0040EFB4: cmp eax, ebx
  0040EFB6: fclex 
  0040EFB8: jnl 40EFC8h
  0040EFBA: push 000000C8h
  0040EFBF: push 00404788h
  0040EFC4: push esi
  0040EFC5: push eax
  0040EFC6: call edi
  0040EFC8: mov edx, var_140
  0040EFCE: mov ebx, MSVBVM60.DLL.__vbaStrI2
  0040EFD4: push 00404774h ; "PDKS1A V"
  0040EFD9: push edx
  0040EFDA: call ebx
  0040EFDC: mov esi, [0040131Ch]
  0040EFE2: mov edx, eax
  0040EFE4: lea ecx, var_44
  0040EFE7: call MSVBVM60.DLL.__vbaStrMove
  0040EFE9: mov edi, [0040106Ch] ; &
  0040EFEF: push eax
  0040EFF0: call edi
  0040EFF2: mov edx, eax
  0040EFF4: lea ecx, var_48
  0040EFF7: call MSVBVM60.DLL.__vbaStrMove
  0040EFF9: push eax
  0040EFFA: push 004047F4h
  0040EFFF: call edi
  0040F001: mov edx, eax
  0040F003: lea ecx, var_4C
  0040F006: call MSVBVM60.DLL.__vbaStrMove
  0040F008: push eax
  0040F009: mov eax, var_144
  0040F00F: push eax
  0040F010: call ebx
  0040F012: mov edx, eax
  0040F014: lea ecx, var_50
  0040F017: call MSVBVM60.DLL.__vbaStrMove
  0040F019: push eax
  0040F01A: call edi
  0040F01C: mov edx, eax
  0040F01E: lea ecx, var_54
  0040F021: call MSVBVM60.DLL.__vbaStrMove
  0040F023: push eax
  0040F024: push 004047F4h
  0040F029: call edi
  0040F02B: mov edx, eax
  0040F02D: lea ecx, var_58
  0040F030: call MSVBVM60.DLL.__vbaStrMove
  0040F032: mov ecx, var_148
  0040F038: push eax
  0040F039: push ecx
  0040F03A: call ebx
  0040F03C: mov edx, eax
  0040F03E: lea ecx, var_5C
  0040F041: call MSVBVM60.DLL.__vbaStrMove
  0040F043: push eax
  0040F044: call edi
  0040F046: mov edx, eax
  0040F048: mov ecx, 004190F0h
  0040F04D: call MSVBVM60.DLL.__vbaStrMove
  0040F04F: lea edx, var_5C
  0040F052: lea eax, var_58
  0040F055: push edx
  0040F056: lea ecx, var_54
  0040F059: push eax
  0040F05A: lea edx, var_50
  0040F05D: push ecx
  0040F05E: lea eax, var_4C
  0040F061: push edx
  0040F062: lea ecx, var_48
  0040F065: push eax
  0040F066: lea edx, var_44
  0040F069: push ecx
  0040F06A: push edx
  0040F06B: push 00000007h
  0040F06D: call MSVBVM60.DLL.__vbaFreeStrList
  0040F073: lea eax, var_6C
  0040F076: lea ecx, var_68
  0040F079: push eax
  0040F07A: lea edx, var_64
  0040F07D: push ecx
  0040F07E: push edx
  0040F07F: push 00000003h
  0040F081: call MSVBVM60.DLL.__vbaFreeObjList
  0040F087: add esp, 00000030h
  0040F08A: mov edx, 00404690h ; "JOGJASOFTWARE"
  0040F08F: mov ecx, 004190F4h
  0040F094: call MSVBVM60.DLL.__vbaStrCopy
  0040F09A: mov var_104, 00403644h
  0040F0A4: mov var_10C, 00000008h
  0040F0AE: mov ebx, MSVBVM60.DLL.__vbaVarDup
  0040F0B4: lea edx, var_10C
  0040F0BA: lea ecx, var_8C
  0040F0C0: call ebx
  0040F0C2: lea edx, var_FC
  0040F0C8: lea ecx, var_7C
  0040F0CB: mov var_F4, 00403558h
  0040F0D5: mov var_FC, 00000008h
  0040F0DF: call ebx
  0040F0E1: lea eax, var_8C
  0040F0E7: lea ecx, var_7C
  0040F0EA: push eax
  0040F0EB: lea edx, var_EC
  0040F0F1: push ecx
  0040F0F2: lea eax, var_9C
  0040F0F8: push edx
  0040F0F9: push eax
  0040F0FA: mov var_E4, 004190F0h
  0040F104: mov var_EC, 00004008h
  0040F10E: call 00416310h
  0040F113: lea ecx, var_9C
  0040F119: push ecx
  0040F11A: call MSVBVM60.DLL.__vbaStrVarMove
  0040F120: mov edx, eax
  0040F122: mov ecx, 004191D0h
  0040F127: call MSVBVM60.DLL.__vbaStrMove
  0040F129: mov ebx, MSVBVM60.DLL.__vbaFreeVarList
  0040F12F: lea edx, var_9C
  0040F135: lea eax, var_8C
  0040F13B: push edx
  0040F13C: lea ecx, var_7C
  0040F13F: push eax
  0040F140: push ecx
  0040F141: push 00000003h
  0040F143: call ebx
  0040F145: add esp, 00000010h
  0040F148: lea edx, var_10C
  0040F14E: lea ecx, var_8C
  0040F154: mov var_104, 00404704h
  0040F15E: mov var_10C, 00000008h
  0040F168: call MSVBVM60.DLL.__vbaVarDup
  0040F16E: lea edx, var_FC
  0040F174: lea ecx, var_7C
  0040F177: mov var_F4, 00403558h
  0040F181: mov var_FC, 00000008h
  0040F18B: call MSVBVM60.DLL.__vbaVarDup
  0040F191: lea edx, var_8C
  0040F197: lea eax, var_7C
  0040F19A: push edx
  0040F19B: lea ecx, var_EC
  0040F1A1: push eax
  0040F1A2: lea edx, var_9C
  0040F1A8: push ecx
  0040F1A9: push edx
  0040F1AA: mov var_E4, 004190F0h
  0040F1B4: mov var_EC, 00004008h
  0040F1BE: call 00416310h
  0040F1C3: lea eax, var_9C
  0040F1C9: push eax
  0040F1CA: call MSVBVM60.DLL.__vbaStrVarMove
  0040F1D0: mov edx, eax
  0040F1D2: mov ecx, 004191D4h
  0040F1D7: call MSVBVM60.DLL.__vbaStrMove
  0040F1D9: lea ecx, var_9C
  0040F1DF: lea edx, var_8C
  0040F1E5: push ecx
  0040F1E6: lea eax, var_7C
  0040F1E9: push edx
  0040F1EA: push eax
  0040F1EB: push 00000003h
  0040F1ED: call ebx
  0040F1EF: add esp, 00000010h
  0040F1F2: mov edx, 00404800h ; "C:\windows"
  0040F1F7: lea ecx, var_44
  0040F1FA: call MSVBVM60.DLL.__vbaStrCopy
  0040F200: lea ecx, var_44
  0040F203: push ecx
  0040F204: call 0040DDF0h
  0040F209: lea ecx, var_44
  0040F20C: call MSVBVM60.DLL.__vbaFreeStr
  0040F212: mov var_F4, 0040481Ch ; "c:\windows\set"
  0040F21C: lea edx, var_EC
  0040F222: push 00000006h
  0040F224: lea eax, var_7C
  0040F227: push edx
  0040F228: push eax
  0040F229: mov var_FC, 00000008h
  0040F233: mov var_E4, 004190F0h
  0040F23D: mov var_EC, 00004008h
  0040F247: call [0040130Ch] ; Left(arg_1, arg_2)
  0040F24D: lea ecx, var_FC
  0040F253: lea edx, var_7C
  0040F256: push ecx
  0040F257: lea eax, var_8C
  0040F25D: push edx
  0040F25E: push eax
  0040F25F: mov var_104, 00404840h ; ".ini"
  0040F269: mov var_10C, 00000008h
  0040F273: call MSVBVM60.DLL.__vbaVarCat
  0040F279: lea ecx, var_10C
  0040F27F: push eax
  0040F280: lea edx, var_9C
  0040F286: push ecx
  0040F287: push edx
  0040F288: call MSVBVM60.DLL.__vbaVarCat
  0040F28E: push eax
  0040F28F: call MSVBVM60.DLL.__vbaStrVarMove
  0040F295: mov edx, eax
  0040F297: mov ecx, 004191D8h
  0040F29C: call MSVBVM60.DLL.__vbaStrMove
  0040F29E: lea eax, var_9C
  0040F2A4: lea ecx, var_8C
  0040F2AA: push eax
  0040F2AB: lea edx, var_7C
  0040F2AE: push ecx
  0040F2AF: push edx
  0040F2B0: push 00000003h
  0040F2B2: call ebx
  0040F2B4: add esp, 00000010h
  0040F2B7: lea eax, var_7C
  0040F2BA: push eax
  0040F2BB: call [004012C0h] ; Date 'arg_1
  0040F2C1: push 00404850h ; "1/1/2200"
  0040F2C6: call MSVBVM60.DLL.__vbaDateStr
  0040F2CC: fstp real8 ptr var_E4
  0040F2D2: lea ecx, var_7C
  0040F2D5: lea edx, var_EC
  0040F2DB: push ecx
  0040F2DC: push edx
  0040F2DD: mov var_EC, 00008007h
  0040F2E7: call MSVBVM60.DLL.__vbaVarTstGt
  0040F2ED: lea ecx, var_7C
  0040F2F0: mov var_14C, ax
  0040F2F7: call MSVBVM60.DLL.__vbaFreeVar
  0040F2FD: cmp word ptr var_14C, 0000h
  0040F305: jz 0040F3BDh
  0040F30B: mov ecx, 80020004h
  0040F310: mov eax, 0000000Ah
  0040F315: mov var_A4, ecx
  0040F31B: mov var_94, ecx
  0040F321: lea edx, var_FC
  0040F327: lea ecx, var_8C
  0040F32D: mov var_AC, eax
  0040F333: mov var_9C, eax
  0040F339: mov var_F4, 004048ACh ; "Jogjasoftware.net"
  0040F343: mov var_FC, 00000008h
  0040F34D: call MSVBVM60.DLL.__vbaVarDup
  0040F353: lea edx, var_EC
  0040F359: lea ecx, var_7C
  0040F35C: mov var_E4, 00404868h ; "File Pendukung tidak ditemukan"
  0040F366: mov var_EC, 00000008h
  0040F370: call MSVBVM60.DLL.__vbaVarDup
  0040F376: lea eax, var_AC
  0040F37C: lea ecx, var_9C
  0040F382: push eax
  0040F383: lea edx, var_8C
  0040F389: push ecx
  0040F38A: push edx
  0040F38B: lea eax, var_7C
  0040F38E: push 00000010h
  0040F390: push eax
  0040F391: call [004010C0h] ; MsgBox(arg_1, arg_2, arg_3, arg_4, arg_5)
  0040F397: lea ecx, var_AC
  0040F39D: lea edx, var_9C
  0040F3A3: push ecx
  0040F3A4: lea eax, var_8C
  0040F3AA: push edx
  0040F3AB: lea ecx, var_7C
  0040F3AE: push eax
  0040F3AF: push ecx
  0040F3B0: push 00000004h
  0040F3B2: call ebx
  0040F3B4: add esp, 00000014h
  0040F3B7: call [00401038h] ; End
  0040F3BD: mov edx, [004191D8h] ; 
  0040F3C3: push edx
  0040F3C4: call 0040D930h
  0040F3C9: test ax, ax
  0040F3CC: FindClose(%x1)
  0040F3D2: mov eax, [4190F0h]
  0040F3D7: push 004048D4h ; "Your Product ID :"
  0040F3DC: push eax
  0040F3DD: call edi
  0040F3DF: mov edx, eax
  0040F3E1: lea ecx, var_44
  0040F3E4: call MSVBVM60.DLL.__vbaStrMove
  0040F3E6: push eax
  0040F3E7: push 00404900h ; "13#1"
  0040F3EC: call edi
  0040F3EE: mov edx, eax
  0040F3F0: lea ecx, var_48
  0040F3F3: call MSVBVM60.DLL.__vbaStrMove
  0040F3F5: push eax
  0040F3F6: push 0040490Ch ; "Lisence :"
  0040F3FB: call edi
  0040F3FD: mov edx, eax
  0040F3FF: lea ecx, var_4C
  0040F402: call MSVBVM60.DLL.__vbaStrMove
  0040F404: mov ecx, [004190F4h] ; 
  0040F40A: push eax
  0040F40B: push ecx
  0040F40C: call edi
  0040F40E: mov edx, eax
  0040F410: lea ecx, var_50
  0040F413: call MSVBVM60.DLL.__vbaStrMove
  0040F415: push eax
  0040F416: push 00404900h ; "13#1"
  0040F41B: call edi
  0040F41D: mov edx, eax
  0040F41F: lea ecx, var_54
  0040F422: call MSVBVM60.DLL.__vbaStrMove
  0040F424: push eax
  0040F425: push 00404928h ; "Enter your CD Key :"
  0040F42A: call edi
  0040F42C: lea edx, var_7C
  0040F42F: lea ecx, var_30
  0040F432: mov var_74, eax
  0040F435: mov var_7C, 00000008h
  0040F43C: call MSVBVM60.DLL.__vbaVarMove
  0040F442: lea edx, var_54
  0040F445: lea eax, var_50
  0040F448: push edx
  0040F449: lea ecx, var_4C
  0040F44C: push eax
  0040F44D: lea edx, var_48
  0040F450: push ecx
  0040F451: lea eax, var_44
  0040F454: push edx
  0040F455: push eax
  0040F456: push 00000005h
  0040F458: call MSVBVM60.DLL.__vbaFreeStrList
  0040F45E: mov eax, [419640h]
  0040F463: add esp, 00000018h
  0040F466: test eax, eax
  0040F468: jnz 40F47Ah
  0040F46A: push 00419640h
  0040F46F: push 004039B0h
  0040F474: call MSVBVM60.DLL.__vbaNew2
  0040F47A: mov eax, [419640h]
  0040F47F: lea edx, var_64
  0040F482: push edx
  0040F483: push eax
  0040F484: mov ecx, [eax]
  0040F486: mov var_14C, eax
  0040F48C: call [ecx+14h]
  0040F48F: test eax, eax
  0040F491: fclex 
  0040F493: jnl 40F4AAh
  0040F495: mov ecx, var_14C
  0040F49B: push 00000014h
  0040F49D: push 004039A0h
  0040F4A2: push ecx
  0040F4A3: push eax
  0040F4A4: call MSVBVM60.DLL.__vbaHresultCheckObj
  0040F4AA: mov eax, var_64
  0040F4AD: lea ecx, var_44
  0040F4B0: push ecx
  0040F4B1: push eax
  0040F4B2: mov edx, [eax]
  0040F4B4: mov var_154, eax
  0040F4BA: call [edx+50h]
  0040F4BD: test eax, eax
  0040F4BF: fclex 
  0040F4C1: jnl 40F4D8h
  0040F4C3: mov edx, var_154
  0040F4C9: push 00000050h
  0040F4CB: push 00404788h
  0040F4D0: push edx
  0040F4D1: push eax
  0040F4D2: call MSVBVM60.DLL.__vbaHresultCheckObj
  0040F4D8: mov eax, var_44
  0040F4DB: push 00000001h
  0040F4DD: push eax
  0040F4DE: push 00404954h ; "jsgen"
  0040F4E3: push 00000000h
  0040F4E5: call [00401254h] ; InStr
  0040F4EB: xor ecx, ecx
  0040F4ED: test eax, eax
  0040F4EF: setnle cl
  0040F4F2: neg ecx
  0040F4F4: mov var_15C, ecx
  0040F4FA: lea ecx, var_44
  0040F4FD: call MSVBVM60.DLL.__vbaFreeStr
  0040F503: lea ecx, var_64
  0040F506: call MSVBVM60.DLL.__vbaFreeObj
  0040F50C: cmp word ptr var_15C, 0000h
  0040F514: jz 0040F67Ah
  0040F51A: mov edx, [004190F0h] ; 
  0040F520: push 00404964h ; "Product ID :"
  0040F525: push edx
  0040F526: call edi
  0040F528: lea edx, var_7C
  0040F52B: lea ecx, var_40
  0040F52E: mov var_74, eax
  0040F531: mov var_7C, 00000008h
  0040F538: call MSVBVM60.DLL.__vbaVarMove
  0040F53E: mov ecx, [004191D0h] ; 
  0040F544: mov eax, 00000008h
  0040F549: mov var_EC, eax
  0040F54F: mov var_FC, eax
  0040F555: mov var_10C, eax
  0040F55B: lea edx, var_40
  0040F55E: mov var_104, ecx
  0040F564: lea eax, var_EC
  0040F56A: push edx
  0040F56B: lea ecx, var_7C
  0040F56E: push eax
  0040F56F: push ecx
  0040F570: mov var_E4, 00404900h ; "13#1"
  0040F57A: mov var_F4, 00404984h ; "CD Key 1X (TODAY):"
  0040F584: call MSVBVM60.DLL.__vbaVarCat
  0040F58A: push eax
  0040F58B: lea edx, var_FC
  0040F591: lea eax, var_8C
  0040F597: push edx
  0040F598: push eax
  0040F599: call MSVBVM60.DLL.__vbaVarCat
  0040F59F: lea ecx, var_10C
  0040F5A5: push eax
  0040F5A6: lea edx, var_9C
  0040F5AC: push ecx
  0040F5AD: push edx
  0040F5AE: call MSVBVM60.DLL.__vbaVarCat
  0040F5B4: mov edx, eax
  0040F5B6: lea ecx, var_40
  0040F5B9: call MSVBVM60.DLL.__vbaVarMove
  0040F5BF: lea eax, var_8C
  0040F5C5: lea ecx, var_7C
  0040F5C8: push eax
  0040F5C9: push ecx
  0040F5CA: push 00000002h
  0040F5CC: call ebx
  0040F5CE: mov edx, [004191D4h] ; 
  0040F5D4: mov eax, 00000008h
  0040F5D9: add esp, 0000000Ch
  0040F5DC: mov var_EC, eax
  0040F5E2: mov var_FC, eax
  0040F5E8: mov var_10C, eax
  0040F5EE: lea eax, var_40
  0040F5F1: mov var_104, edx
  0040F5F7: lea ecx, var_EC
  0040F5FD: push eax
  0040F5FE: lea edx, var_7C
  0040F601: push ecx
  0040F602: push edx
  0040F603: mov var_E4, 00404900h ; "13#1"
  0040F60D: mov var_F4, 004049C4h ; "CD Key Permanent:"
  0040F617: call MSVBVM60.DLL.__vbaVarCat
  0040F61D: push eax
  0040F61E: lea eax, var_FC
  0040F624: lea ecx, var_8C
  0040F62A: push eax
  0040F62B: push ecx
  0040F62C: call MSVBVM60.DLL.__vbaVarCat
  0040F632: push eax
  0040F633: lea edx, var_10C
  0040F639: lea eax, var_9C
  0040F63F: push edx
  0040F640: push eax
  0040F641: call MSVBVM60.DLL.__vbaVarCat
  0040F647: mov edx, eax
  0040F649: lea ecx, var_40
  0040F64C: call MSVBVM60.DLL.__vbaVarMove
  0040F652: lea ecx, var_8C
  0040F658: lea edx, var_7C
  0040F65B: push ecx
  0040F65C: push edx
  0040F65D: push 00000002h
  0040F65F: call ebx
  0040F661: add esp, 0000000Ch
  0040F664: lea eax, var_40
  0040F667: lea ecx, var_7C
  0040F66A: push eax
  0040F66B: push ecx
  0040F66C: call 00417DE0h
  0040F671: lea ecx, var_7C
  0040F674: call MSVBVM60.DLL.__vbaFreeVar
  0040F67A: mov ecx, 80020004h
  0040F67F: mov eax, 0000000Ah
  0040F684: mov var_C4, ecx
  0040F68A: mov var_B4, ecx
  0040F690: mov var_A4, ecx
  0040F696: mov var_94, ecx
  0040F69C: mov var_84, ecx
  0040F6A2: lea edx, var_EC
  0040F6A8: lea ecx, var_7C
  0040F6AB: mov var_CC, eax
  0040F6B1: mov var_BC, eax
  0040F6B7: mov var_AC, eax
  0040F6BD: mov var_9C, eax
  0040F6C3: mov var_8C, eax
  0040F6C9: mov var_E4, 004049ECh ; "JogjaSoftware"
  0040F6D3: mov var_EC, 00000008h
  0040F6DD: call MSVBVM60.DLL.__vbaVarDup
  0040F6E3: lea edx, var_CC
  0040F6E9: lea eax, var_BC
  0040F6EF: push edx
  0040F6F0: lea ecx, var_AC
  0040F6F6: push eax
  0040F6F7: lea edx, var_9C
  0040F6FD: push ecx
  0040F6FE: lea eax, var_8C
  0040F704: push edx
  0040F705: lea ecx, var_7C
  0040F708: push eax
  0040F709: lea edx, var_30
  0040F70C: push ecx
  0040F70D: push edx
  0040F70E: call [004010C8h] ; InputBox
  0040F714: lea edx, var_DC
  0040F71A: lea ecx, var_20
  0040F71D: mov var_D4, eax
  0040F723: mov var_DC, 00000008h
  0040F72D: call MSVBVM60.DLL.__vbaVarMove
  0040F733: lea eax, var_CC
  0040F739: lea ecx, var_BC
  0040F73F: push eax
  0040F740: lea edx, var_AC
  0040F746: push ecx
  0040F747: lea eax, var_9C
  0040F74D: push edx
  0040F74E: lea ecx, var_8C
  0040F754: push eax
  0040F755: lea edx, var_7C
  0040F758: push ecx
  0040F759: push edx
  0040F75A: push 00000006h
  0040F75C: call ebx
  0040F75E: add esp, 0000001Ch
  0040F761: lea eax, var_20
  0040F764: lea ecx, var_7C
  0040F767: push eax
  0040F768: push ecx
  0040F769: call [00401140h] ; arg_1 = Ucase(arg_2)
  0040F76F: mov edx, [004191D0h] ; 
  0040F775: lea eax, var_20
  0040F778: lea ecx, var_9C
  0040F77E: push eax
  0040F77F: push ecx
  0040F780: mov var_E4, edx
  0040F786: mov var_EC, 00008008h
  0040F790: call [00401140h] ; arg_1 = Ucase(arg_2)
  0040F796: mov edx, [004191D4h] ; 
  0040F79C: lea eax, var_7C
  0040F79F: mov var_F4, edx
  0040F7A5: lea ecx, var_EC
  0040F7AB: push eax
  0040F7AC: lea edx, var_8C
  0040F7B2: push ecx
  0040F7B3: push edx
  0040F7B4: mov var_FC, 00008008h
  0040F7BE: call MSVBVM60.DLL.__vbaVarCmpEq
  0040F7C4: push eax
  0040F7C5: lea eax, var_9C
  0040F7CB: lea ecx, var_FC
  0040F7D1: push eax
  0040F7D2: push ecx
  0040F7D3: lea edx, var_AC
  0040F7D9: push edx
  0040F7DA: call MSVBVM60.DLL.__vbaVarCmpEq
  0040F7E0: push eax
  0040F7E1: lea eax, var_BC
  0040F7E7: push eax
  0040F7E8: call [00401180h] ; Or
  0040F7EE: push eax
  0040F7EF: call MSVBVM60.DLL.__vbaBoolVarNull
  0040F7F5: lea ecx, var_9C
  0040F7FB: lea edx, var_7C
  0040F7FE: push ecx
  0040F7FF: push edx
  0040F800: push 00000002h
  0040F802: mov var_14C, ax
  0040F809: call ebx
  0040F80B: add esp, 0000000Ch
  0040F80E: cmp word ptr var_14C, 0000h
  0040F816: jz 0040F914h
  0040F81C: mov esi, [0040122Ch]
  0040F822: lea eax, var_30
  0040F825: push 004191D8h
  0040F82A: lea ecx, var_EC
  0040F830: push eax
  0040F831: lea edx, var_7C
  0040F834: mov edi, 00000008h
  0040F839: push ecx
  0040F83A: push edx
  0040F83B: mov var_E4, 00404900h ; "13#1"
  0040F845: mov var_EC, edi
  0040F84B: call MSVBVM60.DLL.__vbaVarCat
  0040F84D: push eax
  0040F84E: lea eax, var_20
  0040F851: lea ecx, var_8C
  0040F857: push eax
  0040F858: push ecx
  0040F859: call MSVBVM60.DLL.__vbaVarCat
  0040F85B: lea edx, var_9C
  0040F861: push eax
  0040F862: push edx
  0040F863: call 0040DAF0h
  0040F868: lea eax, var_9C
  0040F86E: lea ecx, var_8C
  0040F874: push eax
  0040F875: lea edx, var_7C
  0040F878: push ecx
  0040F879: push edx
  0040F87A: push 00000003h
  0040F87C: call ebx
  0040F87E: mov ecx, 0000000Ah
  0040F883: mov eax, 80020004h
  0040F888: mov var_AC, ecx
  0040F88E: mov var_9C, ecx
  0040F894: mov var_8C, ecx
  0040F89A: add esp, 00000010h
  0040F89D: lea edx, var_EC
  0040F8A3: lea ecx, var_7C
  0040F8A6: mov var_A4, eax
  0040F8AC: mov var_94, eax
  0040F8B2: mov var_84, eax
  0040F8B8: mov var_E4, 00404A0Ch ; "CD Key Accepted"
  0040F8C2: mov var_EC, edi
  0040F8C8: call MSVBVM60.DLL.__vbaVarDup
  0040F8CE: lea eax, var_AC
  0040F8D4: lea ecx, var_9C
  0040F8DA: push eax
  0040F8DB: lea edx, var_8C
  0040F8E1: push ecx
  0040F8E2: push edx
  0040F8E3: lea eax, var_7C
  0040F8E6: push 00000040h
  0040F8E8: push eax
  0040F8E9: call [004010C0h] ; MsgBox(arg_1, arg_2, arg_3, arg_4, arg_5)
  0040F8EF: lea ecx, var_AC
  0040F8F5: lea edx, var_9C
  0040F8FB: push ecx
  0040F8FC: lea eax, var_8C
  0040F902: push edx
  0040F903: lea ecx, var_7C
  0040F906: push eax
  0040F907: push ecx
  0040F908: push 00000004h
  0040F90A: call ebx
  0040F90C: add esp, 00000014h
  0040F90F: FindClose(%x1)
  0040F914: mov ecx, 80020004h
  0040F919: mov eax, 0000000Ah
  0040F91E: push 00404A30h ; "CD Key tidak valid....."
  0040F923: push 00404900h ; "13#1"
  0040F928: mov var_A4, ecx
  0040F92E: mov var_AC, eax
  0040F934: mov var_94, ecx
  0040F93A: mov var_9C, eax
  0040F940: mov var_84, ecx
  0040F946: mov var_8C, eax
  0040F94C: call edi
  0040F94E: mov edx, eax
  0040F950: lea ecx, var_44
  0040F953: call MSVBVM60.DLL.__vbaVarCat
  0040F955: push eax
  0040F956: push 00404900h ; "13#1"
  0040F95B: call edi
  0040F95D: mov edx, eax
  0040F95F: lea ecx, var_48
  0040F962: call MSVBVM60.DLL.__vbaVarCat
  0040F964: push eax
  0040F965: push 00404A64h ; "Call jogjasoftware.net"
  0040F96A: call edi
  0040F96C: mov edx, eax
  0040F96E: lea ecx, var_4C
  0040F971: call MSVBVM60.DLL.__vbaVarCat
  0040F973: push eax
  0040F974: push 00404900h ; "13#1"
  0040F979: call edi
  0040F97B: mov edx, eax
  0040F97D: lea ecx, var_50
  0040F980: call MSVBVM60.DLL.__vbaVarCat
  0040F982: push eax
  0040F983: push 00404A98h ; "Web : http://www.jogjasoftware.net"
  0040F988: call edi
  0040F98A: mov edx, eax
  0040F98C: lea ecx, var_54
  0040F98F: call MSVBVM60.DLL.__vbaVarCat
  0040F991: push eax
  0040F992: push 00404900h ; "13#1"
  0040F997: call edi
  0040F999: mov edx, eax
  0040F99B: lea ecx, var_58
  0040F99E: call MSVBVM60.DLL.__vbaVarCat
  0040F9A0: push eax
  0040F9A1: push 00404AE4h ; "Email:  info@ jogjasoftware.net"
  0040F9A6: call edi
  0040F9A8: mov edx, eax
  0040F9AA: lea ecx, var_5C
  0040F9AD: call MSVBVM60.DLL.__vbaVarCat
  0040F9AF: push eax
  0040F9B0: push 00404900h ; "13#1"
  0040F9B5: call edi
  0040F9B7: mov edx, eax
  0040F9B9: lea ecx, var_60
  0040F9BC: call MSVBVM60.DLL.__vbaVarCat
  0040F9BE: push eax
  0040F9BF: push 00404B28h ; "Telp : 0812 2993 464"
  0040F9C4: call edi
  0040F9C6: mov var_74, eax
  0040F9C9: lea edx, var_AC
  0040F9CF: lea eax, var_9C
  0040F9D5: push edx
  0040F9D6: lea ecx, var_8C
  0040F9DC: push eax
  0040F9DD: push ecx
  0040F9DE: lea edx, var_7C
  0040F9E1: push 00000010h
  0040F9E3: push edx
  0040F9E4: mov var_7C, 00000008h
  0040F9EB: call [004010C0h] ; MsgBox(arg_1, arg_2, arg_3, arg_4, arg_5)
  0040F9F1: lea eax, var_60
  0040F9F4: lea ecx, var_5C
  0040F9F7: push eax
  0040F9F8: lea edx, var_58
  0040F9FB: push ecx
  0040F9FC: lea eax, var_54
  0040F9FF: push edx
  0040FA00: lea ecx, var_50
  0040FA03: push eax
  0040FA04: push ecx
  0040FA05: lea edx, var_4C
  0040FA08: lea eax, var_48
  0040FA0B: push edx
  0040FA0C: lea ecx, var_44
  0040FA0F: push eax
  0040FA10: push ecx
  0040FA11: push 00000008h
  0040FA13: call MSVBVM60.DLL.__vbaFreeStrList
  0040FA19: lea edx, var_AC
  0040FA1F: lea eax, var_9C
  0040FA25: push edx
  0040FA26: lea ecx, var_8C
  0040FA2C: push eax
  0040FA2D: lea edx, var_7C
  0040FA30: push ecx
  0040FA31: push edx
  0040FA32: push 00000004h
  0040FA34: call ebx
  0040FA36: add esp, 00000038h
  0040FA39: call [00401038h] ; End
  0040FA3F: mov eax, [4191E8h]
  0040FA44: test eax, eax
  0040FA46: jnz 40FA58h
  0040FA48: push 004191E8h
  0040FA4D: push 0040252Ch
  0040FA52: call MSVBVM60.DLL.__vbaNew2
  0040FA58: sub esp, 00000010h
  0040FA5B: mov ecx, 0000000Ah
  0040FA60: mov ebx, esp
  0040FA62: mov var_FC, ecx
  0040FA68: mov eax, 80020004h
  0040FA6D: sub esp, 00000010h
  0040FA70: mov [ebx], ecx
  0040FA72: mov ecx, var_F8
  0040FA78: mov var_F4, eax
  0040FA7E: mov esi, [004191E8h] ; 
  0040FA84: mov [ebx+04h], ecx
  0040FA87: mov var_EC, 00000002h
  0040FA91: mov ecx, esp
  0040FA93: mov edx, 00000001h
  0040FA98: mov [ebx+08h], eax
  0040FA9B: mov eax, var_F0
  0040FAA1: mov var_E4, edx
  0040FAA7: mov edi, [esi]
  0040FAA9: mov [ebx+0Ch], eax
  0040FAAC: mov eax, var_EC
  0040FAB2: mov [ecx], eax
  0040FAB4: mov eax, var_E8
  0040FABA: push esi
  0040FABB: mov [ecx+04h], eax
  0040FABE: mov [ecx+08h], edx
  0040FAC1: mov edx, var_E0
  0040FAC7: mov [ecx+0Ch], edx
  0040FACA: call [edi+000002B0h]
  0040FAD0: test eax, eax
  0040FAD2: fclex 
  0040FAD4: jnl 40FAE8h
  0040FAD6: push 000002B0h
  0040FADB: push 00404B54h
  0040FAE0: push esi
  0040FAE1: push eax
  0040FAE2: call MSVBVM60.DLL.__vbaHresultCheckObj
  0040FAE8: wait 
  0040FAE9: push 0040FB7Ch ; "‹Mð_^d‰'#1"
  0040FAEE: jmp 40FB66h
  0040FAF0: lea eax, var_60
  0040FAF3: lea ecx, var_5C
  0040FAF6: push eax
  0040FAF7: lea edx, var_58
  0040FAFA: push ecx
  0040FAFB: lea eax, var_54
  0040FAFE: push edx
  0040FAFF: lea ecx, var_50
  0040FB02: push eax
  0040FB03: lea edx, var_4C
  0040FB06: push ecx
  0040FB07: lea eax, var_48
  0040FB0A: push edx
  0040FB0B: lea ecx, var_44
  0040FB0E: push eax
  0040FB0F: push ecx
  0040FB10: push 00000008h
  0040FB12: call MSVBVM60.DLL.__vbaFreeStrList
  0040FB18: lea edx, var_6C
  0040FB1B: lea eax, var_68
  0040FB1E: push edx
  0040FB1F: lea ecx, var_64
  0040FB22: push eax
  0040FB23: push ecx
  0040FB24: push 00000003h
  0040FB26: call MSVBVM60.DLL.__vbaFreeObjList
  0040FB2C: lea edx, var_DC
  0040FB32: lea eax, var_CC
  0040FB38: push edx
  0040FB39: lea ecx, var_BC
  0040FB3F: push eax
  0040FB40: lea edx, var_AC
  0040FB46: push ecx
  0040FB47: lea eax, var_9C
  0040FB4D: push edx
  0040FB4E: lea ecx, var_8C
  0040FB54: push eax
  0040FB55: lea edx, var_7C
  0040FB58: push ecx
  0040FB59: push edx
  0040FB5A: push 00000007h
  0040FB5C: call MSVBVM60.DLL.__vbaFreeVarList
  0040FB62: add esp, 00000054h
  0040FB65: ret 
End Sub

