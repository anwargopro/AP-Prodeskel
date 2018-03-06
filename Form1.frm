VERSION 5.00
Begin VB.Form Form1
  Caption = "Entri Data DDK - JogjaSoftware.net"
  BackColor = &H404000&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "Form1.frx":0
  LinkTopic = "Form1"
  ClientLeft = 120
  ClientTop = 450
  ClientWidth = 15900
  ClientHeight = 8370
  StartUpPosition = 3 'Windows Default
  Begin PictureBox pcB
    BackColor = &H4000&
    Left = 0
    Top = 5835
    Width = 15900
    Height = 2535
    TabIndex = 5
    ScaleMode = 1
    AutoRedraw = False
    FontTransparent = True
    Align = 2 'Align Bottom
    Begin PictureBox Picture1
      BackColor = &HFF8080&
      ForeColor = &H80000008&
      Left = 9240
      Top = 120
      Width = 2295
      Height = 375
      Visible = 0   'False
      TabIndex = 24
      ScaleMode = 1
      AutoRedraw = False
      FontTransparent = True
      Appearance = 0 'Flat
      Begin OptionButton Option1
        Caption = "All"
        Index = 2
        BackColor = &HFF8080&
        ForeColor = &H80000008&
        Left = 1320
        Top = 40
        Width = 735
        Height = 255
        TabIndex = 27
        Value = 255
        Appearance = 0 'Flat
      End
      Begin OptionButton Option1
        Caption = "AK"
        Index = 1
        BackColor = &HFF8080&
        ForeColor = &H80000008&
        Left = 720
        Top = 40
        Width = 615
        Height = 255
        TabIndex = 26
        Appearance = 0 'Flat
      End
      Begin OptionButton Option1
        Caption = "KK"
        Index = 0
        BackColor = &HFF8080&
        ForeColor = &H80000008&
        Left = 120
        Top = 40
        Width = 615
        Height = 255
        TabIndex = 25
        Appearance = 0 'Flat
      End
    End
    Begin CommandButton tbMasukkan
      Caption = "Masukkan Data AK"
      Index = 1
      Left = 4680
      Top = 120
      Width = 2055
      Height = 375
      TabIndex = 23
    End
    Begin CommandButton tbMasukkan
      Caption = "Masukkan Data KK"
      Index = 0
      Left = 2520
      Top = 120
      Width = 2055
      Height = 375
      TabIndex = 18
    End
    Begin CommandButton tbGet
      Caption = "Paste  Data From Clipboard"
      Left = 120
      Top = 120
      Width = 2175
      Height = 375
      TabIndex = 7
    End
    Begin MSHFlexGrid msh1
      Left = 0
      Top = 600
      Width = 15855
      Height = 1935
      TabIndex = 6
    End
  End
  Begin PictureBox pcR
    BackColor = &HFFC0C0&
    Left = 12885
    Top = 0
    Width = 3015
    Height = 5835
    TabIndex = 1
    ScaleMode = 1
    AutoRedraw = False
    FontTransparent = True
    Align = 4 'Align Right
    Begin CommandButton tbRCdKey
      Caption = "Remove CDKey"
      Left = 480
      Top = 1920
      Width = 2055
      Height = 375
      Visible = 0   'False
      TabIndex = 29
    End
    Begin CommandButton Command1
      Caption = ">> Daftar KK"
      Index = 4
      Left = 240
      Top = 1200
      Width = 1335
      Height = 375
      TabIndex = 28
    End
    Begin CommandButton Command1
      Caption = ">>Menu"
      Index = 2
      Left = 1680
      Top = 720
      Width = 1215
      Height = 375
      TabIndex = 19
    End
    Begin TextBox txtDusun
      Left = 1440
      Top = 2760
      Width = 1455
      Height = 375
      TabIndex = 12
    End
    Begin TextBox txtSumber
      Left = 1440
      Top = 4680
      Width = 1455
      Height = 375
      TabIndex = 11
    End
    Begin TextBox txtJab
      Left = 1440
      Top = 4200
      Width = 1455
      Height = 375
      TabIndex = 10
    End
    Begin TextBox txtPek
      Left = 1440
      Top = 3720
      Width = 1455
      Height = 375
      TabIndex = 9
    End
    Begin TextBox txtPengisi
      Left = 1440
      Top = 3240
      Width = 1455
      Height = 375
      TabIndex = 8
    End
    Begin CommandButton Command1
      Caption = ">> Login"
      Index = 0
      Left = 240
      Top = 720
      Width = 1335
      Height = 375
      TabIndex = 4
    End
    Begin CommandButton tbInfo2
      Caption = "Informasi Penggunaan"
      Left = 240
      Top = 120
      Width = 2655
      Height = 375
      TabIndex = 3
    End
    Begin CommandButton Command1
      Caption = ">> Form KK"
      Index = 1
      Left = 1680
      Top = 1200
      Width = 1215
      Height = 375
      TabIndex = 2
    End
    Begin Label Label1
      Caption = "Sumber Data"
      Index = 4
      Left = 240
      Top = 4680
      Width = 1215
      Height = 255
      TabIndex = 17
      BackStyle = 0 'Transparent
    End
    Begin Label Label1
      Caption = "Jabatan"
      Index = 3
      Left = 240
      Top = 4200
      Width = 735
      Height = 255
      TabIndex = 16
      BackStyle = 0 'Transparent
    End
    Begin Label Label1
      Caption = "Pekerjaan"
      Index = 2
      Left = 240
      Top = 3720
      Width = 735
      Height = 255
      TabIndex = 15
      BackStyle = 0 'Transparent
    End
    Begin Label Label1
      Caption = "Pengisi"
      Index = 1
      Left = 240
      Top = 3240
      Width = 855
      Height = 255
      TabIndex = 14
      BackStyle = 0 'Transparent
    End
    Begin Label Label1
      Caption = "Dusun"
      Index = 0
      Left = 240
      Top = 2880
      Width = 495
      Height = 255
      TabIndex = 13
      BackStyle = 0 'Transparent
    End
  End
  Begin PictureBox pcInfo
    BackColor = &HFFC0C0&
    ForeColor = &H80000008&
    Left = 1560
    Top = 240
    Width = 10335
    Height = 6615
    TabIndex = 20
    ScaleMode = 1
    AutoRedraw = False
    FontTransparent = True
    Appearance = 0 'Flat
    Begin TextBox txtInfo
      BackColor = &HC0FFFF&
      Left = 120
      Top = 480
      Width = 10095
      Height = 6015
      Text = "Text1"
      TabIndex = 22
      MultiLine = -1  'True
      ScrollBars = 2
      Locked = -1  'True
      Appearance = 0 'Flat
    End
    Begin CommandButton Command4
      Caption = "x"
      Left = 9840
      Top = 120
      Width = 375
      Height = 255
      TabIndex = 21
    End
  End
  Begin WebBrowser WebBrowser1
    Left = 0
    Top = 0
    Width = 12495
    Height = 6255
    Visible = 0   'False
    TabIndex = 0
  End
End

Attribute VB_Name = "Form1"


Private Sub tbInfo2_Click() '415AE0
  00415AE0: push ebp
  00415AE1: mov ebp, esp
  00415AE3: sub esp, 0000000Ch
  00415AE6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00415AEB: mov eax, fs:[00h]
  00415AF1: push eax
  00415AF2: mov fs:[00000000h], esp
  00415AF9: sub esp, 00000014h
  00415AFC: push ebx
  00415AFD: push esi
  00415AFE: push edi
  00415AFF: mov var_C, esp
  00415B02: mov var_8, 00401510h
  00415B09: mov esi, arg_8
  00415B0C: mov eax, esi
  00415B0E: and eax, 00000001h
  00415B11: mov var_4, eax
  00415B14: and esi, FFFFFFFEh
  00415B17: push esi
  00415B18: mov arg_8, esi
  00415B1B: mov ecx, [esi]
  00415B1D: call [ecx+04h]
  00415B20: mov edx, [esi]
  00415B22: xor edi, edi
  00415B24: push esi
  00415B25: mov var_18, edi
  00415B28: call [edx+00000334h]
  00415B2E: push eax
  00415B2F: lea eax, var_18
  00415B32: push eax
  00415B33: call [004010BCh] ; Set (object)
  00415B39: mov esi, eax
  00415B3B: push FFFFFFFFh
  00415B3D: push esi
  00415B3E: mov ecx, [esi]
  00415B40: call [ecx+0000009Ch]
  00415B46: cmp eax, edi
  00415B48: fclex 
  00415B4A: jnl 415B5Eh
  00415B4C: push 0000009Ch
  00415B51: push 004055A4h
  00415B56: push esi
  00415B57: push eax
  00415B58: call MSVBVM60.DLL.__vbaHresultCheckObj
  00415B5E: lea ecx, var_18
  00415B61: call MSVBVM60.DLL.__vbaFreeObj
  00415B67: mov var_4, edi
  00415B6A: push 00415B7Ch
  00415B6F: jmp 415B7Bh
  00415B71: lea ecx, var_18
  00415B74: call MSVBVM60.DLL.__vbaFreeObj
  00415B7A: ret 
End Sub

Private Sub Command1_Change(Index As Integer) '4144A0
  004144A0: push ebp
  004144A1: mov ebp, esp
  004144A3: sub esp, 0000000Ch
  004144A6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  004144AB: mov eax, fs:[00h]
  004144B1: push eax
  004144B2: mov fs:[00000000h], esp
  004144B9: sub esp, 00000048h
  004144BC: push ebx
  004144BD: push esi
  004144BE: push edi
  004144BF: mov var_C, esp
  004144C2: mov var_8, 00401470h
  004144C9: mov esi, arg_8
  004144CC: mov eax, esi
  004144CE: and eax, 00000001h
  004144D1: mov var_4, eax
  004144D4: and esi, FFFFFFFEh
  004144D7: push esi
  004144D8: mov arg_8, esi
  004144DB: mov ecx, [esi]
  004144DD: call [ecx+04h]
  004144E0: mov edx, Index
  004144E3: xor ecx, ecx
  004144E5: mov var_18, ecx
  004144E8: mov var_1C, ecx
  004144EB: mov ax, [edx]
  004144EE: mov var_2C, ecx
  004144F1: cmp ax, cx
  004144F4: jnz 414512h
  004144F6: mov edi, MSVBVM60.DLL.__vbaStrCopy
  004144FC: mov edx, 00406528h ; "http://prodeskel.binapemdes.kemendagri.go.id/app_Login/"
  00414501: lea ecx, var_18
  00414504: call edi
  00414506: mov edx, 0040659Ch ; "http://prodeskel.binapemdes.kemendagri.go.id/mpublik/mpublik_form_php.php?sc_item_menu=item_3&sc_apl_menu=app_Login&sc_apl_link=%2F&sc_usa_grupo="
  0041450B: lea ecx, var_18
  0041450E: call edi
  00414510: jmp 41454Dh
  00414512: cmp ax, 0001h
  00414516: jnz 41451Fh
  00414518: mov edx, 00406704h ; "http://prodeskel.binapemdes.kemendagri.go.id/form_ddk01/"
  0041451D: jmp 414544h
  0041451F: cmp ax, 0002h
  00414523: jnz 41452Ch
  00414525: mov edx, 0040677Ch ; "http://prodeskel.binapemdes.kemendagri.go.id/mdesa/"
  0041452A: jmp 414544h
  0041452C: cmp ax, 0003h
  00414530: jnz 414539h
  00414532: mov edx, 004067E8h ; "http://prodeskel.binapemdes.kemendagri.go.id/form_ddk02/"
  00414537: jmp 414544h
  00414539: cmp ax, 0004h
  0041453D: jnz 41454Dh
  0041453F: mov edx, 00406860h ; "http://prodeskel.binapemdes.kemendagri.go.id/grid_ddk01/"
  00414544: lea ecx, var_18
  00414547: call MSVBVM60.DLL.__vbaStrCopy
  0041454D: sub esp, 00000010h
  00414550: mov ecx, 00004008h
  00414555: mov edx, esp
  00414557: lea eax, var_18
  0041455A: push 00000001h
  0041455C: push 00000068h
  0041455E: mov [edx], ecx
  00414560: mov ecx, var_38
  00414563: push esi
  00414564: mov [edx+04h], ecx
  00414567: mov ecx, [esi]
  00414569: mov [edx+08h], eax
  0041456C: mov eax, var_30
  0041456F: mov [edx+0Ch], eax
  00414572: call [ecx+00000344h]
  00414578: mov edi, [004010BCh] ; Set (object)
  0041457E: lea edx, var_1C
  00414581: push eax
  00414582: push edx
  00414583: call edi
  00414585: push eax
  00414586: call MSVBVM60.DLL.__vbaLateIdCall
  0041458C: add esp, 0000001Ch
  0041458F: lea ecx, var_1C
  00414592: call MSVBVM60.DLL.__vbaFreeObj
  00414598: mov eax, [esi]
  0041459A: push esi
  0041459B: call [eax+00000334h]
  004145A1: lea ecx, var_1C
  004145A4: push eax
  004145A5: push ecx
  004145A6: call edi
  004145A8: mov ebx, eax
  004145AA: push 00000000h
  004145AC: push ebx
  004145AD: mov edx, [ebx]
  004145AF: call [edx+0000009Ch]
  004145B5: test eax, eax
  004145B7: fclex 
  004145B9: jnl 4145CDh
  004145BB: push 0000009Ch
  004145C0: push 004055A4h
  004145C5: push ebx
  004145C6: push eax
  004145C7: call MSVBVM60.DLL.__vbaHresultCheckObj
  004145CD: lea ecx, var_1C
  004145D0: call MSVBVM60.DLL.__vbaFreeObj
  004145D6: mov eax, [esi]
  004145D8: push 00000000h
  004145DA: push 80010007h
  004145DF: push esi
  004145E0: call [eax+00000344h]
  004145E6: lea ecx, var_1C
  004145E9: push eax
  004145EA: push ecx
  004145EB: call edi
  004145ED: lea edx, var_2C
  004145F0: push eax
  004145F1: push edx
  004145F2: call MSVBVM60.DLL.__vbaLateIdCallLd
  004145F8: add esp, 00000010h
  004145FB: push eax
  004145FC: call MSVBVM60.DLL.__vbaBoolVar
  00414602: mov bx, ax
  00414605: lea ecx, var_1C
  00414608: not ebx
  0041460A: call MSVBVM60.DLL.__vbaFreeObj
  00414610: lea ecx, var_2C
  00414613: call MSVBVM60.DLL.__vbaFreeVar
  00414619: test bx, bx
  0041461C: jz 414661h
  0041461E: sub esp, 00000010h
  00414621: mov ecx, 0000000Bh
  00414626: mov edx, esp
  00414628: or eax, FFFFFFFFh
  0041462B: push 80010007h
  00414630: push esi
  00414631: mov [edx], ecx
  00414633: mov ecx, var_38
  00414636: mov [edx+04h], ecx
  00414639: mov ecx, [esi]
  0041463B: mov [edx+08h], eax
  0041463E: mov eax, var_30
  00414641: mov [edx+0Ch], eax
  00414644: call [ecx+00000344h]
  0041464A: lea edx, var_1C
  0041464D: push eax
  0041464E: push edx
  0041464F: call edi
  00414651: push eax
  00414652: call MSVBVM60.DLL.__vbaLateIdSt
  00414658: lea ecx, var_1C
  0041465B: call MSVBVM60.DLL.__vbaFreeObj
  00414661: mov var_4, 00000000h
  00414668: push 0041468Ch
  0041466D: jmp 414682h
  0041466F: lea ecx, var_1C
  00414672: call MSVBVM60.DLL.__vbaFreeObj
  00414678: lea ecx, var_2C
  0041467B: call MSVBVM60.DLL.__vbaFreeVar
  00414681: ret 
End Sub

Private Sub tbMasukkan_Change(Index As Integer) '415BA0
  00415BA0: push ebp
  00415BA1: mov ebp, esp
  00415BA3: sub esp, 0000000Ch
  00415BA6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00415BAB: mov eax, fs:[00h]
  00415BB1: push eax
  00415BB2: mov fs:[00000000h], esp
  00415BB9: sub esp, 00000008h
  00415BBC: push ebx
  00415BBD: push esi
  00415BBE: push edi
  00415BBF: mov var_C, esp
  00415BC2: mov var_8, 00401520h
  00415BC9: mov esi, arg_8
  00415BCC: mov eax, esi
  00415BCE: and eax, 00000001h
  00415BD1: mov var_4, eax
  00415BD4: and esi, FFFFFFFEh
  00415BD7: push esi
  00415BD8: mov arg_8, esi
  00415BDB: mov ecx, [esi]
  00415BDD: call [ecx+04h]
  00415BE0: mov eax, Index
  00415BE3: mov edx, [esi]
  00415BE5: push eax
  00415BE6: push esi
  00415BE7: call [edx+00000704h]
  00415BED: mov var_4, 00000000h
  00415BF4: mov eax, arg_8
  00415BF7: push eax
  00415BF8: mov ecx, [eax]
  00415BFA: call [ecx+08h]
  00415BFD: mov eax, var_4
  00415C00: mov ecx, var_14
  00415C03: pop edi
  00415C04: pop esi
  00415C05: mov fs:[00000000h], ecx
  00415C0C: pop ebx
  00415C0D: mov esp, ebp
  00415C0F: pop ebp
  00415C10: retn 0008h
End Sub

Private Sub tbGet_Click() '415470
  00415470: push ebp
  00415471: mov ebp, esp
  00415473: sub esp, 0000000Ch
  00415476: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0041547B: mov eax, fs:[00h]
  00415481: push eax
  00415482: mov fs:[00000000h], esp
  00415489: sub esp, 000000F4h
  0041548F: push ebx
  00415490: push esi
  00415491: push edi
  00415492: mov var_C, esp
  00415495: mov var_8, 00401500h
  0041549C: mov esi, arg_8
  0041549F: mov eax, esi
  004154A1: and eax, 00000001h
  004154A4: mov var_4, eax
  004154A7: and esi, FFFFFFFEh
  004154AA: push esi
  004154AB: mov arg_8, esi
  004154AE: mov ecx, [esi]
  004154B0: call [ecx+04h]
  004154B3: mov edx, [esi]
  004154B5: xor eax, eax
  004154B7: push esi
  004154B8: mov var_28, eax
  004154BB: mov var_34, eax
  004154BE: mov var_38, eax
  004154C1: mov var_3C, eax
  004154C4: mov var_40, eax
  004154C7: mov var_50, eax
  004154CA: mov var_60, eax
  004154CD: mov var_70, eax
  004154D0: mov var_90, eax
  004154D6: mov var_B0, eax
  004154DC: mov var_D0, eax
  004154E2: call [edx+00000340h]
  004154E8: mov edi, [004010BCh] ; Set (object)
  004154EE: push eax
  004154EF: lea eax, var_3C
  004154F2: push eax
  004154F3: call edi
  004154F5: lea ecx, var_3C
  004154F8: lea edx, var_50
  004154FB: push ecx
  004154FC: push edx
  004154FD: call 0040DF60h
  00415502: mov ebx, MSVBVM60.DLL.__vbaFreeObj
  00415508: lea ecx, var_3C
  0041550B: call ebx
  0041550D: lea ecx, var_50
  00415510: call MSVBVM60.DLL.__vbaFreeVar
  00415516: mov eax, [esi]
  00415518: push 00000000h
  0041551A: push 00000004h
  0041551C: push esi
  0041551D: call [eax+00000340h]
  00415523: lea ecx, var_3C
  00415526: push eax
  00415527: push ecx
  00415528: call edi
  0041552A: lea edx, var_50
  0041552D: push eax
  0041552E: push edx
  0041552F: call MSVBVM60.DLL.__vbaLateIdCallLd
  00415535: add esp, 00000010h
  00415538: push eax
  00415539: call MSVBVM60.DLL.__vbaI4Var
  0041553F: mov ecx, eax
  00415541: call MSVBVM60.DLL.__vbaI2I4
  00415547: lea ecx, var_3C
  0041554A: mov var_2C, eax
  0041554D: call ebx
  0041554F: lea ecx, var_50
  00415552: call MSVBVM60.DLL.__vbaFreeVar
  00415558: mov eax, [esi]
  0041555A: push 00000000h
  0041555C: push 00000005h
  0041555E: push esi
  0041555F: call [eax+00000340h]
  00415565: lea ecx, var_3C
  00415568: push eax
  00415569: push ecx
  0041556A: call edi
  0041556C: lea edx, var_50
  0041556F: push eax
  00415570: push edx
  00415571: call MSVBVM60.DLL.__vbaLateIdCallLd
  00415577: add esp, 00000010h
  0041557A: push eax
  0041557B: call MSVBVM60.DLL.__vbaI4Var
  00415581: add eax, 00000001h
  00415584: lea edx, var_70
  00415587: jo 00415AD5h
  0041558D: lea ecx, var_28
  00415590: mov var_68, eax
  00415593: mov var_70, 00000003h
  0041559A: call MSVBVM60.DLL.__vbaVarMove
  004155A0: lea ecx, var_3C
  004155A3: call ebx
  004155A5: lea ecx, var_50
  004155A8: call MSVBVM60.DLL.__vbaFreeVar
  004155AE: lea eax, var_28
  004155B1: lea ecx, var_70
  004155B4: push eax
  004155B5: lea edx, var_50
  004155B8: push ecx
  004155B9: push edx
  004155BA: mov var_68, 00000001h
  004155C1: mov var_70, 00000002h
  004155C8: call MSVBVM60.DLL.__vbaVarAdd
  004155CE: push eax
  004155CF: call MSVBVM60.DLL.__vbaI4Var
  004155D5: sub esp, 00000010h
  004155D8: mov ecx, 00000003h
  004155DD: mov edx, esp
  004155DF: push 00000005h
  004155E1: push esi
  004155E2: mov [edx], ecx
  004155E4: mov ecx, var_7C
  004155E7: mov [edx+04h], ecx
  004155EA: mov ecx, [esi]
  004155EC: mov [edx+08h], eax
  004155EF: mov eax, var_74
  004155F2: mov [edx+0Ch], eax
  004155F5: call [ecx+00000340h]
  004155FB: lea edx, var_3C
  004155FE: push eax
  004155FF: push edx
  00415600: call edi
  00415602: push eax
  00415603: call MSVBVM60.DLL.__vbaLateIdSt
  00415609: lea ecx, var_3C
  0041560C: call ebx
  0041560E: lea ecx, var_50
  00415611: call MSVBVM60.DLL.__vbaFreeVar
  00415617: lea eax, var_28
  0041561A: lea ecx, var_70
  0041561D: push eax
  0041561E: lea edx, var_50
  00415621: push ecx
  00415622: push edx
  00415623: mov var_68, 00000001h
  0041562A: mov var_70, 00000002h
  00415631: call MSVBVM60.DLL.__vbaVarSub
  00415637: push eax
  00415638: call MSVBVM60.DLL.__vbaI4Var
  0041563E: sub esp, 00000010h
  00415641: mov ecx, 00000003h
  00415646: mov edx, esp
  00415648: sub esp, 00000010h
  0041564B: mov [edx], ecx
  0041564D: mov ecx, var_7C
  00415650: mov [edx+04h], ecx
  00415653: xor ecx, ecx
  00415655: mov [edx+08h], ecx
  00415658: mov ecx, var_74
  0041565B: mov [edx+0Ch], ecx
  0041565E: mov edx, esp
  00415660: mov ecx, 00000003h
  00415665: sub esp, 00000010h
  00415668: mov [edx], ecx
  0041566A: mov ecx, var_9C
  00415670: mov [edx+04h], ecx
  00415673: mov ecx, esp
  00415675: mov [edx+08h], eax
  00415678: mov eax, var_94
  0041567E: mov [edx+0Ch], eax
  00415681: mov eax, 00000008h
  00415686: mov [ecx], eax
  00415688: mov edx, var_BC
  0041568E: mov eax, 004069C8h ; "NOURUT"
  00415693: mov [ecx+04h], edx
  00415696: push 00000002h
  00415698: push 00000041h
  0041569A: push esi
  0041569B: mov [ecx+08h], eax
  0041569E: mov eax, var_B4
  004156A4: mov [ecx+0Ch], eax
  004156A7: mov ecx, [esi]
  004156A9: call [ecx+00000340h]
  004156AF: lea edx, var_3C
  004156B2: push eax
  004156B3: push edx
  004156B4: call edi
  004156B6: push eax
  004156B7: call MSVBVM60.DLL.__vbaLateIdCallSt
  004156BD: add esp, 0000003Ch
  004156C0: lea ecx, var_3C
  004156C3: call ebx
  004156C5: mov eax, var_2C
  004156C8: mov var_18, 00000000h
  004156CF: sub ax, 0001h
  004156D3: jo 00415AD5h
  004156D9: mov var_F4, eax
  004156DF: mov eax, 00000001h
  004156E4: mov var_30, eax
  004156E7: cmp ax, var_F4
  004156EE: jnle 00415A69h
  004156F4: sub ax, 0001h
  004156F8: mov ecx, 00000003h
  004156FD: jo 00415AD5h
  00415703: sub esp, 00000010h
  00415706: mov var_D0, ecx
  0041570C: mov edx, esp
  0041570E: sub esp, 00000010h
  00415711: movsx eax, ax
  00415714: mov [edx], ecx
  00415716: mov ecx, var_AC
  0041571C: mov var_C8, 00000001h
  00415726: mov [edx+04h], ecx
  00415729: mov ecx, esp
  0041572B: push 00000002h
  0041572D: push 00000041h
  0041572F: mov [edx+08h], eax
  00415732: mov eax, var_A4
  00415738: push esi
  00415739: mov [edx+0Ch], eax
  0041573C: mov edx, var_D0
  00415742: mov eax, var_CC
  00415748: mov [ecx], edx
  0041574A: mov edx, var_C8
  00415750: mov [ecx+04h], eax
  00415753: mov eax, var_C4
  00415759: mov [ecx+08h], edx
  0041575C: mov [ecx+0Ch], eax
  0041575F: mov ecx, [esi]
  00415761: call [ecx+00000340h]
  00415767: lea edx, var_40
  0041576A: push eax
  0041576B: push edx
  0041576C: call edi
  0041576E: push eax
  0041576F: lea eax, var_60
  00415772: push eax
  00415773: call MSVBVM60.DLL.__vbaLateIdCallLd
  00415779: add esp, 00000030h
  0041577C: push eax
  0041577D: call MSVBVM60.DLL.__vbaStrVarMove
  00415783: mov edx, eax
  00415785: lea ecx, var_38
  00415788: call MSVBVM60.DLL.__vbaStrMove
  0041578E: push eax
  0041578F: call [00401370h] ; Val(arg_1)
  00415795: movsx eax, word ptr var_30
  00415799: fstp real8 ptr var_E8
  0041579F: sub esp, 00000010h
  004157A2: mov ecx, 00000003h
  004157A7: mov edx, esp
  004157A9: mov var_70, ecx
  004157AC: mov var_90, ecx
  004157B2: mov var_68, eax
  004157B5: mov [edx], ecx
  004157B7: mov ecx, var_6C
  004157BA: sub esp, 00000010h
  004157BD: mov [edx+04h], ecx
  004157C0: mov ecx, esp
  004157C2: push 00000002h
  004157C4: push 00000041h
  004157C6: mov [edx+08h], eax
  004157C9: mov eax, var_64
  004157CC: push esi
  004157CD: mov [edx+0Ch], eax
  004157D0: mov edx, var_90
  004157D6: mov eax, var_8C
  004157DC: mov [ecx], edx
  004157DE: mov edx, var_84
  004157E4: mov [ecx+04h], eax
  004157E7: mov eax, 00000001h
  004157EC: mov [ecx+08h], eax
  004157EF: mov eax, [esi]
  004157F1: mov [ecx+0Ch], edx
  004157F4: call [eax+00000340h]
  004157FA: lea ecx, var_3C
  004157FD: push eax
  004157FE: push ecx
  004157FF: call edi
  00415801: lea edx, var_50
  00415804: push eax
  00415805: push edx
  00415806: call MSVBVM60.DLL.__vbaLateIdCallLd
  0041580C: add esp, 00000030h
  0041580F: push eax
  00415810: call MSVBVM60.DLL.__vbaStrVarMove
  00415816: mov edx, eax
  00415818: lea ecx, var_34
  0041581B: call MSVBVM60.DLL.__vbaStrMove
  00415821: push eax
  00415822: call [00401370h] ; Val(arg_1)
  00415828: call MSVBVM60.DLL.__vbaFpR8
  0041582E: fstp real8 ptr var_104
  00415834: fld real8 ptr var_E8
  0041583A: call MSVBVM60.DLL.__vbaFpR8
  00415840: fcomp real8 ptr var_104
  00415846: mov var_108, 00000001h
  00415850: fstsw ax
  00415852: test ah, 40h
  00415855: jz 415861h
  00415857: mov var_108, 00000000h
  00415861: lea eax, var_38
  00415864: lea ecx, var_34
  00415867: push eax
  00415868: push ecx
  00415869: push 00000002h
  0041586B: call MSVBVM60.DLL.__vbaFreeStrList
  00415871: lea edx, var_40
  00415874: lea eax, var_3C
  00415877: push edx
  00415878: push eax
  00415879: push 00000002h
  0041587B: call MSVBVM60.DLL.__vbaFreeObjList
  00415881: lea ecx, var_60
  00415884: lea edx, var_50
  00415887: push ecx
  00415888: push edx
  00415889: push 00000002h
  0041588B: call MSVBVM60.DLL.__vbaFreeVarList
  00415891: mov eax, var_108
  00415897: add esp, 00000024h
  0041589A: neg eax
  0041589C: test ax, ax
  0041589F: jz 00415979h
  004158A5: movsx eax, word ptr var_30
  004158A9: sub esp, 00000010h
  004158AC: mov ecx, 00000003h
  004158B1: mov edx, esp
  004158B3: mov var_70, ecx
  004158B6: mov var_68, eax
  004158B9: push 0000000Ah
  004158BB: mov [edx], ecx
  004158BD: mov ecx, var_6C
  004158C0: push esi
  004158C1: mov [edx+04h], ecx
  004158C4: mov ecx, [esi]
  004158C6: mov [edx+08h], eax
  004158C9: mov eax, var_64
  004158CC: mov [edx+0Ch], eax
  004158CF: call [ecx+00000340h]
  004158D5: lea edx, var_3C
  004158D8: push eax
  004158D9: push edx
  004158DA: call edi
  004158DC: push eax
  004158DD: call MSVBVM60.DLL.__vbaLateIdSt
  004158E3: lea ecx, var_3C
  004158E6: call ebx
  004158E8: sub esp, 00000010h
  004158EB: mov ecx, 00000003h
  004158F0: mov edx, esp
  004158F2: mov var_70, ecx
  004158F5: mov eax, 00000001h
  004158FA: push 0000000Bh
  004158FC: mov [edx], ecx
  004158FE: mov ecx, var_6C
  00415901: mov var_68, eax
  00415904: push esi
  00415905: mov [edx+04h], ecx
  00415908: mov ecx, [esi]
  0041590A: mov [edx+08h], eax
  0041590D: mov eax, var_64
  00415910: mov [edx+0Ch], eax
  00415913: call [ecx+00000340h]
  00415919: lea edx, var_3C
  0041591C: push eax
  0041591D: push edx
  0041591E: call edi
  00415920: push eax
  00415921: call MSVBVM60.DLL.__vbaLateIdSt
  00415927: lea ecx, var_3C
  0041592A: call ebx
  0041592C: sub esp, 00000010h
  0041592F: mov ecx, 00000003h
  00415934: mov edx, esp
  00415936: mov var_70, ecx
  00415939: mov eax, 0000FF00h
  0041593E: push 00000026h
  00415940: mov [edx], ecx
  00415942: mov ecx, var_6C
  00415945: mov var_68, eax
  00415948: push esi
  00415949: mov [edx+04h], ecx
  0041594C: mov ecx, [esi]
  0041594E: mov [edx+08h], eax
  00415951: mov eax, var_64
  00415954: mov [edx+0Ch], eax
  00415957: call [ecx+00000340h]
  0041595D: lea edx, var_3C
  00415960: push eax
  00415961: push edx
  00415962: call edi
  00415964: push eax
  00415965: call MSVBVM60.DLL.__vbaLateIdSt
  0041596B: lea ecx, var_3C
  0041596E: call ebx
  00415970: mov var_18, 00000001h
  00415977: jmp 41598Ah
  00415979: mov ax, var_18
  0041597D: add ax, 0001h
  00415981: jo 00415AD5h
  00415987: mov var_18, eax
  0041598A: lea ecx, var_28
  0041598D: lea edx, var_70
  00415990: push ecx
  00415991: lea eax, var_50
  00415994: push edx
  00415995: push eax
  00415996: mov var_68, 00000001h
  0041599D: mov var_70, 00000002h
  004159A4: call MSVBVM60.DLL.__vbaVarSub
  004159AA: push eax
  004159AB: call MSVBVM60.DLL.__vbaI4Var
  004159B1: mov ecx, var_18
  004159B4: mov var_98, eax
  004159BA: push ecx
  004159BB: call MSVBVM60.DLL.__vbaStrI2
  004159C1: sub esp, 00000010h
  004159C4: mov ecx, 00000003h
  004159C9: mov edx, esp
  004159CB: sub esp, 00000010h
  004159CE: mov var_60, 00000008h
  004159D5: mov var_58, eax
  004159D8: mov [edx], ecx
  004159DA: mov ecx, var_7C
  004159DD: mov [edx+04h], ecx
  004159E0: movsx ecx, word ptr var_30
  004159E4: mov [edx+08h], ecx
  004159E7: mov ecx, var_74
  004159EA: mov [edx+0Ch], ecx
  004159ED: mov edx, esp
  004159EF: mov ecx, 00000003h
  004159F4: sub esp, 00000010h
  004159F7: mov [edx], ecx
  004159F9: mov ecx, var_9C
  004159FF: mov [edx+04h], ecx
  00415A02: mov ecx, var_98
  00415A08: mov [edx+08h], ecx
  00415A0B: mov ecx, var_94
  00415A11: mov [edx+0Ch], ecx
  00415A14: mov ecx, var_60
  00415A17: mov edx, esp
  00415A19: push 00000002h
  00415A1B: push 00000041h
  00415A1D: push esi
  00415A1E: mov [edx], ecx
  00415A20: mov ecx, var_5C
  00415A23: mov [edx+04h], ecx
  00415A26: mov ecx, [esi]
  00415A28: mov [edx+08h], eax
  00415A2B: mov eax, var_54
  00415A2E: mov [edx+0Ch], eax
  00415A31: call [ecx+00000340h]
  00415A37: lea edx, var_3C
  00415A3A: push eax
  00415A3B: push edx
  00415A3C: call edi
  00415A3E: push eax
  00415A3F: call MSVBVM60.DLL.__vbaLateIdCallSt
  00415A45: add esp, 0000003Ch
  00415A48: lea ecx, var_3C
  00415A4B: call ebx
  00415A4D: lea ecx, var_60
  00415A50: call MSVBVM60.DLL.__vbaFreeVar
  00415A56: mov eax, 00000001h
  00415A5B: add ax, var_30
  00415A5F: jo 415AD5h
  00415A61: mov var_30, eax
  00415A64: jmp 004156E7h
  00415A69: mov var_4, 00000000h
  00415A70: wait 
  00415A71: push 00415AB6h
  00415A76: jmp 415AACh
  00415A78: lea eax, var_38
  00415A7B: lea ecx, var_34
  00415A7E: push eax
  00415A7F: push ecx
  00415A80: push 00000002h
  00415A82: call MSVBVM60.DLL.__vbaFreeStrList
  00415A88: lea edx, var_40
  00415A8B: lea eax, var_3C
  00415A8E: push edx
  00415A8F: push eax
  00415A90: push 00000002h
  00415A92: call MSVBVM60.DLL.__vbaFreeObjList
  00415A98: lea ecx, var_60
  00415A9B: lea edx, var_50
  00415A9E: push ecx
  00415A9F: push edx
  00415AA0: push 00000002h
  00415AA2: call MSVBVM60.DLL.__vbaFreeVarList
  00415AA8: add esp, 00000024h
  00415AAB: ret 
End Sub

Private Sub Command4_Click() '412780
  00412780: push ebp
  00412781: mov ebp, esp
  00412783: sub esp, 0000000Ch
  00412786: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0041278B: mov eax, fs:[00h]
  00412791: push eax
  00412792: mov fs:[00000000h], esp
  00412799: sub esp, 00000014h
  0041279C: push ebx
  0041279D: push esi
  0041279E: push edi
  0041279F: mov var_C, esp
  004127A2: mov var_8, 00401450h
  004127A9: mov esi, arg_8
  004127AC: mov eax, esi
  004127AE: and eax, 00000001h
  004127B1: mov var_4, eax
  004127B4: and esi, FFFFFFFEh
  004127B7: push esi
  004127B8: mov arg_8, esi
  004127BB: mov ecx, [esi]
  004127BD: call [ecx+04h]
  004127C0: mov edx, [esi]
  004127C2: xor edi, edi
  004127C4: push esi
  004127C5: mov var_18, edi
  004127C8: call [edx+00000334h]
  004127CE: push eax
  004127CF: lea eax, var_18
  004127D2: push eax
  004127D3: call [004010BCh] ; Set (object)
  004127D9: mov esi, eax
  004127DB: push edi
  004127DC: push esi
  004127DD: mov ecx, [esi]
  004127DF: call [ecx+0000009Ch]
  004127E5: cmp eax, edi
  004127E7: fclex 
  004127E9: jnl 4127FDh
  004127EB: push 0000009Ch
  004127F0: push 004055A4h
  004127F5: push esi
  004127F6: push eax
  004127F7: call MSVBVM60.DLL.__vbaHresultCheckObj
  004127FD: lea ecx, var_18
  00412800: call MSVBVM60.DLL.__vbaFreeObj
  00412806: mov var_4, edi
  00412809: push 0041281Bh
  0041280E: jmp 41281Ah
  00412810: lea ecx, var_18
  00412813: call MSVBVM60.DLL.__vbaFreeObj
  00412819: ret 
End Sub

Private Sub tbRCdKey_Click() '4161E0
  004161E0: push ebp
  004161E1: mov ebp, esp
  004161E3: sub esp, 0000000Ch
  004161E6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  004161EB: mov eax, fs:[00h]
  004161F1: push eax
  004161F2: mov fs:[00000000h], esp
  004161F9: sub esp, 0000008Ch
  004161FF: push ebx
  00416200: push esi
  00416201: push edi
  00416202: mov var_C, esp
  00416205: mov var_8, 00401538h
  0041620C: mov eax, arg_8
  0041620F: mov ecx, eax
  00416211: and ecx, 00000001h
  00416214: mov var_4, ecx
  00416217: and al, FEh
  00416219: push eax
  0041621A: mov arg_8, eax
  0041621D: mov edx, [eax]
  0041621F: call [edx+04h]
  00416222: mov ecx, 80020004h
  00416227: xor esi, esi
  00416229: mov var_4C, ecx
  0041622C: mov eax, 0000000Ah
  00416231: mov var_3C, ecx
  00416234: mov var_2C, ecx
  00416237: mov var_34, esi
  0041623A: mov var_44, esi
  0041623D: mov var_54, esi
  00416240: mov var_64, esi
  00416243: lea edx, var_64
  00416246: lea ecx, var_24
  00416249: mov var_24, esi
  0041624C: mov var_54, eax
  0041624F: mov var_44, eax
  00416252: mov var_34, eax
  00416255: mov var_5C, 004069DCh ; "Yakin akan hapus cdkey?"
  0041625C: mov var_64, 00000008h
  00416263: call MSVBVM60.DLL.__vbaVarDup
  00416269: lea eax, var_54
  0041626C: lea ecx, var_44
  0041626F: push eax
  00416270: lea edx, var_34
  00416273: push ecx
  00416274: push edx
  00416275: lea eax, var_24
  00416278: push 00000124h
  0041627D: push eax
  0041627E: call [004010C0h] ; MsgBox(arg_1, arg_2, arg_3, arg_4, arg_5)
  00416284: xor ecx, ecx
  00416286: cmp eax, 00000006h
  00416289: setz cl
  0041628C: neg ecx
  0041628E: lea edx, var_54
  00416291: mov edi, ecx
  00416293: lea eax, var_44
  00416296: push edx
  00416297: lea ecx, var_34
  0041629A: push eax
  0041629B: lea edx, var_24
  0041629E: push ecx
  0041629F: push edx
  004162A0: push 00000004h
  004162A2: call MSVBVM60.DLL.__vbaFreeVarList
  004162A8: add esp, 00000014h
  004162AB: cmp di, si
  004162AE: jz 4162C8h
  004162B0: lea eax, var_64
  004162B3: mov var_5C, 004191D8h
  004162BA: push eax
  004162BB: mov var_64, 00004008h
  004162C2: call [00401154h] ; Kill(arg_1)
  004162C8: mov var_4, esi
  004162CB: push 004162EFh
  004162D0: jmp 4162EEh
  004162D2: lea ecx, var_54
  004162D5: lea edx, var_44
  004162D8: push ecx
  004162D9: lea eax, var_34
  004162DC: push edx
  004162DD: lea ecx, var_24
  004162E0: push eax
  004162E1: push ecx
  004162E2: push 00000004h
  004162E4: call MSVBVM60.DLL.__vbaFreeVarList
  004162EA: add esp, 00000014h
  004162ED: ret 
End Sub

Private Sub Form_Load() '412840
  00412840: push ebp
  00412841: mov ebp, esp
  00412843: sub esp, 0000000Ch
  00412846: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0041284B: mov eax, fs:[00h]
  00412851: push eax
  00412852: mov fs:[00000000h], esp
  00412859: sub esp, 000000D4h
  0041285F: push ebx
  00412860: push esi
  00412861: push edi
  00412862: mov var_C, esp
  00412865: mov var_8, 00401460h
  0041286C: mov esi, arg_8
  0041286F: mov eax, esi
  00412871: and eax, 00000001h
  00412874: mov var_4, eax
  00412877: and esi, FFFFFFFEh
  0041287A: push esi
  0041287B: mov arg_8, esi
  0041287E: mov ecx, [esi]
  00412880: call [ecx+04h]
  00412883: mov eax, [419640h]
  00412888: xor ebx, ebx
  0041288A: cmp eax, ebx
  0041288C: mov var_18, ebx
  0041288F: mov var_1C, ebx
  00412892: mov var_20, ebx
  00412895: mov var_24, ebx
  00412898: mov var_28, ebx
  0041289B: mov var_2C, ebx
  0041289E: jnz 4128B0h
  004128A0: push 00419640h
  004128A5: push 004039B0h
  004128AA: call MSVBVM60.DLL.__vbaNew2
  004128B0: mov edi, [00419640h] ; 
  004128B6: lea eax, var_28
  004128B9: push eax
  004128BA: push edi
  004128BB: mov edx, [edi]
  004128BD: call [edx+14h]
  004128C0: cmp eax, ebx
  004128C2: fclex 
  004128C4: jnl 4128D5h
  004128C6: push 00000014h
  004128C8: push 004039A0h
  004128CD: push edi
  004128CE: push eax
  004128CF: call MSVBVM60.DLL.__vbaHresultCheckObj
  004128D5: mov eax, var_28
  004128D8: lea edx, var_18
  004128DB: push edx
  004128DC: push eax
  004128DD: mov ecx, [eax]
  004128DF: mov edi, eax
  004128E1: call [ecx+50h]
  004128E4: cmp eax, ebx
  004128E6: fclex 
  004128E8: jnl 4128F9h
  004128EA: push 00000050h
  004128EC: push 00404788h
  004128F1: push edi
  004128F2: push eax
  004128F3: call MSVBVM60.DLL.__vbaHresultCheckObj
  004128F9: mov eax, var_18
  004128FC: push 00000001h
  004128FE: push eax
  004128FF: push 00404954h ; "jsgen"
  00412904: push ebx
  00412905: call [00401254h] ; InStr
  0041290B: xor ecx, ecx
  0041290D: test eax, eax
  0041290F: setnle cl
  00412912: neg ecx
  00412914: mov edi, ecx
  00412916: lea ecx, var_18
  00412919: call MSVBVM60.DLL.__vbaFreeStr
  0041291F: lea ecx, var_28
  00412922: call MSVBVM60.DLL.__vbaFreeObj
  00412928: cmp di, bx
  0041292B: jz 412973h
  0041292D: mov edx, [esi]
  0041292F: push esi
  00412930: call [edx+00000348h]
  00412936: mov edi, [004010BCh] ; Set (object)
  0041293C: push eax
  0041293D: lea eax, var_28
  00412940: push eax
  00412941: call edi
  00412943: mov ebx, eax
  00412945: push FFFFFFFFh
  00412947: push ebx
  00412948: mov ecx, [ebx]
  0041294A: call [ecx+00000094h]
  00412950: test eax, eax
  00412952: fclex 
  00412954: jnl 412968h
  00412956: push 00000094h
  0041295B: push 004055B4h
  00412960: push ebx
  00412961: push eax
  00412962: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412968: lea ecx, var_28
  0041296B: call MSVBVM60.DLL.__vbaFreeObj
  00412971: jmp 412979h
  00412973: mov edi, [004010BCh] ; Set (object)
  00412979: mov edx, [esi]
  0041297B: push esi
  0041297C: call [edx+00000338h]
  00412982: push eax
  00412983: lea eax, var_28
  00412986: push eax
  00412987: call edi
  00412989: mov ebx, eax
  0041298B: push 004055E8h ; "INFORMASI PENGGUNAAN PROGRAM ENTRI DATA PROFILE DESA DAN KELUARGA"
  00412990: push ebx
  00412991: mov ecx, [ebx]
  00412993: call [ecx+000000A4h]
  00412999: test eax, eax
  0041299B: fclex 
  0041299D: jnl 4129B1h
  0041299F: push 000000A4h
  004129A4: push 004051E8h
  004129A9: push ebx
  004129AA: push eax
  004129AB: call MSVBVM60.DLL.__vbaHresultCheckObj
  004129B1: lea ecx, var_28
  004129B4: call MSVBVM60.DLL.__vbaFreeObj
  004129BA: mov edx, [esi]
  004129BC: push esi
  004129BD: call [edx+00000338h]
  004129C3: push eax
  004129C4: lea eax, var_2C
  004129C7: push eax
  004129C8: call edi
  004129CA: mov ecx, [esi]
  004129CC: push esi
  004129CD: mov var_58, eax
  004129D0: call [ecx+00000338h]
  004129D6: lea edx, var_28
  004129D9: push eax
  004129DA: push edx
  004129DB: call edi
  004129DD: mov ebx, eax
  004129DF: lea ecx, var_18
  004129E2: push ecx
  004129E3: push ebx
  004129E4: mov eax, [ebx]
  004129E6: call [eax+000000A0h]
  004129EC: test eax, eax
  004129EE: fclex 
  004129F0: jnl 412A04h
  004129F2: push 000000A0h
  004129F7: push 004051E8h
  004129FC: push ebx
  004129FD: push eax
  004129FE: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412A04: mov eax, var_18
  00412A07: mov edx, var_58
  00412A0A: push eax
  00412A0B: push 00404900h ; "13#1"
  00412A10: mov ebx, [edx]
  00412A12: call [0040106Ch] ; &
  00412A18: mov edx, eax
  00412A1A: lea ecx, var_1C
  00412A1D: call MSVBVM60.DLL.__vbaStrMove
  00412A23: push eax
  00412A24: push 00405670h ; "1. Jalankan program."
  00412A29: call [0040106Ch] ; &
  00412A2F: mov edx, eax
  00412A31: lea ecx, var_20
  00412A34: call MSVBVM60.DLL.__vbaStrMove
  00412A3A: mov ecx, ebx
  00412A3C: mov ebx, var_58
  00412A3F: push eax
  00412A40: push ebx
  00412A41: call [ecx+000000A4h]
  00412A47: test eax, eax
  00412A49: fclex 
  00412A4B: jnl 412A5Fh
  00412A4D: push 000000A4h
  00412A52: push 004051E8h
  00412A57: push ebx
  00412A58: push eax
  00412A59: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412A5F: lea edx, var_20
  00412A62: lea eax, var_1C
  00412A65: push edx
  00412A66: lea ecx, var_18
  00412A69: push eax
  00412A6A: push ecx
  00412A6B: push 00000003h
  00412A6D: call MSVBVM60.DLL.__vbaFreeStrList
  00412A73: lea edx, var_2C
  00412A76: lea eax, var_28
  00412A79: push edx
  00412A7A: push eax
  00412A7B: push 00000002h
  00412A7D: call MSVBVM60.DLL.__vbaFreeObjList
  00412A83: mov ecx, [esi]
  00412A85: add esp, 0000001Ch
  00412A88: push esi
  00412A89: call [ecx+00000338h]
  00412A8F: lea edx, var_2C
  00412A92: push eax
  00412A93: push edx
  00412A94: call edi
  00412A96: mov var_58, eax
  00412A99: mov eax, [esi]
  00412A9B: push esi
  00412A9C: call [eax+00000338h]
  00412AA2: lea ecx, var_28
  00412AA5: push eax
  00412AA6: push ecx
  00412AA7: call edi
  00412AA9: mov ebx, eax
  00412AAB: lea eax, var_18
  00412AAE: push eax
  00412AAF: push ebx
  00412AB0: mov edx, [ebx]
  00412AB2: call [edx+000000A0h]
  00412AB8: test eax, eax
  00412ABA: fclex 
  00412ABC: jnl 412AD0h
  00412ABE: push 000000A0h
  00412AC3: push 004051E8h
  00412AC8: push ebx
  00412AC9: push eax
  00412ACA: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412AD0: mov edx, var_18
  00412AD3: mov ecx, var_58
  00412AD6: push edx
  00412AD7: push 00404900h ; "13#1"
  00412ADC: mov ebx, [ecx]
  00412ADE: call [0040106Ch] ; &
  00412AE4: mov edx, eax
  00412AE6: lea ecx, var_1C
  00412AE9: call MSVBVM60.DLL.__vbaStrMove
  00412AEF: push eax
  00412AF0: push 004056A0h ; "2. Klik tombol Login."
  00412AF5: call [0040106Ch] ; &
  00412AFB: mov edx, eax
  00412AFD: lea ecx, var_20
  00412B00: call MSVBVM60.DLL.__vbaStrMove
  00412B06: mov var_70, ebx
  00412B09: mov ebx, var_58
  00412B0C: push eax
  00412B0D: mov eax, var_70
  00412B10: push ebx
  00412B11: call [eax+000000A4h]
  00412B17: test eax, eax
  00412B19: fclex 
  00412B1B: jnl 412B2Fh
  00412B1D: push 000000A4h
  00412B22: push 004051E8h
  00412B27: push ebx
  00412B28: push eax
  00412B29: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412B2F: lea ecx, var_20
  00412B32: lea edx, var_1C
  00412B35: push ecx
  00412B36: lea eax, var_18
  00412B39: push edx
  00412B3A: push eax
  00412B3B: push 00000003h
  00412B3D: call MSVBVM60.DLL.__vbaFreeStrList
  00412B43: lea ecx, var_2C
  00412B46: lea edx, var_28
  00412B49: push ecx
  00412B4A: push edx
  00412B4B: push 00000002h
  00412B4D: call MSVBVM60.DLL.__vbaFreeObjList
  00412B53: mov eax, [esi]
  00412B55: add esp, 0000001Ch
  00412B58: push esi
  00412B59: call [eax+00000338h]
  00412B5F: lea ecx, var_2C
  00412B62: push eax
  00412B63: push ecx
  00412B64: call edi
  00412B66: mov edx, [esi]
  00412B68: push esi
  00412B69: mov var_58, eax
  00412B6C: call [edx+00000338h]
  00412B72: push eax
  00412B73: lea eax, var_28
  00412B76: push eax
  00412B77: call edi
  00412B79: mov ebx, eax
  00412B7B: lea edx, var_18
  00412B7E: push edx
  00412B7F: push ebx
  00412B80: mov ecx, [ebx]
  00412B82: call [ecx+000000A0h]
  00412B88: test eax, eax
  00412B8A: fclex 
  00412B8C: jnl 412BA0h
  00412B8E: push 000000A0h
  00412B93: push 004051E8h
  00412B98: push ebx
  00412B99: push eax
  00412B9A: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412BA0: mov ecx, var_18
  00412BA3: mov eax, var_58
  00412BA6: push ecx
  00412BA7: push 00404900h ; "13#1"
  00412BAC: mov ebx, [eax]
  00412BAE: call [0040106Ch] ; &
  00412BB4: mov edx, eax
  00412BB6: lea ecx, var_1C
  00412BB9: call MSVBVM60.DLL.__vbaStrMove
  00412BBF: push eax
  00412BC0: push 004056D0h ; "3. Isikan user dan password yang sudah diberikan di web prodeskel.binapemdes.kemendagri.go.id."
  00412BC5: call [0040106Ch] ; &
  00412BCB: mov edx, eax
  00412BCD: lea ecx, var_20
  00412BD0: call MSVBVM60.DLL.__vbaStrMove
  00412BD6: mov edx, ebx
  00412BD8: mov ebx, var_58
  00412BDB: push eax
  00412BDC: push ebx
  00412BDD: call [edx+000000A4h]
  00412BE3: test eax, eax
  00412BE5: fclex 
  00412BE7: jnl 412BFBh
  00412BE9: push 000000A4h
  00412BEE: push 004051E8h
  00412BF3: push ebx
  00412BF4: push eax
  00412BF5: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412BFB: lea eax, var_20
  00412BFE: lea ecx, var_1C
  00412C01: push eax
  00412C02: lea edx, var_18
  00412C05: push ecx
  00412C06: push edx
  00412C07: push 00000003h
  00412C09: call MSVBVM60.DLL.__vbaFreeStrList
  00412C0F: lea eax, var_2C
  00412C12: lea ecx, var_28
  00412C15: push eax
  00412C16: push ecx
  00412C17: push 00000002h
  00412C19: call MSVBVM60.DLL.__vbaFreeObjList
  00412C1F: mov edx, [esi]
  00412C21: add esp, 0000001Ch
  00412C24: push esi
  00412C25: call [edx+00000338h]
  00412C2B: push eax
  00412C2C: lea eax, var_2C
  00412C2F: push eax
  00412C30: call edi
  00412C32: mov ecx, [esi]
  00412C34: push esi
  00412C35: mov var_58, eax
  00412C38: call [ecx+00000338h]
  00412C3E: lea edx, var_28
  00412C41: push eax
  00412C42: push edx
  00412C43: call edi
  00412C45: mov ebx, eax
  00412C47: lea ecx, var_18
  00412C4A: push ecx
  00412C4B: push ebx
  00412C4C: mov eax, [ebx]
  00412C4E: call [eax+000000A0h]
  00412C54: test eax, eax
  00412C56: fclex 
  00412C58: jnl 412C6Ch
  00412C5A: push 000000A0h
  00412C5F: push 004051E8h
  00412C64: push ebx
  00412C65: push eax
  00412C66: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412C6C: mov eax, var_18
  00412C6F: mov edx, var_58
  00412C72: push eax
  00412C73: push 00404900h ; "13#1"
  00412C78: mov ebx, [edx]
  00412C7A: call [0040106Ch] ; &
  00412C80: mov edx, eax
  00412C82: lea ecx, var_1C
  00412C85: call MSVBVM60.DLL.__vbaStrMove
  00412C8B: push eax
  00412C8C: push 004057D0h ; "4. tekan tombol kirim yang ada diweb"
  00412C91: call [0040106Ch] ; &
  00412C97: mov edx, eax
  00412C99: lea ecx, var_20
  00412C9C: call MSVBVM60.DLL.__vbaStrMove
  00412CA2: mov ecx, ebx
  00412CA4: mov ebx, var_58
  00412CA7: push eax
  00412CA8: push ebx
  00412CA9: call [ecx+000000A4h]
  00412CAF: test eax, eax
  00412CB1: fclex 
  00412CB3: jnl 412CC7h
  00412CB5: push 000000A4h
  00412CBA: push 004051E8h
  00412CBF: push ebx
  00412CC0: push eax
  00412CC1: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412CC7: lea edx, var_20
  00412CCA: lea eax, var_1C
  00412CCD: push edx
  00412CCE: lea ecx, var_18
  00412CD1: push eax
  00412CD2: push ecx
  00412CD3: push 00000003h
  00412CD5: call MSVBVM60.DLL.__vbaFreeStrList
  00412CDB: lea edx, var_2C
  00412CDE: lea eax, var_28
  00412CE1: push edx
  00412CE2: push eax
  00412CE3: push 00000002h
  00412CE5: call MSVBVM60.DLL.__vbaFreeObjList
  00412CEB: mov ecx, [esi]
  00412CED: add esp, 0000001Ch
  00412CF0: push esi
  00412CF1: call [ecx+00000338h]
  00412CF7: lea edx, var_2C
  00412CFA: push eax
  00412CFB: push edx
  00412CFC: call edi
  00412CFE: mov var_58, eax
  00412D01: mov eax, [esi]
  00412D03: push esi
  00412D04: call [eax+00000338h]
  00412D0A: lea ecx, var_28
  00412D0D: push eax
  00412D0E: push ecx
  00412D0F: call edi
  00412D11: mov ebx, eax
  00412D13: lea eax, var_18
  00412D16: push eax
  00412D17: push ebx
  00412D18: mov edx, [ebx]
  00412D1A: call [edx+000000A0h]
  00412D20: test eax, eax
  00412D22: fclex 
  00412D24: jnl 412D38h
  00412D26: push 000000A0h
  00412D2B: push 004051E8h
  00412D30: push ebx
  00412D31: push eax
  00412D32: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412D38: mov edx, var_18
  00412D3B: mov ecx, var_58
  00412D3E: push edx
  00412D3F: push 00404900h ; "13#1"
  00412D44: mov ebx, [ecx]
  00412D46: call [0040106Ch] ; &
  00412D4C: mov edx, eax
  00412D4E: lea ecx, var_1C
  00412D51: call MSVBVM60.DLL.__vbaStrMove
  00412D57: push eax
  00412D58: push 00404900h ; "13#1"
  00412D5D: call [0040106Ch] ; &
  00412D63: mov edx, eax
  00412D65: lea ecx, var_20
  00412D68: call MSVBVM60.DLL.__vbaStrMove
  00412D6E: push eax
  00412D6F: push 00405820h ; "MENGKOPI DATA DARI MICROSOFT EXCEL"
  00412D74: call [0040106Ch] ; &
  00412D7A: mov edx, eax
  00412D7C: lea ecx, var_24
  00412D7F: call MSVBVM60.DLL.__vbaStrMove
  00412D85: mov var_7C, ebx
  00412D88: mov ebx, var_58
  00412D8B: push eax
  00412D8C: mov eax, var_7C
  00412D8F: push ebx
  00412D90: call [eax+000000A4h]
  00412D96: test eax, eax
  00412D98: fclex 
  00412D9A: jnl 412DAEh
  00412D9C: push 000000A4h
  00412DA1: push 004051E8h
  00412DA6: push ebx
  00412DA7: push eax
  00412DA8: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412DAE: lea ecx, var_24
  00412DB1: lea edx, var_20
  00412DB4: push ecx
  00412DB5: lea eax, var_1C
  00412DB8: push edx
  00412DB9: lea ecx, var_18
  00412DBC: push eax
  00412DBD: push ecx
  00412DBE: push 00000004h
  00412DC0: call MSVBVM60.DLL.__vbaFreeStrList
  00412DC6: lea edx, var_2C
  00412DC9: lea eax, var_28
  00412DCC: push edx
  00412DCD: push eax
  00412DCE: push 00000002h
  00412DD0: call MSVBVM60.DLL.__vbaFreeObjList
  00412DD6: mov ecx, [esi]
  00412DD8: add esp, 00000020h
  00412DDB: push esi
  00412DDC: call [ecx+00000338h]
  00412DE2: lea edx, var_2C
  00412DE5: push eax
  00412DE6: push edx
  00412DE7: call edi
  00412DE9: mov var_58, eax
  00412DEC: mov eax, [esi]
  00412DEE: push esi
  00412DEF: call [eax+00000338h]
  00412DF5: lea ecx, var_28
  00412DF8: push eax
  00412DF9: push ecx
  00412DFA: call edi
  00412DFC: mov ebx, eax
  00412DFE: lea eax, var_18
  00412E01: push eax
  00412E02: push ebx
  00412E03: mov edx, [ebx]
  00412E05: call [edx+000000A0h]
  00412E0B: test eax, eax
  00412E0D: fclex 
  00412E0F: jnl 412E23h
  00412E11: push 000000A0h
  00412E16: push 004051E8h
  00412E1B: push ebx
  00412E1C: push eax
  00412E1D: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412E23: mov edx, var_18
  00412E26: mov ecx, var_58
  00412E29: push edx
  00412E2A: push 00404900h ; "13#1"
  00412E2F: mov ebx, [ecx]
  00412E31: call [0040106Ch] ; &
  00412E37: mov edx, eax
  00412E39: lea ecx, var_1C
  00412E3C: call MSVBVM60.DLL.__vbaStrMove
  00412E42: push eax
  00412E43: push 0040586Ch ; "1. Buka dahulu file data excel yang akan dimasukkan ke dalam program"
  00412E48: call [0040106Ch] ; &
  00412E4E: mov edx, eax
  00412E50: lea ecx, var_20
  00412E53: call MSVBVM60.DLL.__vbaStrMove
  00412E59: mov var_80, ebx
  00412E5C: mov ebx, var_58
  00412E5F: push eax
  00412E60: mov eax, var_80
  00412E63: push ebx
  00412E64: call [eax+000000A4h]
  00412E6A: test eax, eax
  00412E6C: fclex 
  00412E6E: jnl 412E82h
  00412E70: push 000000A4h
  00412E75: push 004051E8h
  00412E7A: push ebx
  00412E7B: push eax
  00412E7C: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412E82: lea ecx, var_20
  00412E85: lea edx, var_1C
  00412E88: push ecx
  00412E89: lea eax, var_18
  00412E8C: push edx
  00412E8D: push eax
  00412E8E: push 00000003h
  00412E90: call MSVBVM60.DLL.__vbaFreeStrList
  00412E96: lea ecx, var_2C
  00412E99: lea edx, var_28
  00412E9C: push ecx
  00412E9D: push edx
  00412E9E: push 00000002h
  00412EA0: call MSVBVM60.DLL.__vbaFreeObjList
  00412EA6: mov eax, [esi]
  00412EA8: add esp, 0000001Ch
  00412EAB: push esi
  00412EAC: call [eax+00000338h]
  00412EB2: lea ecx, var_2C
  00412EB5: push eax
  00412EB6: push ecx
  00412EB7: call edi
  00412EB9: mov edx, [esi]
  00412EBB: push esi
  00412EBC: mov var_58, eax
  00412EBF: call [edx+00000338h]
  00412EC5: push eax
  00412EC6: lea eax, var_28
  00412EC9: push eax
  00412ECA: call edi
  00412ECC: mov ebx, eax
  00412ECE: lea edx, var_18
  00412ED1: push edx
  00412ED2: push ebx
  00412ED3: mov ecx, [ebx]
  00412ED5: call [ecx+000000A0h]
  00412EDB: test eax, eax
  00412EDD: fclex 
  00412EDF: jnl 412EF3h
  00412EE1: push 000000A0h
  00412EE6: push 004051E8h
  00412EEB: push ebx
  00412EEC: push eax
  00412EED: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412EF3: mov ecx, var_18
  00412EF6: mov eax, var_58
  00412EF9: push ecx
  00412EFA: push 00404900h ; "13#1"
  00412EFF: mov ebx, [eax]
  00412F01: call [0040106Ch] ; &
  00412F07: mov edx, eax
  00412F09: lea ecx, var_1C
  00412F0C: call MSVBVM60.DLL.__vbaStrMove
  00412F12: push eax
  00412F13: push 004058FCh ; "2. Blok data yang akan dimasukkan mulai dari kolom NOMOR KK sampai dengan Alamat"
  00412F18: call [0040106Ch] ; &
  00412F1E: mov edx, eax
  00412F20: lea ecx, var_20
  00412F23: call MSVBVM60.DLL.__vbaStrMove
  00412F29: mov edx, ebx
  00412F2B: mov ebx, var_58
  00412F2E: push eax
  00412F2F: push ebx
  00412F30: call [edx+000000A4h]
  00412F36: test eax, eax
  00412F38: fclex 
  00412F3A: jnl 412F4Eh
  00412F3C: push 000000A4h
  00412F41: push 004051E8h
  00412F46: push ebx
  00412F47: push eax
  00412F48: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412F4E: lea eax, var_20
  00412F51: lea ecx, var_1C
  00412F54: push eax
  00412F55: lea edx, var_18
  00412F58: push ecx
  00412F59: push edx
  00412F5A: push 00000003h
  00412F5C: call MSVBVM60.DLL.__vbaFreeStrList
  00412F62: lea eax, var_2C
  00412F65: lea ecx, var_28
  00412F68: push eax
  00412F69: push ecx
  00412F6A: push 00000002h
  00412F6C: call MSVBVM60.DLL.__vbaFreeObjList
  00412F72: mov edx, [esi]
  00412F74: add esp, 0000001Ch
  00412F77: push esi
  00412F78: call [edx+00000338h]
  00412F7E: push eax
  00412F7F: lea eax, var_2C
  00412F82: push eax
  00412F83: call edi
  00412F85: mov ecx, [esi]
  00412F87: push esi
  00412F88: mov var_58, eax
  00412F8B: call [ecx+00000338h]
  00412F91: lea edx, var_28
  00412F94: push eax
  00412F95: push edx
  00412F96: call edi
  00412F98: mov ebx, eax
  00412F9A: lea ecx, var_18
  00412F9D: push ecx
  00412F9E: push ebx
  00412F9F: mov eax, [ebx]
  00412FA1: call [eax+000000A0h]
  00412FA7: test eax, eax
  00412FA9: fclex 
  00412FAB: jnl 412FBFh
  00412FAD: push 000000A0h
  00412FB2: push 004051E8h
  00412FB7: push ebx
  00412FB8: push eax
  00412FB9: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412FBF: mov eax, var_18
  00412FC2: mov edx, var_58
  00412FC5: push eax
  00412FC6: push 00404900h ; "13#1"
  00412FCB: mov ebx, [edx]
  00412FCD: call [0040106Ch] ; &
  00412FD3: mov edx, eax
  00412FD5: lea ecx, var_1C
  00412FD8: call MSVBVM60.DLL.__vbaStrMove
  00412FDE: push eax
  00412FDF: push 004059B4h ; "3. Pilih menu Edit>>copy atau dengan short cut ctrl+c"
  00412FE4: call [0040106Ch] ; &
  00412FEA: mov edx, eax
  00412FEC: lea ecx, var_20
  00412FEF: call MSVBVM60.DLL.__vbaStrMove
  00412FF5: mov ecx, ebx
  00412FF7: mov ebx, var_58
  00412FFA: push eax
  00412FFB: push ebx
  00412FFC: call [ecx+000000A4h]
  00413002: test eax, eax
  00413004: fclex 
  00413006: jnl 41301Ah
  00413008: push 000000A4h
  0041300D: push 004051E8h
  00413012: push ebx
  00413013: push eax
  00413014: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041301A: lea edx, var_20
  0041301D: lea eax, var_1C
  00413020: push edx
  00413021: lea ecx, var_18
  00413024: push eax
  00413025: push ecx
  00413026: push 00000003h
  00413028: call MSVBVM60.DLL.__vbaFreeStrList
  0041302E: lea edx, var_2C
  00413031: lea eax, var_28
  00413034: push edx
  00413035: push eax
  00413036: push 00000002h
  00413038: call MSVBVM60.DLL.__vbaFreeObjList
  0041303E: mov ecx, [esi]
  00413040: add esp, 0000001Ch
  00413043: push esi
  00413044: call [ecx+00000338h]
  0041304A: lea edx, var_2C
  0041304D: push eax
  0041304E: push edx
  0041304F: call edi
  00413051: mov var_58, eax
  00413054: mov eax, [esi]
  00413056: push esi
  00413057: call [eax+00000338h]
  0041305D: lea ecx, var_28
  00413060: push eax
  00413061: push ecx
  00413062: call edi
  00413064: mov ebx, eax
  00413066: lea eax, var_18
  00413069: push eax
  0041306A: push ebx
  0041306B: mov edx, [ebx]
  0041306D: call [edx+000000A0h]
  00413073: test eax, eax
  00413075: fclex 
  00413077: jnl 41308Bh
  00413079: push 000000A0h
  0041307E: push 004051E8h
  00413083: push ebx
  00413084: push eax
  00413085: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041308B: mov edx, var_18
  0041308E: mov ecx, var_58
  00413091: push edx
  00413092: push 00404900h ; "13#1"
  00413097: mov ebx, [ecx]
  00413099: call [0040106Ch] ; &
  0041309F: mov edx, eax
  004130A1: lea ecx, var_1C
  004130A4: call MSVBVM60.DLL.__vbaStrMove
  004130AA: push eax
  004130AB: push 00405A24h ; "4. dari program yang sudah dijalankan, tekan tombol Paste data from Clipboard, sehingga data akan masuk di grid"
  004130B0: call [0040106Ch] ; &
  004130B6: mov edx, eax
  004130B8: lea ecx, var_20
  004130BB: call MSVBVM60.DLL.__vbaStrMove
  004130C1: mov var_8C, ebx
  004130C7: mov ebx, var_58
  004130CA: push eax
  004130CB: mov eax, var_8C
  004130D1: push ebx
  004130D2: call [eax+000000A4h]
  004130D8: test eax, eax
  004130DA: fclex 
  004130DC: jnl 4130F0h
  004130DE: push 000000A4h
  004130E3: push 004051E8h
  004130E8: push ebx
  004130E9: push eax
  004130EA: call MSVBVM60.DLL.__vbaHresultCheckObj
  004130F0: lea ecx, var_20
  004130F3: lea edx, var_1C
  004130F6: push ecx
  004130F7: lea eax, var_18
  004130FA: push edx
  004130FB: push eax
  004130FC: push 00000003h
  004130FE: call MSVBVM60.DLL.__vbaFreeStrList
  00413104: lea ecx, var_2C
  00413107: lea edx, var_28
  0041310A: push ecx
  0041310B: push edx
  0041310C: push 00000002h
  0041310E: call MSVBVM60.DLL.__vbaFreeObjList
  00413114: mov eax, [esi]
  00413116: add esp, 0000001Ch
  00413119: push esi
  0041311A: call [eax+00000338h]
  00413120: lea ecx, var_2C
  00413123: push eax
  00413124: push ecx
  00413125: call edi
  00413127: mov edx, [esi]
  00413129: push esi
  0041312A: mov var_58, eax
  0041312D: call [edx+00000338h]
  00413133: push eax
  00413134: lea eax, var_28
  00413137: push eax
  00413138: call edi
  0041313A: mov ebx, eax
  0041313C: lea edx, var_18
  0041313F: push edx
  00413140: push ebx
  00413141: mov ecx, [ebx]
  00413143: call [ecx+000000A0h]
  00413149: test eax, eax
  0041314B: fclex 
  0041314D: jnl 413161h
  0041314F: push 000000A0h
  00413154: push 004051E8h
  00413159: push ebx
  0041315A: push eax
  0041315B: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413161: mov ecx, var_18
  00413164: mov eax, var_58
  00413167: push ecx
  00413168: push 00404900h ; "13#1"
  0041316D: mov ebx, [eax]
  0041316F: call [0040106Ch] ; &
  00413175: mov edx, eax
  00413177: lea ecx, var_1C
  0041317A: call MSVBVM60.DLL.__vbaStrMove
  00413180: push eax
  00413181: push 00405BA4h ; "5. klik baris mana yang akan dimasukkan ke dalam form web, klik tombol masukkan data"
  00413186: call [0040106Ch] ; &
  0041318C: mov edx, eax
  0041318E: lea ecx, var_20
  00413191: call MSVBVM60.DLL.__vbaStrMove
  00413197: mov edx, ebx
  00413199: mov ebx, var_58
  0041319C: push eax
  0041319D: push ebx
  0041319E: call [edx+000000A4h]
  004131A4: test eax, eax
  004131A6: fclex 
  004131A8: jnl 4131BCh
  004131AA: push 000000A4h
  004131AF: push 004051E8h
  004131B4: push ebx
  004131B5: push eax
  004131B6: call MSVBVM60.DLL.__vbaHresultCheckObj
  004131BC: lea eax, var_20
  004131BF: lea ecx, var_1C
  004131C2: push eax
  004131C3: lea edx, var_18
  004131C6: push ecx
  004131C7: push edx
  004131C8: push 00000003h
  004131CA: call MSVBVM60.DLL.__vbaFreeStrList
  004131D0: lea eax, var_2C
  004131D3: lea ecx, var_28
  004131D6: push eax
  004131D7: push ecx
  004131D8: push 00000002h
  004131DA: call MSVBVM60.DLL.__vbaFreeObjList
  004131E0: mov edx, [esi]
  004131E2: add esp, 0000001Ch
  004131E5: push esi
  004131E6: call [edx+00000338h]
  004131EC: push eax
  004131ED: lea eax, var_2C
  004131F0: push eax
  004131F1: call edi
  004131F3: mov ecx, [esi]
  004131F5: push esi
  004131F6: mov var_58, eax
  004131F9: call [ecx+00000338h]
  004131FF: lea edx, var_28
  00413202: push eax
  00413203: push edx
  00413204: call edi
  00413206: mov ebx, eax
  00413208: lea ecx, var_18
  0041320B: push ecx
  0041320C: push ebx
  0041320D: mov eax, [ebx]
  0041320F: call [eax+000000A0h]
  00413215: test eax, eax
  00413217: fclex 
  00413219: jnl 41322Dh
  0041321B: push 000000A0h
  00413220: push 004051E8h
  00413225: push ebx
  00413226: push eax
  00413227: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041322D: mov eax, var_18
  00413230: mov edx, var_58
  00413233: push eax
  00413234: push 00404900h ; "13#1"
  00413239: mov ebx, [edx]
  0041323B: call [0040106Ch] ; &
  00413241: mov edx, eax
  00413243: lea ecx, var_1C
  00413246: call MSVBVM60.DLL.__vbaStrMove
  0041324C: push eax
  0041324D: push 00404900h ; "13#1"
  00413252: call [0040106Ch] ; &
  00413258: mov edx, eax
  0041325A: lea ecx, var_20
  0041325D: call MSVBVM60.DLL.__vbaStrMove
  00413263: push eax
  00413264: push 00405C54h ; "MEMASUKKAN DATA KK"
  00413269: call [0040106Ch] ; &
  0041326F: mov edx, eax
  00413271: lea ecx, var_24
  00413274: call MSVBVM60.DLL.__vbaStrMove
  0041327A: mov ecx, ebx
  0041327C: mov ebx, var_58
  0041327F: push eax
  00413280: push ebx
  00413281: call [ecx+000000A4h]
  00413287: test eax, eax
  00413289: fclex 
  0041328B: jnl 41329Fh
  0041328D: push 000000A4h
  00413292: push 004051E8h
  00413297: push ebx
  00413298: push eax
  00413299: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041329F: lea edx, var_24
  004132A2: lea eax, var_20
  004132A5: push edx
  004132A6: lea ecx, var_1C
  004132A9: push eax
  004132AA: lea edx, var_18
  004132AD: push ecx
  004132AE: push edx
  004132AF: push 00000004h
  004132B1: call MSVBVM60.DLL.__vbaFreeStrList
  004132B7: lea eax, var_2C
  004132BA: lea ecx, var_28
  004132BD: push eax
  004132BE: push ecx
  004132BF: push 00000002h
  004132C1: call MSVBVM60.DLL.__vbaFreeObjList
  004132C7: mov edx, [esi]
  004132C9: add esp, 00000020h
  004132CC: push esi
  004132CD: call [edx+00000338h]
  004132D3: push eax
  004132D4: lea eax, var_2C
  004132D7: push eax
  004132D8: call edi
  004132DA: mov ecx, [esi]
  004132DC: push esi
  004132DD: mov var_58, eax
  004132E0: call [ecx+00000338h]
  004132E6: lea edx, var_28
  004132E9: push eax
  004132EA: push edx
  004132EB: call edi
  004132ED: mov ebx, eax
  004132EF: lea ecx, var_18
  004132F2: push ecx
  004132F3: push ebx
  004132F4: mov eax, [ebx]
  004132F6: call [eax+000000A0h]
  004132FC: test eax, eax
  004132FE: fclex 
  00413300: jnl 413314h
  00413302: push 000000A0h
  00413307: push 004051E8h
  0041330C: push ebx
  0041330D: push eax
  0041330E: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413314: mov eax, var_18
  00413317: mov edx, var_58
  0041331A: push eax
  0041331B: push 00404900h ; "13#1"
  00413320: mov ebx, [edx]
  00413322: call [0040106Ch] ; &
  00413328: mov edx, eax
  0041332A: lea ecx, var_1C
  0041332D: call MSVBVM60.DLL.__vbaStrMove
  00413333: push eax
  00413334: push 00405C80h ; "1. klik tombol form KK yang ada diprogram"
  00413339: call [0040106Ch] ; &
  0041333F: mov edx, eax
  00413341: lea ecx, var_20
  00413344: call MSVBVM60.DLL.__vbaStrMove
  0041334A: mov ecx, ebx
  0041334C: mov ebx, var_58
  0041334F: push eax
  00413350: push ebx
  00413351: call [ecx+000000A4h]
  00413357: test eax, eax
  00413359: fclex 
  0041335B: jnl 41336Fh
  0041335D: push 000000A4h
  00413362: push 004051E8h
  00413367: push ebx
  00413368: push eax
  00413369: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041336F: lea edx, var_20
  00413372: lea eax, var_1C
  00413375: push edx
  00413376: lea ecx, var_18
  00413379: push eax
  0041337A: push ecx
  0041337B: push 00000003h
  0041337D: call MSVBVM60.DLL.__vbaFreeStrList
  00413383: lea edx, var_2C
  00413386: lea eax, var_28
  00413389: push edx
  0041338A: push eax
  0041338B: push 00000002h
  0041338D: call MSVBVM60.DLL.__vbaFreeObjList
  00413393: mov ecx, [esi]
  00413395: add esp, 0000001Ch
  00413398: push esi
  00413399: call [ecx+00000338h]
  0041339F: lea edx, var_2C
  004133A2: push eax
  004133A3: push edx
  004133A4: call edi
  004133A6: mov var_58, eax
  004133A9: mov eax, [esi]
  004133AB: push esi
  004133AC: call [eax+00000338h]
  004133B2: lea ecx, var_28
  004133B5: push eax
  004133B6: push ecx
  004133B7: call edi
  004133B9: mov ebx, eax
  004133BB: lea eax, var_18
  004133BE: push eax
  004133BF: push ebx
  004133C0: mov edx, [ebx]
  004133C2: call [edx+000000A0h]
  004133C8: test eax, eax
  004133CA: fclex 
  004133CC: jnl 4133E0h
  004133CE: push 000000A0h
  004133D3: push 004051E8h
  004133D8: push ebx
  004133D9: push eax
  004133DA: call MSVBVM60.DLL.__vbaHresultCheckObj
  004133E0: mov edx, var_18
  004133E3: mov ecx, var_58
  004133E6: push edx
  004133E7: push 00404900h ; "13#1"
  004133EC: mov ebx, [ecx]
  004133EE: call [0040106Ch] ; &
  004133F4: mov edx, eax
  004133F6: lea ecx, var_1C
  004133F9: call MSVBVM60.DLL.__vbaStrMove
  004133FF: push eax
  00413400: push 00405CD8h ; "2. klik baris data yang ada digrid"
  00413405: call [0040106Ch] ; &
  0041340B: mov edx, eax
  0041340D: lea ecx, var_20
  00413410: call MSVBVM60.DLL.__vbaStrMove
  00413416: mov var_9C, ebx
  0041341C: mov ebx, var_58
  0041341F: push eax
  00413420: mov eax, var_9C
  00413426: push ebx
  00413427: call [eax+000000A4h]
  0041342D: test eax, eax
  0041342F: fclex 
  00413431: jnl 413445h
  00413433: push 000000A4h
  00413438: push 004051E8h
  0041343D: push ebx
  0041343E: push eax
  0041343F: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413445: lea ecx, var_20
  00413448: lea edx, var_1C
  0041344B: push ecx
  0041344C: lea eax, var_18
  0041344F: push edx
  00413450: push eax
  00413451: push 00000003h
  00413453: call MSVBVM60.DLL.__vbaFreeStrList
  00413459: lea ecx, var_2C
  0041345C: lea edx, var_28
  0041345F: push ecx
  00413460: push edx
  00413461: push 00000002h
  00413463: call MSVBVM60.DLL.__vbaFreeObjList
  00413469: mov eax, [esi]
  0041346B: add esp, 0000001Ch
  0041346E: push esi
  0041346F: call [eax+00000338h]
  00413475: lea ecx, var_2C
  00413478: push eax
  00413479: push ecx
  0041347A: call edi
  0041347C: mov edx, [esi]
  0041347E: push esi
  0041347F: mov var_58, eax
  00413482: call [edx+00000338h]
  00413488: push eax
  00413489: lea eax, var_28
  0041348C: push eax
  0041348D: call edi
  0041348F: mov ebx, eax
  00413491: lea edx, var_18
  00413494: push edx
  00413495: push ebx
  00413496: mov ecx, [ebx]
  00413498: call [ecx+000000A0h]
  0041349E: test eax, eax
  004134A0: fclex 
  004134A2: jnl 4134B6h
  004134A4: push 000000A0h
  004134A9: push 004051E8h
  004134AE: push ebx
  004134AF: push eax
  004134B0: call MSVBVM60.DLL.__vbaHresultCheckObj
  004134B6: mov ecx, var_18
  004134B9: mov eax, var_58
  004134BC: push ecx
  004134BD: push 00404900h ; "13#1"
  004134C2: mov ebx, [eax]
  004134C4: call [0040106Ch] ; &
  004134CA: mov edx, eax
  004134CC: lea ecx, var_1C
  004134CF: call MSVBVM60.DLL.__vbaStrMove
  004134D5: push eax
  004134D6: push 00405D24h ; "3. klik tombol masukkan data KK"
  004134DB: call [0040106Ch] ; &
  004134E1: mov edx, eax
  004134E3: lea ecx, var_20
  004134E6: call MSVBVM60.DLL.__vbaStrMove
  004134EC: mov edx, ebx
  004134EE: mov ebx, var_58
  004134F1: push eax
  004134F2: push ebx
  004134F3: call [edx+000000A4h]
  004134F9: test eax, eax
  004134FB: fclex 
  004134FD: jnl 413511h
  004134FF: push 000000A4h
  00413504: push 004051E8h
  00413509: push ebx
  0041350A: push eax
  0041350B: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413511: lea eax, var_20
  00413514: lea ecx, var_1C
  00413517: push eax
  00413518: lea edx, var_18
  0041351B: push ecx
  0041351C: push edx
  0041351D: push 00000003h
  0041351F: call MSVBVM60.DLL.__vbaFreeStrList
  00413525: lea eax, var_2C
  00413528: lea ecx, var_28
  0041352B: push eax
  0041352C: push ecx
  0041352D: push 00000002h
  0041352F: call MSVBVM60.DLL.__vbaFreeObjList
  00413535: mov edx, [esi]
  00413537: add esp, 0000001Ch
  0041353A: push esi
  0041353B: call [edx+00000338h]
  00413541: push eax
  00413542: lea eax, var_2C
  00413545: push eax
  00413546: call edi
  00413548: mov ecx, [esi]
  0041354A: push esi
  0041354B: mov var_58, eax
  0041354E: call [ecx+00000338h]
  00413554: lea edx, var_28
  00413557: push eax
  00413558: push edx
  00413559: call edi
  0041355B: mov ebx, eax
  0041355D: lea ecx, var_18
  00413560: push ecx
  00413561: push ebx
  00413562: mov eax, [ebx]
  00413564: call [eax+000000A0h]
  0041356A: test eax, eax
  0041356C: fclex 
  0041356E: jnl 413582h
  00413570: push 000000A0h
  00413575: push 004051E8h
  0041357A: push ebx
  0041357B: push eax
  0041357C: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413582: mov eax, var_18
  00413585: mov edx, var_58
  00413588: push eax
  00413589: push 00404900h ; "13#1"
  0041358E: mov ebx, [edx]
  00413590: call [0040106Ch] ; &
  00413596: mov edx, eax
  00413598: lea ecx, var_1C
  0041359B: call MSVBVM60.DLL.__vbaStrMove
  004135A1: push eax
  004135A2: push 00405B08h ; "4. jika ada data yang mau diperbaiki, ubah data sesuai keinginan..."
  004135A7: call [0040106Ch] ; &
  004135AD: mov edx, eax
  004135AF: lea ecx, var_20
  004135B2: call MSVBVM60.DLL.__vbaStrMove
  004135B8: mov ecx, ebx
  004135BA: mov ebx, var_58
  004135BD: push eax
  004135BE: push ebx
  004135BF: call [ecx+000000A4h]
  004135C5: test eax, eax
  004135C7: fclex 
  004135C9: jnl 4135DDh
  004135CB: push 000000A4h
  004135D0: push 004051E8h
  004135D5: push ebx
  004135D6: push eax
  004135D7: call MSVBVM60.DLL.__vbaHresultCheckObj
  004135DD: lea edx, var_20
  004135E0: lea eax, var_1C
  004135E3: push edx
  004135E4: lea ecx, var_18
  004135E7: push eax
  004135E8: push ecx
  004135E9: push 00000003h
  004135EB: call MSVBVM60.DLL.__vbaFreeStrList
  004135F1: lea edx, var_2C
  004135F4: lea eax, var_28
  004135F7: push edx
  004135F8: push eax
  004135F9: push 00000002h
  004135FB: call MSVBVM60.DLL.__vbaFreeObjList
  00413601: mov ecx, [esi]
  00413603: add esp, 0000001Ch
  00413606: push esi
  00413607: call [ecx+00000338h]
  0041360D: lea edx, var_2C
  00413610: push eax
  00413611: push edx
  00413612: call edi
  00413614: mov var_58, eax
  00413617: mov eax, [esi]
  00413619: push esi
  0041361A: call [eax+00000338h]
  00413620: lea ecx, var_28
  00413623: push eax
  00413624: push ecx
  00413625: call edi
  00413627: mov ebx, eax
  00413629: lea eax, var_18
  0041362C: push eax
  0041362D: push ebx
  0041362E: mov edx, [ebx]
  00413630: call [edx+000000A0h]
  00413636: test eax, eax
  00413638: fclex 
  0041363A: jnl 41364Eh
  0041363C: push 000000A0h
  00413641: push 004051E8h
  00413646: push ebx
  00413647: push eax
  00413648: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041364E: mov edx, var_18
  00413651: mov ecx, var_58
  00413654: push edx
  00413655: push 00404900h ; "13#1"
  0041365A: mov ebx, [ecx]
  0041365C: call [0040106Ch] ; &
  00413662: mov edx, eax
  00413664: lea ecx, var_1C
  00413667: call MSVBVM60.DLL.__vbaStrMove
  0041366D: push eax
  0041366E: push 00405D84h ; "5. klik tombo simpan yang ada di web."
  00413673: call [0040106Ch] ; &
  00413679: mov edx, eax
  0041367B: lea ecx, var_20
  0041367E: call MSVBVM60.DLL.__vbaStrMove
  00413684: mov var_A8, ebx
  0041368A: mov ebx, var_58
  0041368D: push eax
  0041368E: mov eax, var_A8
  00413694: push ebx
  00413695: call [eax+000000A4h]
  0041369B: test eax, eax
  0041369D: fclex 
  0041369F: jnl 4136B3h
  004136A1: push 000000A4h
  004136A6: push 004051E8h
  004136AB: push ebx
  004136AC: push eax
  004136AD: call MSVBVM60.DLL.__vbaHresultCheckObj
  004136B3: lea ecx, var_20
  004136B6: lea edx, var_1C
  004136B9: push ecx
  004136BA: lea eax, var_18
  004136BD: push edx
  004136BE: push eax
  004136BF: push 00000003h
  004136C1: call MSVBVM60.DLL.__vbaFreeStrList
  004136C7: lea ecx, var_2C
  004136CA: lea edx, var_28
  004136CD: push ecx
  004136CE: push edx
  004136CF: push 00000002h
  004136D1: call MSVBVM60.DLL.__vbaFreeObjList
  004136D7: mov eax, [esi]
  004136D9: add esp, 0000001Ch
  004136DC: push esi
  004136DD: call [eax+00000338h]
  004136E3: lea ecx, var_2C
  004136E6: push eax
  004136E7: push ecx
  004136E8: call edi
  004136EA: mov edx, [esi]
  004136EC: push esi
  004136ED: mov var_58, eax
  004136F0: call [edx+00000338h]
  004136F6: push eax
  004136F7: lea eax, var_28
  004136FA: push eax
  004136FB: call edi
  004136FD: mov ebx, eax
  004136FF: lea edx, var_18
  00413702: push edx
  00413703: push ebx
  00413704: mov ecx, [ebx]
  00413706: call [ecx+000000A0h]
  0041370C: test eax, eax
  0041370E: fclex 
  00413710: jnl 413724h
  00413712: push 000000A0h
  00413717: push 004051E8h
  0041371C: push ebx
  0041371D: push eax
  0041371E: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413724: mov ecx, var_18
  00413727: mov eax, var_58
  0041372A: push ecx
  0041372B: push 00404900h ; "13#1"
  00413730: mov ebx, [eax]
  00413732: call [0040106Ch] ; &
  00413738: mov edx, eax
  0041373A: lea ecx, var_1C
  0041373D: call MSVBVM60.DLL.__vbaStrMove
  00413743: push eax
  00413744: push 00405DD4h ; "6. data berhasil disimpan. dengan cara yang sama klik baris selanjutnya yang ingin dimasukkan, klik masukkan data dan simpan."
  00413749: call [0040106Ch] ; &
  0041374F: mov edx, eax
  00413751: lea ecx, var_20
  00413754: call MSVBVM60.DLL.__vbaStrMove
  0041375A: mov edx, ebx
  0041375C: mov ebx, var_58
  0041375F: push eax
  00413760: push ebx
  00413761: call [edx+000000A4h]
  00413767: test eax, eax
  00413769: fclex 
  0041376B: jnl 41377Fh
  0041376D: push 000000A4h
  00413772: push 004051E8h
  00413777: push ebx
  00413778: push eax
  00413779: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041377F: lea eax, var_20
  00413782: lea ecx, var_1C
  00413785: push eax
  00413786: lea edx, var_18
  00413789: push ecx
  0041378A: push edx
  0041378B: push 00000003h
  0041378D: call MSVBVM60.DLL.__vbaFreeStrList
  00413793: lea eax, var_2C
  00413796: lea ecx, var_28
  00413799: push eax
  0041379A: push ecx
  0041379B: push 00000002h
  0041379D: call MSVBVM60.DLL.__vbaFreeObjList
  004137A3: mov edx, [esi]
  004137A5: add esp, 0000001Ch
  004137A8: push esi
  004137A9: call [edx+00000338h]
  004137AF: push eax
  004137B0: lea eax, var_2C
  004137B3: push eax
  004137B4: call edi
  004137B6: mov ecx, [esi]
  004137B8: push esi
  004137B9: mov var_58, eax
  004137BC: call [ecx+00000338h]
  004137C2: lea edx, var_28
  004137C5: push eax
  004137C6: push edx
  004137C7: call edi
  004137C9: mov ebx, eax
  004137CB: lea ecx, var_18
  004137CE: push ecx
  004137CF: push ebx
  004137D0: mov eax, [ebx]
  004137D2: call [eax+000000A0h]
  004137D8: test eax, eax
  004137DA: fclex 
  004137DC: jnl 4137F0h
  004137DE: push 000000A0h
  004137E3: push 004051E8h
  004137E8: push ebx
  004137E9: push eax
  004137EA: call MSVBVM60.DLL.__vbaHresultCheckObj
  004137F0: mov eax, var_18
  004137F3: mov edx, var_58
  004137F6: push eax
  004137F7: push 00404900h ; "13#1"
  004137FC: mov ebx, [edx]
  004137FE: call [0040106Ch] ; &
  00413804: mov edx, eax
  00413806: lea ecx, var_1C
  00413809: call MSVBVM60.DLL.__vbaStrMove
  0041380F: push eax
  00413810: push 00404900h ; "13#1"
  00413815: call [0040106Ch] ; &
  0041381B: mov edx, eax
  0041381D: lea ecx, var_20
  00413820: call MSVBVM60.DLL.__vbaStrMove
  00413826: push eax
  00413827: push 00405ED4h ; "MEMASUKKAN DATA AK"
  0041382C: call [0040106Ch] ; &
  00413832: mov edx, eax
  00413834: lea ecx, var_24
  00413837: call MSVBVM60.DLL.__vbaStrMove
  0041383D: mov ecx, ebx
  0041383F: mov ebx, var_58
  00413842: push eax
  00413843: push ebx
  00413844: call [ecx+000000A4h]
  0041384A: test eax, eax
  0041384C: fclex 
  0041384E: jnl 413862h
  00413850: push 000000A4h
  00413855: push 004051E8h
  0041385A: push ebx
  0041385B: push eax
  0041385C: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413862: lea edx, var_24
  00413865: lea eax, var_20
  00413868: push edx
  00413869: lea ecx, var_1C
  0041386C: push eax
  0041386D: lea edx, var_18
  00413870: push ecx
  00413871: push edx
  00413872: push 00000004h
  00413874: call MSVBVM60.DLL.__vbaFreeStrList
  0041387A: lea eax, var_2C
  0041387D: lea ecx, var_28
  00413880: push eax
  00413881: push ecx
  00413882: push 00000002h
  00413884: call MSVBVM60.DLL.__vbaFreeObjList
  0041388A: mov edx, [esi]
  0041388C: add esp, 00000020h
  0041388F: push esi
  00413890: call [edx+00000338h]
  00413896: push eax
  00413897: lea eax, var_2C
  0041389A: push eax
  0041389B: call edi
  0041389D: mov ecx, [esi]
  0041389F: push esi
  004138A0: mov var_58, eax
  004138A3: call [ecx+00000338h]
  004138A9: lea edx, var_28
  004138AC: push eax
  004138AD: push edx
  004138AE: call edi
  004138B0: mov ebx, eax
  004138B2: lea ecx, var_18
  004138B5: push ecx
  004138B6: push ebx
  004138B7: mov eax, [ebx]
  004138B9: call [eax+000000A0h]
  004138BF: test eax, eax
  004138C1: fclex 
  004138C3: jnl 4138D7h
  004138C5: push 000000A0h
  004138CA: push 004051E8h
  004138CF: push ebx
  004138D0: push eax
  004138D1: call MSVBVM60.DLL.__vbaHresultCheckObj
  004138D7: mov eax, var_18
  004138DA: mov edx, var_58
  004138DD: push eax
  004138DE: push 00404900h ; "13#1"
  004138E3: mov ebx, [edx]
  004138E5: call [0040106Ch] ; &
  004138EB: mov edx, eax
  004138ED: lea ecx, var_1C
  004138F0: call MSVBVM60.DLL.__vbaStrMove
  004138F6: push eax
  004138F7: push 00405F00h ; "1. klik tombol daftar KK yang ada diprogram"
  004138FC: call [0040106Ch] ; &
  00413902: mov edx, eax
  00413904: lea ecx, var_20
  00413907: call MSVBVM60.DLL.__vbaStrMove
  0041390D: mov ecx, ebx
  0041390F: mov ebx, var_58
  00413912: push eax
  00413913: push ebx
  00413914: call [ecx+000000A4h]
  0041391A: test eax, eax
  0041391C: fclex 
  0041391E: jnl 413932h
  00413920: push 000000A4h
  00413925: push 004051E8h
  0041392A: push ebx
  0041392B: push eax
  0041392C: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413932: lea edx, var_20
  00413935: lea eax, var_1C
  00413938: push edx
  00413939: lea ecx, var_18
  0041393C: push eax
  0041393D: push ecx
  0041393E: push 00000003h
  00413940: call MSVBVM60.DLL.__vbaFreeStrList
  00413946: lea edx, var_2C
  00413949: lea eax, var_28
  0041394C: push edx
  0041394D: push eax
  0041394E: push 00000002h
  00413950: call MSVBVM60.DLL.__vbaFreeObjList
  00413956: mov ecx, [esi]
  00413958: add esp, 0000001Ch
  0041395B: push esi
  0041395C: call [ecx+00000338h]
  00413962: lea edx, var_2C
  00413965: push eax
  00413966: push edx
  00413967: call edi
  00413969: mov var_58, eax
  0041396C: mov eax, [esi]
  0041396E: push esi
  0041396F: call [eax+00000338h]
  00413975: lea ecx, var_28
  00413978: push eax
  00413979: push ecx
  0041397A: call edi
  0041397C: mov ebx, eax
  0041397E: lea eax, var_18
  00413981: push eax
  00413982: push ebx
  00413983: mov edx, [ebx]
  00413985: call [edx+000000A0h]
  0041398B: test eax, eax
  0041398D: fclex 
  0041398F: jnl 4139A3h
  00413991: push 000000A0h
  00413996: push 004051E8h
  0041399B: push ebx
  0041399C: push eax
  0041399D: call MSVBVM60.DLL.__vbaHresultCheckObj
  004139A3: mov edx, var_18
  004139A6: mov ecx, var_58
  004139A9: push edx
  004139AA: push 00404900h ; "13#1"
  004139AF: mov ebx, [ecx]
  004139B1: call [0040106Ch] ; &
  004139B7: mov edx, eax
  004139B9: lea ecx, var_1C
  004139BC: call MSVBVM60.DLL.__vbaStrMove
  004139C2: push eax
  004139C3: push 00405F70h ; "2. dari web, cari KK yang akan diedit daftar AKnya, klik tombol edit AK"
  004139C8: call [0040106Ch] ; &
  004139CE: mov edx, eax
  004139D0: lea ecx, var_20
  004139D3: call MSVBVM60.DLL.__vbaStrMove
  004139D9: mov var_B8, ebx
  004139DF: mov ebx, var_58
  004139E2: push eax
  004139E3: mov eax, var_B8
  004139E9: push ebx
  004139EA: call [eax+000000A4h]
  004139F0: test eax, eax
  004139F2: fclex 
  004139F4: jnl 413A08h
  004139F6: push 000000A4h
  004139FB: push 004051E8h
  00413A00: push ebx
  00413A01: push eax
  00413A02: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413A08: lea ecx, var_20
  00413A0B: lea edx, var_1C
  00413A0E: push ecx
  00413A0F: lea eax, var_18
  00413A12: push edx
  00413A13: push eax
  00413A14: push 00000003h
  00413A16: call MSVBVM60.DLL.__vbaFreeStrList
  00413A1C: lea ecx, var_2C
  00413A1F: lea edx, var_28
  00413A22: push ecx
  00413A23: push edx
  00413A24: push 00000002h
  00413A26: call MSVBVM60.DLL.__vbaFreeObjList
  00413A2C: mov eax, [esi]
  00413A2E: add esp, 0000001Ch
  00413A31: push esi
  00413A32: call [eax+00000338h]
  00413A38: lea ecx, var_2C
  00413A3B: push eax
  00413A3C: push ecx
  00413A3D: call edi
  00413A3F: mov edx, [esi]
  00413A41: push esi
  00413A42: mov var_58, eax
  00413A45: call [edx+00000338h]
  00413A4B: push eax
  00413A4C: lea eax, var_28
  00413A4F: push eax
  00413A50: call edi
  00413A52: mov ebx, eax
  00413A54: lea edx, var_18
  00413A57: push edx
  00413A58: push ebx
  00413A59: mov ecx, [ebx]
  00413A5B: call [ecx+000000A0h]
  00413A61: test eax, eax
  00413A63: fclex 
  00413A65: jnl 413A79h
  00413A67: push 000000A0h
  00413A6C: push 004051E8h
  00413A71: push ebx
  00413A72: push eax
  00413A73: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413A79: mov ecx, var_18
  00413A7C: mov eax, var_58
  00413A7F: push ecx
  00413A80: push 00404900h ; "13#1"
  00413A85: mov ebx, [eax]
  00413A87: call [0040106Ch] ; &
  00413A8D: mov edx, eax
  00413A8F: lea ecx, var_1C
  00413A92: call MSVBVM60.DLL.__vbaStrMove
  00413A98: push eax
  00413A99: push 00406008h ; "3. klik baris data yang ada digrid"
  00413A9E: call [0040106Ch] ; &
  00413AA4: mov edx, eax
  00413AA6: lea ecx, var_20
  00413AA9: call MSVBVM60.DLL.__vbaStrMove
  00413AAF: mov edx, ebx
  00413AB1: mov ebx, var_58
  00413AB4: push eax
  00413AB5: push ebx
  00413AB6: call [edx+000000A4h]
  00413ABC: test eax, eax
  00413ABE: fclex 
  00413AC0: jnl 413AD4h
  00413AC2: push 000000A4h
  00413AC7: push 004051E8h
  00413ACC: push ebx
  00413ACD: push eax
  00413ACE: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413AD4: lea eax, var_20
  00413AD7: lea ecx, var_1C
  00413ADA: push eax
  00413ADB: lea edx, var_18
  00413ADE: push ecx
  00413ADF: push edx
  00413AE0: push 00000003h
  00413AE2: call MSVBVM60.DLL.__vbaFreeStrList
  00413AE8: lea eax, var_2C
  00413AEB: lea ecx, var_28
  00413AEE: push eax
  00413AEF: push ecx
  00413AF0: push 00000002h
  00413AF2: call MSVBVM60.DLL.__vbaFreeObjList
  00413AF8: mov edx, [esi]
  00413AFA: add esp, 0000001Ch
  00413AFD: push esi
  00413AFE: call [edx+00000338h]
  00413B04: push eax
  00413B05: lea eax, var_2C
  00413B08: push eax
  00413B09: call edi
  00413B0B: mov ecx, [esi]
  00413B0D: push esi
  00413B0E: mov var_58, eax
  00413B11: call [ecx+00000338h]
  00413B17: lea edx, var_28
  00413B1A: push eax
  00413B1B: push edx
  00413B1C: call edi
  00413B1E: mov ebx, eax
  00413B20: lea ecx, var_18
  00413B23: push ecx
  00413B24: push ebx
  00413B25: mov eax, [ebx]
  00413B27: call [eax+000000A0h]
  00413B2D: test eax, eax
  00413B2F: fclex 
  00413B31: jnl 413B45h
  00413B33: push 000000A0h
  00413B38: push 004051E8h
  00413B3D: push ebx
  00413B3E: push eax
  00413B3F: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413B45: mov eax, var_18
  00413B48: mov edx, var_58
  00413B4B: push eax
  00413B4C: push 00404900h ; "13#1"
  00413B51: mov ebx, [edx]
  00413B53: call [0040106Ch] ; &
  00413B59: mov edx, eax
  00413B5B: lea ecx, var_1C
  00413B5E: call MSVBVM60.DLL.__vbaStrMove
  00413B64: push eax
  00413B65: push 00406054h ; "4. klik tombol masukkan data AK"
  00413B6A: call [0040106Ch] ; &
  00413B70: mov edx, eax
  00413B72: lea ecx, var_20
  00413B75: call MSVBVM60.DLL.__vbaStrMove
  00413B7B: mov ecx, ebx
  00413B7D: mov ebx, var_58
  00413B80: push eax
  00413B81: push ebx
  00413B82: call [ecx+000000A4h]
  00413B88: test eax, eax
  00413B8A: fclex 
  00413B8C: jnl 413BA0h
  00413B8E: push 000000A4h
  00413B93: push 004051E8h
  00413B98: push ebx
  00413B99: push eax
  00413B9A: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413BA0: lea edx, var_20
  00413BA3: lea eax, var_1C
  00413BA6: push edx
  00413BA7: lea ecx, var_18
  00413BAA: push eax
  00413BAB: push ecx
  00413BAC: push 00000003h
  00413BAE: call MSVBVM60.DLL.__vbaFreeStrList
  00413BB4: lea edx, var_2C
  00413BB7: lea eax, var_28
  00413BBA: push edx
  00413BBB: push eax
  00413BBC: push 00000002h
  00413BBE: call MSVBVM60.DLL.__vbaFreeObjList
  00413BC4: mov ecx, [esi]
  00413BC6: add esp, 0000001Ch
  00413BC9: push esi
  00413BCA: call [ecx+00000338h]
  00413BD0: lea edx, var_2C
  00413BD3: push eax
  00413BD4: push edx
  00413BD5: call edi
  00413BD7: mov var_58, eax
  00413BDA: mov eax, [esi]
  00413BDC: push esi
  00413BDD: call [eax+00000338h]
  00413BE3: lea ecx, var_28
  00413BE6: push eax
  00413BE7: push ecx
  00413BE8: call edi
  00413BEA: mov ebx, eax
  00413BEC: lea eax, var_18
  00413BEF: push eax
  00413BF0: push ebx
  00413BF1: mov edx, [ebx]
  00413BF3: call [edx+000000A0h]
  00413BF9: test eax, eax
  00413BFB: fclex 
  00413BFD: jnl 413C11h
  00413BFF: push 000000A0h
  00413C04: push 004051E8h
  00413C09: push ebx
  00413C0A: push eax
  00413C0B: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413C11: mov edx, var_18
  00413C14: mov ecx, var_58
  00413C17: push edx
  00413C18: push 00404900h ; "13#1"
  00413C1D: mov ebx, [ecx]
  00413C1F: call [0040106Ch] ; &
  00413C25: mov edx, eax
  00413C27: lea ecx, var_1C
  00413C2A: call MSVBVM60.DLL.__vbaStrMove
  00413C30: push eax
  00413C31: push 00406098h ; "5. jika ada data yang mau diperbaiki, ubah data sesuai keinginan..."
  00413C36: call [0040106Ch] ; &
  00413C3C: mov edx, eax
  00413C3E: lea ecx, var_20
  00413C41: call MSVBVM60.DLL.__vbaStrMove
  00413C47: mov var_C4, ebx
  00413C4D: mov ebx, var_58
  00413C50: push eax
  00413C51: mov eax, var_C4
  00413C57: push ebx
  00413C58: call [eax+000000A4h]
  00413C5E: test eax, eax
  00413C60: fclex 
  00413C62: jnl 413C76h
  00413C64: push 000000A4h
  00413C69: push 004051E8h
  00413C6E: push ebx
  00413C6F: push eax
  00413C70: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413C76: lea ecx, var_20
  00413C79: lea edx, var_1C
  00413C7C: push ecx
  00413C7D: lea eax, var_18
  00413C80: push edx
  00413C81: push eax
  00413C82: push 00000003h
  00413C84: call MSVBVM60.DLL.__vbaFreeStrList
  00413C8A: lea ecx, var_2C
  00413C8D: lea edx, var_28
  00413C90: push ecx
  00413C91: push edx
  00413C92: push 00000002h
  00413C94: call MSVBVM60.DLL.__vbaFreeObjList
  00413C9A: mov eax, [esi]
  00413C9C: add esp, 0000001Ch
  00413C9F: push esi
  00413CA0: call [eax+00000338h]
  00413CA6: lea ecx, var_2C
  00413CA9: push eax
  00413CAA: push ecx
  00413CAB: call edi
  00413CAD: mov edx, [esi]
  00413CAF: push esi
  00413CB0: mov var_58, eax
  00413CB3: call [edx+00000338h]
  00413CB9: push eax
  00413CBA: lea eax, var_28
  00413CBD: push eax
  00413CBE: call edi
  00413CC0: mov ebx, eax
  00413CC2: lea edx, var_18
  00413CC5: push edx
  00413CC6: push ebx
  00413CC7: mov ecx, [ebx]
  00413CC9: call [ecx+000000A0h]
  00413CCF: test eax, eax
  00413CD1: fclex 
  00413CD3: jnl 413CE7h
  00413CD5: push 000000A0h
  00413CDA: push 004051E8h
  00413CDF: push ebx
  00413CE0: push eax
  00413CE1: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413CE7: mov ecx, var_18
  00413CEA: mov eax, var_58
  00413CED: push ecx
  00413CEE: push 00404900h ; "13#1"
  00413CF3: mov ebx, [eax]
  00413CF5: call [0040106Ch] ; &
  00413CFB: mov edx, eax
  00413CFD: lea ecx, var_1C
  00413D00: call MSVBVM60.DLL.__vbaStrMove
  00413D06: push eax
  00413D07: push 00406154h ; "6. klik tombol simpan yang ada di web."
  00413D0C: call [0040106Ch] ; &
  00413D12: mov edx, eax
  00413D14: lea ecx, var_20
  00413D17: call MSVBVM60.DLL.__vbaStrMove
  00413D1D: mov edx, ebx
  00413D1F: mov ebx, var_58
  00413D22: push eax
  00413D23: push ebx
  00413D24: call [edx+000000A4h]
  00413D2A: test eax, eax
  00413D2C: fclex 
  00413D2E: jnl 413D42h
  00413D30: push 000000A4h
  00413D35: push 004051E8h
  00413D3A: push ebx
  00413D3B: push eax
  00413D3C: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413D42: lea eax, var_20
  00413D45: lea ecx, var_1C
  00413D48: push eax
  00413D49: lea edx, var_18
  00413D4C: push ecx
  00413D4D: push edx
  00413D4E: push 00000003h
  00413D50: call MSVBVM60.DLL.__vbaFreeStrList
  00413D56: lea eax, var_2C
  00413D59: lea ecx, var_28
  00413D5C: push eax
  00413D5D: push ecx
  00413D5E: push 00000002h
  00413D60: call MSVBVM60.DLL.__vbaFreeObjList
  00413D66: mov edx, [esi]
  00413D68: add esp, 0000001Ch
  00413D6B: push esi
  00413D6C: call [edx+00000338h]
  00413D72: push eax
  00413D73: lea eax, var_2C
  00413D76: push eax
  00413D77: call edi
  00413D79: mov ecx, [esi]
  00413D7B: push esi
  00413D7C: mov var_58, eax
  00413D7F: call [ecx+00000338h]
  00413D85: lea edx, var_28
  00413D88: push eax
  00413D89: push edx
  00413D8A: call edi
  00413D8C: mov ebx, eax
  00413D8E: lea ecx, var_18
  00413D91: push ecx
  00413D92: push ebx
  00413D93: mov eax, [ebx]
  00413D95: call [eax+000000A0h]
  00413D9B: test eax, eax
  00413D9D: fclex 
  00413D9F: jnl 413DB3h
  00413DA1: push 000000A0h
  00413DA6: push 004051E8h
  00413DAB: push ebx
  00413DAC: push eax
  00413DAD: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413DB3: mov eax, var_18
  00413DB6: mov edx, var_58
  00413DB9: push eax
  00413DBA: push 00404900h ; "13#1"
  00413DBF: mov ebx, [edx]
  00413DC1: call [0040106Ch] ; &
  00413DC7: mov edx, eax
  00413DC9: lea ecx, var_1C
  00413DCC: call MSVBVM60.DLL.__vbaStrMove
  00413DD2: push eax
  00413DD3: push 004061A8h ; "7. data berhasil disimpan. dengan cara yang sama klik baris selanjutnya yang ingin dimasukkan, klik masukkan data dan simpan."
  00413DD8: call [0040106Ch] ; &
  00413DDE: mov edx, eax
  00413DE0: lea ecx, var_20
  00413DE3: call MSVBVM60.DLL.__vbaStrMove
  00413DE9: mov ecx, ebx
  00413DEB: mov ebx, var_58
  00413DEE: push eax
  00413DEF: push ebx
  00413DF0: call [ecx+000000A4h]
  00413DF6: test eax, eax
  00413DF8: fclex 
  00413DFA: jnl 413E0Eh
  00413DFC: push 000000A4h
  00413E01: push 004051E8h
  00413E06: push ebx
  00413E07: push eax
  00413E08: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413E0E: lea edx, var_20
  00413E11: lea eax, var_1C
  00413E14: push edx
  00413E15: lea ecx, var_18
  00413E18: push eax
  00413E19: push ecx
  00413E1A: push 00000003h
  00413E1C: call MSVBVM60.DLL.__vbaFreeStrList
  00413E22: lea edx, var_2C
  00413E25: lea eax, var_28
  00413E28: push edx
  00413E29: push eax
  00413E2A: push 00000002h
  00413E2C: call MSVBVM60.DLL.__vbaFreeObjList
  00413E32: mov ecx, [esi]
  00413E34: add esp, 0000001Ch
  00413E37: push esi
  00413E38: call [ecx+00000338h]
  00413E3E: lea edx, var_2C
  00413E41: push eax
  00413E42: push edx
  00413E43: call edi
  00413E45: mov var_58, eax
  00413E48: mov eax, [esi]
  00413E4A: push esi
  00413E4B: call [eax+00000338h]
  00413E51: lea ecx, var_28
  00413E54: push eax
  00413E55: push ecx
  00413E56: call edi
  00413E58: mov ebx, eax
  00413E5A: lea eax, var_18
  00413E5D: push eax
  00413E5E: push ebx
  00413E5F: mov edx, [ebx]
  00413E61: call [edx+000000A0h]
  00413E67: test eax, eax
  00413E69: fclex 
  00413E6B: jnl 413E7Fh
  00413E6D: push 000000A0h
  00413E72: push 004051E8h
  00413E77: push ebx
  00413E78: push eax
  00413E79: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413E7F: mov edx, var_18
  00413E82: mov ecx, var_58
  00413E85: push edx
  00413E86: push 00404900h ; "13#1"
  00413E8B: mov ebx, [ecx]
  00413E8D: call [0040106Ch] ; &
  00413E93: mov edx, eax
  00413E95: lea ecx, var_1C
  00413E98: call MSVBVM60.DLL.__vbaStrMove
  00413E9E: push eax
  00413E9F: push 0040633Ch ; "8. jika diinginkan mengisi data AK di KK yang lain, klik tombol daftar KK,dan dengan cara yang sama ulangi langkah Memasukkan data AK dari awal"
  00413EA4: call [0040106Ch] ; &
  00413EAA: mov edx, eax
  00413EAC: lea ecx, var_20
  00413EAF: call MSVBVM60.DLL.__vbaStrMove
  00413EB5: mov var_D0, ebx
  00413EBB: mov ebx, var_58
  00413EBE: push eax
  00413EBF: mov eax, var_D0
  00413EC5: push ebx
  00413EC6: call [eax+000000A4h]
  00413ECC: test eax, eax
  00413ECE: fclex 
  00413ED0: jnl 413EE4h
  00413ED2: push 000000A4h
  00413ED7: push 004051E8h
  00413EDC: push ebx
  00413EDD: push eax
  00413EDE: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413EE4: lea ecx, var_20
  00413EE7: lea edx, var_1C
  00413EEA: push ecx
  00413EEB: lea eax, var_18
  00413EEE: push edx
  00413EEF: push eax
  00413EF0: push 00000003h
  00413EF2: call MSVBVM60.DLL.__vbaFreeStrList
  00413EF8: lea ecx, var_2C
  00413EFB: lea edx, var_28
  00413EFE: push ecx
  00413EFF: push edx
  00413F00: push 00000002h
  00413F02: call MSVBVM60.DLL.__vbaFreeObjList
  00413F08: mov eax, [esi]
  00413F0A: add esp, 0000001Ch
  00413F0D: push esi
  00413F0E: call [eax+00000338h]
  00413F14: lea ecx, var_2C
  00413F17: push eax
  00413F18: push ecx
  00413F19: call edi
  00413F1B: mov edx, [esi]
  00413F1D: push esi
  00413F1E: mov var_58, eax
  00413F21: call [edx+00000338h]
  00413F27: push eax
  00413F28: lea eax, var_28
  00413F2B: push eax
  00413F2C: call edi
  00413F2E: mov ebx, eax
  00413F30: lea edx, var_18
  00413F33: push edx
  00413F34: push ebx
  00413F35: mov ecx, [ebx]
  00413F37: call [ecx+000000A0h]
  00413F3D: test eax, eax
  00413F3F: fclex 
  00413F41: jnl 413F55h
  00413F43: push 000000A0h
  00413F48: push 004051E8h
  00413F4D: push ebx
  00413F4E: push eax
  00413F4F: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413F55: mov ecx, var_18
  00413F58: mov eax, var_58
  00413F5B: push ecx
  00413F5C: push 00404900h ; "13#1"
  00413F61: mov ebx, [eax]
  00413F63: call [0040106Ch] ; &
  00413F69: mov edx, eax
  00413F6B: lea ecx, var_1C
  00413F6E: call MSVBVM60.DLL.__vbaStrMove
  00413F74: push eax
  00413F75: push 00404900h ; "13#1"
  00413F7A: call [0040106Ch] ; &
  00413F80: mov edx, eax
  00413F82: lea ecx, var_20
  00413F85: call MSVBVM60.DLL.__vbaStrMove
  00413F8B: push eax
  00413F8C: push 00406460h ; "INFORMASI PROGRAM"
  00413F91: call [0040106Ch] ; &
  00413F97: mov edx, eax
  00413F99: lea ecx, var_24
  00413F9C: call MSVBVM60.DLL.__vbaStrMove
  00413FA2: mov edx, ebx
  00413FA4: mov ebx, var_58
  00413FA7: push eax
  00413FA8: push ebx
  00413FA9: call [edx+000000A4h]
  00413FAF: test eax, eax
  00413FB1: fclex 
  00413FB3: jnl 413FC7h
  00413FB5: push 000000A4h
  00413FBA: push 004051E8h
  00413FBF: push ebx
  00413FC0: push eax
  00413FC1: call MSVBVM60.DLL.__vbaHresultCheckObj
  00413FC7: lea eax, var_24
  00413FCA: lea ecx, var_20
  00413FCD: push eax
  00413FCE: lea edx, var_1C
  00413FD1: push ecx
  00413FD2: lea eax, var_18
  00413FD5: push edx
  00413FD6: push eax
  00413FD7: push 00000004h
  00413FD9: call MSVBVM60.DLL.__vbaFreeStrList
  00413FDF: lea ecx, var_2C
  00413FE2: lea edx, var_28
  00413FE5: push ecx
  00413FE6: push edx
  00413FE7: push 00000002h
  00413FE9: call MSVBVM60.DLL.__vbaFreeObjList
  00413FEF: mov eax, [esi]
  00413FF1: add esp, 00000020h
  00413FF4: push esi
  00413FF5: call [eax+00000338h]
  00413FFB: lea ecx, var_2C
  00413FFE: push eax
  00413FFF: push ecx
  00414000: call edi
  00414002: mov edx, [esi]
  00414004: push esi
  00414005: mov var_58, eax
  00414008: call [edx+00000338h]
  0041400E: push eax
  0041400F: lea eax, var_28
  00414012: push eax
  00414013: call edi
  00414015: mov ebx, eax
  00414017: lea edx, var_18
  0041401A: push edx
  0041401B: push ebx
  0041401C: mov ecx, [ebx]
  0041401E: call [ecx+000000A0h]
  00414024: test eax, eax
  00414026: fclex 
  00414028: jnl 41403Ch
  0041402A: push 000000A0h
  0041402F: push 004051E8h
  00414034: push ebx
  00414035: push eax
  00414036: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041403C: mov ecx, var_18
  0041403F: mov eax, var_58
  00414042: push ecx
  00414043: push 00404900h ; "13#1"
  00414048: mov ebx, [eax]
  0041404A: call [0040106Ch] ; &
  00414050: mov edx, eax
  00414052: lea ecx, var_1C
  00414055: call MSVBVM60.DLL.__vbaStrMove
  0041405B: push eax
  0041405C: push 00406488h ; "Program ini dikembangkan oleh :"
  00414061: call [0040106Ch] ; &
  00414067: mov edx, eax
  00414069: lea ecx, var_20
  0041406C: call MSVBVM60.DLL.__vbaStrMove
  00414072: mov edx, ebx
  00414074: mov ebx, var_58
  00414077: push eax
  00414078: push ebx
  00414079: call [edx+000000A4h]
  0041407F: test eax, eax
  00414081: fclex 
  00414083: jnl 414097h
  00414085: push 000000A4h
  0041408A: push 004051E8h
  0041408F: push ebx
  00414090: push eax
  00414091: call MSVBVM60.DLL.__vbaHresultCheckObj
  00414097: lea eax, var_20
  0041409A: lea ecx, var_1C
  0041409D: push eax
  0041409E: lea edx, var_18
  004140A1: push ecx
  004140A2: push edx
  004140A3: push 00000003h
  004140A5: call MSVBVM60.DLL.__vbaFreeStrList
  004140AB: lea eax, var_2C
  004140AE: lea ecx, var_28
  004140B1: push eax
  004140B2: push ecx
  004140B3: push 00000002h
  004140B5: call MSVBVM60.DLL.__vbaFreeObjList
  004140BB: mov edx, [esi]
  004140BD: add esp, 0000001Ch
  004140C0: push esi
  004140C1: call [edx+00000338h]
  004140C7: push eax
  004140C8: lea eax, var_2C
  004140CB: push eax
  004140CC: call edi
  004140CE: mov ecx, [esi]
  004140D0: push esi
  004140D1: mov var_58, eax
  004140D4: call [ecx+00000338h]
  004140DA: lea edx, var_28
  004140DD: push eax
  004140DE: push edx
  004140DF: call edi
  004140E1: mov ebx, eax
  004140E3: lea ecx, var_18
  004140E6: push ecx
  004140E7: push ebx
  004140E8: mov eax, [ebx]
  004140EA: call [eax+000000A0h]
  004140F0: test eax, eax
  004140F2: fclex 
  004140F4: jnl 414108h
  004140F6: push 000000A0h
  004140FB: push 004051E8h
  00414100: push ebx
  00414101: push eax
  00414102: call MSVBVM60.DLL.__vbaHresultCheckObj
  00414108: mov eax, var_18
  0041410B: mov edx, var_58
  0041410E: push eax
  0041410F: push 00404900h ; "13#1"
  00414114: mov ebx, [edx]
  00414116: call [0040106Ch] ; &
  0041411C: mov edx, eax
  0041411E: lea ecx, var_1C
  00414121: call MSVBVM60.DLL.__vbaStrMove
  00414127: push eax
  00414128: push 00404690h ; "JOGJASOFTWARE"
  0041412D: call [0040106Ch] ; &
  00414133: mov edx, eax
  00414135: lea ecx, var_20
  00414138: call MSVBVM60.DLL.__vbaStrMove
  0041413E: mov ecx, ebx
  00414140: mov ebx, var_58
  00414143: push eax
  00414144: push ebx
  00414145: call [ecx+000000A4h]
  0041414B: test eax, eax
  0041414D: fclex 
  0041414F: jnl 414163h
  00414151: push 000000A4h
  00414156: push 004051E8h
  0041415B: push ebx
  0041415C: push eax
  0041415D: call MSVBVM60.DLL.__vbaHresultCheckObj
  00414163: lea edx, var_20
  00414166: lea eax, var_1C
  00414169: push edx
  0041416A: lea ecx, var_18
  0041416D: push eax
  0041416E: push ecx
  0041416F: push 00000003h
  00414171: call MSVBVM60.DLL.__vbaFreeStrList
  00414177: lea edx, var_2C
  0041417A: lea eax, var_28
  0041417D: push edx
  0041417E: push eax
  0041417F: push 00000002h
  00414181: call MSVBVM60.DLL.__vbaFreeObjList
  00414187: mov ecx, [esi]
  00414189: add esp, 0000001Ch
  0041418C: push esi
  0041418D: call [ecx+00000338h]
  00414193: lea edx, var_2C
  00414196: push eax
  00414197: push edx
  00414198: call edi
  0041419A: mov var_58, eax
  0041419D: mov eax, [esi]
  0041419F: push esi
  004141A0: call [eax+00000338h]
  004141A6: lea ecx, var_28
  004141A9: push eax
  004141AA: push ecx
  004141AB: call edi
  004141AD: mov ebx, eax
  004141AF: lea eax, var_18
  004141B2: push eax
  004141B3: push ebx
  004141B4: mov edx, [ebx]
  004141B6: call [edx+000000A0h]
  004141BC: test eax, eax
  004141BE: fclex 
  004141C0: jnl 4141D4h
  004141C2: push 000000A0h
  004141C7: push 004051E8h
  004141CC: push ebx
  004141CD: push eax
  004141CE: call MSVBVM60.DLL.__vbaHresultCheckObj
  004141D4: mov edx, var_18
  004141D7: mov ecx, var_58
  004141DA: push edx
  004141DB: push 00404900h ; "13#1"
  004141E0: mov ebx, [ecx]
  004141E2: call [0040106Ch] ; &
  004141E8: mov edx, eax
  004141EA: lea ecx, var_1C
  004141ED: call MSVBVM60.DLL.__vbaStrMove
  004141F3: push eax
  004141F4: push 004064CCh ; "Web : http://www.jogjasoftware.net"
  004141F9: call [0040106Ch] ; &
  004141FF: mov edx, eax
  00414201: lea ecx, var_20
  00414204: call MSVBVM60.DLL.__vbaStrMove
  0041420A: mov var_E0, ebx
  00414210: mov ebx, var_58
  00414213: push eax
  00414214: mov eax, var_E0
  0041421A: push ebx
  0041421B: call [eax+000000A4h]
  00414221: test eax, eax
  00414223: fclex 
  00414225: jnl 414239h
  00414227: push 000000A4h
  0041422C: push 004051E8h
  00414231: push ebx
  00414232: push eax
  00414233: call MSVBVM60.DLL.__vbaHresultCheckObj
  00414239: lea ecx, var_20
  0041423C: lea edx, var_1C
  0041423F: push ecx
  00414240: lea eax, var_18
  00414243: push edx
  00414244: push eax
  00414245: push 00000003h
  00414247: call MSVBVM60.DLL.__vbaFreeStrList
  0041424D: lea ecx, var_2C
  00414250: lea edx, var_28
  00414253: push ecx
  00414254: push edx
  00414255: push 00000002h
  00414257: call MSVBVM60.DLL.__vbaFreeObjList
  0041425D: mov eax, [esi]
  0041425F: add esp, 0000001Ch
  00414262: push esi
  00414263: call [eax+00000338h]
  00414269: lea ecx, var_2C
  0041426C: push eax
  0041426D: push ecx
  0041426E: call edi
  00414270: mov edx, [esi]
  00414272: push esi
  00414273: mov var_58, eax
  00414276: call [edx+00000338h]
  0041427C: push eax
  0041427D: lea eax, var_28
  00414280: push eax
  00414281: call edi
  00414283: mov ebx, eax
  00414285: lea edx, var_18
  00414288: push edx
  00414289: push ebx
  0041428A: mov ecx, [ebx]
  0041428C: call [ecx+000000A0h]
  00414292: test eax, eax
  00414294: fclex 
  00414296: jnl 4142AAh
  00414298: push 000000A0h
  0041429D: push 004051E8h
  004142A2: push ebx
  004142A3: push eax
  004142A4: call MSVBVM60.DLL.__vbaHresultCheckObj
  004142AA: mov ecx, var_18
  004142AD: mov eax, var_58
  004142B0: push ecx
  004142B1: push 00404900h ; "13#1"
  004142B6: mov ebx, [eax]
  004142B8: call [0040106Ch] ; &
  004142BE: mov edx, eax
  004142C0: lea ecx, var_1C
  004142C3: call MSVBVM60.DLL.__vbaStrMove
  004142C9: push eax
  004142CA: push 004062A8h ; "Email : info@jogjasoftware.net"
  004142CF: call [0040106Ch] ; &
  004142D5: mov edx, eax
  004142D7: lea ecx, var_20
  004142DA: call MSVBVM60.DLL.__vbaStrMove
  004142E0: mov edx, ebx
  004142E2: mov ebx, var_58
  004142E5: push eax
  004142E6: push ebx
  004142E7: call [edx+000000A4h]
  004142ED: test eax, eax
  004142EF: fclex 
  004142F1: jnl 414305h
  004142F3: push 000000A4h
  004142F8: push 004051E8h
  004142FD: push ebx
  004142FE: push eax
  004142FF: call MSVBVM60.DLL.__vbaHresultCheckObj
  00414305: lea eax, var_20
  00414308: lea ecx, var_1C
  0041430B: push eax
  0041430C: lea edx, var_18
  0041430F: push ecx
  00414310: push edx
  00414311: push 00000003h
  00414313: call MSVBVM60.DLL.__vbaFreeStrList
  00414319: lea eax, var_2C
  0041431C: lea ecx, var_28
  0041431F: push eax
  00414320: push ecx
  00414321: push 00000002h
  00414323: call MSVBVM60.DLL.__vbaFreeObjList
  00414329: mov edx, [esi]
  0041432B: add esp, 0000001Ch
  0041432E: push esi
  0041432F: call [edx+00000338h]
  00414335: push eax
  00414336: lea eax, var_2C
  00414339: push eax
  0041433A: call edi
  0041433C: mov ecx, [esi]
  0041433E: push esi
  0041433F: mov var_58, eax
  00414342: call [ecx+00000338h]
  00414348: lea edx, var_28
  0041434B: push eax
  0041434C: push edx
  0041434D: call edi
  0041434F: mov ebx, eax
  00414351: lea ecx, var_18
  00414354: push ecx
  00414355: push ebx
  00414356: mov eax, [ebx]
  00414358: call [eax+000000A0h]
  0041435E: test eax, eax
  00414360: fclex 
  00414362: jnl 414376h
  00414364: push 000000A0h
  00414369: push 004051E8h
  0041436E: push ebx
  0041436F: push eax
  00414370: call MSVBVM60.DLL.__vbaHresultCheckObj
  00414376: mov eax, var_18
  00414379: mov edx, var_58
  0041437C: push eax
  0041437D: push 00404900h ; "13#1"
  00414382: mov ebx, [edx]
  00414384: call [0040106Ch] ; &
  0041438A: mov edx, eax
  0041438C: lea ecx, var_1C
  0041438F: call MSVBVM60.DLL.__vbaStrMove
  00414395: push eax
  00414396: push 00404B28h ; "Telp : 0812 2993 464"
  0041439B: call [0040106Ch] ; &
  004143A1: mov edx, eax
  004143A3: lea ecx, var_20
  004143A6: call MSVBVM60.DLL.__vbaStrMove
  004143AC: mov ecx, ebx
  004143AE: mov ebx, var_58
  004143B1: push eax
  004143B2: push ebx
  004143B3: call [ecx+000000A4h]
  004143B9: test eax, eax
  004143BB: fclex 
  004143BD: jnl 4143D1h
  004143BF: push 000000A4h
  004143C4: push 004051E8h
  004143C9: push ebx
  004143CA: push eax
  004143CB: call MSVBVM60.DLL.__vbaHresultCheckObj
  004143D1: lea edx, var_20
  004143D4: lea eax, var_1C
  004143D7: push edx
  004143D8: lea ecx, var_18
  004143DB: push eax
  004143DC: push ecx
  004143DD: push 00000003h
  004143DF: call MSVBVM60.DLL.__vbaFreeStrList
  004143E5: lea edx, var_2C
  004143E8: lea eax, var_28
  004143EB: push edx
  004143EC: push eax
  004143ED: push 00000002h
  004143EF: call MSVBVM60.DLL.__vbaFreeObjList
  004143F5: add esp, 0000000Ch
  004143F8: mov ecx, 0000000Bh
  004143FD: mov edx, esp
  004143FF: xor eax, eax
  00414401: push 80010007h
  00414406: push esi
  00414407: mov [edx], ecx
  00414409: mov ecx, var_38
  0041440C: mov [edx+04h], ecx
  0041440F: mov ecx, [esi]
  00414411: mov [edx+08h], eax
  00414414: mov eax, var_30
  00414417: mov [edx+0Ch], eax
  0041441A: call [ecx+00000344h]
  00414420: lea edx, var_28
  00414423: push eax
  00414424: push edx
  00414425: call edi
  00414427: push eax
  00414428: call MSVBVM60.DLL.__vbaLateIdSt
  0041442E: lea ecx, var_28
  00414431: call MSVBVM60.DLL.__vbaFreeObj
  00414437: mov var_4, 00000000h
  0041443E: push 00414472h
  00414443: jmp 414471h
  00414445: lea eax, var_24
  00414448: lea ecx, var_20
  0041444B: push eax
  0041444C: lea edx, var_1C
  0041444F: push ecx
  00414450: lea eax, var_18
  00414453: push edx
  00414454: push eax
  00414455: push 00000004h
  00414457: call MSVBVM60.DLL.__vbaFreeStrList
  0041445D: lea ecx, var_2C
  00414460: lea edx, var_28
  00414463: push ecx
  00414464: push edx
  00414465: push 00000002h
  00414467: call MSVBVM60.DLL.__vbaFreeObjList
  0041446D: add esp, 00000020h
  00414470: ret 
End Sub

Private Sub Form_Resize() '414FC0
  00414FC0: push ebp
  00414FC1: mov ebp, esp
  00414FC3: sub esp, 00000018h
  00414FC6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00414FCB: mov eax, fs:[00h]
  00414FD1: push eax
  00414FD2: mov fs:[00000000h], esp
  00414FD9: mov eax, 00000068h
  00414FDE: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00414FE3: push ebx
  00414FE4: push esi
  00414FE5: push edi
  00414FE6: mov var_18, esp
  00414FE9: mov var_14, 004014B0h
  00414FF0: mov eax, arg_8
  00414FF3: and eax, 00000001h
  00414FF6: mov var_10, eax
  00414FF9: mov ecx, arg_8
  00414FFC: and ecx, FFFFFFFEh
  00414FFF: mov arg_8, ecx
  00415002: mov var_C, 00000000h
  00415009: mov edx, arg_8
  0041500C: mov eax, [edx]
  0041500E: mov ecx, arg_8
  00415011: push ecx
  00415012: call [eax+04h]
  00415015: mov var_4, 00000001h
  0041501C: mov var_4, 00000002h
  00415023: push FFFFFFFFh
  00415025: call On Error ...
  0041502B: mov var_4, 00000003h
  00415032: mov edx, arg_8
  00415035: mov eax, [edx]
  00415037: mov ecx, arg_8
  0041503A: push ecx
  0041503B: call [eax+00000310h]
  00415041: push eax
  00415042: lea edx, var_24
  00415045: push edx
  00415046: call [004010BCh] ; Set (object)
  0041504C: mov var_54, eax
  0041504F: lea eax, var_50
  00415052: push eax
  00415053: mov ecx, var_54
  00415056: mov edx, [ecx]
  00415058: mov eax, var_54
  0041505B: push eax
  0041505C: call [edx+00000080h]
  00415062: fclex 
  00415064: mov var_58, eax
  00415067: cmp var_58, 00000000h
  0041506B: jnl 41508Ah
  0041506D: push 00000080h
  00415072: push 004055A4h
  00415077: mov ecx, var_54
  0041507A: push ecx
  0041507B: mov edx, var_58
  0041507E: push edx
  0041507F: call MSVBVM60.DLL.__vbaHresultCheckObj
  00415085: mov var_74, eax
  00415088: jmp 415091h
  0041508A: mov var_74, 00000000h
  00415091: lea eax, var_4C
  00415094: push eax
  00415095: mov ecx, arg_8
  00415098: mov edx, [ecx]
  0041509A: mov eax, arg_8
  0041509D: push eax
  0041509E: call [edx+00000080h]
  004150A4: fclex 
  004150A6: mov var_5C, eax
  004150A9: cmp var_5C, 00000000h
  004150AD: jnl 4150CCh
  004150AF: push 00000080h
  004150B4: push 00404B54h
  004150B9: mov ecx, arg_8
  004150BC: push ecx
  004150BD: mov edx, var_5C
  004150C0: push edx
  004150C1: call MSVBVM60.DLL.__vbaHresultCheckObj
  004150C7: mov var_78, eax
  004150CA: jmp 4150D3h
  004150CC: mov var_78, 00000000h
  004150D3: fld real4 ptr var_4C
  004150D6: fsub real4 ptr var_50
  004150D9: fsub real4 ptr [004014F8h] ; 
  004150DF: fstp real4 ptr var_30
  004150E2: fstsw ax
  004150E4: test al, 0Dh
  004150E6: jnz 0041545Ch
  004150EC: mov var_38, 00000004h
  004150F3: mov eax, 00000010h
  004150F8: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  004150FD: mov eax, esp
  004150FF: mov ecx, var_38
  00415102: mov [eax], ecx
  00415104: mov edx, var_34
  00415107: mov [eax+04h], edx
  0041510A: mov ecx, var_30
  0041510D: mov [eax+08h], ecx
  00415110: mov edx, var_2C
  00415113: mov [eax+0Ch], edx
  00415116: push 80010005h
  0041511B: mov eax, arg_8
  0041511E: mov ecx, [eax]
  00415120: mov edx, arg_8
  00415123: push edx
  00415124: call [ecx+00000344h]
  0041512A: push eax
  0041512B: lea eax, var_28
  0041512E: push eax
  0041512F: call [004010BCh] ; Set (object)
  00415135: push eax
  00415136: call MSVBVM60.DLL.__vbaLateIdSt
  0041513C: lea ecx, var_28
  0041513F: push ecx
  00415140: lea edx, var_24
  00415143: push edx
  00415144: push 00000002h
  00415146: call MSVBVM60.DLL.__vbaFreeObjList
  0041514C: add esp, 0000000Ch
  0041514F: mov var_4, 00000004h
  00415156: mov eax, arg_8
  00415159: mov ecx, [eax]
  0041515B: mov edx, arg_8
  0041515E: push edx
  0041515F: call [ecx+000002FCh]
  00415165: push eax
  00415166: lea eax, var_24
  00415169: push eax
  0041516A: call [004010BCh] ; Set (object)
  00415170: mov var_54, eax
  00415173: lea ecx, var_50
  00415176: push ecx
  00415177: mov edx, var_54
  0041517A: mov eax, [edx]
  0041517C: mov ecx, var_54
  0041517F: push ecx
  00415180: call [eax+00000088h]
  00415186: fclex 
  00415188: mov var_58, eax
  0041518B: cmp var_58, 00000000h
  0041518F: jnl 4151AEh
  00415191: push 00000088h
  00415196: push 004055A4h
  0041519B: mov edx, var_54
  0041519E: push edx
  0041519F: mov eax, var_58
  004151A2: push eax
  004151A3: call MSVBVM60.DLL.__vbaHresultCheckObj
  004151A9: mov var_7C, eax
  004151AC: jmp 4151B5h
  004151AE: mov var_7C, 00000000h
  004151B5: lea ecx, var_4C
  004151B8: push ecx
  004151B9: mov edx, arg_8
  004151BC: mov eax, [edx]
  004151BE: mov ecx, arg_8
  004151C1: push ecx
  004151C2: call [eax+00000088h]
  004151C8: fclex 
  004151CA: mov var_5C, eax
  004151CD: cmp var_5C, 00000000h
  004151D1: jnl 4151F0h
  004151D3: push 00000088h
  004151D8: push 00404B54h
  004151DD: mov edx, arg_8
  004151E0: push edx
  004151E1: mov eax, var_5C
  004151E4: push eax
  004151E5: call MSVBVM60.DLL.__vbaHresultCheckObj
  004151EB: mov var_80, eax
  004151EE: jmp 4151F7h
  004151F0: mov var_80, 00000000h
  004151F7: fld real4 ptr var_4C
  004151FA: fsub real4 ptr var_50
  004151FD: fsub real4 ptr [004014F4h] ; 
  00415203: fstp real4 ptr var_30
  00415206: fstsw ax
  00415208: test al, 0Dh
  0041520A: jnz 0041545Ch
  00415210: mov var_38, 00000004h
  00415217: mov eax, 00000010h
  0041521C: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00415221: mov ecx, esp
  00415223: mov edx, var_38
  00415226: mov [ecx], edx
  00415228: mov eax, var_34
  0041522B: mov [ecx+04h], eax
  0041522E: mov edx, var_30
  00415231: mov [ecx+08h], edx
  00415234: mov eax, var_2C
  00415237: mov [ecx+0Ch], eax
  0041523A: push 80010006h
  0041523F: mov ecx, arg_8
  00415242: mov edx, [ecx]
  00415244: mov eax, arg_8
  00415247: push eax
  00415248: call [edx+00000344h]
  0041524E: push eax
  0041524F: lea ecx, var_28
  00415252: push ecx
  00415253: call [004010BCh] ; Set (object)
  00415259: push eax
  0041525A: call MSVBVM60.DLL.__vbaLateIdSt
  00415260: lea edx, var_28
  00415263: push edx
  00415264: lea eax, var_24
  00415267: push eax
  00415268: push 00000002h
  0041526A: call MSVBVM60.DLL.__vbaFreeObjList
  00415270: add esp, 0000000Ch
  00415273: mov var_4, 00000005h
  0041527A: lea ecx, var_4C
  0041527D: push ecx
  0041527E: mov edx, arg_8
  00415281: mov eax, [edx]
  00415283: mov ecx, arg_8
  00415286: push ecx
  00415287: call [eax+00000080h]
  0041528D: fclex 
  0041528F: mov var_54, eax
  00415292: cmp var_54, 00000000h
  00415296: jnl 4152B8h
  00415298: push 00000080h
  0041529D: push 00404B54h
  004152A2: mov edx, arg_8
  004152A5: push edx
  004152A6: mov eax, var_54
  004152A9: push eax
  004152AA: call MSVBVM60.DLL.__vbaHresultCheckObj
  004152B0: mov var_84, eax
  004152B6: jmp 4152C2h
  004152B8: mov var_84, 00000000h
  004152C2: fld real4 ptr var_4C
  004152C5: fsub real4 ptr [004014F0h] ; 
  004152CB: fstp real4 ptr var_30
  004152CE: fstsw ax
  004152D0: test al, 0Dh
  004152D2: jnz 0041545Ch
  004152D8: mov var_38, 00000004h
  004152DF: mov eax, 00000010h
  004152E4: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  004152E9: mov ecx, esp
  004152EB: mov edx, var_38
  004152EE: mov [ecx], edx
  004152F0: mov eax, var_34
  004152F3: mov [ecx+04h], eax
  004152F6: mov edx, var_30
  004152F9: mov [ecx+08h], edx
  004152FC: mov eax, var_2C
  004152FF: mov [ecx+0Ch], eax
  00415302: push 80010005h
  00415307: mov ecx, arg_8
  0041530A: mov edx, [ecx]
  0041530C: mov eax, arg_8
  0041530F: push eax
  00415310: call [edx+00000340h]
  00415316: push eax
  00415317: lea ecx, var_24
  0041531A: push ecx
  0041531B: call [004010BCh] ; Set (object)
  00415321: push eax
  00415322: call MSVBVM60.DLL.__vbaLateIdSt
  00415328: lea ecx, var_24
  0041532B: call MSVBVM60.DLL.__vbaFreeObj
  00415331: mov var_4, 00000006h
  00415338: mov edx, arg_8
  0041533B: mov eax, [edx]
  0041533D: mov ecx, arg_8
  00415340: push ecx
  00415341: call [eax+000002FCh]
  00415347: push eax
  00415348: lea edx, var_24
  0041534B: push edx
  0041534C: call [004010BCh] ; Set (object)
  00415352: mov var_54, eax
  00415355: lea eax, var_4C
  00415358: push eax
  00415359: mov ecx, var_54
  0041535C: mov edx, [ecx]
  0041535E: mov eax, var_54
  00415361: push eax
  00415362: call [edx+00000088h]
  00415368: fclex 
  0041536A: mov var_58, eax
  0041536D: cmp var_58, 00000000h
  00415371: jnl 415393h
  00415373: push 00000088h
  00415378: push 004055A4h
  0041537D: mov ecx, var_54
  00415380: push ecx
  00415381: mov edx, var_58
  00415384: push edx
  00415385: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041538B: mov var_88, eax
  00415391: jmp 41539Dh
  00415393: mov var_88, 00000000h
  0041539D: fld real4 ptr var_4C
  004153A0: fsub real4 ptr [004014ECh] ; 
  004153A6: fstp real4 ptr var_30
  004153A9: fstsw ax
  004153AB: test al, 0Dh
  004153AD: jnz 0041545Ch
  004153B3: mov var_38, 00000004h
  004153BA: mov eax, 00000010h
  004153BF: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  004153C4: mov eax, esp
  004153C6: mov ecx, var_38
  004153C9: mov [eax], ecx
  004153CB: mov edx, var_34
  004153CE: mov [eax+04h], edx
  004153D1: mov ecx, var_30
  004153D4: mov [eax+08h], ecx
  004153D7: mov edx, var_2C
  004153DA: mov [eax+0Ch], edx
  004153DD: push 80010006h
  004153E2: mov eax, arg_8
  004153E5: mov ecx, [eax]
  004153E7: mov edx, arg_8
  004153EA: push edx
  004153EB: call [ecx+00000340h]
  004153F1: push eax
  004153F2: lea eax, var_28
  004153F5: push eax
  004153F6: call [004010BCh] ; Set (object)
  004153FC: push eax
  004153FD: call MSVBVM60.DLL.__vbaLateIdSt
  00415403: lea ecx, var_28
  00415406: push ecx
  00415407: lea edx, var_24
  0041540A: push edx
  0041540B: push 00000002h
  0041540D: call MSVBVM60.DLL.__vbaFreeObjList
  00415413: add esp, 0000000Ch
  00415416: mov var_10, 00000000h
  0041541D: wait 
  0041541E: push 0041543Ah
  00415423: jmp 415439h
  00415425: lea eax, var_28
  00415428: push eax
  00415429: lea ecx, var_24
  0041542C: push ecx
  0041542D: push 00000002h
  0041542F: call MSVBVM60.DLL.__vbaFreeObjList
  00415435: add esp, 0000000Ch
  00415438: ret 
End Sub

Public Sub test1() '4146B0
  004146B0: push ebp
  004146B1: mov ebp, esp
  004146B3: sub esp, 0000000Ch
  004146B6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  004146BB: mov eax, fs:[00h]
  004146C1: push eax
  004146C2: mov fs:[00000000h], esp
  004146C9: sub esp, 00000044h
  004146CC: push ebx
  004146CD: push esi
  004146CE: push edi
  004146CF: mov var_C, esp
  004146D2: mov var_8, 00401480h
  004146D9: xor esi, esi
  004146DB: mov var_4, esi
  004146DE: mov eax, arg_8
  004146E1: push eax
  004146E2: mov ecx, [eax]
  004146E4: call [ecx+04h]
  004146E7: push 004066C0h
  004146EC: push 004051F8h
  004146F1: mov edi, var_3C
  004146F4: sub esp, 00000010h
  004146F7: mov edx, esp
  004146F9: mov ecx, 00000008h
  004146FE: mov ebx, var_34
  00414701: mov eax, 004068D8h ; "upc"
  00414706: mov [edx], ecx
  00414708: push 00000001h
  0041470A: push 00000440h
  0041470F: push esi
  00414710: mov [edx+04h], edi
  00414713: mov var_18, esi
  00414716: mov var_1C, esi
  00414719: mov var_20, esi
  0041471C: mov [edx+08h], eax
  0041471F: lea eax, var_30
  00414722: push eax
  00414723: mov var_30, esi
  00414726: mov [edx+0Ch], ebx
  00414729: call MSVBVM60.DLL.__vbaLateIdCallLd
  0041472F: add esp, 00000020h
  00414732: push eax
  00414733: call MSVBVM60.DLL.__vbaCastObjVar
  00414739: lea ecx, var_20
  0041473C: push eax
  0041473D: push ecx
  0041473E: call [004010BCh] ; Set (object)
  00414744: push eax
  00414745: call MSVBVM60.DLL.__vbaCastObj
  0041474B: lea edx, var_1C
  0041474E: push eax
  0041474F: push edx
  00414750: call [004010BCh] ; Set (object)
  00414756: lea ecx, var_20
  00414759: call MSVBVM60.DLL.__vbaFreeObj
  0041475F: lea ecx, var_30
  00414762: call MSVBVM60.DLL.__vbaFreeVar
  00414768: mov eax, var_1C
  0041476B: push esi
  0041476C: push 800107D0h
  00414771: push eax
  00414772: call MSVBVM60.DLL.__vbaLateIdCall
  00414778: mov ecx, 00000008h
  0041477D: mov eax, 004066D4h ; "07800008846"
  00414782: push ecx
  00414783: mov edx, esp
  00414785: push 800113EDh
  0041478A: mov [edx], ecx
  0041478C: mov [edx+04h], edi
  0041478F: mov [edx+08h], eax
  00414792: mov eax, var_1C
  00414795: push eax
  00414796: mov [edx+0Ch], ebx
  00414799: call MSVBVM60.DLL.__vbaLateIdSt
  0041479F: lea ecx, var_30
  004147A2: mov var_28, FFFFFFFFh
  004147A9: push ecx
  004147AA: push 004066F0h ; "{ENTER}"
  004147AF: mov var_30, 0000000Bh
  004147B6: call [004010E4h] ; SendKeys
  004147BC: lea ecx, var_30
  004147BF: call MSVBVM60.DLL.__vbaFreeVar
  004147C5: push 004147F0h
  004147CA: jmp 4147DFh
  004147CC: lea ecx, var_20
  004147CF: call MSVBVM60.DLL.__vbaFreeObj
  004147D5: lea ecx, var_30
  004147D8: call MSVBVM60.DLL.__vbaFreeVar
  004147DE: ret 
End Sub

Public Function cekRadioHtml(nmradio) '415C20
  00415C20: push ebp
  00415C21: mov ebp, esp
  00415C23: sub esp, 0000000Ch
  00415C26: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00415C2B: mov eax, fs:[00h]
  00415C31: push eax
  00415C32: mov fs:[00000000h], esp
  00415C39: sub esp, 0000011Ch
  00415C3F: push ebx
  00415C40: push esi
  00415C41: push edi
  00415C42: mov var_C, esp
  00415C45: mov var_8, 00401528h
  00415C4C: xor esi, esi
  00415C4E: mov var_4, esi
  00415C51: mov edi, arg_8
  00415C54: push edi
  00415C55: mov eax, [edi]
  00415C57: call [eax+04h]
  00415C5A: mov ecx, arg_14
  00415C5D: mov edx, nmradio
  00415C60: mov var_24, esi
  00415C63: mov var_34, esi
  00415C66: mov [ecx], esi
  00415C68: lea ecx, var_B4
  00415C6E: mov var_44, esi
  00415C71: mov var_54, esi
  00415C74: mov var_64, esi
  00415C77: mov var_68, esi
  00415C7A: mov var_6C, esi
  00415C7D: mov var_70, esi
  00415C80: mov var_74, esi
  00415C83: mov var_84, esi
  00415C89: mov var_94, esi
  00415C8F: mov var_A4, esi
  00415C95: mov var_B4, esi
  00415C9B: mov var_C4, esi
  00415CA1: mov var_D4, esi
  00415CA7: mov var_E8, esi
  00415CAD: mov var_10C, esi
  00415CB3: mov var_11C, esi
  00415CB9: call MSVBVM60.DLL.__vbaVarVargNofree
  00415CBF: mov edx, eax
  00415CC1: lea ecx, var_44
  00415CC4: call MSVBVM60.DLL.__vbaVarCopy
  00415CCA: lea edx, var_44
  00415CCD: push edx
  00415CCE: call MSVBVM60.DLL.__vbaStrVarCopy
  00415CD4: push 00404EF8h
  00415CD9: mov ecx, 00000008h
  00415CDE: sub esp, 00000010h
  00415CE1: mov var_84, ecx
  00415CE7: mov edx, esp
  00415CE9: mov var_7C, eax
  00415CEC: push 00000001h
  00415CEE: push 0000043Eh
  00415CF3: mov [edx], ecx
  00415CF5: mov ecx, var_80
  00415CF8: mov [edx+04h], ecx
  00415CFB: mov ecx, [edi+34h]
  00415CFE: push ecx
  00415CFF: mov [edx+08h], eax
  00415D02: mov eax, var_78
  00415D05: mov [edx+0Ch], eax
  00415D08: lea edx, var_94
  00415D0E: push edx
  00415D0F: call MSVBVM60.DLL.__vbaLateIdCallLd
  00415D15: add esp, 00000020h
  00415D18: push eax
  00415D19: call MSVBVM60.DLL.__vbaCastObjVar
  00415D1F: push eax
  00415D20: lea eax, var_70
  00415D23: push eax
  00415D24: call [004010BCh] ; Set (object)
  00415D2A: mov ebx, eax
  00415D2C: lea edx, var_E8
  00415D32: push edx
  00415D33: push ebx
  00415D34: mov ecx, [ebx]
  00415D36: call [ecx+24h]
  00415D39: cmp eax, esi
  00415D3B: fclex 
  00415D3D: jnl 415D4Eh
  00415D3F: push 00000024h
  00415D41: push 00404EF8h
  00415D46: push ebx
  00415D47: push eax
  00415D48: call MSVBVM60.DLL.__vbaHresultCheckObj
  00415D4E: mov eax, var_E8
  00415D54: lea edx, var_C4
  00415D5A: lea ecx, var_24
  00415D5D: mov var_BC, eax
  00415D63: mov var_C4, 00000003h
  00415D6D: call MSVBVM60.DLL.__vbaVarMove
  00415D73: lea ecx, var_70
  00415D76: call MSVBVM60.DLL.__vbaFreeObj
  00415D7C: lea ecx, var_94
  00415D82: lea edx, var_84
  00415D88: push ecx
  00415D89: mov ebx, 00000002h
  00415D8E: push edx
  00415D8F: push ebx
  00415D90: call MSVBVM60.DLL.__vbaFreeVarList
  00415D96: mov edx, arg_10
  00415D99: add esp, 0000000Ch
  00415D9C: lea ecx, var_B4
  00415DA2: call MSVBVM60.DLL.__vbaVarVargNofree
  00415DA8: mov edx, eax
  00415DAA: lea ecx, var_54
  00415DAD: call MSVBVM60.DLL.__vbaVarCopy
  00415DB3: mov eax, 00000001h
  00415DB8: lea ecx, var_24
  00415DBB: mov var_BC, eax
  00415DC1: mov var_AC, eax
  00415DC7: lea eax, var_C4
  00415DCD: lea edx, var_B4
  00415DD3: push eax
  00415DD4: push ecx
  00415DD5: lea eax, var_84
  00415DDB: push edx
  00415DDC: push eax
  00415DDD: mov var_C4, ebx
  00415DE3: mov var_B4, ebx
  00415DE9: mov var_CC, esi
  00415DEF: mov var_D4, ebx
  00415DF5: call MSVBVM60.DLL.__vbaVarSub
  00415DFB: lea ecx, var_D4
  00415E01: push eax
  00415E02: lea edx, var_11C
  00415E08: push ecx
  00415E09: lea eax, var_10C
  00415E0F: push edx
  00415E10: lea ecx, var_64
  00415E13: push eax
  00415E14: push ecx
  00415E15: call [004010B4h] ; For
  00415E1B: mov ebx, MSVBVM60.DLL.__vbaStrVarVal
  00415E21: cmp eax, esi
  00415E23: jz 0041611Ah
  00415E29: lea edx, var_54
  00415E2C: lea eax, var_6C
  00415E2F: push edx
  00415E30: push eax
  00415E31: call ebx
  00415E33: push eax
  00415E34: call [00401370h] ; Val(arg_1)
  00415E3A: fstp real8 ptr var_F0
  00415E40: lea ecx, var_44
  00415E43: push ecx
  00415E44: call MSVBVM60.DLL.__vbaStrVarCopy
  00415E4A: push 00404EF8h
  00415E4F: mov ecx, 00000008h
  00415E54: sub esp, 00000010h
  00415E57: mov var_84, ecx
  00415E5D: mov edx, esp
  00415E5F: mov var_7C, eax
  00415E62: push 00000001h
  00415E64: push 0000043Eh
  00415E69: mov [edx], ecx
  00415E6B: mov ecx, var_80
  00415E6E: mov [edx+04h], ecx
  00415E71: mov ecx, [edi+34h]
  00415E74: push ecx
  00415E75: mov [edx+08h], eax
  00415E78: mov eax, var_78
  00415E7B: mov [edx+0Ch], eax
  00415E7E: lea edx, var_94
  00415E84: push edx
  00415E85: call MSVBVM60.DLL.__vbaLateIdCallLd
  00415E8B: add esp, 00000020h
  00415E8E: push eax
  00415E8F: call MSVBVM60.DLL.__vbaCastObjVar
  00415E95: push eax
  00415E96: lea eax, var_70
  00415E99: push eax
  00415E9A: call [004010BCh] ; Set (object)
  00415EA0: lea ebx, var_74
  00415EA3: mov ecx, 0000000Ah
  00415EA8: push ebx
  00415EA9: mov var_C4, ecx
  00415EAF: sub esp, 00000010h
  00415EB2: mov edi, eax
  00415EB4: mov ebx, esp
  00415EB6: mov eax, 80020004h
  00415EBB: mov var_BC, eax
  00415EC1: sub esp, 00000010h
  00415EC4: mov [ebx], ecx
  00415EC6: mov ecx, var_C0
  00415ECC: mov edx, [edi]
  00415ECE: mov [ebx+04h], ecx
  00415ED1: mov ecx, esp
  00415ED3: push edi
  00415ED4: mov [ebx+08h], eax
  00415ED7: mov eax, var_B8
  00415EDD: mov [ebx+0Ch], eax
  00415EE0: mov eax, var_64
  00415EE3: mov [ecx], eax
  00415EE5: mov eax, var_60
  00415EE8: mov [ecx+04h], eax
  00415EEB: mov eax, var_5C
  00415EEE: mov [ecx+08h], eax
  00415EF1: mov eax, var_58
  00415EF4: mov [ecx+0Ch], eax
  00415EF7: call [edx+2Ch]
  00415EFA: cmp eax, esi
  00415EFC: fclex 
  00415EFE: jnl 415F0Fh
  00415F00: push 0000002Ch
  00415F02: push 00404EF8h
  00415F07: push edi
  00415F08: push eax
  00415F09: call MSVBVM60.DLL.__vbaHresultCheckObj
  00415F0F: mov ecx, var_74
  00415F12: push esi
  00415F13: push 00404FC8h ; "Value"
  00415F18: lea edx, var_A4
  00415F1E: push ecx
  00415F1F: push edx
  00415F20: call MSVBVM60.DLL.__vbaLateMemCallLd
  00415F26: mov ebx, MSVBVM60.DLL.__vbaStrVarVal
  00415F2C: add esp, 00000010h
  00415F2F: push eax
  00415F30: lea eax, var_68
  00415F33: push eax
  00415F34: call ebx
  00415F36: push eax
  00415F37: call [00401370h] ; Val(arg_1)
  00415F3D: call MSVBVM60.DLL.__vbaFpR8
  00415F43: fstp real8 ptr var_130
  00415F49: fld real8 ptr var_F0
  00415F4F: call MSVBVM60.DLL.__vbaFpR8
  00415F55: fcomp real8 ptr var_130
  00415F5B: fstsw ax
  00415F5D: test ah, 40h
  00415F60: jz 415F69h
  00415F62: mov eax, 00000001h
  00415F67: jmp 415F6Bh
  00415F69: xor eax, eax
  00415F6B: lea ecx, var_6C
  00415F6E: lea edx, var_68
  00415F71: push ecx
  00415F72: push edx
  00415F73: neg eax
  00415F75: push 00000002h
  00415F77: mov edi, eax
  00415F79: call MSVBVM60.DLL.__vbaFreeStrList
  00415F7F: lea eax, var_74
  00415F82: lea ecx, var_70
  00415F85: push eax
  00415F86: push ecx
  00415F87: push 00000002h
  00415F89: call MSVBVM60.DLL.__vbaFreeObjList
  00415F8F: lea edx, var_A4
  00415F95: lea eax, var_94
  00415F9B: push edx
  00415F9C: lea ecx, var_84
  00415FA2: push eax
  00415FA3: push ecx
  00415FA4: push 00000003h
  00415FA6: call MSVBVM60.DLL.__vbaFreeVarList
  00415FAC: add esp, 00000028h
  00415FAF: cmp di, si
  00415FB2: jnz 415FD4h
  00415FB4: lea edx, var_11C
  00415FBA: lea eax, var_10C
  00415FC0: push edx
  00415FC1: lea ecx, var_64
  00415FC4: push eax
  00415FC5: push ecx
  00415FC6: call [00401360h] ; Next
  00415FCC: mov edi, arg_8
  00415FCF: jmp 00415E21h
  00415FD4: lea edx, var_44
  00415FD7: mov var_CC, FFFFFFFFh
  00415FE1: push edx
  00415FE2: mov var_D4, 0000000Bh
  00415FEC: call MSVBVM60.DLL.__vbaStrVarCopy
  00415FF2: push 00404EF8h
  00415FF7: mov ecx, 00000008h
  00415FFC: sub esp, 00000010h
  00415FFF: mov var_84, ecx
  00416005: mov edx, esp
  00416007: mov var_7C, eax
  0041600A: push 00000001h
  0041600C: push 0000043Eh
  00416011: mov [edx], ecx
  00416013: mov ecx, var_80
  00416016: mov [edx+04h], ecx
  00416019: mov ecx, arg_8
  0041601C: mov [edx+08h], eax
  0041601F: mov eax, var_78
  00416022: mov [edx+0Ch], eax
  00416025: mov edx, [ecx+34h]
  00416028: lea eax, var_94
  0041602E: push edx
  0041602F: push eax
  00416030: call MSVBVM60.DLL.__vbaLateIdCallLd
  00416036: add esp, 00000020h
  00416039: push eax
  0041603A: call MSVBVM60.DLL.__vbaCastObjVar
  00416040: lea ecx, var_70
  00416043: push eax
  00416044: push ecx
  00416045: call [004010BCh] ; Set (object)
  0041604B: lea ebx, var_74
  0041604E: mov ecx, 0000000Ah
  00416053: push ebx
  00416054: mov var_C4, ecx
  0041605A: sub esp, 00000010h
  0041605D: mov edi, eax
  0041605F: mov ebx, esp
  00416061: mov eax, 80020004h
  00416066: mov var_BC, eax
  0041606C: sub esp, 00000010h
  0041606F: mov [ebx], ecx
  00416071: mov ecx, var_C0
  00416077: mov edx, [edi]
  00416079: mov [ebx+04h], ecx
  0041607C: mov ecx, esp
  0041607E: push edi
  0041607F: mov [ebx+08h], eax
  00416082: mov eax, var_B8
  00416088: mov [ebx+0Ch], eax
  0041608B: mov eax, var_64
  0041608E: mov [ecx], eax
  00416090: mov eax, var_60
  00416093: mov [ecx+04h], eax
  00416096: mov eax, var_5C
  00416099: mov [ecx+08h], eax
  0041609C: mov eax, var_58
  0041609F: mov [ecx+0Ch], eax
  004160A2: call [edx+2Ch]
  004160A5: cmp eax, esi
  004160A7: fclex 
  004160A9: jnl 4160BAh
  004160AB: push 0000002Ch
  004160AD: push 00404EF8h
  004160B2: push edi
  004160B3: push eax
  004160B4: call MSVBVM60.DLL.__vbaHresultCheckObj
  004160BA: mov edx, var_D4
  004160C0: mov eax, var_D0
  004160C6: sub esp, 00000010h
  004160C9: mov ecx, esp
  004160CB: push 004054B0h ; "Checked"
  004160D0: mov [ecx], edx
  004160D2: mov edx, var_CC
  004160D8: mov [ecx+04h], eax
  004160DB: mov eax, var_C8
  004160E1: mov [ecx+08h], edx
  004160E4: mov [ecx+0Ch], eax
  004160E7: mov ecx, var_74
  004160EA: push ecx
  004160EB: call MSVBVM60.DLL.__vbaLateMemSt
  004160F1: lea edx, var_74
  004160F4: lea eax, var_70
  004160F7: push edx
  004160F8: push eax
  004160F9: push 00000002h
  004160FB: call MSVBVM60.DLL.__vbaFreeObjList
  00416101: lea ecx, var_94
  00416107: lea edx, var_84
  0041610D: push ecx
  0041610E: push edx
  0041610F: push 00000002h
  00416111: call MSVBVM60.DLL.__vbaFreeVarList
  00416117: add esp, 00000018h
  0041611A: wait 
  0041611B: push 004161A6h
  00416120: jmp 416172h
  00416122: test byte ptr var_4, 04h
  00416126: jz 416131h
  00416128: lea ecx, var_34
  0041612B: call MSVBVM60.DLL.__vbaFreeVar
  00416131: lea eax, var_6C
  00416134: lea ecx, var_68
  00416137: push eax
  00416138: push ecx
  00416139: push 00000002h
  0041613B: call MSVBVM60.DLL.__vbaFreeStrList
  00416141: lea edx, var_74
  00416144: lea eax, var_70
  00416147: push edx
  00416148: push eax
  00416149: push 00000002h
  0041614B: call MSVBVM60.DLL.__vbaFreeObjList
  00416151: lea ecx, var_A4
  00416157: lea edx, var_94
  0041615D: push ecx
  0041615E: lea eax, var_84
  00416164: push edx
  00416165: push eax
  00416166: push 00000003h
  00416168: call MSVBVM60.DLL.__vbaFreeVarList
  0041616E: add esp, 00000028h
  00416171: ret 
End Function

Private sub Proc_3_10_40FB90
  0040FB90: push ebp
  0040FB91: mov ebp, esp
  0040FB93: sub esp, 00000008h
  0040FB96: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0040FB9B: mov eax, fs:[00h]
  0040FBA1: push eax
  0040FBA2: mov fs:[00000000h], esp
  0040FBA9: sub esp, 00000100h
  0040FBAF: push ebx
  0040FBB0: push esi
  0040FBB1: push edi
  0040FBB2: mov var_8, esp
  0040FBB5: mov var_4, 00401418h
  0040FBBC: mov eax, arg_8
  0040FBBF: xor esi, esi
  0040FBC1: push 00404ED8h
  0040FBC6: push esi
  0040FBC7: mov ecx, [eax]
  0040FBC9: push 000000CBh
  0040FBCE: push eax
  0040FBCF: mov var_20, esi
  0040FBD2: mov var_24, esi
  0040FBD5: mov var_34, esi
  0040FBD8: mov var_44, esi
  0040FBDB: mov var_48, esi
  0040FBDE: mov var_4C, esi
  0040FBE1: mov var_5C, esi
  0040FBE4: mov var_6C, esi
  0040FBE7: mov var_7C, esi
  0040FBEA: mov var_8C, esi
  0040FBF0: mov var_9C, esi
  0040FBF6: mov var_BC, esi
  0040FBFC: mov var_F8, esi
  0040FC02: mov var_FC, esi
  0040FC08: mov var_100, esi
  0040FC0E: mov var_104, esi
  0040FC14: lea edi, [eax+34h]
  0040FC17: call [ecx+00000344h]
  0040FC1D: mov ebx, [004010BCh] ; Set (object)
  0040FC23: lea edx, var_48
  0040FC26: push eax
  0040FC27: push edx
  0040FC28: call ebx
  0040FC2A: push eax
  0040FC2B: lea eax, var_5C
  0040FC2E: push eax
  0040FC2F: call MSVBVM60.DLL.__vbaLateIdCallLd
  0040FC35: add esp, 00000010h
  0040FC38: push eax
  0040FC39: call MSVBVM60.DLL.__vbaCastObjVar
  0040FC3F: lea ecx, var_4C
  0040FC42: push eax
  0040FC43: push ecx
  0040FC44: call ebx
  0040FC46: push eax
  0040FC47: push edi
  0040FC48: call MSVBVM60.DLL.__vbaObjSetAddref
  0040FC4E: lea edx, var_4C
  0040FC51: lea eax, var_48
  0040FC54: push edx
  0040FC55: push eax
  0040FC56: push 00000002h
  0040FC58: call MSVBVM60.DLL.__vbaFreeObjList
  0040FC5E: add esp, 0000000Ch
  0040FC61: lea ecx, var_5C
  0040FC64: call MSVBVM60.DLL.__vbaFreeVar
  0040FC6A: push 00404EF8h
  0040FC6F: mov ecx, 00000008h
  0040FC74: sub esp, 00000010h
  0040FC77: mov eax, 00404EECh ; "table"
  0040FC7C: mov edx, esp
  0040FC7E: push 00000001h
  0040FC80: push 0000043Fh
  0040FC85: mov [edx], ecx
  0040FC87: mov ecx, var_A8
  0040FC8D: mov [edx+04h], ecx
  0040FC90: mov ecx, [edi]
  0040FC92: push ecx
  0040FC93: mov [edx+08h], eax
  0040FC96: mov eax, var_A0
  0040FC9C: mov [edx+0Ch], eax
  0040FC9F: lea edx, var_5C
  0040FCA2: push edx
  0040FCA3: call MSVBVM60.DLL.__vbaLateIdCallLd
  0040FCA9: add esp, 00000020h
  0040FCAC: push eax
  0040FCAD: call MSVBVM60.DLL.__vbaCastObjVar
  0040FCB3: push eax
  0040FCB4: lea eax, var_48
  0040FCB7: push eax
  0040FCB8: call ebx
  0040FCBA: mov edx, [eax]
  0040FCBC: mov var_F0, eax
  0040FCC2: lea eax, var_4C
  0040FCC5: mov var_110, edx
  0040FCCB: push eax
  0040FCCC: mov ecx, 0000000Ah
  0040FCD1: sub esp, 00000010h
  0040FCD4: mov eax, 80020004h
  0040FCD9: mov edx, esp
  0040FCDB: sub esp, 00000010h
  0040FCDE: mov [edx], ecx
  0040FCE0: mov ecx, var_D8
  0040FCE6: mov [edx+04h], ecx
  0040FCE9: mov ecx, esp
  0040FCEB: mov [edx+08h], eax
  0040FCEE: mov eax, var_D0
  0040FCF4: mov [edx+0Ch], eax
  0040FCF7: mov edx, var_C8
  0040FCFD: mov eax, 00000002h
  0040FD02: mov [ecx], eax
  0040FD04: xor eax, eax
  0040FD06: mov [ecx+04h], edx
  0040FD09: mov edx, var_110
  0040FD0F: mov [ecx+08h], eax
  0040FD12: mov eax, var_C0
  0040FD18: mov [ecx+0Ch], eax
  0040FD1B: mov ecx, var_F0
  0040FD21: push ecx
  0040FD22: call [edx+2Ch]
  0040FD25: cmp eax, esi
  0040FD27: fclex 
  0040FD29: jnl 40FD40h
  0040FD2B: mov ecx, var_F0
  0040FD31: push 0000002Ch
  0040FD33: push 00404EF8h
  0040FD38: push ecx
  0040FD39: push eax
  0040FD3A: call MSVBVM60.DLL.__vbaHresultCheckObj
  0040FD40: mov edx, var_4C
  0040FD43: push 00404F08h
  0040FD48: push edx
  0040FD49: call MSVBVM60.DLL.__vbaCastObj
  0040FD4F: push eax
  0040FD50: lea eax, var_24
  0040FD53: push eax
  0040FD54: call ebx
  0040FD56: lea ecx, var_4C
  0040FD59: lea edx, var_48
  0040FD5C: push ecx
  0040FD5D: push edx
  0040FD5E: push 00000002h
  0040FD60: call MSVBVM60.DLL.__vbaFreeObjList
  0040FD66: add esp, 0000000Ch
  0040FD69: lea ecx, var_5C
  0040FD6C: call MSVBVM60.DLL.__vbaFreeVar
  0040FD72: push 00404EF8h
  0040FD77: mov ecx, 00000008h
  0040FD7C: sub esp, 00000010h
  0040FD7F: mov eax, 00404EECh ; "table"
  0040FD84: mov edx, esp
  0040FD86: push 00000001h
  0040FD88: push 0000043Fh
  0040FD8D: mov [edx], ecx
  0040FD8F: mov ecx, var_A8
  0040FD95: mov [edx+04h], ecx
  0040FD98: mov ecx, [edi]
  0040FD9A: push ecx
  0040FD9B: mov [edx+08h], eax
  0040FD9E: mov eax, var_A0
  0040FDA4: mov [edx+0Ch], eax
  0040FDA7: lea edx, var_5C
  0040FDAA: push edx
  0040FDAB: call MSVBVM60.DLL.__vbaLateIdCallLd
  0040FDB1: add esp, 00000020h
  0040FDB4: push eax
  0040FDB5: call MSVBVM60.DLL.__vbaCastObjVar
  0040FDBB: push eax
  0040FDBC: lea eax, var_20
  0040FDBF: push eax
  0040FDC0: call MSVBVM60.DLL.__vbaVarSetObj
  0040FDC6: lea ecx, var_5C
  0040FDC9: call MSVBVM60.DLL.__vbaFreeVar
  0040FDCF: lea ecx, var_20
  0040FDD2: lea edx, var_44
  0040FDD5: push ecx
  0040FDD6: lea eax, var_F8
  0040FDDC: push edx
  0040FDDD: lea ecx, var_100
  0040FDE3: push eax
  0040FDE4: lea edx, var_104
  0040FDEA: push ecx
  0040FDEB: lea eax, var_FC
  0040FDF1: push edx
  0040FDF2: push eax
  0040FDF3: call [00401324h] ; For Each
  0040FDF9: mov edi, MSVBVM60.DLL.__vbaVarLateMemCallLdRf
  0040FDFF: cmp eax, esi
  0040FE01: jz 40FE79h
  0040FE03: sub esp, 00000010h
  0040FE06: mov ecx, 00000008h
  0040FE0B: mov edx, esp
  0040FE0D: mov eax, 00404F1Ch ; "none"
  0040FE12: push 00404F34h ; "display"
  0040FE17: push esi
  0040FE18: mov [edx], ecx
  0040FE1A: mov ecx, var_A8
  0040FE20: push 00404F28h ; "Style"
  0040FE25: mov [edx+04h], ecx
  0040FE28: lea ecx, var_44
  0040FE2B: push ecx
  0040FE2C: mov [edx+08h], eax
  0040FE2F: mov eax, var_A0
  0040FE35: mov [edx+0Ch], eax
  0040FE38: lea edx, var_5C
  0040FE3B: push edx
  0040FE3C: call edi
  0040FE3E: add esp, 00000010h
  0040FE41: push eax
  0040FE42: call MSVBVM60.DLL.__vbaVarLateMemSt
  0040FE48: lea ecx, var_5C
  0040FE4B: call MSVBVM60.DLL.__vbaFreeVar
  0040FE51: lea eax, var_44
  0040FE54: lea ecx, var_F8
  0040FE5A: push eax
  0040FE5B: lea edx, var_100
  0040FE61: push ecx
  0040FE62: lea eax, var_104
  0040FE68: push edx
  0040FE69: lea ecx, var_FC
  0040FE6F: push eax
  0040FE70: push ecx
  0040FE71: call [00401048h] ; Next (For Each)
  0040FE77: jmp 40FDFFh
  0040FE79: mov edx, arg_8
  0040FE7C: push 00404F84h
  0040FE81: push esi
  0040FE82: push 000003FBh
  0040FE87: mov eax, [edx+34h]
  0040FE8A: lea ecx, var_5C
  0040FE8D: push eax
  0040FE8E: push ecx
  0040FE8F: mov var_BC, 00000008h
  0040FE99: call MSVBVM60.DLL.__vbaLateIdCallLd
  0040FE9F: add esp, 00000010h
  0040FEA2: push eax
  0040FEA3: call MSVBVM60.DLL.__vbaCastObjVar
  0040FEA9: lea edx, var_48
  0040FEAC: push eax
  0040FEAD: push edx
  0040FEAE: call ebx
  0040FEB0: mov ebx, eax
  0040FEB2: lea ecx, var_7C
  0040FEB5: lea edx, var_6C
  0040FEB8: mov var_64, esi
  0040FEBB: mov var_6C, 00000002h
  0040FEC2: mov eax, [ebx]
  0040FEC4: push ecx
  0040FEC5: push edx
  0040FEC6: push ebx
  0040FEC7: call [eax+1Ch]
  0040FECA: cmp eax, esi
  0040FECC: fclex 
  0040FECE: jnl 40FEDFh
  0040FED0: push 0000001Ch
  0040FED2: push 00404F84h
  0040FED7: push ebx
  0040FED8: push eax
  0040FED9: call MSVBVM60.DLL.__vbaHresultCheckObj
  0040FEDF: mov edx, var_D8
  0040FEE5: sub esp, 00000010h
  0040FEE8: mov ecx, esp
  0040FEEA: mov eax, 00000008h
  0040FEEF: push 00404FC8h ; "Value"
  0040FEF4: mov [ecx], eax
  0040FEF6: mov eax, 00404F48h
  0040FEFB: sub esp, 00000010h
  0040FEFE: mov [ecx+04h], edx
  0040FF01: mov edx, var_BC
  0040FF07: mov [ecx+08h], eax
  0040FF0A: mov eax, var_D0
  0040FF10: mov [ecx+0Ch], eax
  0040FF13: mov eax, var_B8
  0040FF19: mov ecx, esp
  0040FF1B: push 00000001h
  0040FF1D: push 00404FA8h ; "getElementById"
  0040FF22: push esi
  0040FF23: mov [ecx], edx
  0040FF25: mov edx, var_B0
  0040FF2B: push 00404F94h ; "Document"
  0040FF30: mov [ecx+04h], eax
  0040FF33: mov eax, 00404F50h ; "id_sc_field_kode_keluarga"
  0040FF38: mov [ecx+08h], eax
  0040FF3B: lea eax, var_7C
  0040FF3E: push eax
  0040FF3F: mov [ecx+0Ch], edx
  0040FF42: lea ecx, var_8C
  0040FF48: push ecx
  0040FF49: call edi
  0040FF4B: add esp, 00000010h
  0040FF4E: lea edx, var_9C
  0040FF54: push eax
  0040FF55: push edx
  0040FF56: call edi
  0040FF58: add esp, 00000020h
  0040FF5B: push eax
  0040FF5C: call MSVBVM60.DLL.__vbaVarLateMemSt
  0040FF62: lea ecx, var_48
  0040FF65: call MSVBVM60.DLL.__vbaFreeObj
  0040FF6B: lea eax, var_9C
  0040FF71: lea ecx, var_8C
  0040FF77: push eax
  0040FF78: lea edx, var_7C
  0040FF7B: push ecx
  0040FF7C: lea eax, var_6C
  0040FF7F: push edx
  0040FF80: lea ecx, var_5C
  0040FF83: push eax
  0040FF84: push ecx
  0040FF85: push 00000005h
  0040FF87: call MSVBVM60.DLL.__vbaFreeVarList
  0040FF8D: add esp, 00000018h
  0040FF90: push 00410003h ; "‹Mð_^3Àd‰'#1"
  0040FF95: jmp 40FFCDh
  0040FF97: lea edx, var_4C
  0040FF9A: lea eax, var_48
  0040FF9D: push edx
  0040FF9E: push eax
  0040FF9F: push 00000002h
  0040FFA1: call MSVBVM60.DLL.__vbaFreeObjList
  0040FFA7: lea ecx, var_9C
  0040FFAD: lea edx, var_8C
  0040FFB3: push ecx
  0040FFB4: lea eax, var_7C
  0040FFB7: push edx
  0040FFB8: lea ecx, var_6C
  0040FFBB: push eax
  0040FFBC: lea edx, var_5C
  0040FFBF: push ecx
  0040FFC0: push edx
  0040FFC1: push 00000005h
  0040FFC3: call MSVBVM60.DLL.__vbaFreeVarList
  0040FFC9: add esp, 00000024h
  0040FFCC: ret 
End Sub

Private sub Proc_3_11_410020
  00410020: push ebp
  00410021: mov ebp, esp
  00410023: sub esp, 00000014h
  00410026: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0041002B: mov eax, fs:[00h]
  00410031: push eax
  00410032: mov fs:[00000000h], esp
  00410039: sub esp, 000001E4h
  0041003F: push ebx
  00410040: push esi
  00410041: push edi
  00410042: mov var_14, esp
  00410045: mov var_10, 00401428h
  0041004C: xor esi, esi
  0041004E: mov var_C, esi
  00410051: mov var_8, esi
  00410054: mov var_2C, esi
  00410057: mov var_3C, esi
  0041005A: mov var_4C, esi
  0041005D: mov var_5C, esi
  00410060: mov var_6C, esi
  00410063: mov var_7C, esi
  00410066: mov var_8C, esi
  0041006C: mov var_9C, esi
  00410072: mov var_AC, esi
  00410078: mov var_B0, esi
  0041007E: mov var_B4, esi
  00410084: mov var_B8, esi
  0041008A: mov var_C8, esi
  00410090: mov var_D8, esi
  00410096: mov var_E8, esi
  0041009C: mov var_F8, esi
  004100A2: mov var_108, esi
  004100A8: mov var_118, esi
  004100AE: mov var_128, esi
  004100B4: mov var_138, esi
  004100BA: mov var_148, esi
  004100C0: mov var_158, esi
  004100C6: mov var_168, esi
  004100CC: mov var_178, esi
  004100D2: mov var_188, esi
  004100D8: mov var_1C4, esi
  004100DE: mov var_1C8, esi
  004100E4: mov var_1CC, esi
  004100EA: mov var_1D0, esi
  004100F0: push 00000001h
  004100F2: call On Error ...
  004100F8: push esi
  004100F9: push 0000000Ah
  004100FB: mov esi, arg_8
  004100FE: mov eax, [esi]
  00410100: push esi
  00410101: call [eax+00000340h]
  00410107: push eax
  00410108: lea ecx, var_B4
  0041010E: push ecx
  0041010F: mov edi, [004010BCh] ; Set (object)
  00410115: call edi
  00410117: push eax
  00410118: lea edx, var_C8
  0041011E: push edx
  0041011F: mov ebx, MSVBVM60.DLL.__vbaLateIdCallLd
  00410125: call ebx
  00410127: add esp, 00000010h
  0041012A: push eax
  0041012B: call MSVBVM60.DLL.__vbaI4Var
  00410131: mov var_140, eax
  00410137: mov var_148, 00000003h
  00410141: lea edx, var_148
  00410147: lea ecx, var_4C
  0041014A: call MSVBVM60.DLL.__vbaVarMove
  00410150: lea ecx, var_B4
  00410156: call MSVBVM60.DLL.__vbaFreeObj
  0041015C: lea ecx, var_C8
  00410162: call MSVBVM60.DLL.__vbaFreeVar
  00410168: lea eax, var_4C
  0041016B: push eax
  0041016C: call MSVBVM60.DLL.__vbaI4Var
  00410172: mov var_140, eax
  00410178: mov ecx, 00000003h
  0041017D: mov var_148, ecx
  00410183: mov var_160, 00000001h
  0041018D: mov var_168, ecx
  00410193: sub esp, 00000010h
  00410196: mov edx, esp
  00410198: mov [edx], ecx
  0041019A: mov ecx, var_144
  004101A0: mov [edx+04h], ecx
  004101A3: mov [edx+08h], eax
  004101A6: mov eax, var_13C
  004101AC: mov [edx+0Ch], eax
  004101AF: sub esp, 00000010h
  004101B2: mov ecx, esp
  004101B4: mov edx, var_168
  004101BA: mov [ecx], edx
  004101BC: mov eax, var_164
  004101C2: mov [ecx+04h], eax
  004101C5: mov edx, var_160
  004101CB: mov [ecx+08h], edx
  004101CE: mov eax, var_15C
  004101D4: mov [ecx+0Ch], eax
  004101D7: push 00000002h
  004101D9: push 00000041h
  004101DB: mov ecx, [esi]
  004101DD: push esi
  004101DE: call [ecx+00000340h]
  004101E4: push eax
  004101E5: lea edx, var_B8
  004101EB: push edx
  004101EC: call edi
  004101EE: push eax
  004101EF: lea eax, var_D8
  004101F5: push eax
  004101F6: call ebx
  004101F8: add esp, 00000030h
  004101FB: push eax
  004101FC: call MSVBVM60.DLL.__vbaStrVarMove
  00410202: mov edx, eax
  00410204: lea ecx, var_B0
  0041020A: call MSVBVM60.DLL.__vbaStrMove
  00410210: push eax
  00410211: push 00403558h
  00410216: call MSVBVM60.DLL.__vbaStrCmp
  0041021C: neg eax
  0041021E: sbb eax, eax
  00410220: inc eax
  00410221: neg eax
  00410223: mov var_1FC, ax
  0041022A: push 00000000h
  0041022C: push 0000000Ah
  0041022E: mov ecx, [esi]
  00410230: push esi
  00410231: call [ecx+00000340h]
  00410237: push eax
  00410238: lea edx, var_B4
  0041023E: push edx
  0041023F: call edi
  00410241: push eax
  00410242: lea eax, var_C8
  00410248: push eax
  00410249: call ebx
  0041024B: add esp, 00000010h
  0041024E: push eax
  0041024F: call MSVBVM60.DLL.__vbaI4Var
  00410255: xor ecx, ecx
  00410257: test eax, eax
  00410259: setle cl
  0041025C: neg ecx
  0041025E: mov edx, var_1FC
  00410264: or edx, ecx
  00410266: mov var_1BC, dx
  0041026D: lea ecx, var_B0
  00410273: call MSVBVM60.DLL.__vbaFreeStr
  00410279: lea eax, var_B8
  0041027F: push eax
  00410280: lea ecx, var_B4
  00410286: push ecx
  00410287: push 00000002h
  00410289: call MSVBVM60.DLL.__vbaFreeObjList
  0041028F: lea edx, var_D8
  00410295: push edx
  00410296: lea eax, var_C8
  0041029C: push eax
  0041029D: push 00000002h
  0041029F: call MSVBVM60.DLL.__vbaFreeVarList
  004102A5: add esp, 00000018h
  004102A8: cmp word ptr var_1BC, 0000h
  004102B0: jz 00410365h
  004102B6: mov eax, 80020004h
  004102BB: mov var_F0, eax
  004102C1: mov ecx, 0000000Ah
  004102C6: mov var_F8, ecx
  004102CC: mov var_E0, eax
  004102D2: mov var_E8, ecx
  004102D8: mov var_D0, eax
  004102DE: mov var_D8, ecx
  004102E4: mov var_140, 00404FD8h ; "Pilih data terlebih dahulu"
  004102EE: mov var_148, 00000008h
  004102F8: lea edx, var_148
  004102FE: lea ecx, var_C8
  00410304: call MSVBVM60.DLL.__vbaVarDup
  0041030A: lea ecx, var_F8
  00410310: push ecx
  00410311: lea edx, var_E8
  00410317: push edx
  00410318: lea eax, var_D8
  0041031E: push eax
  0041031F: push 00000000h
  00410321: lea ecx, var_C8
  00410327: push ecx
  00410328: call [004010C0h] ; MsgBox(arg_1, arg_2, arg_3, arg_4, arg_5)
  0041032E: lea edx, var_F8
  00410334: push edx
  00410335: lea eax, var_E8
  0041033B: push eax
  0041033C: lea ecx, var_D8
  00410342: push ecx
  00410343: lea edx, var_C8
  00410349: push edx
  0041034A: push 00000004h
  0041034C: call MSVBVM60.DLL.__vbaFreeVarList
  00410352: add esp, 00000014h
  00410355: call MSVBVM60.DLL.__vbaExitProc
  0041035B: Method_8964E44D
  00410360: jmp 0041270Bh
  00410365: mov eax, arg_C
  00410368: cmp word ptr [eax], 0000h
  0041036C: jnz 004110EAh
  00410372: mov var_140, 00405064h ; "id_sc_field_kode_keluarga,id_sc_field_namakk,id_sc_field_alamat,id_sc_field_rt,id_sc_field_rw,id_sc_field_nama_dusun"
  0041037C: mov var_148, 00000008h
  00410386: lea edx, var_148
  0041038C: lea ecx, var_5C
  0041038F: call MSVBVM60.DLL.__vbaVarCopy
  00410395: mov var_140, 00405154h ; ",id_sc_field_d014,id_sc_field_d016,id_sc_field_d015,id_sc_field_d017"
  0041039F: mov var_148, 00000008h
  004103A9: lea ecx, var_5C
  004103AC: push ecx
  004103AD: lea edx, var_148
  004103B3: push edx
  004103B4: lea eax, var_C8
  004103BA: push eax
  004103BB: call MSVBVM60.DLL.__vbaVarCat
  004103C1: mov edx, eax
  004103C3: lea ecx, var_5C
  004103C6: call MSVBVM60.DLL.__vbaVarMove
  004103CC: mov var_140, 00403E1Ch
  004103D6: mov var_148, 00000008h
  004103E0: lea edx, var_148
  004103E6: lea ecx, var_C8
  004103EC: call MSVBVM60.DLL.__vbaVarDup
  004103F2: push 00000000h
  004103F4: push FFFFFFFFh
  004103F6: lea ecx, var_C8
  004103FC: push ecx
  004103FD: lea edx, var_5C
  00410400: push edx
  00410401: lea eax, var_B0
  00410407: push eax
  00410408: call MSVBVM60.DLL.__vbaStrVarVal
  0041040E: push eax
  0041040F: lea ecx, var_D8
  00410415: push ecx
  00410416: call [004011E8h] ; arg_1 = Split(arg_2, arg_3, arg_4, arg_5)
  0041041C: lea edx, var_D8
  00410422: lea ecx, var_3C
  00410425: call MSVBVM60.DLL.__vbaVarMove
  0041042B: lea ecx, var_B0
  00410431: call MSVBVM60.DLL.__vbaFreeStr
  00410437: lea ecx, var_C8
  0041043D: call MSVBVM60.DLL.__vbaFreeVar
  00410443: lea edx, var_4C
  00410446: push edx
  00410447: call MSVBVM60.DLL.__vbaI4Var
  0041044D: mov var_140, eax
  00410453: mov ecx, 00000003h
  00410458: mov var_148, ecx
  0041045E: mov var_160, 00000001h
  00410468: mov var_168, ecx
  0041046E: sub esp, 00000010h
  00410471: mov edx, esp
  00410473: mov [edx], ecx
  00410475: mov ecx, var_144
  0041047B: mov [edx+04h], ecx
  0041047E: mov [edx+08h], eax
  00410481: mov eax, var_13C
  00410487: mov [edx+0Ch], eax
  0041048A: sub esp, 00000010h
  0041048D: mov ecx, esp
  0041048F: mov edx, var_168
  00410495: mov [ecx], edx
  00410497: mov eax, var_164
  0041049D: mov [ecx+04h], eax
  004104A0: mov edx, var_160
  004104A6: mov [ecx+08h], edx
  004104A9: mov eax, var_15C
  004104AF: mov [ecx+0Ch], eax
  004104B2: push 00000002h
  004104B4: push 00000041h
  004104B6: mov ecx, [esi]
  004104B8: push esi
  004104B9: call [ecx+00000340h]
  004104BF: push eax
  004104C0: lea edx, var_B4
  004104C6: push edx
  004104C7: call edi
  004104C9: push eax
  004104CA: lea eax, var_C8
  004104D0: push eax
  004104D1: call ebx
  004104D3: add esp, 00000030h
  004104D6: push eax
  004104D7: call MSVBVM60.DLL.__vbaStrVarMove
  004104DD: mov var_D0, eax
  004104E3: mov var_D8, 00000008h
  004104ED: lea edx, var_D8
  004104F3: lea ecx, var_9C
  004104F9: call MSVBVM60.DLL.__vbaVarMove
  004104FF: lea ecx, var_B4
  00410505: call MSVBVM60.DLL.__vbaFreeObj
  0041050B: lea ecx, var_C8
  00410511: call MSVBVM60.DLL.__vbaFreeVar
  00410517: mov var_140, 004051E4h
  00410521: mov var_148, 00000008h
  0041052B: lea ecx, var_4C
  0041052E: push ecx
  0041052F: call MSVBVM60.DLL.__vbaI4Var
  00410535: mov var_150, eax
  0041053B: mov ecx, 00000003h
  00410540: mov var_158, ecx
  00410546: mov var_170, ecx
  0041054C: mov var_178, ecx
  00410552: sub esp, 00000010h
  00410555: mov edx, esp
  00410557: mov [edx], ecx
  00410559: mov ecx, var_154
  0041055F: mov [edx+04h], ecx
  00410562: mov [edx+08h], eax
  00410565: mov eax, var_14C
  0041056B: mov [edx+0Ch], eax
  0041056E: sub esp, 00000010h
  00410571: mov ecx, esp
  00410573: mov edx, var_178
  00410579: mov [ecx], edx
  0041057B: mov eax, var_174
  00410581: mov [ecx+04h], eax
  00410584: mov edx, var_170
  0041058A: mov [ecx+08h], edx
  0041058D: mov eax, var_16C
  00410593: mov [ecx+0Ch], eax
  00410596: push 00000002h
  00410598: push 00000041h
  0041059A: mov ecx, [esi]
  0041059C: push esi
  0041059D: call [ecx+00000340h]
  004105A3: push eax
  004105A4: lea edx, var_B4
  004105AA: push edx
  004105AB: call edi
  004105AD: push eax
  004105AE: lea eax, var_D8
  004105B4: push eax
  004105B5: call ebx
  004105B7: add esp, 00000030h
  004105BA: push eax
  004105BB: call MSVBVM60.DLL.__vbaStrVarMove
  004105C1: mov var_E0, eax
  004105C7: mov var_E8, 00000008h
  004105D1: lea ecx, var_9C
  004105D7: push ecx
  004105D8: lea edx, var_148
  004105DE: push edx
  004105DF: lea eax, var_C8
  004105E5: push eax
  004105E6: call MSVBVM60.DLL.__vbaVarCat
  004105EC: push eax
  004105ED: lea ecx, var_E8
  004105F3: push ecx
  004105F4: lea edx, var_F8
  004105FA: push edx
  004105FB: call MSVBVM60.DLL.__vbaVarCat
  00410601: mov edx, eax
  00410603: lea ecx, var_9C
  00410609: call MSVBVM60.DLL.__vbaVarMove
  0041060F: lea ecx, var_B4
  00410615: call MSVBVM60.DLL.__vbaFreeObj
  0041061B: lea eax, var_E8
  00410621: push eax
  00410622: lea ecx, var_C8
  00410628: push ecx
  00410629: lea edx, var_D8
  0041062F: push edx
  00410630: push 00000003h
  00410632: call MSVBVM60.DLL.__vbaFreeVarList
  00410638: add esp, 00000010h
  0041063B: mov var_140, 004051E4h
  00410645: mov var_148, 00000008h
  0041064F: lea eax, var_4C
  00410652: push eax
  00410653: call MSVBVM60.DLL.__vbaI4Var
  00410659: mov var_150, eax
  0041065F: mov ecx, 00000003h
  00410664: mov var_158, ecx
  0041066A: mov var_170, 00000012h
  00410674: mov var_178, ecx
  0041067A: sub esp, 00000010h
  0041067D: mov edx, esp
  0041067F: mov [edx], ecx
  00410681: mov ecx, var_154
  00410687: mov [edx+04h], ecx
  0041068A: mov [edx+08h], eax
  0041068D: mov eax, var_14C
  00410693: mov [edx+0Ch], eax
  00410696: sub esp, 00000010h
  00410699: mov ecx, esp
  0041069B: mov edx, var_178
  004106A1: mov [ecx], edx
  004106A3: mov eax, var_174
  004106A9: mov [ecx+04h], eax
  004106AC: mov edx, var_170
  004106B2: mov [ecx+08h], edx
  004106B5: mov eax, var_16C
  004106BB: mov [ecx+0Ch], eax
  004106BE: push 00000002h
  004106C0: push 00000041h
  004106C2: mov ecx, [esi]
  004106C4: push esi
  004106C5: call [ecx+00000340h]
  004106CB: push eax
  004106CC: lea edx, var_B4
  004106D2: push edx
  004106D3: call edi
  004106D5: push eax
  004106D6: lea eax, var_D8
  004106DC: push eax
  004106DD: call ebx
  004106DF: add esp, 00000030h
  004106E2: push eax
  004106E3: call MSVBVM60.DLL.__vbaStrVarMove
  004106E9: mov var_E0, eax
  004106EF: mov var_E8, 00000008h
  004106F9: lea ecx, var_9C
  004106FF: push ecx
  00410700: lea edx, var_148
  00410706: push edx
  00410707: lea eax, var_C8
  0041070D: push eax
  0041070E: call MSVBVM60.DLL.__vbaVarCat
  00410714: push eax
  00410715: lea ecx, var_E8
  0041071B: push ecx
  0041071C: lea edx, var_F8
  00410722: push edx
  00410723: call MSVBVM60.DLL.__vbaVarCat
  00410729: mov edx, eax
  0041072B: lea ecx, var_9C
  00410731: call MSVBVM60.DLL.__vbaVarMove
  00410737: lea ecx, var_B4
  0041073D: call MSVBVM60.DLL.__vbaFreeObj
  00410743: lea eax, var_E8
  00410749: push eax
  0041074A: lea ecx, var_C8
  00410750: push ecx
  00410751: lea edx, var_D8
  00410757: push edx
  00410758: push 00000003h
  0041075A: call MSVBVM60.DLL.__vbaFreeVarList
  00410760: add esp, 00000010h
  00410763: mov var_140, 004051E4h
  0041076D: mov var_148, 00000008h
  00410777: lea eax, var_4C
  0041077A: push eax
  0041077B: call MSVBVM60.DLL.__vbaI4Var
  00410781: mov var_150, eax
  00410787: mov ecx, 00000003h
  0041078C: mov var_158, ecx
  00410792: mov var_170, 00000010h
  0041079C: mov var_178, ecx
  004107A2: sub esp, 00000010h
  004107A5: mov edx, esp
  004107A7: mov [edx], ecx
  004107A9: mov ecx, var_154
  004107AF: mov [edx+04h], ecx
  004107B2: mov [edx+08h], eax
  004107B5: mov eax, var_14C
  004107BB: mov [edx+0Ch], eax
  004107BE: sub esp, 00000010h
  004107C1: mov ecx, esp
  004107C3: mov edx, var_178
  004107C9: mov [ecx], edx
  004107CB: mov eax, var_174
  004107D1: mov [ecx+04h], eax
  004107D4: mov edx, var_170
  004107DA: mov [ecx+08h], edx
  004107DD: mov eax, var_16C
  004107E3: mov [ecx+0Ch], eax
  004107E6: push 00000002h
  004107E8: push 00000041h
  004107EA: mov ecx, [esi]
  004107EC: push esi
  004107ED: call [ecx+00000340h]
  004107F3: push eax
  004107F4: lea edx, var_B4
  004107FA: push edx
  004107FB: call edi
  004107FD: push eax
  004107FE: lea eax, var_D8
  00410804: push eax
  00410805: call ebx
  00410807: add esp, 00000030h
  0041080A: push eax
  0041080B: call MSVBVM60.DLL.__vbaStrVarMove
  00410811: mov var_E0, eax
  00410817: mov var_E8, 00000008h
  00410821: lea ecx, var_9C
  00410827: push ecx
  00410828: lea edx, var_148
  0041082E: push edx
  0041082F: lea eax, var_C8
  00410835: push eax
  00410836: call MSVBVM60.DLL.__vbaVarCat
  0041083C: push eax
  0041083D: lea ecx, var_E8
  00410843: push ecx
  00410844: lea edx, var_F8
  0041084A: push edx
  0041084B: call MSVBVM60.DLL.__vbaVarCat
  00410851: mov edx, eax
  00410853: lea ecx, var_9C
  00410859: call MSVBVM60.DLL.__vbaVarMove
  0041085F: lea ecx, var_B4
  00410865: call MSVBVM60.DLL.__vbaFreeObj
  0041086B: lea eax, var_E8
  00410871: push eax
  00410872: lea ecx, var_C8
  00410878: push ecx
  00410879: lea edx, var_D8
  0041087F: push edx
  00410880: push 00000003h
  00410882: call MSVBVM60.DLL.__vbaFreeVarList
  00410888: add esp, 00000010h
  0041088B: mov var_140, 004051E4h
  00410895: mov var_148, 00000008h
  0041089F: lea eax, var_4C
  004108A2: push eax
  004108A3: call MSVBVM60.DLL.__vbaI4Var
  004108A9: mov var_150, eax
  004108AF: mov ecx, 00000003h
  004108B4: mov var_158, ecx
  004108BA: mov var_170, 00000011h
  004108C4: mov var_178, ecx
  004108CA: sub esp, 00000010h
  004108CD: mov edx, esp
  004108CF: mov [edx], ecx
  004108D1: mov ecx, var_154
  004108D7: mov [edx+04h], ecx
  004108DA: mov [edx+08h], eax
  004108DD: mov eax, var_14C
  004108E3: mov [edx+0Ch], eax
  004108E6: sub esp, 00000010h
  004108E9: mov ecx, esp
  004108EB: mov edx, var_178
  004108F1: mov [ecx], edx
  004108F3: mov eax, var_174
  004108F9: mov [ecx+04h], eax
  004108FC: mov edx, var_170
  00410902: mov [ecx+08h], edx
  00410905: mov eax, var_16C
  0041090B: mov [ecx+0Ch], eax
  0041090E: push 00000002h
  00410910: push 00000041h
  00410912: mov ecx, [esi]
  00410914: push esi
  00410915: call [ecx+00000340h]
  0041091B: push eax
  0041091C: lea edx, var_B4
  00410922: push edx
  00410923: call edi
  00410925: push eax
  00410926: lea eax, var_D8
  0041092C: push eax
  0041092D: call ebx
  0041092F: add esp, 00000030h
  00410932: push eax
  00410933: call MSVBVM60.DLL.__vbaStrVarMove
  00410939: mov var_E0, eax
  0041093F: mov var_E8, 00000008h
  00410949: lea ecx, var_9C
  0041094F: push ecx
  00410950: lea edx, var_148
  00410956: push edx
  00410957: lea eax, var_C8
  0041095D: push eax
  0041095E: call MSVBVM60.DLL.__vbaVarCat
  00410964: push eax
  00410965: lea ecx, var_E8
  0041096B: push ecx
  0041096C: lea edx, var_F8
  00410972: push edx
  00410973: call MSVBVM60.DLL.__vbaVarCat
  00410979: mov edx, eax
  0041097B: lea ecx, var_9C
  00410981: call MSVBVM60.DLL.__vbaVarMove
  00410987: lea ecx, var_B4
  0041098D: call MSVBVM60.DLL.__vbaFreeObj
  00410993: lea eax, var_E8
  00410999: push eax
  0041099A: lea ecx, var_C8
  004109A0: push ecx
  004109A1: lea edx, var_D8
  004109A7: push edx
  004109A8: push 00000003h
  004109AA: call MSVBVM60.DLL.__vbaFreeVarList
  004109B0: add esp, 00000010h
  004109B3: mov var_140, 004051E4h
  004109BD: mov var_148, 00000008h
  004109C7: mov eax, [esi]
  004109C9: push esi
  004109CA: call [eax+00000318h]
  004109D0: push eax
  004109D1: lea ecx, var_B4
  004109D7: push ecx
  004109D8: call edi
  004109DA: mov var_1BC, eax
  004109E0: mov edx, [eax]
  004109E2: lea ecx, var_B0
  004109E8: push ecx
  004109E9: push eax
  004109EA: call [edx+000000A0h]
  004109F0: fclex 
  004109F2: test eax, eax
  004109F4: jnl 410A0Eh
  004109F6: push 000000A0h
  004109FB: push 004051E8h
  00410A00: mov edx, var_1BC
  00410A06: push edx
  00410A07: push eax
  00410A08: call MSVBVM60.DLL.__vbaHresultCheckObj
  00410A0E: mov eax, var_B0
  00410A14: mov var_B0, 00000000h
  00410A1E: mov var_D0, eax
  00410A24: mov var_D8, 00000008h
  00410A2E: lea eax, var_9C
  00410A34: push eax
  00410A35: lea ecx, var_148
  00410A3B: push ecx
  00410A3C: lea edx, var_C8
  00410A42: push edx
  00410A43: call MSVBVM60.DLL.__vbaVarCat
  00410A49: push eax
  00410A4A: lea eax, var_D8
  00410A50: push eax
  00410A51: lea ecx, var_E8
  00410A57: push ecx
  00410A58: call MSVBVM60.DLL.__vbaVarCat
  00410A5E: mov edx, eax
  00410A60: lea ecx, var_9C
  00410A66: call MSVBVM60.DLL.__vbaVarMove
  00410A6C: lea ecx, var_B4
  00410A72: call MSVBVM60.DLL.__vbaFreeObj
  00410A78: lea edx, var_D8
  00410A7E: push edx
  00410A7F: lea eax, var_C8
  00410A85: push eax
  00410A86: push 00000002h
  00410A88: call MSVBVM60.DLL.__vbaFreeVarList
  00410A8E: add esp, 0000000Ch
  00410A91: mov var_140, 004051E4h
  00410A9B: mov var_148, 00000008h
  00410AA5: mov ecx, [esi]
  00410AA7: push esi
  00410AA8: call [ecx+00000328h]
  00410AAE: push eax
  00410AAF: lea edx, var_B4
  00410AB5: push edx
  00410AB6: call edi
  00410AB8: mov var_1BC, eax
  00410ABE: mov ecx, [eax]
  00410AC0: lea edx, var_B0
  00410AC6: push edx
  00410AC7: push eax
  00410AC8: call [ecx+000000A0h]
  00410ACE: fclex 
  00410AD0: test eax, eax
  00410AD2: jnl 410AECh
  00410AD4: push 000000A0h
  00410AD9: push 004051E8h
  00410ADE: mov ecx, var_1BC
  00410AE4: push ecx
  00410AE5: push eax
  00410AE6: call MSVBVM60.DLL.__vbaHresultCheckObj
  00410AEC: mov eax, var_B0
  00410AF2: mov var_B0, 00000000h
  00410AFC: mov var_D0, eax
  00410B02: mov var_D8, 00000008h
  00410B0C: lea edx, var_9C
  00410B12: push edx
  00410B13: lea eax, var_148
  00410B19: push eax
  00410B1A: lea ecx, var_C8
  00410B20: push ecx
  00410B21: call MSVBVM60.DLL.__vbaVarCat
  00410B27: push eax
  00410B28: lea edx, var_D8
  00410B2E: push edx
  00410B2F: lea eax, var_E8
  00410B35: push eax
  00410B36: call MSVBVM60.DLL.__vbaVarCat
  00410B3C: mov edx, eax
  00410B3E: lea ecx, var_9C
  00410B44: call MSVBVM60.DLL.__vbaVarMove
  00410B4A: lea ecx, var_B4
  00410B50: call MSVBVM60.DLL.__vbaFreeObj
  00410B56: lea ecx, var_D8
  00410B5C: push ecx
  00410B5D: lea edx, var_C8
  00410B63: push edx
  00410B64: push 00000002h
  00410B66: call MSVBVM60.DLL.__vbaFreeVarList
  00410B6C: add esp, 0000000Ch
  00410B6F: mov var_140, 004051E4h
  00410B79: mov var_148, 00000008h
  00410B83: mov eax, [esi]
  00410B85: push esi
  00410B86: call [eax+00000324h]
  00410B8C: push eax
  00410B8D: lea ecx, var_B4
  00410B93: push ecx
  00410B94: call edi
  00410B96: mov var_1BC, eax
  00410B9C: mov edx, [eax]
  00410B9E: lea ecx, var_B0
  00410BA4: push ecx
  00410BA5: push eax
  00410BA6: call [edx+000000A0h]
  00410BAC: fclex 
  00410BAE: test eax, eax
  00410BB0: jnl 410BCAh
  00410BB2: push 000000A0h
  00410BB7: push 004051E8h
  00410BBC: mov edx, var_1BC
  00410BC2: push edx
  00410BC3: push eax
  00410BC4: call MSVBVM60.DLL.__vbaHresultCheckObj
  00410BCA: mov eax, var_B0
  00410BD0: mov var_B0, 00000000h
  00410BDA: mov var_D0, eax
  00410BE0: mov var_D8, 00000008h
  00410BEA: lea eax, var_9C
  00410BF0: push eax
  00410BF1: lea ecx, var_148
  00410BF7: push ecx
  00410BF8: lea edx, var_C8
  00410BFE: push edx
  00410BFF: call MSVBVM60.DLL.__vbaVarCat
  00410C05: push eax
  00410C06: lea eax, var_D8
  00410C0C: push eax
  00410C0D: lea ecx, var_E8
  00410C13: push ecx
  00410C14: call MSVBVM60.DLL.__vbaVarCat
  00410C1A: mov edx, eax
  00410C1C: lea ecx, var_9C
  00410C22: call MSVBVM60.DLL.__vbaVarMove
  00410C28: lea ecx, var_B4
  00410C2E: call MSVBVM60.DLL.__vbaFreeObj
  00410C34: lea edx, var_D8
  00410C3A: push edx
  00410C3B: lea eax, var_C8
  00410C41: push eax
  00410C42: push 00000002h
  00410C44: call MSVBVM60.DLL.__vbaFreeVarList
  00410C4A: add esp, 0000000Ch
  00410C4D: mov var_140, 004051E4h
  00410C57: mov var_148, 00000008h
  00410C61: mov ecx, [esi]
  00410C63: push esi
  00410C64: call [ecx+00000320h]
  00410C6A: push eax
  00410C6B: lea edx, var_B4
  00410C71: push edx
  00410C72: call edi
  00410C74: mov var_1BC, eax
  00410C7A: mov ecx, [eax]
  00410C7C: lea edx, var_B0
  00410C82: push edx
  00410C83: push eax
  00410C84: call [ecx+000000A0h]
  00410C8A: fclex 
  00410C8C: test eax, eax
  00410C8E: jnl 410CA8h
  00410C90: push 000000A0h
  00410C95: push 004051E8h
  00410C9A: mov ecx, var_1BC
  00410CA0: push ecx
  00410CA1: push eax
  00410CA2: call MSVBVM60.DLL.__vbaHresultCheckObj
  00410CA8: mov eax, var_B0
  00410CAE: mov var_B0, 00000000h
  00410CB8: mov var_D0, eax
  00410CBE: mov var_D8, 00000008h
  00410CC8: lea edx, var_9C
  00410CCE: push edx
  00410CCF: lea eax, var_148
  00410CD5: push eax
  00410CD6: lea ecx, var_C8
  00410CDC: push ecx
  00410CDD: call MSVBVM60.DLL.__vbaVarCat
  00410CE3: push eax
  00410CE4: lea edx, var_D8
  00410CEA: push edx
  00410CEB: lea eax, var_E8
  00410CF1: push eax
  00410CF2: call MSVBVM60.DLL.__vbaVarCat
  00410CF8: mov edx, eax
  00410CFA: lea ecx, var_9C
  00410D00: call MSVBVM60.DLL.__vbaVarMove
  00410D06: lea ecx, var_B4
  00410D0C: call MSVBVM60.DLL.__vbaFreeObj
  00410D12: lea ecx, var_D8
  00410D18: push ecx
  00410D19: lea edx, var_C8
  00410D1F: push edx
  00410D20: push 00000002h
  00410D22: call MSVBVM60.DLL.__vbaFreeVarList
  00410D28: add esp, 0000000Ch
  00410D2B: mov var_140, 004051E4h
  00410D35: mov var_148, 00000008h
  00410D3F: mov eax, [esi]
  00410D41: push esi
  00410D42: call [eax+0000031Ch]
  00410D48: push eax
  00410D49: lea ecx, var_B4
  00410D4F: push ecx
  00410D50: call edi
  00410D52: mov var_1BC, eax
  00410D58: mov edx, [eax]
  00410D5A: lea ecx, var_B0
  00410D60: push ecx
  00410D61: push eax
  00410D62: call [edx+000000A0h]
  00410D68: fclex 
  00410D6A: test eax, eax
  00410D6C: jnl 410D86h
  00410D6E: push 000000A0h
  00410D73: push 004051E8h
  00410D78: mov edx, var_1BC
  00410D7E: push edx
  00410D7F: push eax
  00410D80: call MSVBVM60.DLL.__vbaHresultCheckObj
  00410D86: mov eax, var_B0
  00410D8C: mov var_B0, 00000000h
  00410D96: mov var_D0, eax
  00410D9C: mov var_D8, 00000008h
  00410DA6: lea eax, var_9C
  00410DAC: push eax
  00410DAD: lea ecx, var_148
  00410DB3: push ecx
  00410DB4: lea edx, var_C8
  00410DBA: push edx
  00410DBB: call MSVBVM60.DLL.__vbaVarCat
  00410DC1: push eax
  00410DC2: lea eax, var_D8
  00410DC8: push eax
  00410DC9: lea ecx, var_E8
  00410DCF: push ecx
  00410DD0: call MSVBVM60.DLL.__vbaVarCat
  00410DD6: mov edx, eax
  00410DD8: lea ecx, var_9C
  00410DDE: call MSVBVM60.DLL.__vbaVarMove
  00410DE4: lea ecx, var_B4
  00410DEA: call MSVBVM60.DLL.__vbaFreeObj
  00410DF0: lea edx, var_D8
  00410DF6: push edx
  00410DF7: lea eax, var_C8
  00410DFD: push eax
  00410DFE: push 00000002h
  00410E00: call MSVBVM60.DLL.__vbaFreeVarList
  00410E06: add esp, 0000000Ch
  00410E09: lea ecx, [esi+58h]
  00410E0C: lea edx, var_9C
  00410E12: call MSVBVM60.DLL.__vbaVarCopy
  00410E18: lea edx, [esi+58h]
  00410E1B: lea ecx, var_8C
  00410E21: call MSVBVM60.DLL.__vbaVarCopy
  00410E27: mov var_140, 004051E4h
  00410E31: mov var_148, 00000008h
  00410E3B: lea edx, var_148
  00410E41: lea ecx, var_C8
  00410E47: call MSVBVM60.DLL.__vbaVarDup
  00410E4D: push 00000000h
  00410E4F: push FFFFFFFFh
  00410E51: lea ecx, var_C8
  00410E57: push ecx
  00410E58: lea edx, var_8C
  00410E5E: push edx
  00410E5F: lea eax, var_B0
  00410E65: push eax
  00410E66: call MSVBVM60.DLL.__vbaStrVarVal
  00410E6C: push eax
  00410E6D: lea ecx, var_D8
  00410E73: push ecx
  00410E74: call [004011E8h] ; arg_1 = Split(arg_2, arg_3, arg_4, arg_5)
  00410E7A: lea edx, var_D8
  00410E80: lea ecx, var_AC
  00410E86: call MSVBVM60.DLL.__vbaVarMove
  00410E8C: lea ecx, var_B0
  00410E92: call MSVBVM60.DLL.__vbaFreeStr
  00410E98: lea ecx, var_C8
  00410E9E: call MSVBVM60.DLL.__vbaFreeVar
  00410EA4: lea edx, [esi+34h]
  00410EA7: mov var_200, edx
  00410EAD: push 00404ED8h
  00410EB2: push 00000000h
  00410EB4: push 000000CBh
  00410EB9: mov eax, [esi]
  00410EBB: push esi
  00410EBC: call [eax+00000344h]
  00410EC2: push eax
  00410EC3: lea ecx, var_B4
  00410EC9: push ecx
  00410ECA: call edi
  00410ECC: push eax
  00410ECD: lea edx, var_C8
  00410ED3: push edx
  00410ED4: call ebx
  00410ED6: add esp, 00000010h
  00410ED9: push eax
  00410EDA: mov esi, [00401188h]
  00410EE0: call MSVBVM60.DLL.__vbaCastObjVar
  00410EE2: push eax
  00410EE3: lea eax, var_B8
  00410EE9: push eax
  00410EEA: call edi
  00410EEC: push eax
  00410EED: mov ecx, var_200
  00410EF3: push ecx
  00410EF4: call MSVBVM60.DLL.__vbaObjSetAddref
  00410EFA: lea edx, var_B8
  00410F00: push edx
  00410F01: lea eax, var_B4
  00410F07: push eax
  00410F08: push 00000002h
  00410F0A: call MSVBVM60.DLL.__vbaFreeObjList
  00410F10: add esp, 0000000Ch
  00410F13: lea ecx, var_C8
  00410F19: call MSVBVM60.DLL.__vbaFreeVar
  00410F1F: mov var_140, 00000000h
  00410F29: mov var_148, 00000002h
  00410F33: lea edx, var_148
  00410F39: lea ecx, var_2C
  00410F3C: call MSVBVM60.DLL.__vbaVarMove
  00410F42: lea ecx, var_3C
  00410F45: push ecx
  00410F46: lea edx, var_7C
  00410F49: push edx
  00410F4A: lea eax, var_1C4
  00410F50: push eax
  00410F51: lea ecx, var_1CC
  00410F57: push ecx
  00410F58: lea edx, var_1D0
  00410F5E: push edx
  00410F5F: lea eax, var_1C8
  00410F65: push eax
  00410F66: call [00401324h] ; For Each
  00410F6C: test eax, eax
  00410F6E: jz 00412698h
  00410F74: lea ecx, var_2C
  00410F77: mov var_140, ecx
  00410F7D: mov var_148, 0000400Ch
  00410F87: lea edx, var_7C
  00410F8A: push edx
  00410F8B: call MSVBVM60.DLL.__vbaStrVarCopy
  00410F91: mov var_C0, eax
  00410F97: mov var_C8, 00000008h
  00410FA1: sub esp, 00000010h
  00410FA4: mov eax, esp
  00410FA6: mov ecx, var_148
  00410FAC: mov [eax], ecx
  00410FAE: mov edx, var_144
  00410FB4: mov [eax+04h], edx
  00410FB7: mov ecx, var_140
  00410FBD: mov [eax+08h], ecx
  00410FC0: mov edx, var_13C
  00410FC6: mov [eax+0Ch], edx
  00410FC9: push 00000001h
  00410FCB: lea eax, var_AC
  00410FD1: push eax
  00410FD2: lea ecx, var_E8
  00410FD8: push ecx
  00410FD9: call MSVBVM60.DLL.__vbaVarIndexLoad
  00410FDF: add esp, 0000000Ch
  00410FE2: mov edx, esp
  00410FE4: mov ecx, [eax]
  00410FE6: mov [edx], ecx
  00410FE8: mov ecx, [eax+04h]
  00410FEB: mov [edx+04h], ecx
  00410FEE: mov ecx, [eax+08h]
  00410FF1: mov [edx+08h], ecx
  00410FF4: mov eax, [eax+0Ch]
  00410FF7: mov [edx+0Ch], eax
  00410FFA: push 00404FC8h ; "Value"
  00410FFF: push 004051F8h
  00411004: sub esp, 00000010h
  00411007: mov ecx, esp
  00411009: mov edx, var_C8
  0041100F: mov [ecx], edx
  00411011: mov eax, var_C4
  00411017: mov [ecx+04h], eax
  0041101A: mov edx, var_C0
  00411020: mov [ecx+08h], edx
  00411023: mov eax, var_BC
  00411029: mov [ecx+0Ch], eax
  0041102C: push 00000001h
  0041102E: push 00000440h
  00411033: mov ecx, var_200
  00411039: mov edx, [ecx]
  0041103B: push edx
  0041103C: lea eax, var_D8
  00411042: push eax
  00411043: call ebx
  00411045: add esp, 00000020h
  00411048: push eax
  00411049: call MSVBVM60.DLL.__vbaCastObjVar
  0041104B: push eax
  0041104C: lea ecx, var_B4
  00411052: push ecx
  00411053: call edi
  00411055: push eax
  00411056: call MSVBVM60.DLL.__vbaLateMemSt
  0041105C: lea ecx, var_B4
  00411062: call MSVBVM60.DLL.__vbaFreeObj
  00411068: lea edx, var_E8
  0041106E: push edx
  0041106F: lea eax, var_D8
  00411075: push eax
  00411076: lea ecx, var_C8
  0041107C: push ecx
  0041107D: push 00000003h
  0041107F: call MSVBVM60.DLL.__vbaFreeVarList
  00411085: add esp, 00000010h
  00411088: mov var_140, 00000001h
  00411092: mov var_148, 00000002h
  0041109C: lea edx, var_2C
  0041109F: push edx
  004110A0: lea eax, var_148
  004110A6: push eax
  004110A7: lea ecx, var_C8
  004110AD: push ecx
  004110AE: call MSVBVM60.DLL.__vbaVarAdd
  004110B4: mov edx, eax
  004110B6: lea ecx, var_2C
  004110B9: call MSVBVM60.DLL.__vbaVarMove
  004110BF: lea edx, var_7C
  004110C2: push edx
  004110C3: lea eax, var_1C4
  004110C9: push eax
  004110CA: lea ecx, var_1CC
  004110D0: push ecx
  004110D1: lea edx, var_1D0
  004110D7: push edx
  004110D8: lea eax, var_1C8
  004110DE: push eax
  004110DF: call [00401048h] ; Next (For Each)
  004110E5: jmp 00410F6Ch
  004110EA: mov var_140, 00405240h ; "id_sc_field_no_urut,id_sc_field_nik,id_sc_field_d025,id_sc_field_d025a,d026,d027,id_sc_field_d028,id_sc_field_d030,d031,d032,d033,d034,id_ac_d035,d036,d037"
  004110F4: mov var_148, 00000008h
  004110FE: lea edx, var_148
  00411104: lea ecx, var_5C
  00411107: call MSVBVM60.DLL.__vbaVarCopy
  0041110D: mov var_140, 0040537Ch ; "id_sc_field_d038,d040,d045 [],d047 [],d048 [],d049 []"
  00411117: mov var_148, 00000008h
  00411121: lea ecx, var_5C
  00411124: push ecx
  00411125: lea edx, var_148
  0041112B: push edx
  0041112C: lea eax, var_C8
  00411132: push eax
  00411133: call MSVBVM60.DLL.__vbaVarCat
  00411139: mov edx, eax
  0041113B: lea ecx, var_5C
  0041113E: call MSVBVM60.DLL.__vbaVarMove
  00411144: push 00404ED8h
  00411149: push 00000000h
  0041114B: push 000000CBh
  00411150: mov ecx, [esi]
  00411152: push esi
  00411153: call [ecx+00000344h]
  00411159: push eax
  0041115A: lea edx, var_B4
  00411160: push edx
  00411161: call edi
  00411163: push eax
  00411164: lea eax, var_C8
  0041116A: push eax
  0041116B: call ebx
  0041116D: add esp, 00000010h
  00411170: push eax
  00411171: call MSVBVM60.DLL.__vbaCastObjVar
  00411177: push eax
  00411178: lea ecx, var_B8
  0041117E: push ecx
  0041117F: call edi
  00411181: push eax
  00411182: lea eax, [esi+34h]
  00411185: push eax
  00411186: call MSVBVM60.DLL.__vbaObjSetAddref
  0041118C: lea edx, var_B8
  00411192: push edx
  00411193: lea eax, var_B4
  00411199: push eax
  0041119A: push 00000002h
  0041119C: call MSVBVM60.DLL.__vbaFreeObjList
  004111A2: add esp, 0000000Ch
  004111A5: lea ecx, var_C8
  004111AB: call MSVBVM60.DLL.__vbaFreeVar
  004111B1: lea ecx, var_4C
  004111B4: push ecx
  004111B5: call MSVBVM60.DLL.__vbaI4Var
  004111BB: mov var_140, eax
  004111C1: mov ecx, 00000003h
  004111C6: mov var_148, ecx
  004111CC: mov var_160, 00000013h
  004111D6: mov var_168, ecx
  004111DC: sub esp, 00000010h
  004111DF: mov edx, esp
  004111E1: mov [edx], ecx
  004111E3: mov ecx, var_144
  004111E9: mov [edx+04h], ecx
  004111EC: mov [edx+08h], eax
  004111EF: mov eax, var_13C
  004111F5: mov [edx+0Ch], eax
  004111F8: sub esp, 00000010h
  004111FB: mov ecx, esp
  004111FD: mov edx, var_168
  00411203: mov [ecx], edx
  00411205: mov eax, var_164
  0041120B: mov [ecx+04h], eax
  0041120E: mov edx, var_160
  00411214: mov [ecx+08h], edx
  00411217: mov eax, var_15C
  0041121D: mov [ecx+0Ch], eax
  00411220: push 00000002h
  00411222: push 00000041h
  00411224: mov ecx, [esi]
  00411226: push esi
  00411227: call [ecx+00000340h]
  0041122D: push eax
  0041122E: lea edx, var_B4
  00411234: push edx
  00411235: call edi
  00411237: push eax
  00411238: lea eax, var_C8
  0041123E: push eax
  0041123F: call ebx
  00411241: add esp, 00000030h
  00411244: push eax
  00411245: call MSVBVM60.DLL.__vbaStrVarMove
  0041124B: mov var_E0, eax
  00411251: mov ecx, 00000008h
  00411256: mov var_E8, ecx
  0041125C: mov var_180, 004053F0h ; "id_sc_field_no_urut"
  00411266: mov var_188, ecx
  0041126C: sub esp, 00000010h
  0041126F: mov edx, esp
  00411271: mov [edx], ecx
  00411273: mov ecx, var_E4
  00411279: mov [edx+04h], ecx
  0041127C: mov [edx+08h], eax
  0041127F: mov eax, var_DC
  00411285: mov [edx+0Ch], eax
  00411288: push 00404FC8h ; "Value"
  0041128D: push 004051F8h
  00411292: sub esp, 00000010h
  00411295: mov ecx, esp
  00411297: mov edx, var_188
  0041129D: mov [ecx], edx
  0041129F: mov eax, var_184
  004112A5: mov [ecx+04h], eax
  004112A8: mov edx, var_180
  004112AE: mov [ecx+08h], edx
  004112B1: mov eax, var_17C
  004112B7: mov [ecx+0Ch], eax
  004112BA: push 00000001h
  004112BC: push 00000440h
  004112C1: mov ecx, [esi+34h]
  004112C4: push ecx
  004112C5: lea edx, var_D8
  004112CB: push edx
  004112CC: call ebx
  004112CE: add esp, 00000020h
  004112D1: push eax
  004112D2: call MSVBVM60.DLL.__vbaCastObjVar
  004112D8: push eax
  004112D9: lea eax, var_B8
  004112DF: push eax
  004112E0: call edi
  004112E2: push eax
  004112E3: call MSVBVM60.DLL.__vbaLateMemSt
  004112E9: lea ecx, var_B8
  004112EF: push ecx
  004112F0: lea edx, var_B4
  004112F6: push edx
  004112F7: push 00000002h
  004112F9: call MSVBVM60.DLL.__vbaFreeObjList
  004112FF: lea eax, var_E8
  00411305: push eax
  00411306: lea ecx, var_D8
  0041130C: push ecx
  0041130D: lea edx, var_C8
  00411313: push edx
  00411314: push 00000003h
  00411316: call MSVBVM60.DLL.__vbaFreeVarList
  0041131C: add esp, 0000001Ch
  0041131F: lea eax, var_4C
  00411322: push eax
  00411323: call MSVBVM60.DLL.__vbaI4Var
  00411329: mov var_140, eax
  0041132F: mov ecx, 00000003h
  00411334: mov var_148, ecx
  0041133A: mov var_160, 00000002h
  00411344: mov var_168, ecx
  0041134A: sub esp, 00000010h
  0041134D: mov edx, esp
  0041134F: mov [edx], ecx
  00411351: mov ecx, var_144
  00411357: mov [edx+04h], ecx
  0041135A: mov [edx+08h], eax
  0041135D: mov eax, var_13C
  00411363: mov [edx+0Ch], eax
  00411366: sub esp, 00000010h
  00411369: mov ecx, esp
  0041136B: mov edx, var_168
  00411371: mov [ecx], edx
  00411373: mov eax, var_164
  00411379: mov [ecx+04h], eax
  0041137C: mov edx, var_160
  00411382: mov [ecx+08h], edx
  00411385: mov eax, var_15C
  0041138B: mov [ecx+0Ch], eax
  0041138E: push 00000002h
  00411390: push 00000041h
  00411392: mov ecx, [esi]
  00411394: push esi
  00411395: call [ecx+00000340h]
  0041139B: push eax
  0041139C: lea edx, var_B4
  004113A2: push edx
  004113A3: call edi
  004113A5: push eax
  004113A6: lea eax, var_C8
  004113AC: push eax
  004113AD: call ebx
  004113AF: add esp, 00000030h
  004113B2: push eax
  004113B3: call MSVBVM60.DLL.__vbaStrVarMove
  004113B9: mov var_E0, eax
  004113BF: mov ecx, 00000008h
  004113C4: mov var_E8, ecx
  004113CA: mov var_180, 0040520Ch ; "id_sc_field_nik"
  004113D4: mov var_188, ecx
  004113DA: sub esp, 00000010h
  004113DD: mov edx, esp
  004113DF: mov [edx], ecx
  004113E1: mov ecx, var_E4
  004113E7: mov [edx+04h], ecx
  004113EA: mov [edx+08h], eax
  004113ED: mov eax, var_DC
  004113F3: mov [edx+0Ch], eax
  004113F6: push 00404FC8h ; "Value"
  004113FB: push 004051F8h
  00411400: sub esp, 00000010h
  00411403: mov ecx, esp
  00411405: mov edx, var_188
  0041140B: mov [ecx], edx
  0041140D: mov eax, var_184
  00411413: mov [ecx+04h], eax
  00411416: mov edx, var_180
  0041141C: mov [ecx+08h], edx
  0041141F: mov eax, var_17C
  00411425: mov [ecx+0Ch], eax
  00411428: push 00000001h
  0041142A: push 00000440h
  0041142F: mov ecx, [esi+34h]
  00411432: push ecx
  00411433: lea edx, var_D8
  00411439: push edx
  0041143A: call ebx
  0041143C: add esp, 00000020h
  0041143F: push eax
  00411440: call MSVBVM60.DLL.__vbaCastObjVar
  00411446: push eax
  00411447: lea eax, var_B8
  0041144D: push eax
  0041144E: call edi
  00411450: push eax
  00411451: call MSVBVM60.DLL.__vbaLateMemSt
  00411457: lea ecx, var_B8
  0041145D: push ecx
  0041145E: lea edx, var_B4
  00411464: push edx
  00411465: push 00000002h
  00411467: call MSVBVM60.DLL.__vbaFreeObjList
  0041146D: lea eax, var_E8
  00411473: push eax
  00411474: lea ecx, var_D8
  0041147A: push ecx
  0041147B: lea edx, var_C8
  00411481: push edx
  00411482: push 00000003h
  00411484: call MSVBVM60.DLL.__vbaFreeVarList
  0041148A: add esp, 0000001Ch
  0041148D: lea eax, var_4C
  00411490: push eax
  00411491: call MSVBVM60.DLL.__vbaI4Var
  00411497: mov var_140, eax
  0041149D: mov ecx, 00000003h
  004114A2: mov var_148, ecx
  004114A8: mov var_160, ecx
  004114AE: mov var_168, ecx
  004114B4: sub esp, 00000010h
  004114B7: mov edx, esp
  004114B9: mov [edx], ecx
  004114BB: mov ecx, var_144
  004114C1: mov [edx+04h], ecx
  004114C4: mov [edx+08h], eax
  004114C7: mov eax, var_13C
  004114CD: mov [edx+0Ch], eax
  004114D0: sub esp, 00000010h
  004114D3: mov ecx, esp
  004114D5: mov edx, var_168
  004114DB: mov [ecx], edx
  004114DD: mov eax, var_164
  004114E3: mov [ecx+04h], eax
  004114E6: mov edx, var_160
  004114EC: mov [ecx+08h], edx
  004114EF: mov eax, var_15C
  004114F5: mov [ecx+0Ch], eax
  004114F8: push 00000002h
  004114FA: push 00000041h
  004114FC: mov ecx, [esi]
  004114FE: push esi
  004114FF: call [ecx+00000340h]
  00411505: push eax
  00411506: lea edx, var_B4
  0041150C: push edx
  0041150D: call edi
  0041150F: push eax
  00411510: lea eax, var_C8
  00411516: push eax
  00411517: call ebx
  00411519: add esp, 00000030h
  0041151C: push eax
  0041151D: call MSVBVM60.DLL.__vbaStrVarMove
  00411523: mov var_E0, eax
  00411529: mov ecx, 00000008h
  0041152E: mov var_E8, ecx
  00411534: mov var_180, 00405014h ; "id_sc_field_d025"
  0041153E: mov var_188, ecx
  00411544: sub esp, 00000010h
  00411547: mov edx, esp
  00411549: mov [edx], ecx
  0041154B: mov ecx, var_E4
  00411551: mov [edx+04h], ecx
  00411554: mov [edx+08h], eax
  00411557: mov eax, var_DC
  0041155D: mov [edx+0Ch], eax
  00411560: push 00404FC8h ; "Value"
  00411565: push 004051F8h
  0041156A: sub esp, 00000010h
  0041156D: mov ecx, esp
  0041156F: mov edx, var_188
  00411575: mov [ecx], edx
  00411577: mov eax, var_184
  0041157D: mov [ecx+04h], eax
  00411580: mov edx, var_180
  00411586: mov [ecx+08h], edx
  00411589: mov eax, var_17C
  0041158F: mov [ecx+0Ch], eax
  00411592: push 00000001h
  00411594: push 00000440h
  00411599: mov ecx, [esi+34h]
  0041159C: push ecx
  0041159D: lea edx, var_D8
  004115A3: push edx
  004115A4: call ebx
  004115A6: add esp, 00000020h
  004115A9: push eax
  004115AA: call MSVBVM60.DLL.__vbaCastObjVar
  004115B0: push eax
  004115B1: lea eax, var_B8
  004115B7: push eax
  004115B8: call edi
  004115BA: push eax
  004115BB: call MSVBVM60.DLL.__vbaLateMemSt
  004115C1: lea ecx, var_B8
  004115C7: push ecx
  004115C8: lea edx, var_B4
  004115CE: push edx
  004115CF: push 00000002h
  004115D1: call MSVBVM60.DLL.__vbaFreeObjList
  004115D7: lea eax, var_E8
  004115DD: push eax
  004115DE: lea ecx, var_D8
  004115E4: push ecx
  004115E5: lea edx, var_C8
  004115EB: push edx
  004115EC: push 00000003h
  004115EE: call MSVBVM60.DLL.__vbaFreeVarList
  004115F4: add esp, 0000001Ch
  004115F7: lea eax, var_4C
  004115FA: push eax
  004115FB: call MSVBVM60.DLL.__vbaI4Var
  00411601: mov var_140, eax
  00411607: mov ecx, 00000003h
  0041160C: mov var_148, ecx
  00411612: mov var_160, 0000000Ch
  0041161C: mov var_168, ecx
  00411622: sub esp, 00000010h
  00411625: mov edx, esp
  00411627: mov [edx], ecx
  00411629: mov ecx, var_144
  0041162F: mov [edx+04h], ecx
  00411632: mov [edx+08h], eax
  00411635: mov eax, var_13C
  0041163B: mov [edx+0Ch], eax
  0041163E: sub esp, 00000010h
  00411641: mov ecx, esp
  00411643: mov edx, var_168
  00411649: mov [ecx], edx
  0041164B: mov eax, var_164
  00411651: mov [ecx+04h], eax
  00411654: mov edx, var_160
  0041165A: mov [ecx+08h], edx
  0041165D: mov eax, var_15C
  00411663: mov [ecx+0Ch], eax
  00411666: push 00000002h
  00411668: push 00000041h
  0041166A: mov ecx, [esi]
  0041166C: push esi
  0041166D: call [ecx+00000340h]
  00411673: push eax
  00411674: lea edx, var_B4
  0041167A: push edx
  0041167B: call edi
  0041167D: push eax
  0041167E: lea eax, var_C8
  00411684: push eax
  00411685: call ebx
  00411687: add esp, 00000030h
  0041168A: push eax
  0041168B: call MSVBVM60.DLL.__vbaStrVarMove
  00411691: mov var_E0, eax
  00411697: mov ecx, 00000008h
  0041169C: mov var_E8, ecx
  004116A2: mov var_180, 0040503Ch ; "id_sc_field_d028"
  004116AC: mov var_188, ecx
  004116B2: sub esp, 00000010h
  004116B5: mov edx, esp
  004116B7: mov [edx], ecx
  004116B9: mov ecx, var_E4
  004116BF: mov [edx+04h], ecx
  004116C2: mov [edx+08h], eax
  004116C5: mov eax, var_DC
  004116CB: mov [edx+0Ch], eax
  004116CE: push 00404FC8h ; "Value"
  004116D3: push 004051F8h
  004116D8: sub esp, 00000010h
  004116DB: mov ecx, esp
  004116DD: mov edx, var_188
  004116E3: mov [ecx], edx
  004116E5: mov eax, var_184
  004116EB: mov [ecx+04h], eax
  004116EE: mov edx, var_180
  004116F4: mov [ecx+08h], edx
  004116F7: mov eax, var_17C
  004116FD: mov [ecx+0Ch], eax
  00411700: push 00000001h
  00411702: push 00000440h
  00411707: mov ecx, [esi+34h]
  0041170A: push ecx
  0041170B: lea edx, var_D8
  00411711: push edx
  00411712: call ebx
  00411714: add esp, 00000020h
  00411717: push eax
  00411718: call MSVBVM60.DLL.__vbaCastObjVar
  0041171E: push eax
  0041171F: lea eax, var_B8
  00411725: push eax
  00411726: call edi
  00411728: push eax
  00411729: call MSVBVM60.DLL.__vbaLateMemSt
  0041172F: lea ecx, var_B8
  00411735: push ecx
  00411736: lea edx, var_B4
  0041173C: push edx
  0041173D: push 00000002h
  0041173F: call MSVBVM60.DLL.__vbaFreeObjList
  00411745: lea eax, var_E8
  0041174B: push eax
  0041174C: lea ecx, var_D8
  00411752: push ecx
  00411753: lea edx, var_C8
  00411759: push edx
  0041175A: push 00000003h
  0041175C: call MSVBVM60.DLL.__vbaFreeVarList
  00411762: add esp, 0000001Ch
  00411765: lea eax, var_4C
  00411768: push eax
  00411769: call MSVBVM60.DLL.__vbaI4Var
  0041176F: mov var_140, eax
  00411775: mov ecx, 00000003h
  0041177A: mov var_148, ecx
  00411780: mov var_160, 0000000Dh
  0041178A: mov var_168, ecx
  00411790: sub esp, 00000010h
  00411793: mov edx, esp
  00411795: mov [edx], ecx
  00411797: mov ecx, var_144
  0041179D: mov [edx+04h], ecx
  004117A0: mov [edx+08h], eax
  004117A3: mov eax, var_13C
  004117A9: mov [edx+0Ch], eax
  004117AC: sub esp, 00000010h
  004117AF: mov ecx, esp
  004117B1: mov edx, var_168
  004117B7: mov [ecx], edx
  004117B9: mov eax, var_164
  004117BF: mov [ecx+04h], eax
  004117C2: mov edx, var_160
  004117C8: mov [ecx+08h], edx
  004117CB: mov eax, var_15C
  004117D1: mov [ecx+0Ch], eax
  004117D4: push 00000002h
  004117D6: push 00000041h
  004117D8: mov ecx, [esi]
  004117DA: push esi
  004117DB: call [ecx+00000340h]
  004117E1: push eax
  004117E2: lea edx, var_B4
  004117E8: push edx
  004117E9: call edi
  004117EB: push eax
  004117EC: lea eax, var_C8
  004117F2: push eax
  004117F3: call ebx
  004117F5: add esp, 00000030h
  004117F8: push eax
  004117F9: call MSVBVM60.DLL.__vbaStrVarMove
  004117FF: mov var_E0, eax
  00411805: mov ecx, 00000008h
  0041180A: mov var_E8, ecx
  00411810: mov var_180, 00405430h ; "id_sc_field_d029"
  0041181A: mov var_188, ecx
  00411820: sub esp, 00000010h
  00411823: mov edx, esp
  00411825: mov [edx], ecx
  00411827: mov ecx, var_E4
  0041182D: mov [edx+04h], ecx
  00411830: mov [edx+08h], eax
  00411833: mov eax, var_DC
  00411839: mov [edx+0Ch], eax
  0041183C: push 00404FC8h ; "Value"
  00411841: push 004051F8h
  00411846: sub esp, 00000010h
  00411849: mov ecx, esp
  0041184B: mov edx, var_188
  00411851: mov [ecx], edx
  00411853: mov eax, var_184
  00411859: mov [ecx+04h], eax
  0041185C: mov edx, var_180
  00411862: mov [ecx+08h], edx
  00411865: mov eax, var_17C
  0041186B: mov [ecx+0Ch], eax
  0041186E: push 00000001h
  00411870: push 00000440h
  00411875: mov ecx, [esi+34h]
  00411878: push ecx
  00411879: lea edx, var_D8
  0041187F: push edx
  00411880: call ebx
  00411882: add esp, 00000020h
  00411885: push eax
  00411886: call MSVBVM60.DLL.__vbaCastObjVar
  0041188C: push eax
  0041188D: lea eax, var_B8
  00411893: push eax
  00411894: call edi
  00411896: push eax
  00411897: call MSVBVM60.DLL.__vbaLateMemSt
  0041189D: lea ecx, var_B8
  004118A3: push ecx
  004118A4: lea edx, var_B4
  004118AA: push edx
  004118AB: push 00000002h
  004118AD: call MSVBVM60.DLL.__vbaFreeObjList
  004118B3: lea eax, var_E8
  004118B9: push eax
  004118BA: lea ecx, var_D8
  004118C0: push ecx
  004118C1: lea edx, var_C8
  004118C7: push edx
  004118C8: push 00000003h
  004118CA: call MSVBVM60.DLL.__vbaFreeVarList
  004118D0: add esp, 0000001Ch
  004118D3: lea eax, var_4C
  004118D6: push eax
  004118D7: call MSVBVM60.DLL.__vbaI4Var
  004118DD: mov var_140, eax
  004118E3: mov ecx, 00000003h
  004118E8: mov var_148, ecx
  004118EE: mov var_160, 00000007h
  004118F8: mov var_168, ecx
  004118FE: sub esp, 00000010h
  00411901: mov edx, esp
  00411903: mov [edx], ecx
  00411905: mov ecx, var_144
  0041190B: mov [edx+04h], ecx
  0041190E: mov [edx+08h], eax
  00411911: mov eax, var_13C
  00411917: mov [edx+0Ch], eax
  0041191A: sub esp, 00000010h
  0041191D: mov ecx, esp
  0041191F: mov edx, var_168
  00411925: mov [ecx], edx
  00411927: mov eax, var_164
  0041192D: mov [ecx+04h], eax
  00411930: mov edx, var_160
  00411936: mov [ecx+08h], edx
  00411939: mov eax, var_15C
  0041193F: mov [ecx+0Ch], eax
  00411942: push 00000002h
  00411944: push 00000041h
  00411946: mov ecx, [esi]
  00411948: push esi
  00411949: call [ecx+00000340h]
  0041194F: push eax
  00411950: lea edx, var_B4
  00411956: push edx
  00411957: call edi
  00411959: push eax
  0041195A: lea eax, var_C8
  00411960: push eax
  00411961: call ebx
  00411963: add esp, 00000030h
  00411966: push eax
  00411967: call MSVBVM60.DLL.__vbaStrVarMove
  0041196D: mov var_E0, eax
  00411973: mov ecx, 00000008h
  00411978: mov var_E8, ecx
  0041197E: mov var_180, 00405458h ; "id_ac_d035"
  00411988: mov var_188, ecx
  0041198E: sub esp, 00000010h
  00411991: mov edx, esp
  00411993: mov [edx], ecx
  00411995: mov ecx, var_E4
  0041199B: mov [edx+04h], ecx
  0041199E: mov [edx+08h], eax
  004119A1: mov eax, var_DC
  004119A7: mov [edx+0Ch], eax
  004119AA: push 00404FC8h ; "Value"
  004119AF: push 004051F8h
  004119B4: sub esp, 00000010h
  004119B7: mov ecx, esp
  004119B9: mov edx, var_188
  004119BF: mov [ecx], edx
  004119C1: mov eax, var_184
  004119C7: mov [ecx+04h], eax
  004119CA: mov edx, var_180
  004119D0: mov [ecx+08h], edx
  004119D3: mov eax, var_17C
  004119D9: mov [ecx+0Ch], eax
  004119DC: push 00000001h
  004119DE: push 00000440h
  004119E3: mov ecx, [esi+34h]
  004119E6: push ecx
  004119E7: lea edx, var_D8
  004119ED: push edx
  004119EE: call ebx
  004119F0: add esp, 00000020h
  004119F3: push eax
  004119F4: call MSVBVM60.DLL.__vbaCastObjVar
  004119FA: push eax
  004119FB: lea eax, var_B8
  00411A01: push eax
  00411A02: call edi
  00411A04: push eax
  00411A05: call MSVBVM60.DLL.__vbaLateMemSt
  00411A0B: lea ecx, var_B8
  00411A11: push ecx
  00411A12: lea edx, var_B4
  00411A18: push edx
  00411A19: push 00000002h
  00411A1B: call MSVBVM60.DLL.__vbaFreeObjList
  00411A21: lea eax, var_E8
  00411A27: push eax
  00411A28: lea ecx, var_D8
  00411A2E: push ecx
  00411A2F: lea edx, var_C8
  00411A35: push edx
  00411A36: push 00000003h
  00411A38: call MSVBVM60.DLL.__vbaFreeVarList
  00411A3E: add esp, 0000001Ch
  00411A41: lea eax, var_4C
  00411A44: push eax
  00411A45: call MSVBVM60.DLL.__vbaI4Var
  00411A4B: mov var_140, eax
  00411A51: mov ecx, 00000003h
  00411A56: mov var_148, ecx
  00411A5C: mov var_160, 0000000Bh
  00411A66: mov var_168, ecx
  00411A6C: sub esp, 00000010h
  00411A6F: mov edx, esp
  00411A71: mov [edx], ecx
  00411A73: mov ecx, var_144
  00411A79: mov [edx+04h], ecx
  00411A7C: mov [edx+08h], eax
  00411A7F: mov eax, var_13C
  00411A85: mov [edx+0Ch], eax
  00411A88: sub esp, 00000010h
  00411A8B: mov ecx, esp
  00411A8D: mov edx, var_168
  00411A93: mov [ecx], edx
  00411A95: mov eax, var_164
  00411A9B: mov [ecx+04h], eax
  00411A9E: mov edx, var_160
  00411AA4: mov [ecx+08h], edx
  00411AA7: mov eax, var_15C
  00411AAD: mov [ecx+0Ch], eax
  00411AB0: push 00000002h
  00411AB2: push 00000041h
  00411AB4: mov ecx, [esi]
  00411AB6: push esi
  00411AB7: call [ecx+00000340h]
  00411ABD: push eax
  00411ABE: lea edx, var_B4
  00411AC4: push edx
  00411AC5: call edi
  00411AC7: push eax
  00411AC8: lea eax, var_C8
  00411ACE: push eax
  00411ACF: call ebx
  00411AD1: add esp, 00000030h
  00411AD4: push eax
  00411AD5: call MSVBVM60.DLL.__vbaStrVarMove
  00411ADB: mov var_E0, eax
  00411AE1: mov ecx, 00000008h
  00411AE6: mov var_E8, ecx
  00411AEC: mov var_180, 00405474h ; "id_sc_field_d038"
  00411AF6: mov var_188, ecx
  00411AFC: sub esp, 00000010h
  00411AFF: mov edx, esp
  00411B01: mov [edx], ecx
  00411B03: mov ecx, var_E4
  00411B09: mov [edx+04h], ecx
  00411B0C: mov [edx+08h], eax
  00411B0F: mov eax, var_DC
  00411B15: mov [edx+0Ch], eax
  00411B18: push 00404FC8h ; "Value"
  00411B1D: push 004051F8h
  00411B22: sub esp, 00000010h
  00411B25: mov ecx, esp
  00411B27: mov edx, var_188
  00411B2D: mov [ecx], edx
  00411B2F: mov eax, var_184
  00411B35: mov [ecx+04h], eax
  00411B38: mov edx, var_180
  00411B3E: mov [ecx+08h], edx
  00411B41: mov eax, var_17C
  00411B47: mov [ecx+0Ch], eax
  00411B4A: push 00000001h
  00411B4C: push 00000440h
  00411B51: mov ecx, [esi+34h]
  00411B54: push ecx
  00411B55: lea edx, var_D8
  00411B5B: push edx
  00411B5C: call ebx
  00411B5E: add esp, 00000020h
  00411B61: push eax
  00411B62: call MSVBVM60.DLL.__vbaCastObjVar
  00411B68: push eax
  00411B69: lea eax, var_B8
  00411B6F: push eax
  00411B70: call edi
  00411B72: push eax
  00411B73: call MSVBVM60.DLL.__vbaLateMemSt
  00411B79: lea ecx, var_B8
  00411B7F: push ecx
  00411B80: lea edx, var_B4
  00411B86: push edx
  00411B87: push 00000002h
  00411B89: call MSVBVM60.DLL.__vbaFreeObjList
  00411B8F: lea eax, var_E8
  00411B95: push eax
  00411B96: lea ecx, var_D8
  00411B9C: push ecx
  00411B9D: lea edx, var_C8
  00411BA3: push edx
  00411BA4: push 00000003h
  00411BA6: call MSVBVM60.DLL.__vbaFreeVarList
  00411BAC: add esp, 0000001Ch
  00411BAF: lea eax, var_4C
  00411BB2: push eax
  00411BB3: call MSVBVM60.DLL.__vbaI4Var
  00411BB9: mov var_140, eax
  00411BBF: mov ecx, 00000003h
  00411BC4: mov var_148, ecx
  00411BCA: mov var_160, 00000004h
  00411BD4: mov var_168, ecx
  00411BDA: sub esp, 00000010h
  00411BDD: mov edx, esp
  00411BDF: mov [edx], ecx
  00411BE1: mov ecx, var_144
  00411BE7: mov [edx+04h], ecx
  00411BEA: mov [edx+08h], eax
  00411BED: mov eax, var_13C
  00411BF3: mov [edx+0Ch], eax
  00411BF6: sub esp, 00000010h
  00411BF9: mov ecx, esp
  00411BFB: mov edx, var_168
  00411C01: mov [ecx], edx
  00411C03: mov eax, var_164
  00411C09: mov [ecx+04h], eax
  00411C0C: mov edx, var_160
  00411C12: mov [ecx+08h], edx
  00411C15: mov eax, var_15C
  00411C1B: mov [ecx+0Ch], eax
  00411C1E: push 00000002h
  00411C20: push 00000041h
  00411C22: mov ecx, [esi]
  00411C24: push esi
  00411C25: call [ecx+00000340h]
  00411C2B: push eax
  00411C2C: lea edx, var_B4
  00411C32: push edx
  00411C33: call edi
  00411C35: push eax
  00411C36: lea eax, var_C8
  00411C3C: push eax
  00411C3D: call ebx
  00411C3F: add esp, 00000030h
  00411C42: push eax
  00411C43: call MSVBVM60.DLL.__vbaStrVarMove
  00411C49: mov edx, eax
  00411C4B: lea ecx, var_B0
  00411C51: call MSVBVM60.DLL.__vbaStrMove
  00411C57: push eax
  00411C58: push 0040549Ch
  00411C5D: call MSVBVM60.DLL.__vbaStrCmp
  00411C63: neg eax
  00411C65: sbb eax, eax
  00411C67: inc eax
  00411C68: neg eax
  00411C6A: mov var_1BC, ax
  00411C71: lea ecx, var_B0
  00411C77: call MSVBVM60.DLL.__vbaFreeStr
  00411C7D: lea ecx, var_B4
  00411C83: call MSVBVM60.DLL.__vbaFreeObj
  00411C89: lea ecx, var_C8
  00411C8F: call MSVBVM60.DLL.__vbaFreeVar
  00411C95: cmp word ptr var_1BC, 0000h
  00411C9D: mov var_180, FFFFFFFFh
  00411CA7: mov var_188, 0000000Bh
  00411CB1: mov eax, 004054A4h ; "d026"
  00411CB6: mov var_140, eax
  00411CBC: mov ecx, 00000008h
  00411CC1: mov var_148, ecx
  00411CC7: push 00404EF8h
  00411CCC: jz 00411DB1h
  00411CD2: sub esp, 00000010h
  00411CD5: mov edx, esp
  00411CD7: mov [edx], ecx
  00411CD9: mov ecx, var_144
  00411CDF: mov [edx+04h], ecx
  00411CE2: mov [edx+08h], eax
  00411CE5: mov eax, var_13C
  00411CEB: mov [edx+0Ch], eax
  00411CEE: push 00000001h
  00411CF0: push 0000043Eh
  00411CF5: mov ecx, [esi+34h]
  00411CF8: push ecx
  00411CF9: lea edx, var_C8
  00411CFF: push edx
  00411D00: call ebx
  00411D02: add esp, 00000020h
  00411D05: push eax
  00411D06: call MSVBVM60.DLL.__vbaCastObjVar
  00411D0C: push eax
  00411D0D: lea eax, var_B4
  00411D13: push eax
  00411D14: call edi
  00411D16: mov var_1BC, eax
  00411D1C: mov var_170, 80020004h
  00411D26: mov ecx, 0000000Ah
  00411D2B: mov var_178, ecx
  00411D31: mov var_160, 00000000h
  00411D3B: mov var_168, 00000002h
  00411D45: mov edx, [eax]
  00411D47: lea eax, var_B8
  00411D4D: push eax
  00411D4E: sub esp, 00000010h
  00411D51: mov eax, esp
  00411D53: mov [eax], ecx
  00411D55: mov ecx, var_174
  00411D5B: mov [eax+04h], ecx
  00411D5E: mov ecx, var_170
  00411D64: mov [eax+08h], ecx
  00411D67: mov ecx, var_16C
  00411D6D: mov [eax+0Ch], ecx
  00411D70: sub esp, 00000010h
  00411D73: mov eax, esp
  00411D75: mov ecx, var_168
  00411D7B: mov [eax], ecx
  00411D7D: mov ecx, var_164
  00411D83: mov [eax+04h], ecx
  00411D86: mov ecx, var_160
  00411D8C: mov [eax+08h], ecx
  00411D8F: mov ecx, var_15C
  00411D95: mov [eax+0Ch], ecx
  00411D98: mov eax, var_1BC
  00411D9E: push eax
  00411D9F: call [edx+2Ch]
  00411DA2: fclex 
  00411DA4: test eax, eax
  00411DA6: jnl 00411E9Ch
  00411DAC: jmp 00411E87h
  00411DB1: sub esp, 00000010h
  00411DB4: mov edx, esp
  00411DB6: mov [edx], ecx
  00411DB8: mov ecx, var_144
  00411DBE: mov [edx+04h], ecx
  00411DC1: mov [edx+08h], eax
  00411DC4: mov eax, var_13C
  00411DCA: mov [edx+0Ch], eax
  00411DCD: push 00000001h
  00411DCF: push 0000043Eh
  00411DD4: mov ecx, [esi+34h]
  00411DD7: push ecx
  00411DD8: lea edx, var_C8
  00411DDE: push edx
  00411DDF: call ebx
  00411DE1: add esp, 00000020h
  00411DE4: push eax
  00411DE5: call MSVBVM60.DLL.__vbaCastObjVar
  00411DEB: push eax
  00411DEC: lea eax, var_B4
  00411DF2: push eax
  00411DF3: call edi
  00411DF5: mov var_1BC, eax
  00411DFB: mov var_170, 80020004h
  00411E05: mov ecx, 0000000Ah
  00411E0A: mov var_178, ecx
  00411E10: mov var_160, 00000001h
  00411E1A: mov var_168, 00000002h
  00411E24: mov edx, [eax]
  00411E26: lea eax, var_B8
  00411E2C: push eax
  00411E2D: sub esp, 00000010h
  00411E30: mov eax, esp
  00411E32: mov [eax], ecx
  00411E34: mov ecx, var_174
  00411E3A: mov [eax+04h], ecx
  00411E3D: mov ecx, var_170
  00411E43: mov [eax+08h], ecx
  00411E46: mov ecx, var_16C
  00411E4C: mov [eax+0Ch], ecx
  00411E4F: sub esp, 00000010h
  00411E52: mov eax, esp
  00411E54: mov ecx, var_168
  00411E5A: mov [eax], ecx
  00411E5C: mov ecx, var_164
  00411E62: mov [eax+04h], ecx
  00411E65: mov ecx, var_160
  00411E6B: mov [eax+08h], ecx
  00411E6E: mov ecx, var_15C
  00411E74: mov [eax+0Ch], ecx
  00411E77: mov eax, var_1BC
  00411E7D: push eax
  00411E7E: call [edx+2Ch]
  00411E81: fclex 
  00411E83: test eax, eax
  00411E85: jnl 411E9Ch
  00411E87: push 0000002Ch
  00411E89: push 00404EF8h
  00411E8E: mov ecx, var_1BC
  00411E94: push ecx
  00411E95: push eax
  00411E96: call MSVBVM60.DLL.__vbaHresultCheckObj
  00411E9C: sub esp, 00000010h
  00411E9F: mov edx, esp
  00411EA1: mov eax, var_188
  00411EA7: mov [edx], eax
  00411EA9: mov ecx, var_184
  00411EAF: mov [edx+04h], ecx
  00411EB2: mov eax, var_180
  00411EB8: mov [edx+08h], eax
  00411EBB: mov ecx, var_17C
  00411EC1: mov [edx+0Ch], ecx
  00411EC4: push 004054B0h ; "Checked"
  00411EC9: mov edx, var_B8
  00411ECF: push edx
  00411ED0: call MSVBVM60.DLL.__vbaLateMemSt
  00411ED6: lea eax, var_B8
  00411EDC: push eax
  00411EDD: lea ecx, var_B4
  00411EE3: push ecx
  00411EE4: push 00000002h
  00411EE6: call MSVBVM60.DLL.__vbaFreeObjList
  00411EEC: add esp, 0000000Ch
  00411EEF: lea ecx, var_C8
  00411EF5: call MSVBVM60.DLL.__vbaFreeVar
  00411EFB: lea edx, var_4C
  00411EFE: push edx
  00411EFF: call MSVBVM60.DLL.__vbaI4Var
  00411F05: mov var_140, eax
  00411F0B: mov ecx, 00000003h
  00411F10: mov var_148, ecx
  00411F16: mov var_160, 00000005h
  00411F20: mov var_168, ecx
  00411F26: sub esp, 00000010h
  00411F29: mov edx, esp
  00411F2B: mov [edx], ecx
  00411F2D: mov ecx, var_144
  00411F33: mov [edx+04h], ecx
  00411F36: mov [edx+08h], eax
  00411F39: mov eax, var_13C
  00411F3F: mov [edx+0Ch], eax
  00411F42: sub esp, 00000010h
  00411F45: mov ecx, esp
  00411F47: mov edx, var_168
  00411F4D: mov [ecx], edx
  00411F4F: mov eax, var_164
  00411F55: mov [ecx+04h], eax
  00411F58: mov edx, var_160
  00411F5E: mov [ecx+08h], edx
  00411F61: mov eax, var_15C
  00411F67: mov [ecx+0Ch], eax
  00411F6A: push 00000002h
  00411F6C: push 00000041h
  00411F6E: mov ecx, [esi]
  00411F70: push esi
  00411F71: call [ecx+00000340h]
  00411F77: push eax
  00411F78: lea edx, var_B4
  00411F7E: push edx
  00411F7F: call edi
  00411F81: push eax
  00411F82: lea eax, var_C8
  00411F88: push eax
  00411F89: call ebx
  00411F8B: add esp, 00000030h
  00411F8E: lea ecx, var_C8
  00411F94: push ecx
  00411F95: call MSVBVM60.DLL.__vbaStrVarMove
  00411F9B: mov var_E0, eax
  00411FA1: mov eax, 00000008h
  00411FA6: mov var_E8, eax
  00411FAC: mov var_180, 004054C4h ; "d027"
  00411FB6: mov var_188, eax
  00411FBC: lea edx, var_188
  00411FC2: lea ecx, var_D8
  00411FC8: call MSVBVM60.DLL.__vbaVarDup
  00411FCE: mov edx, [esi]
  00411FD0: lea eax, var_F8
  00411FD6: push eax
  00411FD7: lea ecx, var_E8
  00411FDD: push ecx
  00411FDE: lea eax, var_D8
  00411FE4: push eax
  00411FE5: push esi
  00411FE6: call [edx+000006FCh]
  00411FEC: test eax, eax
  00411FEE: jnl 412002h
  00411FF0: push 000006FCh
  00411FF5: push 004049ACh
  00411FFA: push esi
  00411FFB: push eax
  00411FFC: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412002: lea ecx, var_B4
  00412008: call MSVBVM60.DLL.__vbaFreeObj
  0041200E: lea ecx, var_F8
  00412014: push ecx
  00412015: lea edx, var_E8
  0041201B: push edx
  0041201C: lea eax, var_D8
  00412022: push eax
  00412023: lea ecx, var_C8
  00412029: push ecx
  0041202A: push 00000004h
  0041202C: call MSVBVM60.DLL.__vbaFreeVarList
  00412032: add esp, 00000014h
  00412035: lea edx, var_4C
  00412038: push edx
  00412039: call MSVBVM60.DLL.__vbaI4Var
  0041203F: mov var_140, eax
  00412045: mov ecx, 00000003h
  0041204A: mov var_148, ecx
  00412050: mov var_160, 00000006h
  0041205A: mov var_168, ecx
  00412060: sub esp, 00000010h
  00412063: mov edx, esp
  00412065: mov [edx], ecx
  00412067: mov ecx, var_144
  0041206D: mov [edx+04h], ecx
  00412070: mov [edx+08h], eax
  00412073: mov eax, var_13C
  00412079: mov [edx+0Ch], eax
  0041207C: sub esp, 00000010h
  0041207F: mov ecx, esp
  00412081: mov edx, var_168
  00412087: mov [ecx], edx
  00412089: mov eax, var_164
  0041208F: mov [ecx+04h], eax
  00412092: mov edx, var_160
  00412098: mov [ecx+08h], edx
  0041209B: mov eax, var_15C
  004120A1: mov [ecx+0Ch], eax
  004120A4: push 00000002h
  004120A6: push 00000041h
  004120A8: mov ecx, [esi]
  004120AA: push esi
  004120AB: call [ecx+00000340h]
  004120B1: push eax
  004120B2: lea edx, var_B4
  004120B8: push edx
  004120B9: call edi
  004120BB: push eax
  004120BC: lea eax, var_C8
  004120C2: push eax
  004120C3: call ebx
  004120C5: add esp, 00000030h
  004120C8: lea ecx, var_C8
  004120CE: push ecx
  004120CF: call MSVBVM60.DLL.__vbaStrVarMove
  004120D5: mov var_E0, eax
  004120DB: mov eax, 00000008h
  004120E0: mov var_E8, eax
  004120E6: mov var_180, 004054D4h ; "d031"
  004120F0: mov var_188, eax
  004120F6: lea edx, var_188
  004120FC: lea ecx, var_D8
  00412102: call MSVBVM60.DLL.__vbaVarDup
  00412108: mov edx, [esi]
  0041210A: lea eax, var_F8
  00412110: push eax
  00412111: lea ecx, var_E8
  00412117: push ecx
  00412118: lea eax, var_D8
  0041211E: push eax
  0041211F: push esi
  00412120: call [edx+000006FCh]
  00412126: test eax, eax
  00412128: jnl 41213Ch
  0041212A: push 000006FCh
  0041212F: push 004049ACh
  00412134: push esi
  00412135: push eax
  00412136: call MSVBVM60.DLL.__vbaHresultCheckObj
  0041213C: lea ecx, var_B4
  00412142: call MSVBVM60.DLL.__vbaFreeObj
  00412148: lea ecx, var_F8
  0041214E: push ecx
  0041214F: lea edx, var_E8
  00412155: push edx
  00412156: lea eax, var_D8
  0041215C: push eax
  0041215D: lea ecx, var_C8
  00412163: push ecx
  00412164: push 00000004h
  00412166: call MSVBVM60.DLL.__vbaFreeVarList
  0041216C: add esp, 00000014h
  0041216F: lea edx, var_4C
  00412172: push edx
  00412173: call MSVBVM60.DLL.__vbaI4Var
  00412179: mov var_140, eax
  0041217F: mov ecx, 00000003h
  00412184: mov var_148, ecx
  0041218A: mov var_160, 00000008h
  00412194: mov var_168, ecx
  0041219A: sub esp, 00000010h
  0041219D: mov edx, esp
  0041219F: mov [edx], ecx
  004121A1: mov ecx, var_144
  004121A7: mov [edx+04h], ecx
  004121AA: mov [edx+08h], eax
  004121AD: mov eax, var_13C
  004121B3: mov [edx+0Ch], eax
  004121B6: sub esp, 00000010h
  004121B9: mov ecx, esp
  004121BB: mov edx, var_168
  004121C1: mov [ecx], edx
  004121C3: mov eax, var_164
  004121C9: mov [ecx+04h], eax
  004121CC: mov edx, var_160
  004121D2: mov [ecx+08h], edx
  004121D5: mov eax, var_15C
  004121DB: mov [ecx+0Ch], eax
  004121DE: push 00000002h
  004121E0: push 00000041h
  004121E2: mov ecx, [esi]
  004121E4: push esi
  004121E5: call [ecx+00000340h]
  004121EB: push eax
  004121EC: lea edx, var_B4
  004121F2: push edx
  004121F3: call edi
  004121F5: push eax
  004121F6: lea eax, var_C8
  004121FC: push eax
  004121FD: call ebx
  004121FF: add esp, 00000030h
  00412202: lea ecx, var_C8
  00412208: push ecx
  00412209: call MSVBVM60.DLL.__vbaStrVarMove
  0041220F: mov var_E0, eax
  00412215: mov eax, 00000008h
  0041221A: mov var_E8, eax
  00412220: mov var_180, 004054E4h ; "d032"
  0041222A: mov var_188, eax
  00412230: lea edx, var_188
  00412236: lea ecx, var_D8
  0041223C: call MSVBVM60.DLL.__vbaVarDup
  00412242: mov edx, [esi]
  00412244: lea eax, var_F8
  0041224A: push eax
  0041224B: lea ecx, var_E8
  00412251: push ecx
  00412252: lea eax, var_D8
  00412258: push eax
  00412259: push esi
  0041225A: call [edx+000006FCh]
  00412260: test eax, eax
  00412262: jnl 412276h
  00412264: push 000006FCh
  00412269: push 004049ACh
  0041226E: push esi
  0041226F: push eax
  00412270: call MSVBVM60.DLL.__vbaHresultCheckObj
  00412276: lea ecx, var_B4
  0041227C: call MSVBVM60.DLL.__vbaFreeObj
  00412282: lea ecx, var_F8
  00412288: push ecx
  00412289: lea edx, var_E8
  0041228F: push edx
  00412290: lea eax, var_D8
  00412296: push eax
  00412297: lea ecx, var_C8
  0041229D: push ecx
  0041229E: push 00000004h
  004122A0: call MSVBVM60.DLL.__vbaFreeVarList
  004122A6: add esp, 00000014h
  004122A9: lea edx, var_4C
  004122AC: push edx
  004122AD: call MSVBVM60.DLL.__vbaI4Var
  004122B3: mov var_140, eax
  004122B9: mov ecx, 00000003h
  004122BE: mov var_148, ecx
  004122C4: mov var_160, 00000009h
  004122CE: mov var_168, ecx
  004122D4: sub esp, 00000010h
  004122D7: mov edx, esp
  004122D9: mov [edx], ecx
  004122DB: mov ecx, var_144
  004122E1: mov [edx+04h], ecx
  004122E4: mov [edx+08h], eax
  004122E7: mov eax, var_13C
  004122ED: mov [edx+0Ch], eax
  004122F0: sub esp, 00000010h
  004122F3: mov ecx, esp
  004122F5: mov edx, var_168
  004122FB: mov [ecx], edx
  004122FD: mov eax, var_164
  00412303: mov [ecx+04h], eax
  00412306: mov edx, var_160
  0041230C: mov [ecx+08h], edx
  0041230F: mov eax, var_15C
  00412315: mov [ecx+0Ch], eax
  00412318: push 00000002h
  0041231A: push 00000041h
  0041231C: mov ecx, [esi]
  0041231E: push esi
  0041231F: call [ecx+00000340h]
  00412325: push eax
  00412326: lea edx, var_B4
  0041232C: push edx
  0041232D: call edi
  0041232F: push eax
  00412330: lea eax, var_C8
  00412336: push eax
  00412337: call ebx
  00412339: add esp, 00000030h
  0041233C: lea ecx, var_C8
  00412342: push ecx
  00412343: call MSVBVM60.DLL.__vbaStrVarMove
  00412349: mov var_E0, eax
  0041234F: mov eax, 00000008h
  00412354: mov var_E8, eax
  0041235A: mov var_180, 004054F4h ; "d036"
  00412364: mov var_188, eax
  0041236A: lea edx, var_188
  00412370: lea ecx, var_D8
  00412376: call MSVBVM60.DLL.__vbaVarDup
  0041237C: mov edx, [esi]
  0041237E: lea eax, var_F8
  00412384: push eax
  00412385: lea ecx, var_E8
  0041238B: push ecx
  0041238C: lea eax, var_D8
  00412392: push eax
  00412393: push esi
  00412394: call [edx+000006FCh]
  0041239A: test eax, eax
  0041239C: jnl 4123B0h
  0041239E: push 000006FCh
  004123A3: push 004049ACh
  004123A8: push esi
  004123A9: push eax
  004123AA: call MSVBVM60.DLL.__vbaHresultCheckObj
  004123B0: lea ecx, var_B4
  004123B6: call MSVBVM60.DLL.__vbaFreeObj
  004123BC: lea ecx, var_F8
  004123C2: push ecx
  004123C3: lea edx, var_E8
  004123C9: push edx
  004123CA: lea eax, var_D8
  004123D0: push eax
  004123D1: lea ecx, var_C8
  004123D7: push ecx
  004123D8: push 00000004h
  004123DA: call MSVBVM60.DLL.__vbaFreeVarList
  004123E0: add esp, 00000014h
  004123E3: lea edx, var_4C
  004123E6: push edx
  004123E7: call MSVBVM60.DLL.__vbaI4Var
  004123ED: mov var_140, eax
  004123F3: mov ecx, 00000003h
  004123F8: mov var_148, ecx
  004123FE: mov var_160, 0000000Ah
  00412408: mov var_168, ecx
  0041240E: sub esp, 00000010h
  00412411: mov edx, esp
  00412413: mov [edx], ecx
  00412415: mov ecx, var_144
  0041241B: mov [edx+04h], ecx
  0041241E: mov [edx+08h], eax
  00412421: mov eax, var_13C
  00412427: mov [edx+0Ch], eax
  0041242A: sub esp, 00000010h
  0041242D: mov ecx, esp
  0041242F: mov edx, var_168
  00412435: mov [ecx], edx
  00412437: mov eax, var_164
  0041243D: mov [ecx+04h], eax
  00412440: mov edx, var_160
  00412446: mov [ecx+08h], edx
  00412449: mov eax, var_15C
  0041244F: mov [ecx+0Ch], eax
  00412452: push 00000002h
  00412454: push 00000041h
  00412456: mov ecx, [esi]
  00412458: push esi
  00412459: call [ecx+00000340h]
  0041245F: push eax
  00412460: lea edx, var_B4
  00412466: push edx
  00412467: call edi
  00412469: push eax
  0041246A: lea eax, var_C8
  00412470: push eax
  00412471: call ebx
  00412473: add esp, 00000030h
  00412476: lea ecx, var_C8
  0041247C: push ecx
  0041247D: call MSVBVM60.DLL.__vbaStrVarMove
  00412483: mov var_E0, eax
  00412489: mov eax, 00000008h
  0041248E: mov var_E8, eax
  00412494: mov var_180, 00405504h ; "d037"
  0041249E: mov var_188, eax
  004124A4: lea edx, var_188
  004124AA: lea ecx, var_D8
  004124B0: call MSVBVM60.DLL.__vbaVarDup
  004124B6: mov edx, [esi]
  004124B8: lea eax, var_F8
  004124BE: push eax
  004124BF: lea ecx, var_E8
  004124C5: push ecx
  004124C6: lea eax, var_D8
  004124CC: push eax
  004124CD: push esi
  004124CE: call [edx+000006FCh]
  004124D4: test eax, eax
  004124D6: jnl 4124EAh
  004124D8: push 000006FCh
  004124DD: push 004049ACh
  004124E2: push esi
  004124E3: push eax
  004124E4: call MSVBVM60.DLL.__vbaHresultCheckObj
  004124EA: lea ecx, var_B4
  004124F0: call MSVBVM60.DLL.__vbaFreeObj
  004124F6: lea ecx, var_F8
  004124FC: push ecx
  004124FD: lea edx, var_E8
  00412503: push edx
  00412504: lea eax, var_D8
  0041250A: push eax
  0041250B: lea ecx, var_C8
  00412511: push ecx
  00412512: push 00000004h
  00412514: call MSVBVM60.DLL.__vbaFreeVarList
  0041251A: add esp, 00000014h
  0041251D: call MSVBVM60.DLL.__vbaExitProc
  00412523: Method_8964E44D
  00412528: jmp 0041270Bh
  0041252D: mov var_160, 00405544h ; "AK"
  00412537: mov edi, 00000008h
  0041253C: mov var_168, edi
  00412542: lea edx, var_168
  00412548: lea ecx, var_D8
  0041254E: mov esi, [004012DCh]
  00412554: call MSVBVM60.DLL.__vbaVarDup
  00412556: mov var_150, 00405538h ; "KK"
  00412560: mov var_158, edi
  00412566: lea edx, var_158
  0041256C: lea ecx, var_C8
  00412572: call MSVBVM60.DLL.__vbaVarDup
  00412574: xor edx, edx
  00412576: mov eax, arg_C
  00412579: cmp [eax], dx
  0041257C: setz dl
  0041257F: neg edx
  00412581: mov var_140, dx
  00412588: mov var_148, 0000000Bh
  00412592: lea ecx, var_D8
  00412598: push ecx
  00412599: lea edx, var_C8
  0041259F: push edx
  004125A0: lea eax, var_148
  004125A6: push eax
  004125A7: lea ecx, var_E8
  004125AD: push ecx
  004125AE: call [0040127Ch] ; IIf
  004125B4: mov eax, 80020004h
  004125B9: mov var_130, eax
  004125BF: mov ecx, 0000000Ah
  004125C4: mov var_138, ecx
  004125CA: mov var_120, eax
  004125D0: mov var_128, ecx
  004125D6: mov var_110, eax
  004125DC: mov var_118, ecx
  004125E2: mov var_170, 00405514h ; "Pastikan form"
  004125EC: mov var_178, edi
  004125F2: mov var_180, 00405550h ; "tanpa heading web sudah tampil di layar"
  004125FC: mov var_188, edi
  00412602: lea edx, var_138
  00412608: push edx
  00412609: lea eax, var_128
  0041260F: push eax
  00412610: lea ecx, var_118
  00412616: push ecx
  00412617: push 00000010h
  00412619: lea edx, var_178
  0041261F: push edx
  00412620: lea eax, var_E8
  00412626: push eax
  00412627: lea ecx, var_F8
  0041262D: push ecx
  0041262E: mov esi, [0040122Ch]
  00412634: call MSVBVM60.DLL.__vbaVarCat
  00412636: push eax
  00412637: lea edx, var_188
  0041263D: push edx
  0041263E: lea eax, var_108
  00412644: push eax
  00412645: call MSVBVM60.DLL.__vbaVarCat
  00412647: push eax
  00412648: call [004010C0h] ; MsgBox(arg_1, arg_2, arg_3, arg_4, arg_5)
  0041264E: lea ecx, var_138
  00412654: push ecx
  00412655: lea edx, var_128
  0041265B: push edx
  0041265C: lea eax, var_118
  00412662: push eax
  00412663: lea ecx, var_108
  00412669: push ecx
  0041266A: lea edx, var_F8
  00412670: push edx
  00412671: lea eax, var_E8
  00412677: push eax
  00412678: lea ecx, var_D8
  0041267E: push ecx
  0041267F: lea edx, var_C8
  00412685: push edx
  00412686: lea eax, var_148
  0041268C: push eax
  0041268D: push 00000009h
  0041268F: call MSVBVM60.DLL.__vbaFreeVarList
  00412695: add esp, 00000028h
  00412698: call MSVBVM60.DLL.__vbaExitProc
  0041269E: Method_8964E44D
  004126A3: jmp 41270Bh
  004126A5: lea ecx, var_B0
  004126AB: call MSVBVM60.DLL.__vbaFreeStr
  004126B1: lea ecx, var_B8
  004126B7: push ecx
  004126B8: lea edx, var_B4
  004126BE: push edx
  004126BF: push 00000002h
  004126C1: call MSVBVM60.DLL.__vbaFreeObjList
  004126C7: lea eax, var_138
  004126CD: push eax
  004126CE: lea ecx, var_128
  004126D4: push ecx
  004126D5: lea edx, var_118
  004126DB: push edx
  004126DC: lea eax, var_108
  004126E2: push eax
  004126E3: lea ecx, var_F8
  004126E9: push ecx
  004126EA: lea edx, var_E8
  004126F0: push edx
  004126F1: lea eax, var_D8
  004126F7: push eax
  004126F8: lea ecx, var_C8
  004126FE: push ecx
  004126FF: push 00000008h
  00412701: call MSVBVM60.DLL.__vbaFreeVarList
  00412707: add esp, 00000030h
  0041270A: ret 
End Sub

Private sub Proc_3_12_414810
  00414810: push ebp
  00414811: mov ebp, esp
  00414813: sub esp, 00000008h
  00414816: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0041481B: mov eax, fs:[00h]
  00414821: push eax
  00414822: mov fs:[00000000h], esp
  00414829: sub esp, 0000004Ch
  0041482C: push ebx
  0041482D: push esi
  0041482E: push edi
  0041482F: mov var_8, esp
  00414832: mov var_4, 00401490h
  00414839: mov eax, arg_C
  0041483C: lea ecx, var_18
  0041483F: xor esi, esi
  00414841: push eax
  00414842: push ecx
  00414843: mov var_14, esi
  00414846: mov var_18, esi
  00414849: mov var_1C, esi
  0041484C: mov var_20, esi
  0041484F: mov var_24, esi
  00414852: mov var_34, esi
  00414855: mov var_44, esi
  00414858: mov var_54, esi
  0041485B: call MSVBVM60.DLL.__vbaObjSetAddref
  00414861: mov eax, arg_8
  00414864: push esi
  00414865: push 000000C8h
  0041486A: push eax
  0041486B: mov edx, [eax]
  0041486D: call [edx+00000344h]
  00414873: mov esi, [004010BCh] ; Set (object)
  00414879: push eax
  0041487A: lea eax, var_24
  0041487D: push eax
  0041487E: call Set (object)
  00414880: mov edi, MSVBVM60.DLL.__vbaLateIdCallLd
  00414886: lea ecx, var_34
  00414889: push eax
  0041488A: push ecx
  0041488B: call edi
  0041488D: add esp, 00000010h
  00414890: push eax
  00414891: call MSVBVM60.DLL.__vbaUnkVar
  00414897: mov edx, var_18
  0041489A: push eax
  0041489B: push edx
  0041489C: call MSVBVM60.DLL.__vbaObjIs
  004148A2: mov ebx, MSVBVM60.DLL.__vbaFreeObj
  004148A8: lea ecx, var_24
  004148AB: mov var_58, eax
  004148AE: call ebx
  004148B0: lea ecx, var_34
  004148B3: call MSVBVM60.DLL.__vbaFreeVar
  004148B9: cmp word ptr var_58, 0000h
  004148BE: jz 00414A5Ch
  004148C4: mov edx, arg_10
  004148C7: lea ecx, var_54
  004148CA: mov var_3C, 004062ECh ; "http://www.google.com/"
  004148D1: mov var_44, 00008008h
  004148D8: call MSVBVM60.DLL.__vbaVarVargNofree
  004148DE: push eax
  004148DF: lea eax, var_44
  004148E2: push eax
  004148E3: call MSVBVM60.DLL.__vbaVarTstEq
  004148E9: test ax, ax
  004148EC: jz 00414A5Ch
  004148F2: mov eax, arg_8
  004148F5: push 00404ED8h
  004148FA: push 00000000h
  004148FC: push 000000CBh
  00414901: mov ecx, [eax]
  00414903: push eax
  00414904: call [ecx+00000344h]
  0041490A: lea edx, var_24
  0041490D: push eax
  0041490E: push edx
  0041490F: call Set (object)
  00414911: push eax
  00414912: lea eax, var_34
  00414915: push eax
  00414916: call edi
  00414918: add esp, 00000010h
  0041491B: push eax
  0041491C: call MSVBVM60.DLL.__vbaCastObjVar
  00414922: lea ecx, var_14
  00414925: push eax
  00414926: push ecx
  00414927: call Set (object)
  00414929: lea ecx, var_24
  0041492C: call ebx
  0041492E: lea ecx, var_34
  00414931: call MSVBVM60.DLL.__vbaFreeVar
  00414937: push 00406324h
  0041493C: push 004051F8h
  00414941: sub esp, 00000010h
  00414944: mov ecx, 00000008h
  00414949: mov edx, esp
  0041494B: mov var_44, ecx
  0041494E: mov eax, 00406320h
  00414953: push 00000001h
  00414955: mov [edx], ecx
  00414957: mov ecx, var_40
  0041495A: mov var_3C, eax
  0041495D: push 00000440h
  00414962: mov [edx+04h], ecx
  00414965: mov ecx, var_14
  00414968: push ecx
  00414969: mov [edx+08h], eax
  0041496C: mov eax, var_38
  0041496F: mov [edx+0Ch], eax
  00414972: lea edx, var_34
  00414975: push edx
  00414976: call edi
  00414978: add esp, 00000020h
  0041497B: push eax
  0041497C: call MSVBVM60.DLL.__vbaCastObjVar
  00414982: push eax
  00414983: lea eax, var_24
  00414986: push eax
  00414987: call Set (object)
  00414989: push eax
  0041498A: call MSVBVM60.DLL.__vbaCastObj
  00414990: lea ecx, var_1C
  00414993: push eax
  00414994: push ecx
  00414995: call Set (object)
  00414997: lea ecx, var_24
  0041499A: call ebx
  0041499C: lea ecx, var_34
  0041499F: call MSVBVM60.DLL.__vbaFreeVar
  004149A5: push 00406120h
  004149AA: push 004051F8h
  004149AF: sub esp, 00000010h
  004149B2: mov ecx, 00000008h
  004149B7: mov edx, esp
  004149B9: mov var_44, ecx
  004149BC: mov eax, 00406518h ; "btnG"
  004149C1: push 00000001h
  004149C3: mov [edx], ecx
  004149C5: mov ecx, var_40
  004149C8: mov var_3C, eax
  004149CB: push 00000440h
  004149D0: mov [edx+04h], ecx
  004149D3: mov ecx, var_14
  004149D6: mov [edx+08h], eax
  004149D9: mov eax, var_38
  004149DC: mov [edx+0Ch], eax
  004149DF: lea edx, var_34
  004149E2: push ecx
  004149E3: push edx
  004149E4: call edi
  004149E6: add esp, 00000020h
  004149E9: push eax
  004149EA: call MSVBVM60.DLL.__vbaCastObjVar
  004149F0: push eax
  004149F1: lea eax, var_24
  004149F4: push eax
  004149F5: call Set (object)
  004149F7: push eax
  004149F8: call MSVBVM60.DLL.__vbaCastObj
  004149FE: lea ecx, var_20
  00414A01: push eax
  00414A02: push ecx
  00414A03: call Set (object)
  00414A05: lea ecx, var_24
  00414A08: call ebx
  00414A0A: lea ecx, var_34
  00414A0D: call MSVBVM60.DLL.__vbaFreeVar
  00414A13: sub esp, 00000010h
  00414A16: mov ecx, 00000008h
  00414A1B: mov edx, esp
  00414A1D: mov var_44, ecx
  00414A20: mov eax, 00405794h ; "Looking for answers"
  00414A25: push 800113EDh
  00414A2A: mov [edx], ecx
  00414A2C: mov ecx, var_40
  00414A2F: mov var_3C, eax
  00414A32: mov [edx+04h], ecx
  00414A35: mov ecx, var_1C
  00414A38: push ecx
  00414A39: mov [edx+08h], eax
  00414A3C: mov eax, var_38
  00414A3F: mov [edx+0Ch], eax
  00414A42: call MSVBVM60.DLL.__vbaLateIdSt
  00414A48: mov edx, var_20
  00414A4B: push 00000000h
  00414A4D: push 80010409h
  00414A52: push edx
  00414A53: call MSVBVM60.DLL.__vbaLateIdCall
  00414A59: add esp, 0000000Ch
  00414A5C: push 00414A91h ; "‹Mð_^3Àd‰'#1"
  00414A61: jmp 414A76h
  00414A63: lea ecx, var_24
  00414A66: call MSVBVM60.DLL.__vbaFreeObj
  00414A6C: lea ecx, var_34
  00414A6F: call MSVBVM60.DLL.__vbaFreeVar
  00414A75: ret 
End Sub

Private sub Proc_3_13_414AB0
  00414AB0: push ebp
  00414AB1: mov ebp, esp
  00414AB3: sub esp, 00000008h
  00414AB6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00414ABB: mov eax, fs:[00h]
  00414AC1: push eax
  00414AC2: mov fs:[00000000h], esp
  00414AC9: sub esp, 0000004Ch
  00414ACC: push ebx
  00414ACD: push esi
  00414ACE: push edi
  00414ACF: mov var_8, esp
  00414AD2: mov var_4, 004014A0h
  00414AD9: sub esp, 00000010h
  00414ADC: mov edi, arg_8
  00414ADF: mov edx, esp
  00414AE1: mov ecx, 00000008h
  00414AE6: xor eax, eax
  00414AE8: push 00000001h
  00414AEA: mov [edx], ecx
  00414AEC: mov ecx, var_34
  00414AEF: mov var_14, eax
  00414AF2: mov var_18, eax
  00414AF5: mov var_28, eax
  00414AF8: mov [edx+04h], ecx
  00414AFB: mov ecx, [edi]
  00414AFD: mov eax, 004068E4h ; "http://spiderevaluation.com"
  00414B02: mov [edx+08h], eax
  00414B05: mov eax, var_2C
  00414B08: push 000001F4h
  00414B0D: push edi
  00414B0E: mov [edx+0Ch], eax
  00414B11: call [ecx+00000344h]
  00414B17: mov esi, [004010BCh] ; Set (object)
  00414B1D: lea edx, var_14
  00414B20: push eax
  00414B21: push edx
  00414B22: call Set (object)
  00414B24: push eax
  00414B25: call MSVBVM60.DLL.__vbaLateIdCall
  00414B2B: add esp, 0000001Ch
  00414B2E: lea ecx, var_14
  00414B31: call MSVBVM60.DLL.__vbaFreeObj
  00414B37: mov ebx, MSVBVM60.DLL.__vbaLateIdCallLd
  00414B3D: mov eax, [edi]
  00414B3F: push 00000000h
  00414B41: push FFFFFDF3h
  00414B46: push edi
  00414B47: call [eax+00000344h]
  00414B4D: lea ecx, var_14
  00414B50: push eax
  00414B51: push ecx
  00414B52: call Set (object)
  00414B54: lea edx, var_28
  00414B57: push eax
  00414B58: push edx
  00414B59: call ebx
  00414B5B: add esp, 00000010h
  00414B5E: push eax
  00414B5F: call MSVBVM60.DLL.__vbaI4Var
  00414B65: xor ecx, ecx
  00414B67: cmp eax, 00000004h
  00414B6A: setnz cl
  00414B6D: neg ecx
  00414B6F: mov var_4C, ecx
  00414B72: lea ecx, var_14
  00414B75: call MSVBVM60.DLL.__vbaFreeObj
  00414B7B: lea ecx, var_28
  00414B7E: call MSVBVM60.DLL.__vbaFreeVar
  00414B84: cmp word ptr var_4C, 0000h
  00414B89: jz 414B93h
  00414B8B: call [004010E0h] ; DoEvents
  00414B91: jmp 414B3Dh
  00414B93: mov edx, [edi]
  00414B95: push 00404ED8h
  00414B9A: push 00000000h
  00414B9C: push 000000CBh
  00414BA1: push edi
  00414BA2: call [edx+00000344h]
  00414BA8: push eax
  00414BA9: lea eax, var_14
  00414BAC: push eax
  00414BAD: call Set (object)
  00414BAF: lea ecx, var_28
  00414BB2: push eax
  00414BB3: push ecx
  00414BB4: call ebx
  00414BB6: add esp, 00000010h
  00414BB9: push eax
  00414BBA: call MSVBVM60.DLL.__vbaCastObjVar
  00414BC0: lea edx, var_18
  00414BC3: push eax
  00414BC4: push edx
  00414BC5: call Set (object)
  00414BC7: push eax
  00414BC8: lea eax, [edi+34h]
  00414BCB: push eax
  00414BCC: call MSVBVM60.DLL.__vbaObjSetAddref
  00414BD2: lea eax, var_18
  00414BD5: lea ecx, var_14
  00414BD8: push eax
  00414BD9: push ecx
  00414BDA: push 00000002h
  00414BDC: call MSVBVM60.DLL.__vbaFreeObjList
  00414BE2: add esp, 0000000Ch
  00414BE5: lea ecx, var_28
  00414BE8: call MSVBVM60.DLL.__vbaFreeVar
  00414BEE: push 00406940h
  00414BF3: push 004051F8h
  00414BF8: sub esp, 00000010h
  00414BFB: mov ecx, 00000008h
  00414C00: mov edx, esp
  00414C02: mov eax, 00406920h ; "btnAcknowledge"
  00414C07: push 00000001h
  00414C09: push 00000440h
  00414C0E: mov [edx], ecx
  00414C10: mov ecx, var_34
  00414C13: mov [edx+04h], ecx
  00414C16: mov ecx, [edi+34h]
  00414C19: push ecx
  00414C1A: mov [edx+08h], eax
  00414C1D: mov eax, var_2C
  00414C20: mov [edx+0Ch], eax
  00414C23: lea edx, var_28
  00414C26: push edx
  00414C27: call ebx
  00414C29: add esp, 00000020h
  00414C2C: push eax
  00414C2D: call MSVBVM60.DLL.__vbaCastObjVar
  00414C33: push eax
  00414C34: lea eax, var_14
  00414C37: push eax
  00414C38: call Set (object)
  00414C3A: push eax
  00414C3B: call MSVBVM60.DLL.__vbaCastObj
  00414C41: lea ecx, var_18
  00414C44: push eax
  00414C45: push ecx
  00414C46: call Set (object)
  00414C48: push eax
  00414C49: lea eax, [edi+38h]
  00414C4C: push eax
  00414C4D: call MSVBVM60.DLL.__vbaObjSetAddref
  00414C53: lea edx, var_18
  00414C56: lea eax, var_14
  00414C59: push edx
  00414C5A: push eax
  00414C5B: push 00000002h
  00414C5D: call MSVBVM60.DLL.__vbaFreeObjList
  00414C63: add esp, 0000000Ch
  00414C66: lea ecx, var_28
  00414C69: call MSVBVM60.DLL.__vbaFreeVar
  00414C6F: push 00000000h
  00414C71: push 800107D0h
  00414C76: mov ecx, [edi+38h]
  00414C79: push ecx
  00414C7A: call MSVBVM60.DLL.__vbaLateIdCall
  00414C80: mov edx, [edi+38h]
  00414C83: push 00000000h
  00414C85: push 80010409h
  00414C8A: push edx
  00414C8B: call MSVBVM60.DLL.__vbaLateIdCall
  00414C91: add esp, 00000018h
  00414C94: mov eax, [edi]
  00414C96: push 00000000h
  00414C98: push FFFFFDF3h
  00414C9D: push edi
  00414C9E: call [eax+00000344h]
  00414CA4: lea ecx, var_14
  00414CA7: push eax
  00414CA8: push ecx
  00414CA9: call Set (object)
  00414CAB: lea edx, var_28
  00414CAE: push eax
  00414CAF: push edx
  00414CB0: call ebx
  00414CB2: add esp, 00000010h
  00414CB5: push eax
  00414CB6: call MSVBVM60.DLL.__vbaI4Var
  00414CBC: xor ecx, ecx
  00414CBE: cmp eax, 00000004h
  00414CC1: setnz cl
  00414CC4: neg ecx
  00414CC6: mov var_4C, ecx
  00414CC9: lea ecx, var_14
  00414CCC: call MSVBVM60.DLL.__vbaFreeObj
  00414CD2: lea ecx, var_28
  00414CD5: call MSVBVM60.DLL.__vbaFreeVar
  00414CDB: cmp word ptr var_4C, 0000h
  00414CE0: jz 414CEAh
  00414CE2: call [004010E0h] ; DoEvents
  00414CE8: jmp 414C94h
  00414CEA: mov eax, [edi]
  00414CEC: push 00404ED8h
  00414CF1: push 00000000h
  00414CF3: lea edx, [edi+34h]
  00414CF6: push 000000CBh
  00414CFB: push edi
  00414CFC: mov var_54, edx
  00414CFF: call [eax+00000344h]
  00414D05: lea ecx, var_14
  00414D08: push eax
  00414D09: push ecx
  00414D0A: call Set (object)
  00414D0C: lea edx, var_28
  00414D0F: push eax
  00414D10: push edx
  00414D11: call ebx
  00414D13: add esp, 00000010h
  00414D16: push eax
  00414D17: call MSVBVM60.DLL.__vbaCastObjVar
  00414D1D: push eax
  00414D1E: lea eax, var_18
  00414D21: push eax
  00414D22: call Set (object)
  00414D24: mov ecx, var_54
  00414D27: push eax
  00414D28: push ecx
  00414D29: call MSVBVM60.DLL.__vbaObjSetAddref
  00414D2F: lea edx, var_18
  00414D32: lea eax, var_14
  00414D35: push edx
  00414D36: push eax
  00414D37: push 00000002h
  00414D39: call MSVBVM60.DLL.__vbaFreeObjList
  00414D3F: add esp, 0000000Ch
  00414D42: lea ecx, var_28
  00414D45: call MSVBVM60.DLL.__vbaFreeVar
  00414D4B: push 004066C0h
  00414D50: push 004051F8h
  00414D55: lea edx, [edi+40h]
  00414D58: sub esp, 00000010h
  00414D5B: mov var_58, edx
  00414D5E: mov edx, esp
  00414D60: mov ecx, 00000008h
  00414D65: mov eax, 00406954h ; "txtUsername"
  00414D6A: mov [edx], ecx
  00414D6C: mov ecx, var_34
  00414D6F: push 00000001h
  00414D71: push 00000440h
  00414D76: mov [edx+04h], ecx
  00414D79: mov ecx, var_54
  00414D7C: mov [edx+08h], eax
  00414D7F: mov eax, var_2C
  00414D82: mov [edx+0Ch], eax
  00414D85: mov edx, [ecx]
  00414D87: lea eax, var_28
  00414D8A: push edx
  00414D8B: push eax
  00414D8C: call ebx
  00414D8E: add esp, 00000020h
  00414D91: push eax
  00414D92: call MSVBVM60.DLL.__vbaCastObjVar
  00414D98: lea ecx, var_14
  00414D9B: push eax
  00414D9C: push ecx
  00414D9D: call Set (object)
  00414D9F: push eax
  00414DA0: call MSVBVM60.DLL.__vbaCastObj
  00414DA6: lea edx, var_18
  00414DA9: push eax
  00414DAA: push edx
  00414DAB: call Set (object)
  00414DAD: push eax
  00414DAE: mov eax, var_58
  00414DB1: push eax
  00414DB2: call MSVBVM60.DLL.__vbaObjSetAddref
  00414DB8: lea ecx, var_18
  00414DBB: lea edx, var_14
  00414DBE: push ecx
  00414DBF: push edx
  00414DC0: push 00000002h
  00414DC2: call MSVBVM60.DLL.__vbaFreeObjList
  00414DC8: add esp, 0000000Ch
  00414DCB: lea ecx, var_28
  00414DCE: call MSVBVM60.DLL.__vbaFreeVar
  00414DD4: push 004066C0h
  00414DD9: push 004051F8h
  00414DDE: lea edx, [edi+44h]
  00414DE1: sub esp, 00000010h
  00414DE4: mov var_5C, edx
  00414DE7: mov edx, esp
  00414DE9: mov ecx, 00000008h
  00414DEE: mov eax, 00406970h ; "txtPassword"
  00414DF3: mov [edx], ecx
  00414DF5: mov ecx, var_34
  00414DF8: push 00000001h
  00414DFA: push 00000440h
  00414DFF: mov [edx+04h], ecx
  00414E02: mov ecx, var_54
  00414E05: mov [edx+08h], eax
  00414E08: mov eax, var_2C
  00414E0B: mov [edx+0Ch], eax
  00414E0E: mov edx, [ecx]
  00414E10: lea eax, var_28
  00414E13: push edx
  00414E14: push eax
  00414E15: call ebx
  00414E17: add esp, 00000020h
  00414E1A: push eax
  00414E1B: call MSVBVM60.DLL.__vbaCastObjVar
  00414E21: lea ecx, var_14
  00414E24: push eax
  00414E25: push ecx
  00414E26: call Set (object)
  00414E28: push eax
  00414E29: call MSVBVM60.DLL.__vbaCastObj
  00414E2F: lea edx, var_18
  00414E32: push eax
  00414E33: push edx
  00414E34: call Set (object)
  00414E36: push eax
  00414E37: mov eax, var_5C
  00414E3A: push eax
  00414E3B: call MSVBVM60.DLL.__vbaObjSetAddref
  00414E41: lea ecx, var_18
  00414E44: lea edx, var_14
  00414E47: push ecx
  00414E48: push edx
  00414E49: push 00000002h
  00414E4B: call MSVBVM60.DLL.__vbaFreeObjList
  00414E51: add esp, 0000000Ch
  00414E54: lea ecx, var_28
  00414E57: call MSVBVM60.DLL.__vbaFreeVar
  00414E5D: push 00406940h
  00414E62: push 004051F8h
  00414E67: sub esp, 00000010h
  00414E6A: mov ecx, 00000008h
  00414E6F: mov edx, esp
  00414E71: mov eax, 0040698Ch ; "btnLogin"
  00414E76: push 00000001h
  00414E78: push 00000440h
  00414E7D: mov [edx], ecx
  00414E7F: mov ecx, var_34
  00414E82: add edi, 0000003Ch
  00414E85: mov [edx+04h], ecx
  00414E88: mov ecx, var_54
  00414E8B: mov [edx+08h], eax
  00414E8E: mov eax, var_2C
  00414E91: mov [edx+0Ch], eax
  00414E94: mov edx, [ecx]
  00414E96: lea eax, var_28
  00414E99: push edx
  00414E9A: push eax
  00414E9B: call ebx
  00414E9D: add esp, 00000020h
  00414EA0: push eax
  00414EA1: call MSVBVM60.DLL.__vbaCastObjVar
  00414EA7: lea ecx, var_14
  00414EAA: push eax
  00414EAB: push ecx
  00414EAC: call Set (object)
  00414EAE: push eax
  00414EAF: call MSVBVM60.DLL.__vbaCastObj
  00414EB5: lea edx, var_18
  00414EB8: push eax
  00414EB9: push edx
  00414EBA: call Set (object)
  00414EBC: push eax
  00414EBD: push edi
  00414EBE: call MSVBVM60.DLL.__vbaObjSetAddref
  00414EC4: lea eax, var_18
  00414EC7: lea ecx, var_14
  00414ECA: push eax
  00414ECB: push ecx
  00414ECC: push 00000002h
  00414ECE: call MSVBVM60.DLL.__vbaFreeObjList
  00414ED4: add esp, 0000000Ch
  00414ED7: lea ecx, var_28
  00414EDA: call MSVBVM60.DLL.__vbaFreeVar
  00414EE0: mov ebx, var_58
  00414EE3: mov esi, [0040102Ch]
  00414EE9: push 00000000h
  00414EEB: push 800107D0h
  00414EF0: mov edx, [ebx]
  00414EF2: push edx
  00414EF3: call MSVBVM60.DLL.__vbaLateIdCall
  00414EF5: mov ecx, 00000008h
  00414EFA: mov eax, 004069A4h ; "wolf"
  00414EFF: push ecx
  00414F00: mov edx, esp
  00414F02: push 800113EDh
  00414F07: mov [edx], ecx
  00414F09: mov ecx, var_34
  00414F0C: mov [edx+04h], ecx
  00414F0F: mov ecx, [ebx]
  00414F11: mov ebx, MSVBVM60.DLL.__vbaLateIdSt
  00414F17: push ecx
  00414F18: mov [edx+08h], eax
  00414F1B: mov eax, var_2C
  00414F1E: mov [edx+0Ch], eax
  00414F21: call ebx
  00414F23: mov edx, var_5C
  00414F26: push 00000000h
  00414F28: push 800107D0h
  00414F2D: mov eax, [edx]
  00414F2F: push eax
  00414F30: call MSVBVM60.DLL.__vbaLateIdCall
  00414F32: mov ecx, 00000008h
  00414F37: mov eax, 004069B4h ; "spider"
  00414F3C: push ecx
  00414F3D: mov edx, esp
  00414F3F: push 800113EDh
  00414F44: mov [edx], ecx
  00414F46: mov ecx, var_34
  00414F49: mov [edx+04h], ecx
  00414F4C: mov ecx, var_5C
  00414F4F: mov [edx+08h], eax
  00414F52: mov eax, var_2C
  00414F55: mov [edx+0Ch], eax
  00414F58: mov edx, [ecx]
  00414F5A: push edx
  00414F5B: call ebx
  00414F5D: mov eax, [edi]
  00414F5F: push 00000000h
  00414F61: push 800107D0h
  00414F66: push eax
  00414F67: call MSVBVM60.DLL.__vbaLateIdCall
  00414F69: mov ecx, [edi]
  00414F6B: push 00000000h
  00414F6D: push 80010409h
  00414F72: push ecx
  00414F73: call MSVBVM60.DLL.__vbaLateIdCall
  00414F75: add esp, 00000018h
  00414F78: push 00414F9Dh ; "‹Mð_^3Àd‰'#1"
  00414F7D: jmp 414F9Ch
  00414F7F: lea edx, var_18
  00414F82: lea eax, var_14
  00414F85: push edx
  00414F86: push eax
  00414F87: push 00000002h
  00414F89: call MSVBVM60.DLL.__vbaFreeObjList
  00414F8F: add esp, 0000000Ch
  00414F92: lea ecx, var_28
  00414F95: call MSVBVM60.DLL.__vbaFreeVar
  00414F9B: ret 
End Sub

