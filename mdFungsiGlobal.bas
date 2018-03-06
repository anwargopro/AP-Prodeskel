
Private sub Proc_4_0_416310
  00416310: push ebp
  00416311: mov ebp, esp
  00416313: sub esp, 0000000Ch
  00416316: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0041631B: mov eax, fs:[00h]
  00416321: push eax
  00416322: mov fs:[00000000h], esp
  00416329: sub esp, 000001C4h
  0041632F: push ebx
  00416330: push esi
  00416331: push edi
  00416332: mov var_C, esp
  00416335: mov var_8, 00401548h
  0041633C: mov edi, arg_10
  0041633F: mov ebx, MSVBVM60.DLL.__vbaVarVargNofree
  00416345: xor esi, esi
  00416347: mov edx, edi
  00416349: lea ecx, var_190
  0041634F: mov var_24, esi
  00416352: mov var_34, esi
  00416355: mov var_38, esi
  00416358: mov var_48, esi
  0041635B: mov var_4C, esi
  0041635E: mov var_50, esi
  00416361: mov var_60, esi
  00416364: mov var_70, esi
  00416367: mov var_80, esi
  0041636A: mov var_90, esi
  00416370: mov var_A0, esi
  00416376: mov var_B0, esi
  0041637C: mov var_C0, esi
  00416382: mov var_D0, esi
  00416388: mov var_E0, esi
  0041638E: mov var_F0, esi
  00416394: mov var_100, esi
  0041639A: mov var_110, esi
  004163A0: mov var_120, esi
  004163A6: mov var_130, esi
  004163AC: mov var_140, esi
  004163B2: mov var_150, esi
  004163B8: mov var_160, esi
  004163BE: mov var_170, esi
  004163C4: mov var_190, esi
  004163CA: mov var_1A0, esi
  004163D0: mov var_1B0, esi
  004163D6: mov var_178, 00403558h
  004163E0: mov var_180, 00008008h
  004163EA: call ebx
  004163EC: push eax
  004163ED: lea eax, var_180
  004163F3: push eax
  004163F4: call MSVBVM60.DLL.__vbaVarTstEq
  004163FA: test ax, ax
  004163FD: jz 41640Ah
  004163FF: lea edx, var_24
  00416402: mov ecx, edi
  00416404: call MSVBVM60.DLL.__vbaVargVarCopy
  0041640A: push esi
  0041640B: push FFFFFFFFh
  0041640D: push 00000001h
  0041640F: push 00403558h
  00416414: push 00404034h
  00416419: mov edx, edi
  0041641B: lea ecx, var_180
  00416421: call ebx
  00416423: mov esi, [0040122Ch]
  00416429: lea ecx, var_24
  0041642C: push eax
  0041642D: lea edx, var_60
  00416430: push ecx
  00416431: push edx
  00416432: call MSVBVM60.DLL.__vbaVarCat
  00416434: push eax
  00416435: lea eax, var_4C
  00416438: push eax
  00416439: call MSVBVM60.DLL.__vbaStrVarVal
  0041643F: push eax
  00416440: call [004011ECh] ; Replace(arg_1, arg_2, arg_3, arg_4, arg_5, arg_6)
  00416446: lea edx, var_70
  00416449: mov ecx, edi
  0041644B: mov var_68, eax
  0041644E: mov var_70, 00000008h
  00416455: call MSVBVM60.DLL.__vbaVargVarMove
  0041645B: lea ecx, var_4C
  0041645E: call MSVBVM60.DLL.__vbaFreeStr
  00416464: lea ecx, var_60
  00416467: call MSVBVM60.DLL.__vbaFreeVar
  0041646D: mov edx, arg_14
  00416470: lea ecx, var_190
  00416476: mov var_178, 00403644h
  00416480: mov var_180, 00008008h
  0041648A: call ebx
  0041648C: lea ecx, var_180
  00416492: push eax
  00416493: push ecx
  00416494: call MSVBVM60.DLL.__vbaVarTstEq
  0041649A: test ax, ax
  0041649D: jz 004165A8h
  004164A3: mov edi, [004012C0h] ; Date 'arg_1
  004164A9: lea edx, var_60
  004164AC: push edx
  004164AD: call edi
  004164AF: lea eax, var_A0
  004164B5: push eax
  004164B6: call edi
  004164B8: mov edi, MSVBVM60.DLL.__vbaVarDup
  004164BE: lea edx, var_180
  004164C4: lea ecx, var_70
  004164C7: mov var_178, 00406C78h ; "dd"
  004164D1: mov var_180, 00000008h
  004164DB: call edi
  004164DD: mov ebx, MSVBVM60.DLL.rtcVarFromFormatVar
  004164E3: push 00000001h
  004164E5: lea ecx, var_70
  004164E8: push 00000001h
  004164EA: lea edx, var_60
  004164ED: push ecx
  004164EE: lea eax, var_80
  004164F1: push edx
  004164F2: push eax
  004164F3: call ebx
  004164F5: mov eax, 00000008h
  004164FA: lea edx, var_1A0
  00416500: lea ecx, var_B0
  00416506: mov var_188, 00406C84h ; "-JS"
  00416510: mov var_190, eax
  00416516: mov var_198, 00406C90h ; "yy"
  00416520: mov var_1A0, eax
  00416526: call edi
  00416528: push 00000001h
  0041652A: lea ecx, var_B0
  00416530: push 00000001h
  00416532: lea edx, var_A0
  00416538: push ecx
  00416539: lea eax, var_C0
  0041653F: push edx
  00416540: push eax
  00416541: call ebx
  00416543: lea ecx, var_80
  00416546: lea edx, var_190
  0041654C: push ecx
  0041654D: lea eax, var_90
  00416553: push edx
  00416554: push eax
  00416555: call MSVBVM60.DLL.__vbaVarCat
  00416557: lea ecx, var_C0
  0041655D: push eax
  0041655E: lea edx, var_D0
  00416564: push ecx
  00416565: push edx
  00416566: call MSVBVM60.DLL.__vbaVarCat
  00416568: mov edx, eax
  0041656A: lea ecx, var_34
  0041656D: call MSVBVM60.DLL.__vbaVarMove
  00416573: lea eax, var_C0
  00416579: lea ecx, var_90
  0041657F: push eax
  00416580: lea edx, var_B0
  00416586: push ecx
  00416587: lea eax, var_A0
  0041658D: push edx
  0041658E: lea ecx, var_80
  00416591: push eax
  00416592: lea edx, var_70
  00416595: push ecx
  00416596: lea eax, var_60
  00416599: push edx
  0041659A: push eax
  0041659B: push 00000007h
  0041659D: call MSVBVM60.DLL.__vbaFreeVarList
  004165A3: add esp, 00000020h
  004165A6: jmp 4165CBh
  004165A8: lea edx, var_180
  004165AE: lea ecx, var_34
  004165B1: mov var_178, 00406C9Ch ; "10-JSPM"
  004165BB: mov var_180, 00000008h
  004165C5: call MSVBVM60.DLL.__vbaVarCopy
  004165CB: mov edi, arg_C
  004165CE: push 00000002h
  004165D0: lea ecx, var_60
  004165D3: push edi
  004165D4: push ecx
  004165D5: call [0040130Ch] ; Left(arg_1, arg_2)
  004165DB: mov ebx, [00401124h] ; arg_1 = Mid$(arg_2, arg_3, arg_4)
  004165E1: lea edx, var_80
  004165E4: push edx
  004165E5: push 00000003h
  004165E7: lea eax, var_90
  004165ED: push edi
  004165EE: push eax
  004165EF: mov var_78, 00000001h
  004165F6: mov var_80, 00000002h
  004165FD: call ebx
  004165FF: lea ecx, var_B0
  00416605: lea edx, var_C0
  0041660B: push ecx
  0041660C: push 00000005h
  0041660E: push edi
  0041660F: push edx
  00416610: mov var_A8, 00000001h
  0041661A: mov var_B0, 00000002h
  00416624: call ebx
  00416626: mov ecx, arg_10
  00416629: lea eax, var_F0
  0041662F: push eax
  00416630: push 00000003h
  00416632: lea edx, var_100
  00416638: push ecx
  00416639: push edx
  0041663A: mov var_198, 00406CB0h
  00416644: mov var_1A0, 00000008h
  0041664E: mov var_E8, 00000001h
  00416658: mov var_F0, 00000002h
  00416662: call ebx
  00416664: mov ecx, arg_10
  00416667: lea eax, var_120
  0041666D: push eax
  0041666E: push 00000007h
  00416670: lea edx, var_130
  00416676: push ecx
  00416677: push edx
  00416678: mov var_118, 00000001h
  00416682: mov var_120, 00000002h
  0041668C: call ebx
  0041668E: mov eax, 00000002h
  00416693: lea ecx, var_160
  00416699: mov var_148, eax
  0041669F: mov var_150, eax
  004166A5: lea eax, var_150
  004166AB: push eax
  004166AC: push 00000003h
  004166AE: push edi
  004166AF: push ecx
  004166B0: call ebx
  004166B2: lea edx, var_60
  004166B5: lea eax, var_34
  004166B8: push edx
  004166B9: lea ecx, var_70
  004166BC: push eax
  004166BD: push ecx
  004166BE: call MSVBVM60.DLL.__vbaVarCat
  004166C0: push eax
  004166C1: lea edx, var_90
  004166C7: lea eax, var_A0
  004166CD: push edx
  004166CE: push eax
  004166CF: call MSVBVM60.DLL.__vbaVarCat
  004166D1: lea ecx, var_C0
  004166D7: push eax
  004166D8: lea edx, var_D0
  004166DE: push ecx
  004166DF: push edx
  004166E0: call MSVBVM60.DLL.__vbaVarCat
  004166E2: push eax
  004166E3: lea eax, var_1A0
  004166E9: lea ecx, var_E0
  004166EF: push eax
  004166F0: push ecx
  004166F1: call MSVBVM60.DLL.__vbaVarCat
  004166F3: push eax
  004166F4: lea edx, var_100
  004166FA: lea eax, var_110
  00416700: push edx
  00416701: push eax
  00416702: call MSVBVM60.DLL.__vbaVarCat
  00416704: lea ecx, var_130
  0041670A: push eax
  0041670B: lea edx, var_140
  00416711: push ecx
  00416712: push edx
  00416713: call MSVBVM60.DLL.__vbaVarCat
  00416715: push eax
  00416716: lea eax, var_160
  0041671C: lea ecx, var_170
  00416722: push eax
  00416723: push ecx
  00416724: call MSVBVM60.DLL.__vbaVarCat
  00416726: push eax
  00416727: call MSVBVM60.DLL.__vbaStrVarMove
  0041672D: mov ebx, MSVBVM60.DLL.__vbaStrMove
  00416733: mov edx, eax
  00416735: lea ecx, var_38
  00416738: call ebx
  0041673A: lea edx, var_170
  00416740: lea eax, var_160
  00416746: push edx
  00416747: lea ecx, var_140
  0041674D: push eax
  0041674E: lea edx, var_150
  00416754: push ecx
  00416755: lea eax, var_130
  0041675B: push edx
  0041675C: lea ecx, var_110
  00416762: push eax
  00416763: lea edx, var_120
  00416769: push ecx
  0041676A: lea eax, var_100
  00416770: push edx
  00416771: lea ecx, var_E0
  00416777: push eax
  00416778: lea edx, var_F0
  0041677E: push ecx
  0041677F: lea eax, var_D0
  00416785: push edx
  00416786: lea ecx, var_C0
  0041678C: push eax
  0041678D: lea edx, var_A0
  00416793: push ecx
  00416794: lea eax, var_B0
  0041679A: push edx
  0041679B: lea ecx, var_90
  004167A1: push eax
  004167A2: lea edx, var_70
  004167A5: push ecx
  004167A6: lea eax, var_80
  004167A9: push edx
  004167AA: lea ecx, var_60
  004167AD: push eax
  004167AE: push ecx
  004167AF: push 00000012h
  004167B1: call MSVBVM60.DLL.__vbaFreeVarList
  004167B7: add esp, 0000004Ch
  004167BA: mov edx, edi
  004167BC: lea ecx, var_1A0
  004167C2: mov var_1A8, 00000005h
  004167CC: mov var_1B0, 00000002h
  004167D6: call MSVBVM60.DLL.__vbaVarVargNofree
  004167DC: lea edx, var_70
  004167DF: push eax
  004167E0: push edx
  004167E1: call [00401098h] ; arg_1 = Len(arg_2)
  004167E7: push eax
  004167E8: lea eax, var_1B0
  004167EE: lea ecx, var_80
  004167F1: push eax
  004167F2: push ecx
  004167F3: call MSVBVM60.DLL.__vbaVarSub
  004167F9: mov edx, eax
  004167FB: lea ecx, var_A0
  00416801: call MSVBVM60.DLL.__vbaVarMove
  00416807: lea edx, var_A0
  0041680D: lea eax, var_90
  00416813: push edx
  00416814: lea ecx, var_B0
  0041681A: push eax
  0041681B: push ecx
  0041681C: mov var_88, 00000005h
  00416826: mov var_90, 00000002h
  00416830: call 00416A10h
  00416835: lea edx, var_B0
  0041683B: push edx
  0041683C: call MSVBVM60.DLL.__vbaI4Var
  00416842: mov edx, arg_10
  00416845: push eax
  00416846: lea ecx, var_180
  0041684C: call MSVBVM60.DLL.__vbaVarVargNofree
  00416852: push eax
  00416853: mov edx, edi
  00416855: lea ecx, var_190
  0041685B: call MSVBVM60.DLL.__vbaVarVargNofree
  00416861: push eax
  00416862: lea eax, var_60
  00416865: push eax
  00416866: call MSVBVM60.DLL.__vbaVarCat
  00416868: lea ecx, var_C0
  0041686E: push eax
  0041686F: push ecx
  00416870: call [0040132Ch] ; Right(arg_1, arg_2)
  00416876: mov edx, var_38
  00416879: lea eax, var_C0
  0041687F: push edx
  00416880: push 00000000h
  00416882: push FFFFFFFFh
  00416884: push 00000001h
  00416886: push 00403558h
  0041688B: push 004047F4h
  00416890: lea ecx, var_4C
  00416893: push eax
  00416894: push ecx
  00416895: call MSVBVM60.DLL.__vbaStrVarVal
  0041689B: push eax
  0041689C: call [004011ECh] ; Replace(arg_1, arg_2, arg_3, arg_4, arg_5, arg_6)
  004168A2: mov edx, eax
  004168A4: lea ecx, var_50
  004168A7: call ebx
  004168A9: push eax
  004168AA: call [0040106Ch] ; &
  004168B0: mov edx, eax
  004168B2: lea ecx, var_38
  004168B5: call ebx
  004168B7: lea edx, var_50
  004168BA: lea eax, var_4C
  004168BD: push edx
  004168BE: push eax
  004168BF: push 00000002h
  004168C1: call MSVBVM60.DLL.__vbaFreeStrList
  004168C7: lea ecx, var_C0
  004168CD: lea edx, var_B0
  004168D3: push ecx
  004168D4: lea eax, var_60
  004168D7: push edx
  004168D8: lea ecx, var_A0
  004168DE: push eax
  004168DF: lea edx, var_90
  004168E5: push ecx
  004168E6: push edx
  004168E7: push 00000005h
  004168E9: call MSVBVM60.DLL.__vbaFreeVarList
  004168EF: mov eax, var_38
  004168F2: add esp, 00000024h
  004168F5: lea edx, var_180
  004168FB: lea ecx, var_48
  004168FE: mov var_178, eax
  00416904: mov var_180, 00000008h
  0041690E: call MSVBVM60.DLL.__vbaVarCopy
  00416914: push 004169D8h
  00416919: jmp 004169BEh
  0041691E: test byte ptr var_4, 04h
  00416922: jz 41692Dh
  00416924: lea ecx, var_48
  00416927: call MSVBVM60.DLL.__vbaFreeVar
  0041692D: lea ecx, var_50
  00416930: lea edx, var_4C
  00416933: push ecx
  00416934: push edx
  00416935: push 00000002h
  00416937: call MSVBVM60.DLL.__vbaFreeStrList
  0041693D: lea eax, var_170
  00416943: lea ecx, var_160
  00416949: push eax
  0041694A: lea edx, var_150
  00416950: push ecx
  00416951: lea eax, var_140
  00416957: push edx
  00416958: lea ecx, var_130
  0041695E: push eax
  0041695F: lea edx, var_120
  00416965: push ecx
  00416966: lea eax, var_110
  0041696C: push edx
  0041696D: lea ecx, var_100
  00416973: push eax
  00416974: lea edx, var_F0
  0041697A: push ecx
  0041697B: lea eax, var_E0
  00416981: push edx
  00416982: lea ecx, var_D0
  00416988: push eax
  00416989: lea edx, var_C0
  0041698F: push ecx
  00416990: lea eax, var_B0
  00416996: push edx
  00416997: lea ecx, var_A0
  0041699D: push eax
  0041699E: lea edx, var_90
  004169A4: push ecx
  004169A5: lea eax, var_80
  004169A8: push edx
  004169A9: lea ecx, var_70
  004169AC: push eax
  004169AD: lea edx, var_60
  004169B0: push ecx
  004169B1: push edx
  004169B2: push 00000012h
  004169B4: call MSVBVM60.DLL.__vbaFreeVarList
  004169BA: add esp, 00000058h
  004169BD: ret 
End Sub

Private sub Proc_4_1_416A10
  00416A10: push ebp
  00416A11: mov ebp, esp
  00416A13: sub esp, 0000000Ch
  00416A16: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00416A1B: mov eax, fs:[00h]
  00416A21: push eax
  00416A22: mov fs:[00000000h], esp
  00416A29: sub esp, 00000068h
  00416A2C: push ebx
  00416A2D: push esi
  00416A2E: push edi
  00416A2F: mov var_C, esp
  00416A32: mov var_8, 00401558h
  00416A39: mov esi, arg_C
  00416A3C: mov edi, MSVBVM60.DLL.__vbaVarVargNofree
  00416A42: xor eax, eax
  00416A44: mov edx, esi
  00416A46: lea ecx, var_64
  00416A49: mov var_24, eax
  00416A4C: mov var_34, eax
  00416A4F: mov var_44, eax
  00416A52: mov var_54, eax
  00416A55: mov var_64, eax
  00416A58: mov var_74, eax
  00416A5B: call edi
  00416A5D: mov ebx, arg_10
  00416A60: push eax
  00416A61: mov edx, ebx
  00416A63: lea ecx, var_74
  00416A66: call edi
  00416A68: push eax
  00416A69: lea eax, var_34
  00416A6C: push eax
  00416A6D: call MSVBVM60.DLL.__vbaVarCmpLt
  00416A73: mov edi, MSVBVM60.DLL.__vbaVarMove
  00416A79: mov edx, eax
  00416A7B: lea ecx, var_44
  00416A7E: call edi
  00416A80: push ebx
  00416A81: lea ecx, var_44
  00416A84: push esi
  00416A85: lea edx, var_54
  00416A88: push ecx
  00416A89: push edx
  00416A8A: call [0040127Ch] ; IIf
  00416A90: lea edx, var_54
  00416A93: lea ecx, var_24
  00416A96: call edi
  00416A98: lea ecx, var_44
  00416A9B: call MSVBVM60.DLL.__vbaFreeVar
  00416AA1: push 00416AD0h
  00416AA6: jmp 416ACFh
  00416AA8: test byte ptr var_4, 04h
  00416AAC: jz 416AB7h
  00416AAE: lea ecx, var_24
  00416AB1: call MSVBVM60.DLL.__vbaFreeVar
  00416AB7: lea eax, var_54
  00416ABA: lea ecx, var_44
  00416ABD: push eax
  00416ABE: lea edx, var_34
  00416AC1: push ecx
  00416AC2: push edx
  00416AC3: push 00000003h
  00416AC5: call MSVBVM60.DLL.__vbaFreeVarList
  00416ACB: add esp, 00000010h
  00416ACE: ret 
End Sub

Private sub Proc_4_2_416B00
  00416B00: push ebp
  00416B01: mov ebp, esp
  00416B03: sub esp, 0000000Ch
  00416B06: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00416B0B: mov eax, fs:[00h]
  00416B11: push eax
  00416B12: mov fs:[00000000h], esp
  00416B19: sub esp, 00000068h
  00416B1C: push ebx
  00416B1D: push esi
  00416B1E: push edi
  00416B1F: mov var_C, esp
  00416B22: mov var_8, 00401568h
  00416B29: mov esi, arg_C
  00416B2C: mov edi, MSVBVM60.DLL.__vbaVarVargNofree
  00416B32: xor eax, eax
  00416B34: mov edx, esi
  00416B36: lea ecx, var_64
  00416B39: mov var_24, eax
  00416B3C: mov var_34, eax
  00416B3F: mov var_44, eax
  00416B42: mov var_54, eax
  00416B45: mov var_64, eax
  00416B48: mov var_74, eax
  00416B4B: call edi
  00416B4D: mov ebx, arg_10
  00416B50: push eax
  00416B51: mov edx, ebx
  00416B53: lea ecx, var_74
  00416B56: call edi
  00416B58: push eax
  00416B59: lea eax, var_34
  00416B5C: push eax
  00416B5D: call MSVBVM60.DLL.__vbaVarCmpGt
  00416B63: mov edi, MSVBVM60.DLL.__vbaVarMove
  00416B69: mov edx, eax
  00416B6B: lea ecx, var_44
  00416B6E: call edi
  00416B70: push ebx
  00416B71: lea ecx, var_44
  00416B74: push esi
  00416B75: lea edx, var_54
  00416B78: push ecx
  00416B79: push edx
  00416B7A: call [0040127Ch] ; IIf
  00416B80: lea edx, var_54
  00416B83: lea ecx, var_24
  00416B86: call edi
  00416B88: lea ecx, var_44
  00416B8B: call MSVBVM60.DLL.__vbaFreeVar
  00416B91: push 00416BC0h
  00416B96: jmp 416BBFh
  00416B98: test byte ptr var_4, 04h
  00416B9C: jz 416BA7h
  00416B9E: lea ecx, var_24
  00416BA1: call MSVBVM60.DLL.__vbaFreeVar
  00416BA7: lea eax, var_54
  00416BAA: lea ecx, var_44
  00416BAD: push eax
  00416BAE: lea edx, var_34
  00416BB1: push ecx
  00416BB2: push edx
  00416BB3: push 00000003h
  00416BB5: call MSVBVM60.DLL.__vbaFreeVarList
  00416BBB: add esp, 00000010h
  00416BBE: ret 
End Sub

Private sub Proc_4_3_416BF0
  00416BF0: push ebp
  00416BF1: mov ebp, esp
  00416BF3: sub esp, 00000018h
  00416BF6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00416BFB: mov eax, fs:[00h]
  00416C01: push eax
  00416C02: mov fs:[00000000h], esp
  00416C09: mov eax, 00000278h
  00416C0E: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00416C13: push ebx
  00416C14: push esi
  00416C15: push edi
  00416C16: mov var_18, esp
  00416C19: mov var_14, 00401578h
  00416C20: mov var_10, 00000000h
  00416C27: mov var_C, 00000000h
  00416C2E: mov var_4, 00000001h
  00416C35: push 00000005h
  00416C37: push 004095D0h
  00416C3C: lea eax, var_74
  00416C3F: push eax
  00416C40: call MSVBVM60.DLL.__vbaAryConstruct2
  00416C46: mov var_4, 00000002h
  00416C4D: mov var_184, 00403558h
  00416C57: mov var_18C, 00008008h
  00416C61: mov edx, arg_C
  00416C64: lea ecx, var_19C
  00416C6A: call MSVBVM60.DLL.__vbaVarVargNofree
  00416C70: push eax
  00416C71: lea ecx, var_18C
  00416C77: push ecx
  00416C78: call MSVBVM60.DLL.__vbaVarTstEq
  00416C7E: movsx edx, ax
  00416C81: test edx, edx
  00416C83: jz 416CAFh
  00416C85: mov var_4, 00000003h
  00416C8C: mov var_184, 00409438h ; "11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11"
  00416C96: mov var_18C, 00000008h
  00416CA0: lea edx, var_18C
  00416CA6: mov ecx, arg_C
  00416CA9: call MSVBVM60.DLL.__vbaVargVarCopy
  00416CAF: mov var_4, 00000005h
  00416CB6: push 00000000h
  00416CB8: push 00000002h
  00416CBA: push 00409528h
  00416CBF: push 00000000h
  00416CC1: push FFFFFE00h
  00416CC6: mov eax, arg_8
  00416CC9: mov ecx, [eax]
  00416CCB: push ecx
  00416CCC: lea edx, var_AC
  00416CD2: push edx
  00416CD3: call MSVBVM60.DLL.__vbaLateIdCallLd
  00416CD9: add esp, 00000010h
  00416CDC: push eax
  00416CDD: call MSVBVM60.DLL.__vbaCastObjVar
  00416CE3: push eax
  00416CE4: lea eax, var_9C
  00416CEA: push eax
  00416CEB: call [004010BCh] ; Set (object)
  00416CF1: push eax
  00416CF2: lea ecx, var_BC
  00416CF8: push ecx
  00416CF9: call MSVBVM60.DLL.__vbaLateIdCallLd
  00416CFF: add esp, 00000010h
  00416D02: push eax
  00416D03: call MSVBVM60.DLL.__vbaCyVar
  00416D09: push edx
  00416D0A: push eax
  00416D0B: push 0000012Ch
  00416D10: call MSVBVM60.DLL.__vbaCyMulI2
  00416D16: push edx
  00416D17: push eax
  00416D18: call MSVBVM60.DLL.__vbaR8Cy
  00416D1E: cmp [00419000h], 00000000h
  00416D25: jnz 416D2Fh
  00416D27: fdiv real8 ptr [00401608h] ; 
  00416D2D: jmp 416D40h
  00416D2F: push [0040160Ch] ; 
  00416D35: push [00401608h] ; 
  00416D3B: call 00401694h ; MSVBVM60.DLL._adj_fdiv_m64
  00416D40: fstp real8 ptr var_184
  00416D46: fstsw ax
  00416D48: test al, 0Dh
  00416D4A: jnz 00417CEAh
  00416D50: mov var_18C, 00000005h
  00416D5A: lea edx, var_18C
  00416D60: lea ecx, var_88
  00416D66: call MSVBVM60.DLL.__vbaVarMove
  00416D6C: lea ecx, var_9C
  00416D72: call MSVBVM60.DLL.__vbaFreeObj
  00416D78: lea edx, var_BC
  00416D7E: push edx
  00416D7F: lea eax, var_AC
  00416D85: push eax
  00416D86: push 00000002h
  00416D88: call MSVBVM60.DLL.__vbaFreeVarList
  00416D8E: add esp, 0000000Ch
  00416D91: mov var_4, 00000006h
  00416D98: mov var_1A4, 00403E1Ch
  00416DA2: mov var_1AC, 00000008h
  00416DAC: lea edx, var_1AC
  00416DB2: lea ecx, var_BC
  00416DB8: call MSVBVM60.DLL.__vbaVarDup
  00416DBE: mov var_184, 0040953Ch ; ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
  00416DC8: mov var_18C, 00000008h
  00416DD2: push 00000000h
  00416DD4: push FFFFFFFFh
  00416DD6: lea ecx, var_BC
  00416DDC: push ecx
  00416DDD: mov edx, arg_C
  00416DE0: lea ecx, var_19C
  00416DE6: call MSVBVM60.DLL.__vbaVarVargNofree
  00416DEC: push eax
  00416DED: lea edx, var_18C
  00416DF3: push edx
  00416DF4: lea eax, var_AC
  00416DFA: push eax
  00416DFB: call MSVBVM60.DLL.__vbaVarCat
  00416E01: push eax
  00416E02: lea ecx, var_98
  00416E08: push ecx
  00416E09: call MSVBVM60.DLL.__vbaStrVarVal
  00416E0F: push eax
  00416E10: lea edx, var_CC
  00416E16: push edx
  00416E17: call [004011E8h] ; arg_1 = Split(arg_2, arg_3, arg_4, arg_5)
  00416E1D: lea eax, var_CC
  00416E23: push eax
  00416E24: push 00002008h
  00416E29: call MSVBVM60.DLL.__vbaAryVar
  00416E2F: mov var_240, eax
  00416E35: lea ecx, var_240
  00416E3B: push ecx
  00416E3C: lea edx, var_48
  00416E3F: push edx
  00416E40: call MSVBVM60.DLL.__vbaAryCopy
  00416E46: lea ecx, var_98
  00416E4C: call MSVBVM60.DLL.__vbaFreeStr
  00416E52: lea eax, var_CC
  00416E58: push eax
  00416E59: lea ecx, var_BC
  00416E5F: push ecx
  00416E60: lea edx, var_AC
  00416E66: push edx
  00416E67: push 00000003h
  00416E69: call MSVBVM60.DLL.__vbaFreeVarList
  00416E6F: add esp, 00000010h
  00416E72: mov var_4, 00000007h
  00416E79: mov var_194, 00000000h
  00416E83: mov var_19C, 00000003h
  00416E8D: mov var_184, 00000002h
  00416E97: mov var_18C, 00000002h
  00416EA1: lea eax, var_88
  00416EA7: push eax
  00416EA8: lea ecx, var_18C
  00416EAE: push ecx
  00416EAF: lea edx, var_AC
  00416EB5: push edx
  00416EB6: call MSVBVM60.DLL.__vbaVarDiv
  00416EBC: push eax
  00416EBD: call MSVBVM60.DLL.__vbaI4Var
  00416EC3: mov var_1B4, eax
  00416EC9: mov var_1BC, 00000003h
  00416ED3: mov eax, 00000010h
  00416ED8: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00416EDD: mov eax, esp
  00416EDF: mov ecx, var_19C
  00416EE5: mov [eax], ecx
  00416EE7: mov edx, var_198
  00416EED: mov [eax+04h], edx
  00416EF0: mov ecx, var_194
  00416EF6: mov [eax+08h], ecx
  00416EF9: mov edx, var_190
  00416EFF: mov [eax+0Ch], edx
  00416F02: mov eax, 00000010h
  00416F07: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00416F0C: mov eax, esp
  00416F0E: mov ecx, var_1BC
  00416F14: mov [eax], ecx
  00416F16: mov edx, var_1B8
  00416F1C: mov [eax+04h], edx
  00416F1F: mov ecx, var_1B4
  00416F25: mov [eax+08h], ecx
  00416F28: mov edx, var_1B0
  00416F2E: mov [eax+0Ch], edx
  00416F31: push 00000001h
  00416F33: push 00000039h
  00416F35: mov eax, arg_8
  00416F38: mov ecx, [eax]
  00416F3A: push ecx
  00416F3B: call MSVBVM60.DLL.__vbaLateIdCallSt
  00416F41: add esp, 0000002Ch
  00416F44: mov var_4, 00000008h
  00416F4B: mov var_184, 0000012Ch
  00416F55: mov var_18C, 00000002h
  00416F5F: lea edx, var_18C
  00416F65: lea ecx, var_40
  00416F68: call MSVBVM60.DLL.__vbaVarMove
  00416F6E: mov var_4, 00000009h
  00416F75: mov word ptr var_5C, 0000h
  00416F7B: mov var_4, 0000000Ah
  00416F82: push 00000000h
  00416F84: push 00000005h
  00416F86: mov edx, arg_8
  00416F89: mov eax, [edx]
  00416F8B: push eax
  00416F8C: lea ecx, var_AC
  00416F92: push ecx
  00416F93: call MSVBVM60.DLL.__vbaLateIdCallLd
  00416F99: add esp, 00000010h
  00416F9C: push eax
  00416F9D: call MSVBVM60.DLL.__vbaI4Var
  00416FA3: mov ecx, eax
  00416FA5: sub ecx, 00000001h
  00416FA8: jo 00417CEFh
  00416FAE: call MSVBVM60.DLL.__vbaI2I4
  00416FB4: mov var_254, ax
  00416FBB: mov word ptr var_250, 0001h
  00416FC4: mov word ptr var_8C, 0001h
  00416FCD: lea ecx, var_AC
  00416FD3: call MSVBVM60.DLL.__vbaFreeVar
  00416FD9: jmp 416FF6h
  00416FDB: mov dx, var_8C
  00416FE2: add dx, var_250
  00416FE9: jo 00417CEFh
  00416FEF: mov var_8C, dx
  00416FF6: mov ax, var_8C
  00416FFD: cmp ax, var_254
  00417004: jnle 004176A6h
  0041700A: mov var_4, 0000000Bh
  00417011: mov ecx, var_48
  00417014: push ecx
  00417015: lea edx, var_90
  0041701B: push edx
  0041701C: call MSVBVM60.DLL.__vbaAryLock
  00417022: cmp var_90, 00000000h
  00417029: jz 417094h
  0041702B: mov eax, var_90
  00417031: cmp word ptr [eax], 0001h
  00417035: jnz 417094h
  00417037: mov cx, var_8C
  0041703E: sub cx, 0001h
  00417042: jo 00417CEFh
  00417048: movsx edx, cx
  0041704B: mov eax, var_90
  00417051: sub edx, [eax+14h]
  00417054: mov var_244, edx
  0041705A: mov ecx, var_90
  00417060: mov edx, var_244
  00417066: cmp edx, [ecx+10h]
  00417069: jnb 417077h
  0041706B: mov var_274, 00000000h
  00417075: jmp 417083h
  00417077: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0041707D: mov var_274, eax
  00417083: mov eax, var_244
  00417089: shl eax, 02h
  0041708C: mov var_278, eax
  00417092: jmp 4170A0h
  00417094: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0041709A: mov var_278, eax
  004170A0: mov ecx, var_90
  004170A6: mov edx, [ecx+0Ch]
  004170A9: add edx, var_278
  004170AF: mov var_184, edx
  004170B5: mov var_18C, 00004008h
  004170BF: lea eax, var_18C
  004170C5: push eax
  004170C6: lea ecx, var_AC
  004170CC: push ecx
  004170CD: call 00417D00h
  004170D2: lea edx, var_90
  004170D8: push edx
  004170D9: call MSVBVM60.DLL.__vbaAryUnlock
  004170DF: mov eax, var_48
  004170E2: push eax
  004170E3: lea ecx, var_94
  004170E9: push ecx
  004170EA: call MSVBVM60.DLL.__vbaAryLock
  004170F0: cmp var_94, 00000000h
  004170F7: jz 417162h
  004170F9: mov edx, var_94
  004170FF: cmp word ptr [edx], 0001h
  00417103: jnz 417162h
  00417105: mov ax, var_8C
  0041710C: sub ax, 0001h
  00417110: jo 00417CEFh
  00417116: movsx ecx, ax
  00417119: mov edx, var_94
  0041711F: sub ecx, [edx+14h]
  00417122: mov var_248, ecx
  00417128: mov eax, var_94
  0041712E: mov ecx, var_248
  00417134: cmp ecx, [eax+10h]
  00417137: jnb 417145h
  00417139: mov var_27C, 00000000h
  00417143: jmp 417151h
  00417145: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0041714B: mov var_27C, eax
  00417151: mov edx, var_248
  00417157: shl edx, 02h
  0041715A: mov var_280, edx
  00417160: jmp 41716Eh
  00417162: call MSVBVM60.DLL.__vbaGenerateBoundsError
  00417168: mov var_280, eax
  0041716E: mov eax, var_94
  00417174: mov ecx, [eax+0Ch]
  00417177: add ecx, var_280
  0041717D: mov var_1A4, ecx
  00417183: mov var_1AC, 00004008h
  0041718D: lea edx, var_1AC
  00417193: push edx
  00417194: lea eax, var_CC
  0041719A: push eax
  0041719B: call 00417D00h
  004171A0: lea ecx, var_94
  004171A6: push ecx
  004171A7: call MSVBVM60.DLL.__vbaAryUnlock
  004171AD: push 00000000h
  004171AF: push 80010005h
  004171B4: mov edx, arg_8
  004171B7: mov eax, [edx]
  004171B9: push eax
  004171BA: lea ecx, var_DC
  004171C0: push ecx
  004171C1: call MSVBVM60.DLL.__vbaLateIdCallLd
  004171C7: add esp, 00000010h
  004171CA: movsx edx, word ptr var_8C
  004171D1: mov var_204, edx
  004171D7: mov var_20C, 00000003h
  004171E1: mov var_124, 000009C4h
  004171EB: mov var_12C, 00000002h
  004171F5: lea eax, var_DC
  004171FB: push eax
  004171FC: call MSVBVM60.DLL.__vbaR4Var
  00417202: fstp real4 ptr var_1B4
  00417208: mov var_1BC, 00000004h
  00417212: mov var_1C4, 0000012Ch
  0041721C: mov var_1CC, 00000002h
  00417226: lea ecx, var_1BC
  0041722C: push ecx
  0041722D: lea edx, var_40
  00417230: push edx
  00417231: lea eax, var_EC
  00417237: push eax
  00417238: call MSVBVM60.DLL.__vbaVarSub
  0041723E: push eax
  0041723F: lea ecx, var_1CC
  00417245: push ecx
  00417246: lea edx, var_FC
  0041724C: push edx
  0041724D: call MSVBVM60.DLL.__vbaVarSub
  00417253: push eax
  00417254: lea eax, var_10C
  0041725A: push eax
  0041725B: call MSVBVM60.DLL.__vbaVarAbs
  00417261: mov edx, eax
  00417263: lea ecx, var_11C
  00417269: call MSVBVM60.DLL.__vbaVarMove
  0041726F: lea ecx, var_12C
  00417275: push ecx
  00417276: lea edx, var_11C
  0041727C: push edx
  0041727D: lea eax, var_13C
  00417283: push eax
  00417284: call 00416A10h
  00417289: cmp var_48, 00000000h
  0041728D: jz 4172EFh
  0041728F: mov ecx, var_48
  00417292: cmp word ptr [ecx], 0001h
  00417296: jnz 4172EFh
  00417298: mov dx, var_8C
  0041729F: sub dx, 0001h
  004172A3: jo 00417CEFh
  004172A9: movsx eax, dx
  004172AC: mov ecx, var_48
  004172AF: sub eax, [ecx+14h]
  004172B2: mov var_24C, eax
  004172B8: mov edx, var_48
  004172BB: mov eax, var_24C
  004172C1: cmp eax, [edx+10h]
  004172C4: jnb 4172D2h
  004172C6: mov var_284, 00000000h
  004172D0: jmp 4172DEh
  004172D2: call MSVBVM60.DLL.__vbaGenerateBoundsError
  004172D8: mov var_284, eax
  004172DE: mov ecx, var_24C
  004172E4: shl ecx, 02h
  004172E7: mov var_288, ecx
  004172ED: jmp 4172FBh
  004172EF: call MSVBVM60.DLL.__vbaGenerateBoundsError
  004172F5: mov var_288, eax
  004172FB: mov edx, var_48
  004172FE: mov eax, [edx+0Ch]
  00417301: mov ecx, var_288
  00417307: mov edx, [eax+ecx]
  0041730A: push edx
  0041730B: push 004095C8h
  00417310: call MSVBVM60.DLL.__vbaStrCmp
  00417316: neg eax
  00417318: sbb eax, eax
  0041731A: neg eax
  0041731C: neg eax
  0041731E: mov var_1E4, ax
  00417325: mov var_1EC, 0000000Bh
  0041732F: lea eax, var_13C
  00417335: push eax
  00417336: lea ecx, var_CC
  0041733C: push ecx
  0041733D: lea edx, var_1EC
  00417343: push edx
  00417344: lea eax, var_14C
  0041734A: push eax
  0041734B: call [0040127Ch] ; IIf
  00417351: mov var_164, 00000000h
  0041735B: mov var_16C, 00000002h
  00417365: mov var_194, 00000000h
  0041736F: mov var_19C, 00008002h
  00417379: lea ecx, var_AC
  0041737F: push ecx
  00417380: lea edx, var_19C
  00417386: push edx
  00417387: lea eax, var_BC
  0041738D: push eax
  0041738E: call MSVBVM60.DLL.__vbaVarCmpEq
  00417394: mov edx, eax
  00417396: lea ecx, var_15C
  0041739C: call MSVBVM60.DLL.__vbaVarMove
  004173A2: lea ecx, var_14C
  004173A8: push ecx
  004173A9: lea edx, var_16C
  004173AF: push edx
  004173B0: lea eax, var_15C
  004173B6: push eax
  004173B7: lea ecx, var_17C
  004173BD: push ecx
  004173BE: call [0040127Ch] ; IIf
  004173C4: lea edx, var_17C
  004173CA: push edx
  004173CB: call MSVBVM60.DLL.__vbaI4Var
  004173D1: mov var_224, eax
  004173D7: mov var_22C, 00000003h
  004173E1: mov eax, 00000010h
  004173E6: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  004173EB: mov eax, esp
  004173ED: mov ecx, var_20C
  004173F3: mov [eax], ecx
  004173F5: mov edx, var_208
  004173FB: mov [eax+04h], edx
  004173FE: mov ecx, var_204
  00417404: mov [eax+08h], ecx
  00417407: mov edx, var_200
  0041740D: mov [eax+0Ch], edx
  00417410: mov eax, 00000010h
  00417415: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  0041741A: mov eax, esp
  0041741C: mov ecx, var_22C
  00417422: mov [eax], ecx
  00417424: mov edx, var_228
  0041742A: mov [eax+04h], edx
  0041742D: mov ecx, var_224
  00417433: mov [eax+08h], ecx
  00417436: mov edx, var_220
  0041743C: mov [eax+0Ch], edx
  0041743F: push 00000001h
  00417441: push 00000039h
  00417443: mov eax, arg_8
  00417446: mov ecx, [eax]
  00417448: push ecx
  00417449: call MSVBVM60.DLL.__vbaLateIdCallSt
  0041744F: add esp, 0000002Ch
  00417452: lea edx, var_17C
  00417458: push edx
  00417459: lea eax, var_14C
  0041745F: push eax
  00417460: lea ecx, var_16C
  00417466: push ecx
  00417467: lea edx, var_15C
  0041746D: push edx
  0041746E: lea eax, var_13C
  00417474: push eax
  00417475: lea ecx, var_CC
  0041747B: push ecx
  0041747C: lea edx, var_1EC
  00417482: push edx
  00417483: lea eax, var_12C
  00417489: push eax
  0041748A: lea ecx, var_11C
  00417490: push ecx
  00417491: lea edx, var_DC
  00417497: push edx
  00417498: lea eax, var_AC
  0041749E: push eax
  0041749F: push 0000000Bh
  004174A1: call MSVBVM60.DLL.__vbaFreeVarList
  004174A7: add esp, 00000030h
  004174AA: mov var_4, 0000000Ch
  004174B1: movsx ecx, word ptr var_8C
  004174B8: mov var_184, ecx
  004174BE: mov var_18C, 00000003h
  004174C8: mov eax, 00000010h
  004174CD: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  004174D2: mov edx, esp
  004174D4: mov eax, var_18C
  004174DA: mov [edx], eax
  004174DC: mov ecx, var_188
  004174E2: mov [edx+04h], ecx
  004174E5: mov eax, var_184
  004174EB: mov [edx+08h], eax
  004174EE: mov ecx, var_180
  004174F4: mov [edx+0Ch], ecx
  004174F7: push 00000001h
  004174F9: push 00000039h
  004174FB: mov edx, arg_8
  004174FE: mov eax, [edx]
  00417500: push eax
  00417501: lea ecx, var_AC
  00417507: push ecx
  00417508: call MSVBVM60.DLL.__vbaLateIdCallLd
  0041750E: add esp, 00000020h
  00417511: push eax
  00417512: call MSVBVM60.DLL.__vbaI4Var
  00417518: xor edx, edx
  0041751A: test eax, eax
  0041751C: setnle dl
  0041751F: neg edx
  00417521: mov var_244, dx
  00417528: lea ecx, var_AC
  0041752E: call MSVBVM60.DLL.__vbaFreeVar
  00417534: movsx eax, word ptr var_244
  0041753B: test eax, eax
  0041753D: jz 0041769Ah
  00417543: mov var_4, 0000000Dh
  0041754A: mov cx, var_5C
  0041754E: add cx, 0001h
  00417552: jo 00417CEFh
  00417558: mov var_5C, cx
  0041755C: mov var_4, 0000000Eh
  00417563: movsx edx, word ptr var_8C
  0041756A: mov var_184, edx
  00417570: mov var_18C, 00000003h
  0041757A: movsx eax, word ptr var_8C
  00417581: mov var_244, eax
  00417587: cmp var_244, 00000065h
  0041758E: jnb 41759Ch
  00417590: mov var_28C, 00000000h
  0041759A: jmp 4175A8h
  0041759C: call MSVBVM60.DLL.__vbaGenerateBoundsError
  004175A2: mov var_28C, eax
  004175A8: mov eax, 00000010h
  004175AD: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  004175B2: mov ecx, esp
  004175B4: mov edx, var_18C
  004175BA: mov [ecx], edx
  004175BC: mov eax, var_188
  004175C2: mov [ecx+04h], eax
  004175C5: mov edx, var_184
  004175CB: mov [ecx+08h], edx
  004175CE: mov eax, var_180
  004175D4: mov [ecx+0Ch], eax
  004175D7: push 00000001h
  004175D9: push 00000039h
  004175DB: mov ecx, arg_8
  004175DE: mov edx, [ecx]
  004175E0: push edx
  004175E1: lea eax, var_AC
  004175E7: push eax
  004175E8: call MSVBVM60.DLL.__vbaLateIdCallLd
  004175EE: add esp, 00000020h
  004175F1: push eax
  004175F2: call MSVBVM60.DLL.__vbaI4Var
  004175F8: mov var_290, eax
  004175FE: fild dword ptr var_290
  00417604: mov ecx, var_244
  0041760A: mov edx, var_68
  0041760D: fstp real8 ptr [edx+ecx*8]
  00417610: lea ecx, var_AC
  00417616: call MSVBVM60.DLL.__vbaFreeVar
  0041761C: mov var_4, 0000000Fh
  00417623: movsx eax, word ptr var_8C
  0041762A: mov var_244, eax
  00417630: cmp var_244, 00000065h
  00417637: jnb 417645h
  00417639: mov var_294, 00000000h
  00417643: jmp 417651h
  00417645: call MSVBVM60.DLL.__vbaGenerateBoundsError
  0041764B: mov var_294, eax
  00417651: mov ecx, var_244
  00417657: mov edx, var_68
  0041765A: mov eax, [edx+ecx*8]
  0041765D: mov var_184, eax
  00417663: mov ecx, [edx+ecx*8+04h]
  00417667: mov var_180, ecx
  0041766D: mov var_18C, 00000005h
  00417677: lea edx, var_30
  0041767A: push edx
  0041767B: lea eax, var_18C
  00417681: push eax
  00417682: lea ecx, var_AC
  00417688: push ecx
  00417689: call MSVBVM60.DLL.__vbaVarAdd
  0041768F: mov edx, eax
  00417691: lea ecx, var_30
  00417694: call MSVBVM60.DLL.__vbaVarMove
  0041769A: mov var_4, 00000011h
  004176A1: jmp 00416FDBh
  004176A6: mov var_4, 00000012h
  004176AD: mov var_A4, 00000001h
  004176B7: mov var_AC, 00000002h
  004176C1: lea edx, var_5C
  004176C4: mov var_184, edx
  004176CA: mov var_18C, 00004002h
  004176D4: lea eax, var_AC
  004176DA: push eax
  004176DB: lea ecx, var_18C
  004176E1: push ecx
  004176E2: lea edx, var_BC
  004176E8: push edx
  004176E9: call 00416B00h
  004176EE: lea eax, var_BC
  004176F4: push eax
  004176F5: call MSVBVM60.DLL.__vbaI2Var
  004176FB: mov var_5C, ax
  004176FF: lea ecx, var_BC
  00417705: push ecx
  00417706: lea edx, var_AC
  0041770C: push edx
  0041770D: push 00000002h
  0041770F: call MSVBVM60.DLL.__vbaFreeVarList
  00417715: add esp, 0000000Ch
  00417718: mov var_4, 00000014h
  0041771F: push 00000000h
  00417721: push 80010005h
  00417726: mov eax, arg_8
  00417729: mov ecx, [eax]
  0041772B: push ecx
  0041772C: lea edx, var_AC
  00417732: push edx
  00417733: call MSVBVM60.DLL.__vbaLateIdCallLd
  00417739: add esp, 00000010h
  0041773C: push eax
  0041773D: call MSVBVM60.DLL.__vbaR4Var
  00417743: fsub real4 ptr [004014F0h] ; 
  00417749: fstp real4 ptr var_184
  0041774F: fstsw ax
  00417751: test al, 0Dh
  00417753: jnz 00417CEAh
  00417759: mov var_18C, 00000004h
  00417763: lea eax, var_18C
  00417769: push eax
  0041776A: lea ecx, var_88
  00417770: push ecx
  00417771: lea edx, var_BC
  00417777: push edx
  00417778: call MSVBVM60.DLL.__vbaVarSub
  0041777E: mov edx, eax
  00417780: lea ecx, var_58
  00417783: call MSVBVM60.DLL.__vbaVarMove
  00417789: lea ecx, var_AC
  0041778F: call MSVBVM60.DLL.__vbaFreeVar
  00417795: mov var_4, 00000015h
  0041779C: push FFFFFFFFh
  0041779E: call On Error ...
  004177A4: mov var_4, 00000016h
  004177AB: push 00000000h
  004177AD: push 00000005h
  004177AF: mov eax, arg_8
  004177B2: mov ecx, [eax]
  004177B4: push ecx
  004177B5: lea edx, var_AC
  004177BB: push edx
  004177BC: call MSVBVM60.DLL.__vbaLateIdCallLd
  004177C2: add esp, 00000010h
  004177C5: push eax
  004177C6: call MSVBVM60.DLL.__vbaI4Var
  004177CC: mov ecx, eax
  004177CE: sub ecx, 00000001h
  004177D1: jo 00417CEFh
  004177D7: call MSVBVM60.DLL.__vbaI2I4
  004177DD: mov var_25C, ax
  004177E4: mov word ptr var_258, 0001h
  004177ED: mov word ptr var_8C, 0001h
  004177F6: lea ecx, var_AC
  004177FC: call MSVBVM60.DLL.__vbaFreeVar
  00417802: jmp 41781Fh
  00417804: mov ax, var_8C
  0041780B: add ax, var_258
  00417812: jo 00417CEFh
  00417818: mov var_8C, ax
  0041781F: mov cx, var_8C
  00417826: cmp cx, var_25C
  0041782D: jnle 00417BE0h
  00417833: mov var_4, 00000017h
  0041783A: movsx edx, word ptr var_8C
  00417841: mov var_184, edx
  00417847: mov var_18C, 00000003h
  00417851: mov eax, 00000010h
  00417856: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  0041785B: mov eax, esp
  0041785D: mov ecx, var_18C
  00417863: mov [eax], ecx
  00417865: mov edx, var_188
  0041786B: mov [eax+04h], edx
  0041786E: mov ecx, var_184
  00417874: mov [eax+08h], ecx
  00417877: mov edx, var_180
  0041787D: mov [eax+0Ch], edx
  00417880: push 00000001h
  00417882: push 00000039h
  00417884: mov eax, arg_8
  00417887: mov ecx, [eax]
  00417889: push ecx
  0041788A: lea edx, var_AC
  00417890: push edx
  00417891: call MSVBVM60.DLL.__vbaLateIdCallLd
  00417897: add esp, 00000020h
  0041789A: push eax
  0041789B: call MSVBVM60.DLL.__vbaI4Var
  004178A1: xor ecx, ecx
  004178A3: test eax, eax
  004178A5: setnle cl
  004178A8: neg ecx
  004178AA: mov var_244, cx
  004178B1: lea ecx, var_AC
  004178B7: call MSVBVM60.DLL.__vbaFreeVar
  004178BD: movsx edx, word ptr var_244
  004178C4: test edx, edx
  004178C6: jz 004179F0h
  004178CC: mov var_4, 00000018h
  004178D3: movsx eax, word ptr var_8C
  004178DA: mov var_194, eax
  004178E0: mov var_19C, 00000003h
  004178EA: movsx ecx, word ptr var_8C
  004178F1: mov var_244, ecx
  004178F7: cmp var_244, 00000065h
  004178FE: jnb 41790Ch
  00417900: mov var_298, 00000000h
  0041790A: jmp 417918h
  0041790C: call MSVBVM60.DLL.__vbaGenerateBoundsError
  00417912: mov var_298, eax
  00417918: mov edx, var_244
  0041791E: mov eax, var_68
  00417921: mov ecx, [eax+edx*8]
  00417924: mov var_184, ecx
  0041792A: mov edx, [eax+edx*8+04h]
  0041792E: mov var_180, edx
  00417934: mov var_18C, 00000005h
  0041793E: lea eax, var_18C
  00417944: push eax
  00417945: lea ecx, var_30
  00417948: push ecx
  00417949: lea edx, var_AC
  0041794F: push edx
  00417950: call MSVBVM60.DLL.__vbaVarDiv
  00417956: push eax
  00417957: lea eax, var_58
  0041795A: push eax
  0041795B: lea ecx, var_BC
  00417961: push ecx
  00417962: call MSVBVM60.DLL.__vbaVarMul
  00417968: push eax
  00417969: call MSVBVM60.DLL.__vbaI4Var
  0041796F: mov var_1B4, eax
  00417975: mov var_1BC, 00000003h
  0041797F: mov eax, 00000010h
  00417984: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417989: mov edx, esp
  0041798B: mov eax, var_19C
  00417991: mov [edx], eax
  00417993: mov ecx, var_198
  00417999: mov [edx+04h], ecx
  0041799C: mov eax, var_194
  004179A2: mov [edx+08h], eax
  004179A5: mov ecx, var_190
  004179AB: mov [edx+0Ch], ecx
  004179AE: mov eax, 00000010h
  004179B3: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  004179B8: mov edx, esp
  004179BA: mov eax, var_1BC
  004179C0: mov [edx], eax
  004179C2: mov ecx, var_1B8
  004179C8: mov [edx+04h], ecx
  004179CB: mov eax, var_1B4
  004179D1: mov [edx+08h], eax
  004179D4: mov ecx, var_1B0
  004179DA: mov [edx+0Ch], ecx
  004179DD: push 00000001h
  004179DF: push 00000039h
  004179E1: mov edx, arg_8
  004179E4: mov eax, [edx]
  004179E6: push eax
  004179E7: call MSVBVM60.DLL.__vbaLateIdCallSt
  004179ED: add esp, 0000002Ch
  004179F0: mov var_4, 0000001Ah
  004179F7: mov var_1C4, 00000000h
  00417A01: mov var_1CC, 00000003h
  00417A0B: movsx ecx, word ptr var_8C
  00417A12: mov var_1E4, ecx
  00417A18: mov var_1EC, 00000003h
  00417A22: mov var_184, 00000000h
  00417A2C: mov var_18C, 00000003h
  00417A36: movsx edx, word ptr var_8C
  00417A3D: mov var_1A4, edx
  00417A43: mov var_1AC, 00000003h
  00417A4D: mov eax, 00000010h
  00417A52: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417A57: mov eax, esp
  00417A59: mov ecx, var_18C
  00417A5F: mov [eax], ecx
  00417A61: mov edx, var_188
  00417A67: mov [eax+04h], edx
  00417A6A: mov ecx, var_184
  00417A70: mov [eax+08h], ecx
  00417A73: mov edx, var_180
  00417A79: mov [eax+0Ch], edx
  00417A7C: mov eax, 00000010h
  00417A81: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417A86: mov eax, esp
  00417A88: mov ecx, var_1AC
  00417A8E: mov [eax], ecx
  00417A90: mov edx, var_1A8
  00417A96: mov [eax+04h], edx
  00417A99: mov ecx, var_1A4
  00417A9F: mov [eax+08h], ecx
  00417AA2: mov edx, var_1A0
  00417AA8: mov [eax+0Ch], edx
  00417AAB: push 00000002h
  00417AAD: push 00000041h
  00417AAF: mov eax, arg_8
  00417AB2: mov ecx, [eax]
  00417AB4: push ecx
  00417AB5: lea edx, var_AC
  00417ABB: push edx
  00417ABC: call MSVBVM60.DLL.__vbaLateIdCallLd
  00417AC2: add esp, 00000030h
  00417AC5: push eax
  00417AC6: call MSVBVM60.DLL.__vbaStrVarMove
  00417ACC: mov var_B4, eax
  00417AD2: mov var_BC, 00000008h
  00417ADC: lea eax, var_BC
  00417AE2: push eax
  00417AE3: lea ecx, var_CC
  00417AE9: push ecx
  00417AEA: call [00401140h] ; arg_1 = Ucase(arg_2)
  00417AF0: lea edx, var_CC
  00417AF6: push edx
  00417AF7: call MSVBVM60.DLL.__vbaStrVarMove
  00417AFD: mov var_D4, eax
  00417B03: mov var_DC, 00000008h
  00417B0D: mov eax, 00000010h
  00417B12: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417B17: mov eax, esp
  00417B19: mov ecx, var_1CC
  00417B1F: mov [eax], ecx
  00417B21: mov edx, var_1C8
  00417B27: mov [eax+04h], edx
  00417B2A: mov ecx, var_1C4
  00417B30: mov [eax+08h], ecx
  00417B33: mov edx, var_1C0
  00417B39: mov [eax+0Ch], edx
  00417B3C: mov eax, 00000010h
  00417B41: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417B46: mov eax, esp
  00417B48: mov ecx, var_1EC
  00417B4E: mov [eax], ecx
  00417B50: mov edx, var_1E8
  00417B56: mov [eax+04h], edx
  00417B59: mov ecx, var_1E4
  00417B5F: mov [eax+08h], ecx
  00417B62: mov edx, var_1E0
  00417B68: mov [eax+0Ch], edx
  00417B6B: mov eax, 00000010h
  00417B70: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417B75: mov eax, esp
  00417B77: mov ecx, var_DC
  00417B7D: mov [eax], ecx
  00417B7F: mov edx, var_D8
  00417B85: mov [eax+04h], edx
  00417B88: mov ecx, var_D4
  00417B8E: mov [eax+08h], ecx
  00417B91: mov edx, var_D0
  00417B97: mov [eax+0Ch], edx
  00417B9A: push 00000002h
  00417B9C: push 00000041h
  00417B9E: mov eax, arg_8
  00417BA1: mov ecx, [eax]
  00417BA3: push ecx
  00417BA4: call MSVBVM60.DLL.__vbaLateIdCallSt
  00417BAA: add esp, 0000003Ch
  00417BAD: lea edx, var_DC
  00417BB3: push edx
  00417BB4: lea eax, var_CC
  00417BBA: push eax
  00417BBB: lea ecx, var_BC
  00417BC1: push ecx
  00417BC2: lea edx, var_AC
  00417BC8: push edx
  00417BC9: push 00000004h
  00417BCB: call MSVBVM60.DLL.__vbaFreeVarList
  00417BD1: add esp, 00000014h
  00417BD4: mov var_4, 0000001Bh
  00417BDB: jmp 00417804h
  00417BE0: wait 
  00417BE1: push 00417CD7h ; "‹Màd‰'#1"
  00417BE6: jmp 00417C8Bh
  00417BEB: lea eax, var_90
  00417BF1: push eax
  00417BF2: call MSVBVM60.DLL.__vbaAryUnlock
  00417BF8: lea ecx, var_94
  00417BFE: push ecx
  00417BFF: call MSVBVM60.DLL.__vbaAryUnlock
  00417C05: lea ecx, var_98
  00417C0B: call MSVBVM60.DLL.__vbaFreeStr
  00417C11: lea ecx, var_9C
  00417C17: call MSVBVM60.DLL.__vbaFreeObj
  00417C1D: lea edx, var_17C
  00417C23: push edx
  00417C24: lea eax, var_16C
  00417C2A: push eax
  00417C2B: lea ecx, var_15C
  00417C31: push ecx
  00417C32: lea edx, var_14C
  00417C38: push edx
  00417C39: lea eax, var_13C
  00417C3F: push eax
  00417C40: lea ecx, var_12C
  00417C46: push ecx
  00417C47: lea edx, var_11C
  00417C4D: push edx
  00417C4E: lea eax, var_10C
  00417C54: push eax
  00417C55: lea ecx, var_FC
  00417C5B: push ecx
  00417C5C: lea edx, var_EC
  00417C62: push edx
  00417C63: lea eax, var_DC
  00417C69: push eax
  00417C6A: lea ecx, var_CC
  00417C70: push ecx
  00417C71: lea edx, var_BC
  00417C77: push edx
  00417C78: lea eax, var_AC
  00417C7E: push eax
  00417C7F: push 0000000Eh
  00417C81: call MSVBVM60.DLL.__vbaFreeVarList
  00417C87: add esp, 0000003Ch
  00417C8A: ret 
End Sub

Private sub Proc_4_4_417D00
  00417D00: push ebp
  00417D01: mov ebp, esp
  00417D03: sub esp, 0000000Ch
  00417D06: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00417D0B: mov eax, fs:[00h]
  00417D11: push eax
  00417D12: mov fs:[00000000h], esp
  00417D19: sub esp, 0000003Ch
  00417D1C: push ebx
  00417D1D: push esi
  00417D1E: push edi
  00417D1F: mov var_C, esp
  00417D22: mov var_8, 00401610h
  00417D29: xor eax, eax
  00417D2B: lea ecx, var_38
  00417D2E: mov var_24, eax
  00417D31: mov var_28, eax
  00417D34: mov var_38, eax
  00417D37: mov var_48, eax
  00417D3A: mov eax, arg_C
  00417D3D: push eax
  00417D3E: push ecx
  00417D3F: call 00418080h
  00417D44: lea edx, var_38
  00417D47: lea eax, var_28
  00417D4A: push edx
  00417D4B: push eax
  00417D4C: call MSVBVM60.DLL.__vbaStrVarVal
  00417D52: push eax
  00417D53: call [00401370h] ; Val(arg_1)
  00417D59: fstp real8 ptr var_40
  00417D5C: lea edx, var_48
  00417D5F: lea ecx, var_24
  00417D62: mov var_48, 00000005h
  00417D69: call MSVBVM60.DLL.__vbaVarMove
  00417D6F: lea ecx, var_28
  00417D72: call MSVBVM60.DLL.__vbaFreeStr
  00417D78: lea ecx, var_38
  00417D7B: call MSVBVM60.DLL.__vbaFreeVar
  00417D81: wait 
  00417D82: push 00417DACh
  00417D87: jmp 417DABh
  00417D89: test byte ptr var_4, 04h
  00417D8D: jz 417D98h
  00417D8F: lea ecx, var_24
  00417D92: call MSVBVM60.DLL.__vbaFreeVar
  00417D98: lea ecx, var_28
  00417D9B: call MSVBVM60.DLL.__vbaFreeStr
  00417DA1: lea ecx, var_38
  00417DA4: call MSVBVM60.DLL.__vbaFreeVar
  00417DAA: ret 
End Sub

Private sub Proc_4_5_417DE0
  00417DE0: push ebp
  00417DE1: mov ebp, esp
  00417DE3: sub esp, 00000018h
  00417DE6: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  00417DEB: mov eax, fs:[00h]
  00417DF1: push eax
  00417DF2: mov fs:[00000000h], esp
  00417DF9: mov eax, 00000074h
  00417DFE: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417E03: push ebx
  00417E04: push esi
  00417E05: push edi
  00417E06: mov var_18, esp
  00417E09: mov var_14, 00401620h
  00417E10: mov var_10, 00000000h
  00417E17: mov var_C, 00000000h
  00417E1E: mov var_4, 00000001h
  00417E25: mov var_4, 00000002h
  00417E2C: push FFFFFFFFh
  00417E2E: call On Error ...
  00417E34: mov var_4, 00000003h
  00417E3B: cmp [00419640h], 00000000h
  00417E42: jnz 417E5Dh
  00417E44: push 00419640h
  00417E49: push 004039B0h
  00417E4E: call MSVBVM60.DLL.__vbaNew2
  00417E54: mov var_80, 00419640h
  00417E5B: jmp 417E64h
  00417E5D: mov var_80, 00419640h
  00417E64: mov eax, var_80
  00417E67: mov ecx, [eax]
  00417E69: mov var_5C, ecx
  00417E6C: lea edx, var_38
  00417E6F: push edx
  00417E70: mov eax, var_5C
  00417E73: mov ecx, [eax]
  00417E75: mov edx, var_5C
  00417E78: push edx
  00417E79: call [ecx+1Ch]
  00417E7C: fclex 
  00417E7E: mov var_60, eax
  00417E81: cmp var_60, 00000000h
  00417E85: jnl 417EA4h
  00417E87: push 0000001Ch
  00417E89: push 004039A0h
  00417E8E: mov eax, var_5C
  00417E91: push eax
  00417E92: mov ecx, var_60
  00417E95: push ecx
  00417E96: call MSVBVM60.DLL.__vbaHresultCheckObj
  00417E9C: mov var_84, eax
  00417EA2: jmp 417EAEh
  00417EA4: mov var_84, 00000000h
  00417EAE: mov edx, var_38
  00417EB1: mov var_64, edx
  00417EB4: mov eax, var_64
  00417EB7: mov ecx, [eax]
  00417EB9: mov edx, var_64
  00417EBC: push edx
  00417EBD: call [ecx+50h]
  00417EC0: fclex 
  00417EC2: mov var_68, eax
  00417EC5: cmp var_68, 00000000h
  00417EC9: jnl 417EE8h
  00417ECB: push 00000050h
  00417ECD: push 00404594h
  00417ED2: mov eax, var_64
  00417ED5: push eax
  00417ED6: mov ecx, var_68
  00417ED9: push ecx
  00417EDA: call MSVBVM60.DLL.__vbaHresultCheckObj
  00417EE0: mov var_88, eax
  00417EE6: jmp 417EF2h
  00417EE8: mov var_88, 00000000h
  00417EF2: lea ecx, var_38
  00417EF5: call MSVBVM60.DLL.__vbaFreeObj
  00417EFB: mov var_4, 00000004h
  00417F02: cmp [00419640h], 00000000h
  00417F09: jnz 417F27h
  00417F0B: push 00419640h
  00417F10: push 004039B0h
  00417F15: call MSVBVM60.DLL.__vbaNew2
  00417F1B: mov var_8C, 00419640h
  00417F25: jmp 417F31h
  00417F27: mov var_8C, 00419640h
  00417F31: mov edx, var_8C
  00417F37: mov eax, [edx]
  00417F39: mov var_5C, eax
  00417F3C: lea ecx, var_38
  00417F3F: push ecx
  00417F40: mov edx, var_5C
  00417F43: mov eax, [edx]
  00417F45: mov ecx, var_5C
  00417F48: push ecx
  00417F49: call [eax+1Ch]
  00417F4C: fclex 
  00417F4E: mov var_60, eax
  00417F51: cmp var_60, 00000000h
  00417F55: jnl 417F74h
  00417F57: push 0000001Ch
  00417F59: push 004039A0h
  00417F5E: mov edx, var_5C
  00417F61: push edx
  00417F62: mov eax, var_60
  00417F65: push eax
  00417F66: call MSVBVM60.DLL.__vbaHresultCheckObj
  00417F6C: mov var_90, eax
  00417F72: jmp 417F7Eh
  00417F74: mov var_90, 00000000h
  00417F7E: mov ecx, var_38
  00417F81: mov var_64, ecx
  00417F84: mov var_50, 80020004h
  00417F8B: mov var_58, 0000000Ah
  00417F92: mov eax, 00000010h
  00417F97: call 00401670h ; MSVBVM60.DLL.__vbaChkstk
  00417F9C: mov edx, esp
  00417F9E: mov eax, var_58
  00417FA1: mov [edx], eax
  00417FA3: mov ecx, var_54
  00417FA6: mov [edx+04h], ecx
  00417FA9: mov eax, var_50
  00417FAC: mov [edx+08h], eax
  00417FAF: mov ecx, var_4C
  00417FB2: mov [edx+0Ch], ecx
  00417FB5: mov edx, arg_C
  00417FB8: lea ecx, var_48
  00417FBB: call MSVBVM60.DLL.__vbaVarVargNofree
  00417FC1: push eax
  00417FC2: lea edx, var_34
  00417FC5: push edx
  00417FC6: call MSVBVM60.DLL.__vbaStrVarVal
  00417FCC: push eax
  00417FCD: mov eax, var_64
  00417FD0: mov ecx, [eax]
  00417FD2: mov edx, var_64
  00417FD5: push edx
  00417FD6: call [ecx+60h]
  00417FD9: fclex 
  00417FDB: mov var_68, eax
  00417FDE: cmp var_68, 00000000h
  00417FE2: jnl 418001h
  00417FE4: push 00000060h
  00417FE6: push 00404594h
  00417FEB: mov eax, var_64
  00417FEE: push eax
  00417FEF: mov ecx, var_68
  00417FF2: push ecx
  00417FF3: call MSVBVM60.DLL.__vbaHresultCheckObj
  00417FF9: mov var_94, eax
  00417FFF: jmp 41800Bh
  00418001: mov var_94, 00000000h
  0041800B: lea ecx, var_34
  0041800E: call MSVBVM60.DLL.__vbaFreeStr
  00418014: lea ecx, var_38
  00418017: call MSVBVM60.DLL.__vbaFreeObj
  0041801D: push 0041804Bh
  00418022: jmp 41804Ah
  00418024: mov edx, var_10
  00418027: and edx, 00000004h
  0041802A: test edx, edx
  0041802C: jz 418037h
  0041802E: lea ecx, var_30
  00418031: call MSVBVM60.DLL.__vbaFreeVar
  00418037: lea ecx, var_34
  0041803A: call MSVBVM60.DLL.__vbaFreeStr
  00418040: lea ecx, var_38
  00418043: call MSVBVM60.DLL.__vbaFreeObj
  00418049: ret 
End Sub

Private sub Proc_4_6_418080
  00418080: push ebp
  00418081: mov ebp, esp
  00418083: sub esp, 0000000Ch
  00418086: push 00401676h ; MSVBVM60.DLL.__vbaExceptHandler
  0041808B: mov eax, fs:[00h]
  00418091: push eax
  00418092: mov fs:[00000000h], esp
  00418099: sub esp, 0000003Ch
  0041809C: push ebx
  0041809D: push esi
  0041809E: push edi
  0041809F: mov var_C, esp
  004180A2: mov var_8, 00401658h
  004180A9: mov esi, arg_C
  004180AC: mov edi, MSVBVM60.DLL.__vbaVarVargNofree
  004180B2: xor eax, eax
  004180B4: mov edx, esi
  004180B6: push eax
  004180B7: push FFFFFFFFh
  004180B9: push 00000001h
  004180BB: push 00403558h
  004180C0: push 00403E1Ch
  004180C5: lea ecx, var_48
  004180C8: mov var_24, eax
  004180CB: mov var_28, eax
  004180CE: mov var_38, eax
  004180D1: mov var_48, eax
  004180D4: call edi
  004180D6: push eax
  004180D7: lea eax, var_28
  004180DA: push eax
  004180DB: call MSVBVM60.DLL.__vbaStrVarVal
  004180E1: push eax
  004180E2: call [004011ECh] ; Replace(arg_1, arg_2, arg_3, arg_4, arg_5, arg_6)
  004180E8: lea edx, var_38
  004180EB: mov ecx, esi
  004180ED: mov var_30, eax
  004180F0: mov var_38, 00000008h
  004180F7: call MSVBVM60.DLL.__vbaVargVarMove
  004180FD: lea ecx, var_28
  00418100: call MSVBVM60.DLL.__vbaFreeStr
  00418106: mov edx, esi
  00418108: lea ecx, var_48
  0041810B: call edi
  0041810D: mov edx, eax
  0041810F: lea ecx, var_24
  00418112: call MSVBVM60.DLL.__vbaVarCopy
  00418118: push 00418142h
  0041811D: jmp 418141h
  0041811F: test byte ptr var_4, 04h
  00418123: jz 41812Eh
  00418125: lea ecx, var_24
  00418128: call MSVBVM60.DLL.__vbaFreeVar
  0041812E: lea ecx, var_28
  00418131: call MSVBVM60.DLL.__vbaFreeStr
  00418137: lea ecx, var_38
  0041813A: call MSVBVM60.DLL.__vbaFreeVar
  00418140: ret 
End Sub

