Attribute VB_Name = "a"
Option Explicit
Dim fso As New FileSystemObject

Public Function LOADFILE(FilePath As String)
Dim sEXT As String: sEXT = get_ExtensionFromFileName(FilePath)
sEXT = LCase(sEXT)
Select Case sEXT

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "gdp" ' --> Hard Truck Apocalypse
wndGDP.Show
wndGDP.Path.Caption = FilePath
'--------------------------------------------------------
Case "cursor" ' --> Hard Truck Apocalypse
loadtxt FilePath
'--------------------------------------------------------
Case "psys" ' --> Hard Truck Apocalypse
loadtxt FilePath
'--------------------------------------------------------
Case "ssl" ' --> Hard Truck Apocalypse
loadtxt FilePath
'--------------------------------------------------------
Case "vs" ' --> Hard Truck Apocalypse
loadtxt FilePath
'--------------------------------------------------------
Case "ps" ' --> Hard Truck Apocalypse
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "pck" ' --> Commandos II
loadArchive FilePath
'--------------------------------------------------------
Case "mac" ' --> Commandos II
loadtxt FilePath
'--------------------------------------------------------
Case "itl" ' --> Commandos II
loadtxt FilePath
'--------------------------------------------------------
Case "str" ' --> Commandos II
loadtxt FilePath
'--------------------------------------------------------
Case "msb" ' --> Commandos II
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "cns" ' --> RMUGEN 2
loadtxt FilePath
'--------------------------------------------------------
Case "cmd" ' --> RMUGEN 2
loadtxt FilePath
'--------------------------------------------------------
Case "air" ' --> RMUGEN 2
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "big" ' --> ToCA Race Driver #Series
loadBIGF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
'--------------------------------------------------------
Case "jpk" ' --> ToCA Race Driver 3
loadBIGF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
'--------------------------------------------------------
Case "mr5" ' --> ToCA Race Driver 1
loadBIGF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
'--------------------------------------------------------
Case "lng" ' --> ToCa Race Driver #Series
frmLNG.Show
frmLNG.Path.Caption = jrMain.Path.Caption
LNG.LNG_open FilePath, frmLNG.lngData
'--------------------------------------------------------
Case "icz" ' --> ToCA Race Driver 3
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "tbf" ' --> Colin McRae Rally 04
loadBIGF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
'--------------------------------------------------------
Case "bgx" ' --> Colin McRae Rally 04
loadBIGF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
'--------------------------------------------------------
Case "pfx" ' --> Colin McRae Rally 04
loadBIGF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "slf" ' --> Jagged Alliance 2
loadSLF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "dlg" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "ar" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "cm" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "lev" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "lts" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "ftr" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "ftr" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "wpn" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "dsc" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "hro" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "itm" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "are" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "aba" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "pth" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "ba" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "tab" ' --> Gorky 17
loadtxt FilePath
'--------------------------------------------------------
Case "var" ' --> Gorky 17
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "ads" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "ads" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "lst" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "nds" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "opt" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "pxy" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "report" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "rsr" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "sup" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "ai" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "sia" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'--------------------------------------------------------
Case "md" ' --> Cossacks 2 - Napoleonic Wars
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "script" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "ai" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "skin" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "cfg" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "arena" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "shader" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "def" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "menu" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'--------------------------------------------------------
Case "bsp" ' --> Return to castle Wolfenstein SBWL
loadBSP FilePath
'--------------------------------------------------------
Case "h" ' --> Return to castle Wolfenstein SBWL
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "met" ' --> Prison Tycoon
loadtxt FilePath
'--------------------------------------------------------
Case "fds" ' --> Prison Tycoon
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "anm" ' --> Driv3r
loadtxt FilePath
'--------------------------------------------------------
Case "skm" ' --> Driv3r
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "db" ' --> Suffering 2 - Ties That Bind
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "nm" ' --> John Deere - American Builder Deluxe
loadtxt FilePath
'--------------------------------------------------------
Case "x" ' --> John Deere - American Builder Deluxe
loadtxt FilePath
'--------------------------------------------------------
Case "fx" ' --> John Deere - American Builder Deluxe
loadtxt FilePath
'--------------------------------------------------------
Case "gmf" ' --> John Deere - American Builder Deluxe
    If GMF.GMF_GMA_CheckIfValid(FilePath) = True Then
        loadtxt FilePath
        Exit Function
    Else
        If GMF.GMF_GMA_CheckIfValidGMI(FilePath) = True Then
            MsgBox "This is a GMF_GMI file. jRipper supports only GMF_GMA.", vbExclamation, "Wrong .GMF format"
            Exit Function
        Else
            MsgBox "Unknown .GMF format", vbCritical, "Not GMF_GMI/GMF_GMA format"
            Exit Function
        End If
    End If
'--------------------------------------------------------
Case "gma" ' --> John Deere - American Builder Deluxe
    If GMF.GMF_GMA_CheckIfValid(FilePath) = True Then
        loadtxt FilePath
        Exit Function
    Else
        If GMF.GMF_GMA_CheckIfValidGMI(FilePath) = True Then
            MsgBox "This is a GMA_GMI file. jRipper supports only GMA_GMA.", vbExclamation, "Wrong .GMA format"
            Exit Function
        Else
            MsgBox "Unknown .GMA format", vbCritical, "Not GMA_GMI/GMA_GMA format"
            Exit Function
        End If
    End If
'--------------------------------------------------------
Case "gms" ' --> John Deere - American Builder Deluxe
    If GMF.GMF_GMA_CheckIfValid(FilePath) = True Then
        loadtxt FilePath
        Exit Function
    Else
        If GMF.GMF_GMA_CheckIfValidGMI(FilePath) = True Then
            MsgBox "This is a GMS_GMI file. jRipper supports only GMS_GMA.", vbExclamation, "Wrong .GMS format"
            Exit Function
        Else
            MsgBox "Unknown .GMS format", vbCritical, "Not GMS_GMI/GMS_GMA format"
            Exit Function
        End If
    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "ctr" ' --> NFS Porsche
loadtxt FilePath
'--------------------------------------------------------
Case "lay" ' --> NFS Porsche
loadtxt FilePath
'--------------------------------------------------------
Case "clr" ' --> NFS Porsche
loadtxt FilePath
'--------------------------------------------------------
Case "tpg" ' --> NFS Porsche
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "twt" ' --> Carmageddon II
loadTWT FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "wad" ' --> Tomb Raider 3 (AOLC/TLA)
loadWAD FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "blk" ' --> Adrenalin Extreme Show
loadtxt FilePath
'--------------------------------------------------------
Case "ta" ' --> Adrenalin Extreme Show
loadtxt FilePath
'--------------------------------------------------------
Case "lua" ' --> Adrenalin Extreme Show / HT Apocalypse
loadtxt FilePath
'--------------------------------------------------------
Case "gui" ' --> Adrenalin Extreme Show
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "bmp" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "png" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "jpg" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "tga" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "jpeg" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "pcx" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "tif" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "dib" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "ico" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "cur" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "016" ' --> # picture
'loadpic FilePath
'--------------------------------------------------------
Case "256" ' --> # picture
'loadpic FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "ini" ' --> generic text file
loadtxt FilePath
'--------------------------------------------------------
Case "nfo" ' --> generic text file
loadtxt FilePath
frmmain.txtmain.SelStart = 0
frmmain.txtmain.SelLength = Len(frmmain.txtmain.Text)
frmmain.txtmain.SelColor = vbWhite
frmmain.txtmain.SelFontName = "Terminal"
frmmain.txtmain.SelFontSize = 9
frmmain.txtmain.SelStart = 0
'--------------------------------------------------------
Case "diz" ' --> generic text file
loadtxt FilePath
frmmain.txtmain.SelStart = 0
frmmain.txtmain.SelLength = Len(frmmain.txtmain.Text)
frmmain.txtmain.SelColor = vbWhite
frmmain.txtmain.SelFontName = "Terminal"
frmmain.txtmain.SelFontSize = 9
frmmain.txtmain.SelStart = 0
'--------------------------------------------------------
Case "txt" ' --> generic text file
loadtxt FilePath
'--------------------------------------------------------
Case "xml" ' --> generic text file
loadtxt FilePath
'--------------------------------------------------------
Case "doc" ' --> Document File (not Office 2003/2007 !)
loadtxt FilePath
'--------------------------------------------------------
Case "rtf" ' --> Rich Text File
loadtxt FilePath
'--------------------------------------------------------
Case "asm" ' --> ASM code
loadtxt FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "vp6" ' --> EA movie (VP6 codec required)
VP6_Playa.Path.Caption = jrMain.Path.Caption
VP6_Playa.fname.Caption = jrMain.FPath.Caption
VP6_Playa.Show
'--------------------------------------------------------
Case "bik" ' --> BINK Video (.dll required)
loadbink
'--------------------------------------------------------
Case "xmv" ' --> BINK Video (.dll required)
loadbink
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case "dat" ' --> # Various Games
If BIG.BIGF_CheckIfValid(FilePath) = True Then
loadBIGF FilePath
ArchiveMan.Path.Text = jrMain.Path.Caption
ElseIf BIG.BIGF_CheckIfValidBIGC(FilePath) = True Then
    MsgBox "This is a BIGC Archive. jRipper supports only BIGF Archives", vbExclamation, "Not BIGF"
Else
    'put code here!
End If
'--------------------------------------------------------
Case "ast" '--> # Various Games (SChl audio)
loadSChl FilePath
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

End Select
End Function

Public Function EXTRACTFILE(ByVal FilePath, ByVal ExtractAll As Boolean)
Dim dblStart As Double
Dim dblEnd As Double
Dim Dest As String
Dim x
Dim sEXT As String: sEXT = get_ExtensionFromFileName(FilePath)
Dest = InputBox("Where to save?", "Extract Path", App.Path & "\Unpacked")
If Right(Dest, 1) = "\" Then
Dest = Strings.Left(Dest, Len(Dest) - 1)
End If
If fso.FolderExists(Dest) = False Then
fso.CreateFolder Dest
End If
'
sEXT = LCase(sEXT)
Select Case sEXT

Case "pck" ' Commandos II
    If ExtractAll = False Then
        MsgBox "You must extract all!", vbCritical, "jRipper " & IAPPV
    Else
        Dim File As String: File = ArchiveMan.Path.Text
        Dim file2 As String: file2 = ArchiveMan.Label1.Caption
        ArchiveMan.pb1.Visible = True
        If file2 = "DATA2.PCK" Then
           ArchiveMan.pb1.Max = 302
            dblStart = Timer
            For x = 2 To 303
            PCK.PCK_extractOne_1 File, Dest, x
            ArchiveMan.pb1.Value = x - 1
            Next x
            dblEnd = Timer
        ElseIf file2 = "DATA.PCK" Then
            ArchiveMan.pb1.Max = 4914
            dblStart = Timer
            PCK.PCK_extractOne_2 File, Dest, ArchiveMan.pb1
            dblEnd = Timer
        End If
        ArchiveMan.pb1.Visible = False
        ArchiveMan.info_LOAD.Caption = "Saved in: " & FormatNumber(dblEnd - dblStart, 5, vbUseDefault, vbTrue, vbTrue) & " s"
    End If
    
Case "big" ' ToCA's
BIGFextract ExtractAll, Dest
Case "jpk" ' ToCA's
BIGFextract ExtractAll, Dest
Case "mr5" ' ToCA's
BIGFextract ExtractAll, Dest
Case "tbf" ' Colin 04
BIGFextract ExtractAll, Dest
Case "bgx" ' Colin 04
BIGFextract ExtractAll, Dest
Case "pfx" ' Colin 04
BIGFextract ExtractAll, Dest
Case "slf" ' Jagged Alliance 2
Dest = Dest & "\" & Split(ArchiveMan.info_MORE1.Caption, ": ")(1)
If Right(Dest, 1) = "\" Then
Dest = Strings.Left(Dest, Len(Dest) - 1)
End If
If fso.FolderExists(Dest) = False Then
fso.CreateFolder Dest
End If
SLFextract ExtractAll, Dest


End Select
End Function

Function loadtxt(ByVal FilePath As String)
frmmain.Show
TXT.TXT2_load FilePath, frmmain.txtmain
frmmain.txtmain.SelStart = 0
frmmain.txtmain.SelLength = Len(frmmain.txtmain.Text)
frmmain.txtmain.SelColor = vbWhite
frmmain.txtmain.SelStart = 0
frmmain.bbar.Caption = "Lines: " & TXT.lineCount(FilePath) & " File: " & FilePath
frmmain.lines.Caption = TXT.lineCount(FilePath)
End Function

Function loadArchive(ByVal FilePath As String)
ArchiveMan.Show
ArchiveMan.Path.Text = FilePath
DoEvents
End Function

Function loadBIGF(ByVal FilePath As String)
Dim TS As Double
Dim TE As Double
If BIG.BIGF_CheckIfValid(FilePath) = True Then
    loadArchive FilePath
    TS = Timer
    BIG.BIGF_Decode FilePath, ArchiveMan.Port, ArchiveMan.pb1
    TE = Timer
    ArchiveMan.info_LOAD.Caption = "Loaded in: " & FormatNumber(TE - TS, 5, vbUseDefault, vbTrue, vbTrue) & " s"
    ArchiveMan.x.Enabled = True
    ArchiveMan.xa.Enabled = True
    ArchiveMan.Label1.Caption = jrMain.FPath.Caption
    ArchiveMan.selall.Value = 1
    Exit Function
ElseIf BIG.BIGF_CheckIfValidBIGC(FilePath) = True Then
    MsgBox "This is a BIGC Archive. jRipper supports only BIGF Archives", vbExclamation, "Not BIGF"
    Exit Function
Else
    MsgBox "Invalid BIGF/BIGC Archive", vbCritical, "Not BIGF/BIGC"
    Exit Function
End If
End Function

Function loadSLF(ByVal FilePath As String)
Dim TS As Double
Dim TE As Double
    loadArchive FilePath
    TS = Timer
    SLF.SLF_decode FilePath, ArchiveMan.Port, ArchiveMan.pb1
    TE = Timer
    ArchiveMan.info_LOAD.Caption = "Loaded in: " & FormatNumber(TE - TS, 5, vbUseDefault, vbTrue, vbTrue) & " s"
    ArchiveMan.x.Enabled = True
    ArchiveMan.xa.Enabled = True
    ArchiveMan.Label1.Caption = jrMain.FPath.Caption
    ArchiveMan.selall.Value = 1
    ArchiveMan.addinfo "Root Dir: " & SLF.SLF_get_DirName(FilePath), "FileList Entry Offset: " & SLF.SLF_get_FileTableEntryOffset(FilePath)
    Exit Function
End Function

Function loadTWT(ByVal FilePath As String)
Dim TS As Double
Dim TE As Double
    loadArchive FilePath
    TS = Timer
    TWT.TWT_Decode FilePath, ArchiveMan.Port
    TE = Timer
    ArchiveMan.info_LOAD.Caption = "Loaded in: " & FormatNumber(TE - TS, 5, vbUseDefault, vbTrue, vbTrue) & " s"
    ArchiveMan.x.Enabled = True
    ArchiveMan.xa.Enabled = True
    ArchiveMan.Label1.Caption = jrMain.FPath.Caption
    ArchiveMan.selall.Value = 1
    ArchiveMan.addinfo "First File Offset: " & TWT.TWT_get_FirstFileOffset(FilePath)
    Exit Function
End Function

Function loadSChl(ByVal FilePath As String)
Dim TS As Double
Dim TE As Double
Dim SCHlChuncks As Long
    loadArchive FilePath
    TS = Timer
    SCHlChuncks = SCHl_get_Headers(FilePath, ArchiveMan.pb1, ArchiveMan.Port)
    TE = Timer
    ArchiveMan.info_LOAD.Caption = "Loaded in: " & FormatNumber(TE - TS, 5, vbUseDefault, vbTrue, vbTrue) & " s"
    ArchiveMan.x.Enabled = True
    ArchiveMan.xa.Enabled = True
    ArchiveMan.Label1.Caption = jrMain.FPath.Caption
    ArchiveMan.selall.Value = 1
    ArchiveMan.addinfo "First Split Header Offset: " & SChl.SCHl_get_SplitHeader(FilePath)
    Exit Function
End Function

Function loadBSP(ByVal FilePath As String)
Dim TS As Double
Dim TE As Double
    loadArchive FilePath
    TS = Timer
    BSP.BSP_decode FilePath, ArchiveMan.Port, ArchiveMan.pb1
    TE = Timer
    ArchiveMan.info_LOAD.Caption = "Loaded in: " & FormatNumber(TE - TS, 5, vbUseDefault, vbTrue, vbTrue) & " s"
    ArchiveMan.x.Enabled = True
    ArchiveMan.xa.Enabled = True
    ArchiveMan.Label1.Caption = jrMain.FPath.Caption
    ArchiveMan.selall.Value = 1
    ArchiveMan.addinfo "First File Offset: " & BSP.BSP_get_FirstFileOffset(FilePath), "First File Name Entry: " & BSP.BSP_get_FirstFileNameEntry(FilePath)
    Exit Function
End Function

Function loadWAD(ByVal FilePath As String)
Dim TS As Double
Dim TE As Double
    loadArchive FilePath
    TS = Timer
    WAD.WAD_Decode FilePath, ArchiveMan.Port
    TE = Timer
    ArchiveMan.info_LOAD.Caption = "Loaded in: " & FormatNumber(TE - TS, 5, vbUseDefault, vbTrue, vbTrue) & " s"
    ArchiveMan.x.Enabled = True
    ArchiveMan.xa.Enabled = True
    ArchiveMan.Label1.Caption = jrMain.FPath.Caption
    ArchiveMan.selall.Value = 1
    ArchiveMan.addinfo "First File Offset: " & WAD.WAD_get_FirstFileOffset(FilePath)
    Exit Function
End Function

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Function BIGFextract(ExtractAll As Boolean, Dest As String)
' Dims
Dim x
Dim File As String: File = ArchiveMan.Path.Text
Dim dblStart As Double
Dim dblEnd As Double
 ' Do it
ArchiveMan.pb1.Visible = True
ArchiveMan.pb1.Max = ArchiveMan.Port.ListItems.Count
If ExtractAll = False Then
    dblStart = Timer
    For x = 1 To ArchiveMan.Port.ListItems.Count
    If ArchiveMan.Port.ListItems(x).Checked = True Then
    BIG.BIGF_extractOne File, Dest, x
    End If
    ArchiveMan.pb1.Value = x
    Next x
    dblEnd = Timer
Else
    dblStart = Timer
    For x = 1 To ArchiveMan.Port.ListItems.Count
    BIG.BIGF_extractOne File, Dest, x
    ArchiveMan.pb1.Value = x
    Next x
    dblEnd = Timer
End If
ArchiveMan.pb1.Visible = False
ArchiveMan.info_LOAD.Caption = "Extracted in: " & FormatNumber(dblEnd - dblStart, 5, vbUseDefault, vbTrue, vbTrue) & " s"
Exit Function
End Function

Function SLFextract(ExtractAll As Boolean, Dest As String)
' Dims
Dim x
Dim File As String: File = ArchiveMan.Path.Text
Dim dblStart As Double
Dim dblEnd As Double
 ' Do it
ArchiveMan.pb1.Visible = True
ArchiveMan.pb1.Max = ArchiveMan.Port.ListItems.Count
If ExtractAll = False Then
    dblStart = Timer
    For x = 1 To ArchiveMan.Port.ListItems.Count
    If ArchiveMan.Port.ListItems(x).Checked = True Then
    SLF.SLF_extractOne File, Dest, x, ArchiveMan.Port
    End If
    ArchiveMan.pb1.Value = x
    Next x
    dblEnd = Timer
Else
    dblStart = Timer
    For x = 1 To ArchiveMan.Port.ListItems.Count
    SLF.SLF_extractOne File, Dest, x, ArchiveMan.Port
    ArchiveMan.pb1.Value = x
    Next x
    dblEnd = Timer
End If
ArchiveMan.pb1.Visible = False
ArchiveMan.info_LOAD.Caption = "Extracted in: " & FormatNumber(dblEnd - dblStart, 5, vbUseDefault, vbTrue, vbTrue) & " s"
Exit Function
End Function

Function loadbink()
BIK_Playa.Path.Caption = jrMain.Path.Caption
BIK_Playa.always.Value = ReadINI(App.Path & "\bin\jr.ini", "BIKPlayer", "Always")

If BIK_Playa.always.Value = 1 Then
BIK_Playa.Show
DoEvents
BIK_Playa.switches.Caption = ReadINI(App.Path & "\bin\jr.ini", "BIKPlayer", "Filterx")
BIK_Playa.bik_minimizeJR.Value = ReadINI(App.Path & "\bin\jr.ini", "BIKPlayer", "SFilter")
    If BIK_Playa.bik_minimizeJR.Value = 1 Then
    jrMain.WindowState = vbMinimized
    End If
BIK.BIK_play jrMain.Path.Caption, BIK_Playa.switches.Caption
Else
BIK_Playa.Show
End If
End Function
