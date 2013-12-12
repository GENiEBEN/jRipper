Attribute VB_Name = "GamesList"
Option Explicit
Dim fso As New FileSystemObject

Function LoadGames(ComboBox As ComboBox)
With ComboBox
.AddItem "Tomb Raider - Legend"
.AddItem "Need for Speed - Most Wanted [PAL]"
.AddItem "Need for Speed - Most Wanted [NTSC]"
.AddItem "Age of Empires III"
.AddItem "Cabela's Dangerous Hunts 2"
.AddItem "KINGPIN: Life of Crime"
.AddItem "NASCAR SimRacing"
.AddItem "LADA Racing Club"
.AddItem "Holiday World Tycoon"
.AddItem "4x4 EVO2"
.AddItem "ToCA Race Driver 3"
.AddItem "GodFather, The"
.AddItem "Need for Speed - Underground 2"
.AddItem "Heroes of the Pacific"
.AddItem "House of the Dead III"
.AddItem "F1 Challenge '99-'02"
.AddItem "Rogue Trooper"
.AddItem "eRacer"
.AddItem "Syberia II"
.AddItem "Call Of Cthulhu - Dark Corners of the Earth"
.AddItem "Hitman: Contracts"
.AddItem "Simpons, The - Hit & Run"
.AddItem "Condemned - Criminal Origins"
.AddItem "SWAT 4"
.AddItem "SWAT 4 - Stetchkov Syndicate"
.AddItem "Constantine"
.AddItem "Marc Ecko's Getting Up - Contents Under Pressure"
.AddItem "Colin McRae Rally 2"
.AddItem "Colin McRae Rally 4"
.AddItem "Boiling Point - Road to hell"
.AddItem "Tomb Raider - The Angel of Darkness"
.AddItem "FORD Racing 3 *European*"
.AddItem "FORD Street Racing *European*"
.AddItem "Midnight Outlaw - 6 hours to Sun up"
.AddItem "Chronicles of Riddick - Escape From Butcher Bay"
.AddItem "Richard Burns Rally"
.AddItem "Splinter Cell - Pandora Tomorrow [PAL]"
.AddItem "World Soccer Winning Eleven 8 International"
.AddItem "FarCry"
.AddItem "ChampionSheep Rally"
.AddItem "GUN"
.AddItem "Codename - Panzers, Phase Two *ITALIAN*"
.AddItem "Colin McRae Rally 2005 (ULTiMA Crew Mod)"
.AddItem "Colin McRae Rally 2005"
.AddItem "Grand Theft Auto - Vice City"
.AddItem "Alien vs. Predator 2"
.AddItem "Need for Speed - Most Wanted DEMO [NTSC]"
.AddItem "F.E.A.R. First Encounter Assault and Recon"
.AddItem "Grand Theft Auto III"
.AddItem "Street Racing Syndicate"
.AddItem "Sin - Episode 1 Emergence"
.AddItem "Street Hacker"
.AddItem "Liquidator2"
.AddItem "DRIV3R"
.AddItem "Real War"
.AddItem "True Crime - New York City *European*"
.AddItem "Utopia City"
.AddItem "ToCA Race Driver 2"
.AddItem "Stalin Subway, The"
.AddItem "ObsCure"
.AddItem "Lord of the Rings: War of the Ring"
.AddItem "Gorky17"
.AddItem "Odi•um"
.AddItem "Hard Truck - Apocalypse"
.AddItem "Temple of Elemental Evil"
.AddItem "Commandos II - Men of Courage"
.AddItem "Need for Speed - Underground 1"
.AddItem "Adrenaline Extreme Show"
.AddItem "Reservoir Dogs"
.AddItem "Bad Day L.A."



End With
NIMP.Combo1.Text = "Bad Day L.A."
End Function

Function LoadVideos(GameName)
Dim strEXT As String * 3
'
With NIMP.ret
NIMP.list.ListItems.Clear
Select Case GameName ' GameName is case-sensitive!
    
Case "Tomb Raider - Legend"
add "Trailer", "title.bik"
add "nVidia", "nVidia.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Crystal Dynamics\Tomb Raider: Legend", "InstallPath", "\"

Case "Need for Speed - Most Wanted [NTSC]"
add "EA logo", "ealogo_english_ntsc.vp6"
add "Josie Maran speech", "psa_english_ntsc.vp6"
add "E3 Trailer", "attract_movie_english_ntsc.vp6"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\EA GAMES\Need for Speed Most Wanted", "Install Dir", "MOVIES\"

Case "Need for Speed - Most Wanted [PAL]"
add "EA logo", "ealogo_english_pal.vp6"
add "Josie Maran speech", "psa_english_pal.vp6"
add "E3 Trailer", "attract_movie_english_pal.vp6"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\EA GAMES\Need for Speed Most Wanted", "Install Dir", "MOVIES\"

Case "Age of Empires III" 'not tested
add "Logos", "logos.bik"
add "Trailer", "age3.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft Games\Age of Empires 3\1.0", "SetupPath", "avi\"

Case "Cabela's Dangerous Hunts 2"
add "Activision", "atvi.bik"
add "Magic Wand Productions", "magicwand.bik"
add "Legal", "legal.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Activision Value\Magic Wand\Dngh2", "Path1", "Data\Movies\"

Case "KINGPIN: Life of Crime"
add "All intro movies", "kpintro.exe"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\kingpin.exe", "Path", "\"
    
Case "NASCAR SimRacing"
add "EA logo", "EAS.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\EA SPORTS\NASCAR SimRacing", "Install Dir", "Movies\"
add "RANDOM: B.Vickers", "Cameos\BVickers.bik"
add "RANDOM: B.Weber", "Cameos\BWeber1.bik"
add "RANDOM: Dale Jr", "Cameos\DaleJr.bik"
add "RANDOM: Daytona Fans", "Cameos\Daytona_fans1.bik"
add "RANDOM: E.Sadler", "Cameos\ESadler.bik"
add "RANDOM: Im in the game", "Cameos\Intro_Im_in_the_game_Multi.bik"
add "RANDOM: J.Gordon", "Cameos\JGordon_action.bik"
add "RANDOM: J.Johnson", "Cameos\JJohnson_action.bik"
add "RANDOM: K.Busch", "Cameos\KBusch.bik"
add "RANDOM: Team LOWES", "Cameos\Lowes_crew.bik"
add "RANDOM: M.Kenseth", "Cameos\MKenseth.bik"
add "RANDOM: R.Newman", "Cameos\RNewman.bik"
add "RANDOM: T.Stewart", "Cameos\TStewart.bik"

Case "LADA Racing Club"
add "ND [Video]", "nd.ogg"
add "ND [Audio]", "nd.audio.ogg"
add "Geleos [Video]", "geleos.ogg"
add "Geleos [Audio]", "geleos.audio.ogg"
add "Trailer [Video]", "kasta.ogg"
add "Trailer [Audio]", "kasta.audio.ogg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Íîâûé Äèñê\LADA Racing Club", "AppPath", "\resources\video\"

Case "Holiday World Tycoon"
add "All intros (compiled BIK)", "intro.exe"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Holiday World Tycoon", "Path", "\global\grafik\filme\"

Case "4x4 EVO2"
add "All intros", "4x4.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Terminal Reality\4x4 Evo2", "path", "\video\"

Case "ToCA Race Driver 3"
add "Intel PEE [english]", "intel_e.bik"
add "Intel PEE [french]", "intel_f.bik"
add "Intel PEE [german]", "intel_g.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Codemasters\Race Driver 3", "PATH_MAIN", "\video\"
add "nVidia", "nVid.bik"
add "Trailer-DTM [english]", "Intr_dtm_e.bik"
add "Trailer-DTM [french]", "Intr_dtm_f.bik"
add "Trailer-DTM [german]", "Intr_dtm_g.bik"
add "Trailer-DTM [italian]", "Intr_dtm_i.bik"
add "Trailer-DTM [spanish]", "Intr_dtm_s.bik"
add "Trailer-TOCA [english]", "intr_toca_e.bik"
add "Trailer-TOCA [french]", "intr_toca_f.bik"
add "Trailer-TOCA [german]", "intr_toca_g.bik"
add "Trailer-TOCA [italian]", "intr_toca_i.bik"
add "Trailer-TOCA [spanish]", "intr_toca_s.bik"
add "Trailer-V8 [english]", "intr_v8_e.bik"
add "Trailer-V8 [french]", "intr_v8_f.bik"
add "Trailer-V8 [german]", "intr_v8_g.bik"
add "Trailer-V8 [italian]", "intr_v8_i.bik"
add "Trailer-V8 [spanish]", "intr_v8_s.bik"

Case "GodFather, The"
add "Paramount", "paramount.vp6"
add "EA Presents", "load.vp6"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Electronic Arts\The Godfather The Game", "Install Dir", "movies\"

Case "Need for Speed - Underground 2"
add "EA logo", "ealogo.vp6"
add "THX logo", "THX_logo.vp6"
add "Drive carefull...ya'right", "PSA.vp6"
add "Trailer", "FMVOpening.vp6"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\EA GAMES\Need for Speed Underground 2", "Install Dir", "MOVIES\"

Case "Heroes of the Pacific"
add "Codemasters", "introcm.wmv"
add "IRgurus", "irgurus.wmv"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Codemasters\Heroes Of The Pacific", "InstallPath", "\Video\Win32\"

Case "House of the Dead III"
add "Trailer", "TITLE_US_CS.SFD"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\SEGA\hod3\Settings", "Path", "\mvi\"

Case "F1 Challenge '99-'02"
add "EA Sports logo", "Eas.bik"
add "Trailer", "EAIntro.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\EA SPORTS\F1 Challenge 99-02", "Install Dir", "Movies\"
add "Game Title", "GameTitle.bmp"
add "Legal Screen", "Legal.bmp"
dir2 "HKEY_LOCAL_MACHINE\SOFTWARE\EA SPORTS\F1 Challenge 99-02", "Install Dir", "Options\"

Case "Rogue Trooper"
add "eidos", "pub.bik"
add "REBELLION", "rebel.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Eidos\Rogue Trooper", "Directory", "\FMV\Splash\"
add "Prologue", "Prologue.bik"
dir2 "HKEY_LOCAL_MACHINE\SOFTWARE\Eidos\Rogue Trooper", "Directory", "\FMV\Cutscene\"

Case "eRacer"
add "All intros", "intro.avi"
dir "HKEY_CURRENT_USER\Software\Rage Games Ltd\eRacer", "HOVAPPDATA", "\Gui\Avi\"

Case "Syberia II"
add "Microids", "Intro_Microids.bik"
add "Syberia Ice Logo", "CI_PreIntro.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microids\Syberia 2", "CurrentPath", "\Video\"

Case "Call Of Cthulhu - Dark Corners of the Earth"
add "Legal [english]", "legal.wmv"
add "Legal [french]", "legal_french.wmv"
add "Legal [german]", "legal_german.wmv"
add "Bethesda Softworks / UBiSOFT", "beth_logo.wmv"
add "headfirs productions", "hf_logo.wmv"
add "Warning! [english]", "warning.wmv"
add "Warning! [french]", "warning_french.wmv"
add "Warning! [german]", "warning_german.wmv"
add "Game Logo [english]", "cocdcote_logo.wmv"
add "Game Logo [french]", "cocdcote_logo_french.wmv"
add "Game Logo [german]", "cocdcote_logo_german.wmv"
dir "HKEY_CURRENT_USER\Software\Bethesda Softworks\Call Of Cthulhu DCoTE\Settings", "Executable", ""
NIMP.ret.Caption = Replace(NIMP.ret.Caption, "Engine\cocmainwin32.exe", "Development\pcvideo\")

Case "Hitman: Contracts"
add "eidos", "Eidos.bik"
add "Io-Interactive", "Io_logo.bik"
add "nVidia", "nVidia.bik"
add "Ballistic Test", "Intro.bik"
add "Music by JK", "JesperKy.bik"
add "Legal", "Copyrigh.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Eidos\Hitman Contracts", "InstallDir", "\Movies\"

Case "Simpons, The - Hit & Run"
add "vivendi", "vuglogo.rmv"
add "20 Century Fox", "foxlogo.rmv"
add "Gracie", "gracie.rmv"
add "Radical", "radlogo.rmv"
add "Intro", "intro.rmv"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{F79AAB3A-B8B4-4AC7-94AB-1C4C076C6A89}", "InstallLocation", "movies\"

Case "Condemned - Criminal Origins"
add "All intros", "CondemnedA.Arch00"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Monolith Productions\Condemned - Criminal Origins\1.00.0000", "InstallDir", "Game\"

Case "SWAT 4"
add "Sierra", "SierraLogo.bik"
add "nVidia", "Nvidia.bik"
add "Trailer", "swat4.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Sierra\SWAT 4", "InstallPath", "Content\Movies\"

Case "SWAT 4 - Stetchkov Syndicate"
add "Sierra", "SierraLogo.bik"
add "Trailer", "SW4X_intro.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Sierra\SWAT 4", "InstallPath", "Contentexpansion\Movies\"

Case "Constantine"
add "SCi Games", "Logo1.bik"
add "bits studios", "Logo2.bik"
add "VERTIGO", "Logo3.bik"
add "WB Interactive", "Logo4.bik"
add "Legals", "legal.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\SCi Games\Constantine", "InstallPath", "\Movies\Frontend\"

Case "Marc Ecko's Getting Up - Contents Under Pressure"
add "Legals", "Legal.bik"
add "Don't do it in real life", "GrafDis.bik"
add "ATARI", "atari.bik"
add "ECKO Games", "EckoLogo.bik"
add "Collective", "Collectv.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{B8F941EA-FC3E-4915-B5EB-E91A47BF3394}", "InstallLocation", "\engine\Movies\English\"

Case "Colin McRae Rally 2"
add "Codemasters", "cm.bik"
add "Trailer", "intro.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Codemasters\Colin McRae Rally 2", "Game_HDPath", "\FrontEnd\Videos\"

Case "Colin McRae Rally 4"
add "Codemasters [Hi]", "sting.wmv"
add "Codemasters [Lo]", "sting_l.wmv"
add "nVidia [Hi]", "twimtbp.wmv"
add "nVidia [Lo]", "twimtbp_l.wmv"
add "nVidia GFFX [Hi]", "twimtbpfx.wmv"
add "nVidia GFFX [Lo]", "twimtbpfx_l.wmv"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Codemasters\Colin McRae Rally 04", "INSTALL_PATH", "\Data\Videos\"

Case "Boiling Point - Road to hell"
add "nVidia", "nvidia.avi"
add "ATARI", "publisher.avi"
add "Deep Shadows", "deepshadows.avi"
add "Trailer [EU]", "intro.avi"
add "Trailer [US]", "intro_us.avi"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{58AC967F-CE64-4065-AF54-FA66BAF31FE8}", "InstallLocation", "\Avi\"

Case "Tomb Raider - The Angel of Darkness"
add "EiDOS Interactive", "eidos.mpg"
add "CORE Design", "Core.mpg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Core Design\TombRaiderAngelOfDarkness\1.0", "InstalledPath", "Data\FMV\"

Case "FORD Racing 3 *European*"
add "Empire", "empire25.wmv"
add "Razorworks", "razor25.wmv"
add "Trailer", "intro25.wmv"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Empire Interactive\Ford Racing 3", "Installation Path", "\ANIMS\EUROPEAN\"

Case "FORD Street Racing *European*"
add "XPLOSIV", "xplos25.wmv"
add "Razorworks", "razor25.wmv"
add "Trailer", "intro25.wmv"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Empire Interactive\Ford Street Racing", "Installation Path", "\ANIMS\EUROPEAN\"

Case "Midnight Outlaw - 6 hours to Sun up"
add "ESRB", "logo_esrb.mpg"
add "Babylon Soft", "Babylonsoft.mpg"
add "Value Soft", "valuesoft.mpg"
add "Logo Screen", "Midnight_outlaw.mpg"
add "BONUS: RPM Tuning Trailer", "rpmtuning.mpg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Babylon Software\Midnight Outlaw", "Path", "\data\RTC\Videos\"

Case "Chronicles of Riddick - Escape From Butcher Bay"
add "VIVENDI Universal", "vug_new.ogg"
add "Universal *NTSC*", "universal_ntsc.ogg"
add "TIGON", "Tigon.ogg"
add "AMD/nVidia/SoundBlaster", "logo_mixed.ogg"
add "Trailer", "Sbz.ogg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A8DE8C34-7F51-4cc8-B326-C425793EE741}", "InstallLocation", "Content\Videos\"

Case "Richard Burns Rally"
add "All intros", "RBR.wmv"
NIMP.ret.Caption = Get_DefaultValue(vHKEY_LOCAL_MACHINE, "SOFTWARE\SCi Games\Richard Burns Rally\InstallPath", "C:\Program Files\Richard Burns Rally") & "\Video\"

Case "Splinter Cell - Pandora Tomorrow [PAL]"
add "UBISOFT", "UbiLogoPAL.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Ubisoft\Splinter Cell Pandora Tomorrow", "InstalledPath", "\videos\"
add "Intro", "Intro.bik"
dir2 "HKEY_LOCAL_MACHINE\SOFTWARE\Ubisoft\Splinter Cell Pandora Tomorrow", "InstalledPath", "\offline\Videos\"

Case "World Soccer Winning Eleven 8 International"
add "Trailer", "opmov"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\KONAMIWE8IU\WE8IU", "installdir", "dat\"

Case "FarCry"
add "UBISOFT [english]", "Ubi.bik"
add "CRYTEK [english]", "Crytek.bik"
add "Powered by... [english]", "sandbox.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\03E3856B6F673FC44BFEA706451E740B", "A2CDBD6DC27E48246BDAB6B3164BADCB", "English\"
add "UBISOFT [french]", "Ubi.bik"
add "CRYTEK [french]", "Crytek.bik"
add "Powered by... [french]", "sandbox.bik"
dir2 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\03E3856B6F673FC44BFEA706451E740B", "A2CDBD6DC27E48246BDAB6B3164BADCB", "French\"

Case "ChampionSheep Rally"
add "All intros", "LOGO.DBC"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Championsheep Rally", "DisplayIcon", ""
NIMP.ret.Caption = Replace(NIMP.ret.Caption, "CSR.exe", "Data\LOAD\")

Case "GUN"
add "Activision", "ATVI.bik"
add "Beenox", "BEENOX.bik"
add "NeverSoft + Trailer", "NSINTRO.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Activision\GUN", "InstallPath", "\data\movies\"

Case "Codename - Panzers, Phase Two *ITALIAN*"
add "Trailer", "intro.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\FX Interactive\Panzers II", "Path", ""

Case "Colin McRae Rally 2005 (ULTiMA Crew Mod)"
add "StarFORCE3 Defeated! [Hi]", "sting.wmv"
add "StarFORCE3 Defeated! [Lo]", "sting_l.wmv"
add "ULTiMA - We work for you [Hi]", "twimtbp.wmv"
add "ULTiMA - We work for you [Lo]", "twimtbp_l.wmv"
add "Trailer [Hi]", "final.wmv"
add "Trailer [Lo]", "final_l.wmv"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Codemasters\Colin McRae Rally 2005", "INSTALL_PATH", "\Data\Videos\"

Case "Colin McRae Rally 2005"
add "Codemasters [Hi]", "sting.wmv"
add "Codemasters [Lo]", "sting_l.wmv"
add "nVidia [Hi]", "twimtbp.wmv"
add "nVidia [Lo]", "twimtbp_l.wmv"
add "Trailer [Hi]", "final.wmv"
add "Trailer [Lo]", "final_l.wmv"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Codemasters\Colin McRae Rally 2005", "INSTALL_PATH", "\Data\Videos\"

Case "Grand Theft Auto - Vice City"
add "Rockstar", "Logo.mpg"
add "Intro/Credits", "GTAtitles.mpg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{4B35F00C-E63D-40DC-9839-DF15A33EAC46}", "InstallLocation", "\movies\"

Case "Alien vs. Predator 2"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Monolith Productions\Aliens vs. Predator 2\1.0", "InstallDir", "\Movies\"
add "FOX Interactive", "foxlogo.bik"
add "SIERRA", "sierralogo.bik"
add "LithTech", "ltlogo.bik"
add "MonoLith", "lithlogo.bik"

Case "Need for Speed - Most Wanted DEMO [NTSC]"
add "EA logo", "ealogo_english_ntsc.vp6"
add "Josie Maran speech", "psa_english_ntsc.vp6"
add "E3 Trailer", "attract_movie_english_ntsc.vp6"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\EA GAMES\Need for Speed Most Wanted Demo", "Install Dir", "MOVIES\"

Case "F.E.A.R. First Encounter Assault and Recon"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{2B653229-9854-4989-B780-D978F5F13EAB}", "InstallLocation", "\"
add "All intros", "FEAR.Arch00" 'one big 3.9GB archive....

Case "Grand Theft Auto III"
add "Rockstar", "Logo.mpg"
add "Intro/Credits [english]", "GTAtitles.mpg"
add "Intro/Credits [german]", "GTAtitlesGER.mpg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\D:\gta3cdPC\gta3.exe", "Path", "\movies\"

Case "Street Racing Syndicate"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\NAMCO\Street Racing Syndicate", "InstallDir", ""
NIMP.ret.Caption = Replace(NIMP.ret.Caption, "\bin", "\videos\")
add "eutechnyx", "050.dat" 'renamed .avi
add "Trailer", "021.dat" 'renamed .avi

Case "Sin - Episode 1 Emergence"
Dim X: X = InputBox("Please enter complete path of SinEpisodes.exe (only folder, without SinEpisodes.exe)", "Sin - Episode 1 Emergence", "C:\SinEpisodes")
If fso.FolderExists(X) = False Then
    MsgBox "Incorrect folder! Restart NiMP and try again.", vbCritical, "Sin - Episode 1 Emergence"
    End
Else
Dim Y As String: Y = X
    If Not Right(X, 1) = "\" Then
        Y = X & "\"
    End If
NIMP.ret.Caption = Y & "SE1\media\"
add "nVidia", "Nvidia.avi"
add "Ritual", "RitualSin.avi"
End If

Case "Street Hacker"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Street Hacker", "Install DIR", "\media\"
add "Intro", "sh intro video.avi"

Case "Liquidator2"
add "Intro", "intro.ogg"
add "nVidia", "NVIDIA.ogg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Parallax Arts Studio Inc.\Liquidator", "InstallDir", "\data\avi\"

Case "DRIV3R"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{01DBF423-E27B-45DA-B7F3-F9D4DB39B1C9}", "InstallLocation", "\FMV\"
add "ATARI/REFLECTIONS", "ATARI.XMV"
add "nVidia", "NVIDIA.XMV"

Case "Real War"
add "Logos", "logo.mpg"
add "Trailer", "intro.mpg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\REALWAR", "VIDTREE", ""

Case "True Crime - New York City *European*"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Aspyr\True Crime New York City\Install", "Path", "\Movies\"
add "Activision", "atvi.bik"
add "luxoflux", "Luxo.bik"
add "Aspyr", "Aspyr.bik"
add "Method", "Method.bik"
add "Legal", "Title.bik"

Case "Utopia City"
'dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Utopia City", "UninstallString", ""
'nimp.ret.Caption = Replace(nimp.ret.Caption, "C:/WINDOWS/TMUninst.exe ", "")
'nimp.ret.Caption = Mid(nimp.ret.Caption, 4, Len(NIMP.ret.Caption) - 16)
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Parallax Arts Studio Inc.\Utopia City", "InstallDir", "\data\avi\"
add "Logos", "logo.ogg"
add "Intro", "intro.ogg"

Case "ToCA Race Driver 2"
add "Codemasters", "sting.bik"
add "nVidia", "nVid.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Codemasters\Race Driver 2", "PATH_MAIN", "\video\"

Case "Stalin Subway, The"
NIMP.ret.Caption = Get_DefaultValue(vHKEY_LOCAL_MACHINE, "SOFTWARE\Buka\TSS", "C:\Program Files\The Stalin Subway")
NIMP.ret.Caption = Replace(NIMP.ret.Caption, "metro2.exe", "VIDEO\")
add "All logos", "Logo.avi"

Case "ObsCure"
add "Microids presents ... Trailer", "cin01.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\MC2\Obscure", "CurrentPath", "data\_videopc\"

Case "Lord of the Rings: War of the Ring"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Liquid Entertainment\War of the Ring", "Install Path", "\Archives\"
add "All intros", "Movies.H2O"

Case "Gorky17"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Metropolis\G17", "dst", "AVI\"
add "Topware", "topware.asf"
add "Metropolis", "logo.asf"
add "G17 Logo", "intro.asf"

Case "Odi•um"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Metropolis\G17", "dst", "AVI\"
add "Monolith", "monolith.asf"
add "Topware", "topware.asf"
add "Metropolis", "logo.asf"
add "G17 Logo", "intro.asf"

Case "Hard Truck - Apocalypse"
add "CDV", "logocdv.wmv"
add "BUKA", "logobuka.wmv"
add "Targem", "logotargem.wmv"
add "Trailer", "INTRO_RUS_FIELD.wmv"
NIMP.ret.Caption = Get_DefaultValue(vHKEY_LOCAL_MACHINE, "SOFTWARE\Buka\ExMachina", "C:\Hard Truck Apocalypse\")
NIMP.ret.Caption = Replace(NIMP.ret.Caption, "hta.exe", "data\video\")

Case "Temple of Elemental Evil"
add "ATARi", "AtariLogo.bik"
add "TROIKA Games", "TroikaLogo.bik"
add "Wizards of the Coast", "WotCLogo.bik"
add "Trailer", "introcinematic.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{AD80F06B-0F21-4EEE-934D-BEF0D21E6383}", "InstallLocation", "\data\movies\"

Case "Commandos II - Men of Courage"
add "EIDOS", "DATAEI.pop"
add "PYRO", "DATARO.pop"
add "Trailer", "DATALE00.pop"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Pyro\Comm2\1.0", "DirIns", "\DATA\WOFIP\"

Case "Need for Speed - Underground 1"
add "EA/THX Logos", "na_boot.mad"
add "Drive Safe .... ya right", "PSA.mad"
add "E3 Trailer", "e3_title.mad"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\EA GAMES\Need For Speed Underground", "Install Dir", "Movies\"

Case "Adrenaline Extreme Show"
add "1C [video]", "1C.ogg"
add "1C [audio]", "1C.audio.ogg"
add "Gaijin [video]", "gaijin.ogg"
add "Gaijin [audio]", "gaijin.audio.ogg"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Gaijin Entertainment\Adrenalin", "InstallPath", "\adrenalin\video\"

Case "Reservoir Dogs"
add "EIDOS", "eidos.bik"
add "LionsGate", "lions.bik"
add "Volatile Games", "volatile.bik"
add "Random Video", "introvid.bik"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Eidos\Reservoir Dogs", "Location", "Video\"

Case "Bad Day L.A."
add "Enlight Soft", "enlight.wmv"
add "Mauretania", "MAURETANIA.wmv"
add "Trailer", "intro.avi"
dir "HKEY_LOCAL_MACHINE\SOFTWARE\Enlight Software\Bad Day LA", "Path", "\movie\"



Case Else
NIMP.list.ListItems.Clear
NIMP.ret.Caption = ""
End Select
End With
End Function

