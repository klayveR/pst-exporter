#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>

#include "CreatePrfFile.au3"

Opt("WinTitleMatchMode", 2)

Global $sOpmLogFilePath = @LocalAppDataDir & "\Temp\Outlook-Protokoll\OPMLog.log"
Global $sPrfFilePath = @ScriptDir & "\profile.prf"

ShowGUI()

Func ShowGUI()
    Local $GUI = GUICreate("PST Exporter", 400, 180)

    GUICtrlCreateLabel("E-Mail Adresse", 5, 10, 100, 20)
    Local $idInputEmail = GUICtrlCreateInput("", 110, 5, 280, 20)

    GUICtrlCreateLabel("Passwort", 5, 35, 100, 20)
    Local $idInputPassword = GUICtrlCreateInput("", 110, 30, 280, 20)

    GUICtrlCreateLabel("POP3-Server", 5, 60, 100, 20)
    Local $isInputPopServer = GUICtrlCreateInput("pop.ionos.de", 110, 55, 280, 20)

    GUICtrlCreateLabel("SMTP-Server", 5, 85, 100, 20)
    Local $idInputSmtpServer = GUICtrlCreateInput("smtp.ionos.de", 110, 80, 280, 20)

    GUICtrlCreateLabel("Outlook.exe", 5, 110, 100, 20)
    Local $idInputOutlookExePath = GUICtrlCreateInput("C:\Program Files (x86)\Microsoft Office\Office12\outlook.exe", 110, 105, 185, 20)
    Local $idButtonOutlookExePath = GUICtrlCreateButton("Auswählen...", 300, 104, 90, 22)

    GUICtrlCreateLabel("PST-Verzeichnis", 5, 135, 100, 20)
    Local $idInputPstPath = GUICtrlCreateInput("D:\", 110, 130, 185, 20)
    Local $idButtonPstPath = GUICtrlCreateButton("Auswählen...", 300, 129, 90, 22)

    Local $idButtonStart = GUICtrlCreateButton("Start", 5, 155, 385, 20)

    GUISetState(@SW_SHOW, $GUI)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop

            Case $idButtonOutlookExePath
                Local $sSelectedOutlookExePath = FileOpenDialog("Pfad zu Outlook.exe", "C:\", "Ausführbare Datei (*.exe)")
                If Not @error Then
                    GUICtrlSetData ($idInputOutlookExePath, $sSelectedOutlookExePath)
                EndIf

            Case $idButtonPstPath
                Local $sSelectedPstPath = FileSelectFolder("PST-Verzeichnis", "C:\")
                If Not $sSelectedPstPath = "" Then
                    GUICtrlSetData ($idInputPstPath, $sSelectedPstPath)
                EndIf

            Case $idButtonStart
                If GUICtrlRead($idInputEmail) = "" Then ContinueLoop MsgBox(48, "Warnung", "E-Mail Adresse ist erforderlich.")
                If GUICtrlRead($idInputPassword) = "" Then ContinueLoop MsgBox(48, "Warnung", "Passwort ist erforderlich.")
                If GUICtrlRead($isInputPopServer) = "" Then ContinueLoop MsgBox(48, "Warnung", "POP3-Server ist erforderlich.")
                If GUICtrlRead($idInputSmtpServer) = "" Then ContinueLoop MsgBox(48, "Warnung", "SMTP-Server ist erforderlich.")
                If Not FileExists(GUICtrlRead($idInputOutlookExePath)) Then ContinueLoop MsgBox(48, "Warnung", "Outlook.exe existiert nicht")
                If Not FileExists(GUICtrlRead($idInputPstPath)) Then ContinueLoop MsgBox(48, "Warnung", "PST-Verzeichnis existiert nicht")

                GUISetState(@SW_DISABLE, $GUI)
                ExportPST(GUICtrlRead($idInputEmail), GUICtrlRead($idInputPassword), GUICtrlRead($isInputPopServer), GUICtrlRead($idInputSmtpServer), GUICtrlRead($idInputOutlookExePath), GUICtrlRead($idInputPstPath))
                GUISetState(@SW_ENABLE, $GUI)
                WinActivate("PST Exporter")
        EndSwitch
    WEnd

    GUIDelete($GUI)
EndFunc

Func ExportPST($sEmail, $sPassword, $sPopServer, $sSmtpServer, $sOutlookExePath, $sPstPath)
    ; Delete existing prf
    FileDelete($sPrfFilePath)

    ; Create new prf
    Local $iPrfCreated = CreatePrfFile($sPrfFilePath, $sPstPath, $sPopServer, $sSmtpServer, $sEmail)
    If Not $iPrfCreated Then
        MsgBox (16, "Fehler", "PRF konnte nicht erstellt werden.")
        Exit(1)
    EndIf

    ; Delete previous log
    FileDelete($sOpmLogFilePath)

    ; Close Outlook and open again importing prf
    WinClose("Microsoft Outlook")
    ShellExecute($sOutlookExePath, '/importprf "' & $sPrfFilePath & '"')

    Local $hOutlook = WinWait("Microsoft Outlook (Protokollierung aktiviert)", "", 10)
    If $hOutlook = 0 Then
        MsgBox (16, "Fehler", "Microsoft Outlook konnte nicht geöffnet werden.")
        BlockInput($BI_ENABLE)
        Exit(1)
    EndIf

    BlockInput($BI_DISABLE)

    Local $hPassword = WinWait("Netzwerk-Kennwort eingeben", "Geben Sie Ihren Benutzernamen und Ihr Kennwort ein.", 10)
    If Not $hPassword = 0 Then
        WinActivate($hPassword)
        Send($sPassword, $SEND_RAW)
        Send("{ENTER}")
    EndIf

    BlockInput($BI_ENABLE)

    Local $sOutlookTitle = WinGetTitle($hOutlook)
    Local $iLogEnabled = StringInStr($sOutlookTitle, "Protokollierung aktiviert")
    If Not $iLogEnabled Then
        MsgBox (48, "Fehler", "Outlook Protokollierung nicht aktiviert. Bitte das Fenster manuell schließen nachdem die Übertragung abgeschlossen ist." & @CRLF & @CRLF & "Extras -> Optionen -> Weitere -> Erweiterte Optionen... -> Protokollierung aktivieren (Problembehandlung)")
    EndIf

    While 1
        If Not WinExists ($hOutlook) Then
            ExitLoop
        EndIf

        Local $hFileOpen = FileOpen($sOpmLogFilePath, $FO_READ)
        If $hFileOpen = -1 Then
            ; Error, it's whatever, might not exist yet
        Else
            Local $sFileRead = FileRead($hFileOpen)
            Local $iPopSignedOff = StringInStr($sFileRead, "+OK POP server signing off")

            If $iPopSignedOff Then
                Sleep(10000)
                WinClose("Microsoft Outlook")

                Local $iRsfCompleted = StringInStr($sFileRead, "ReportStatus: RSF_COMPLETED")
                If $iRsfCompleted Then
                    Local $iRsfSuccess = StringInStr($sFileRead, "ReportStatus: RSF_COMPLETED, hr = 0x00000000")

                    If Not $iRsfSuccess Then
                        Local $iRsfInvalidLogin = StringInStr($sFileRead, "ReportStatus: RSF_COMPLETED, hr = 0x8004210a")

                        If $iRsfInvalidLogin Then
                            MsgBox (16, "Fehler", "Anmeldedaten sind inkorrekt.")
                        Else
                            MsgBox (16, "Fehler", "Es ist ein Fehler bei der Übertragung aufgetreten.")
                        EndIf
                    Else
                        MsgBox (64, "Erfolg", "PST für " & $sEmail & " erfolgreich exportiert.")
                    EndIf
                EndIf

                ExitLoop
            EndIf

            FileClose($hFileOpen)
        EndIf
        
        Sleep(5000)
    WEnd

    ; Delete prf
    FileDelete($sPrfFilePath)
EndFunc