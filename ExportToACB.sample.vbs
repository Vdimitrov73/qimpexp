Option Explicit
' ==================================================================
'  ExportToACB.sample.vbs  -  SAMPLE QImpExp Quicken 2009 Integration Script
'
'  SAMPLE ONLY — copy to ExportToACB.vbs and fill in your own values.
'  The real ExportToACB.vbs is personal and git-ignored; never commit it.
'  account_periods.json / security_map.json live in the same folder
'  as the ACB workbook.
'
'  Step 1: Drives Quicken 2009 to export a QIF for the current
'          calendar year from your Quicken account to a fixed path.
'  Step 2: Calls qimpexp --import-qif to insert any new/missing
'          Buy/Sell/ROC/Phantom transactions into the ACB workbook.
'
'  Date rules (matches QImpExp conventions):
'    Buy / Sell    -> Quicken stores trade date
'                     qimpexp converts trade -> settlement (T+1) for ACB
'    ROC / Phantom -> record date used in both Quicken and ACB
'
'  Register in QHIMENU.INI:
'    [QHI]
'    ExeName=C:\path\to\QImpExp\ExportToACB.vbs
'    MenuString=ExportToACB
'    IntuitID=1007
'    InformExec=FALSE
' ==================================================================

Dim wsh, fso
Set wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' -- Configuration --------------------------------------------------

Dim yearNow  : yearNow  = CStr(Year(Date))
Dim dateFrom : dateFrom = yearNow & "-01-01"
Dim dateTo   : dateTo   = yearNow & "-12-31"

Const ACCOUNT  = "YOUR QUICKEN ACCOUNT NAME"
' Number of "c" key presses to reach your account in the export dialog's
' account dropdown (depends on your account list; adjust to fit).
Const ACCOUNT_C_PRESSES = 1

Dim baseDir, QIF_PATH, QIF_DIR, ACB_PATH, EXE_PATH, PY_PATH
baseDir = wsh.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Tax Documents\QImpExp\"

QIF_DIR  = baseDir
QIF_PATH = baseDir & ACCOUNT & ".QIF"
ACB_PATH = wsh.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Tax Documents\acb_worksheet.xlsx"
EXE_PATH = baseDir & "qimpexp.exe"
PY_PATH  = baseDir & "qimpexp.py"

' -- Menu navigation key sequences ---------------------------------
'    Quicken 2009 Canadian: File -> Export -> QIF File...
'    Verified accelerators:
'      %f = Alt+F   opens File menu
'       e = Export  sub-menu   (change to "x" if your menu shows "eXport")
'       q = QIF File...        (change to "e" if shown as "Export to QIF")

Const MENU_FILE   = "%f"
Const MENU_EXPORT = "e"
Const MENU_QIF    = "q"

' ==================================================================
'  Helpers
' ==================================================================

Function ActivateWindow(title, retries)
    Dim i
    For i = 1 To retries
        If wsh.AppActivate(title) Then
            WScript.Sleep 300
            ActivateWindow = True
            Exit Function
        End If
        WScript.Sleep 500
    Next
    MsgBox "Cannot activate window: """ & title & """" & vbCrLf & vbCrLf & _
           "Is Quicken running and fully loaded?", _
           vbCritical + vbSystemModal, "ExportToACB"
    WScript.Quit
End Function

Sub SK(keys, ms)
    wsh.SendKeys keys
    WScript.Sleep ms
End Sub

Sub Fatal(msg)
    MsgBox msg, vbCritical + vbSystemModal, "ExportToACB"
    WScript.Quit
End Sub

' ==================================================================
'  PRE-FLIGHT CHECKS
' ==================================================================

If Not fso.FolderExists(QIF_DIR) Then
    fso.CreateFolder QIF_DIR
End If

If Not fso.FileExists(ACB_PATH) Then
    Fatal "ACB workbook not found:" & vbCrLf & ACB_PATH
End If

If Not fso.FileExists(EXE_PATH) And Not fso.FileExists(PY_PATH) Then
    Fatal "qimpexp not found." & vbCrLf & vbCrLf & _
          "Expected exe: " & EXE_PATH & vbCrLf & _
          "Or script:    " & PY_PATH
End If

If fso.FileExists(QIF_PATH) Then
    On Error Resume Next
    fso.DeleteFile QIF_PATH
    If Err.Number <> 0 Then
        Fatal "Cannot delete the existing QIF file:" & vbCrLf & QIF_PATH & vbCrLf & vbCrLf & _
              "It may be open in another application." & vbCrLf & _
              "Close it and retry."
    End If
    On Error GoTo 0
End If

' ==================================================================
'  STEP 1 - Export QIF from Quicken 2009
' ==================================================================

ActivateWindow "Quicken", 10
SK "{ESC}{ESC}{ESC}", 600

SK MENU_FILE,   500
SK MENU_EXPORT, 400
SK MENU_QIF,    800

ActivateWindow "QIF Export", 12

' -- File path
SK "^a{DEL}", 200
SK QIF_PATH, 500

' -- Account dropdown
SK "{TAB}{TAB}", 300
SK "{HOME}", 200

Dim i
For i = 1 To ACCOUNT_C_PRESSES
    SK "c", 180
Next
WScript.Sleep 500

' -- From date
SK "{TAB}", 300
SK "^a{DEL}", 100
SK dateFrom, 300

' -- To date
SK "{TAB}", 300
SK "^a{DEL}", 100
SK dateTo, 300

SK "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}", 300
SK "{ENTER}", 3000
WScript.Sleep 1500


If Not fso.FileExists(QIF_PATH) Then
    Fatal "QIF export failed - file was not created:" & vbCrLf & vbCrLf & _
          QIF_PATH & vbCrLf & vbCrLf & _
          "Troubleshooting tips:" & vbCrLf & _
          "  1. Open Quicken and try File -> Export -> QIF File manually." & vbCrLf & _
          "  2. Verify the account name matches EXACTLY: " & ACCOUNT & vbCrLf & _
          "  3. Check menu accelerators (MENU_EXPORT / MENU_QIF constants)."
End If

If DateDiff("s", fso.GetFile(QIF_PATH).DateLastModified, Now) > 120 Then
    MsgBox "Warning: The QIF file timestamp is older than 2 minutes." & vbCrLf & _
           "The export may not have completed correctly." & vbCrLf & vbCrLf & _
           "Continuing - verify the ACB workbook afterward.", _
           vbExclamation + vbSystemModal, "ExportToACB"
End If

' ==================================================================
'  STEP 2 - Import QIF into ACB workbook via qimpexp
' ==================================================================

Dim runner
If fso.FileExists(EXE_PATH) Then
    runner = """" & EXE_PATH & """"
Else
    runner = "python """ & PY_PATH & """"
End If

Dim qArgs
qArgs = " --acb """         & ACB_PATH & """" & _
        " --import-qif """ & QIF_PATH & """" & _
        " --year "         & yearNow  & _
        " --mode full"     & _
        " --verbose"

Dim batPath : batPath = fso.GetSpecialFolder(2) & "\qimpexp_run.bat"
Dim bat     : Set bat = fso.CreateTextFile(batPath, True)
bat.WriteLine "@echo off"
bat.WriteLine "title QImpExp - Importing " & yearNow & " QIF into ACB Workbook"
bat.WriteLine "echo."
bat.WriteLine "echo  ============================================"
bat.WriteLine "echo   QImpExp - ACB Workbook Update"
bat.WriteLine "echo  ============================================"
bat.WriteLine "echo   Year  :  " & yearNow
bat.WriteLine "echo   QIF   :  " & QIF_PATH
bat.WriteLine "echo   ACB   :  " & ACB_PATH
bat.WriteLine "echo   Mode  :  full  (Buy + Sell + ROC + Phantom)"
bat.WriteLine "echo  ============================================"
bat.WriteLine "echo."
bat.WriteLine runner & qArgs
bat.WriteLine "echo."
bat.WriteLine "echo  ============================================"
bat.WriteLine "echo   Import finished - review output above."
bat.WriteLine "echo   Press any key to close this window."
bat.WriteLine "echo  ============================================"
bat.WriteLine "pause > nul"
bat.Close

wsh.Run """" & batPath & """", 1, True

On Error Resume Next
fso.DeleteFile batPath
On Error GoTo 0

MsgBox "Export + Import complete!" & vbCrLf & vbCrLf & _
       "Year   :  " & yearNow & vbCrLf & _
       "QIF    :  " & QIF_PATH & vbCrLf & _
       "ACB    :  " & ACB_PATH & vbCrLf & vbCrLf & _
       "Check the ACB workbook to verify the new rows.", _
       vbInformation + vbSystemModal, "ExportToACB"