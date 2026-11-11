' =================================================================================================
' Outlook VBA – Retention Column Toolkit (Live + Manual batch + Mailbox batch + Field/View utils)
' Works in Outlook for Windows (Classic). Not supported in "New Outlook".
' =================================================================================================

Option Explicit

' ======== MASTER TOGGLE ========
Private Const ENABLE_LIVE_STAMPING As Boolean = True   ' set False to disable event-driven stamping

' ======== TUNING (Manual) ========
Private Const BATCH_SIZE_MANUAL As Long = 6000         ' per-folder batch cap for manual runs
Private Const YIELD_EVERY As Long = 50                 ' UI breathing room while looping
Private Const SKIP_ALREADY_STAMPED As Boolean = False  ' True = skip items with a non-empty Retention already

' ======== MAPI retention properties (MRM) ========
Private Const PR_POLICY_TAG As String = "http://schemas.microsoft.com/mapi/proptag/0x30190102"        ' PT_BINARY
Private Const PR_RETENTION_PERIOD As String = "http://schemas.microsoft.com/mapi/proptag/0x301A0003"  ' PT_LONG (days)
Private Const PR_RETENTION_DATE As String = "http://schemas.microsoft.com/mapi/proptag/0x301C0040"    ' PT_SYSTIME

' ======== MODULE-LEVEL OBJECTS / EVENTS ========
Private WithEvents olExp As Outlook.Explorer
Attribute olExp.VB_VarHelpID = -1
Private WithEvents olExplorers As Outlook.Explorers
Attribute olExplorers.VB_VarHelpID = -1
Private WithEvents InboxItems As Outlook.Items
Attribute InboxItems.VB_VarHelpID = -1
Private WithEvents SentItems As Outlook.Items
Attribute SentItems.VB_VarHelpID = -1

' =================================================================================================
'                                   STARTUP / EVENT WIRING
' =================================================================================================

Private Sub Application_Startup()
    WireEventSinks
End Sub

Private Sub Application_Quit()
    On Error Resume Next
    Set InboxItems = Nothing
    Set SentItems = Nothing
    Set olExp = Nothing
    Set olExplorers = Nothing
End Sub

Private Sub WireEventSinks()
    On Error Resume Next
    Set olExplorers = Application.Explorers
    If Application.ActiveExplorer Is Nothing Then
        If Application.Explorers.Count > 0 Then
            Set olExp = Application.Explorers(1)
        Else
            Set olExp = Nothing
        End If
    Else
        Set olExp = Application.ActiveExplorer
    End If
    Dim ns As Outlook.NameSpace: Set ns = Application.Session
    Set InboxItems = ns.GetDefaultFolder(olFolderInbox).Items
    Set SentItems = ns.GetDefaultFolder(olFolderSentMail).Items
    On Error GoTo 0
End Sub

' Track newly opened explorers so SelectionChange works even if Outlook started minimized
Private Sub olExplorers_NewExplorer(ByVal Explorer As Outlook.Explorer)
    On Error Resume Next
    Set olExp = Explorer
End Sub

' =================================================================================================
'                                LIVE STAMPING (as you work)
' =================================================================================================

' Stamp when you select messages (covers Quick Step policy changes while triaging)
Private Sub olExp_SelectionChange()
    If Not ENABLE_LIVE_STAMPING Then Exit Sub
    On Error Resume Next
    If olExp Is Nothing Then Exit Sub
    Dim sel As Outlook.Selection: Set sel = olExp.Selection
    If sel Is Nothing Then Exit Sub
    Dim it As Object
    For Each it In sel
        If TypeOf it Is Outlook.MailItem Then StampOneItemRetention it
    Next
End Sub

' Stamp new arrivals in Inbox
Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    If Not ENABLE_LIVE_STAMPING Then Exit Sub
    On Error Resume Next
    If TypeOf Item Is Outlook.MailItem Then StampOneItemRetention Item
End Sub

' Stamp when Inbox items change (e.g., you apply a Retention policy via Quick Step)
Private Sub InboxItems_ItemChange(ByVal Item As Object)
    If Not ENABLE_LIVE_STAMPING Then Exit Sub
    On Error Resume Next
    If TypeOf Item Is Outlook.MailItem Then StampOneItemRetention Item
End Sub

' Stamp when items land in Sent Items
Private Sub SentItems_ItemAdd(ByVal Item As Object)
    If Not ENABLE_LIVE_STAMPING Then Exit Sub
    On Error Resume Next
    If TypeOf Item Is Outlook.MailItem Then StampOneItemRetention Item
End Sub

' Also stamp right before sending (belt-and-suspenders)
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    If Not ENABLE_LIVE_STAMPING Then Exit Sub
    On Error Resume Next
    If TypeOf Item Is Outlook.MailItem Then StampOneItemRetention Item
End Sub

' =================================================================================================
'                                        MANUAL STAMPERS
' =================================================================================================

' Run on the CURRENT folder, up to BATCH_SIZE_MANUAL items in that folder
Public Sub Stamp_Current_Folder_Batched_Retention()
    Dim fld As Outlook.MAPIFolder, itms As Outlook.Items
    Dim i As Long, processed As Long

    If Application.ActiveExplorer Is Nothing Then
        MsgBox "Open a mail folder first.", vbExclamation: Exit Sub
    End If
    Set fld = Application.ActiveExplorer.CurrentFolder
    If fld.DefaultItemType <> olMailItem Then
        MsgBox "This folder isn't a mail folder.", vbExclamation: Exit Sub
    End If

    Set itms = fld.Items
    On Error Resume Next: itms.Sort "[ReceivedTime]", True: On Error GoTo 0

    For i = itms.Count To 1 Step -1
        Dim it As Object
        On Error Resume Next: Set it = itms.Item(i)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextItem
        On Error GoTo 0

        If TypeOf it Is Outlook.MailItem Then
            Dim mi As Outlook.MailItem: Set mi = it
            If Left$(mi.MessageClass, 8) = "IPM.Note" Then
                If Not (SKIP_ALREADY_STAMPED And IsAlreadyStampedValue(mi)) Then
                    If StampOneItemRetention(mi) Then
                        processed = processed + 1
                        If (processed Mod YIELD_EVERY) = 0 Then DoEvents
                        If processed >= BATCH_SIZE_MANUAL Then Exit For
                    End If
                End If
            End If
        End If
NextItem:
    Next

    MsgBox "Stamped " & processed & " item(s) in: " & fld.FolderPath & vbCrLf & _
           "Run again for the next batch.", vbInformation
End Sub

' ===================== Mailbox-wide batch (total cap) =====================

' One-click: stamps up to 6,000 across the entire CURRENT mailbox tree (respects per-folder cap)
Public Sub Stamp_Current_Mailbox_Batched_Retention_6000()
    Stamp_Current_Mailbox_Batched_Retention_Limited 6000
End Sub

' Same as above, but you can pass any total limit (e.g., 3000, 5000, 10000)
Public Sub Stamp_Current_Mailbox_Batched_Retention_Limited(ByVal limitTotal As Long)
    If Application.ActiveExplorer Is Nothing Then
        MsgBox "Open a folder inside the mailbox you want to process.", vbExclamation
        Exit Sub
    End If

    Dim root As Outlook.folder
    On Error Resume Next
    Set root = Application.ActiveExplorer.CurrentFolder.Store.GetRootFolder
    If Err.Number <> 0 Or root Is Nothing Then
        Err.Clear
        MsgBox "Couldn't resolve the mailbox root. Select a folder in the target mailbox and try again.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Dim skipped As String
    Dim totalDone As Long
    totalDone = StampMailboxTreeLimited(root, limitTotal, skipped)

    Dim msg As String
    msg = "Mailbox pass complete." & vbCrLf & _
          "Stamped: " & totalDone & " item(s) (cap " & limitTotal & ")."
    If Len(skipped) > 0 Then msg = msg & vbCrLf & vbCrLf & "Skipped:" & vbCrLf & skipped
    MsgBox msg, vbInformation
End Sub

' Depth-first walk of the mailbox tree with an overall remaining counter.
' Uses the per-folder batcher (StampFolderBatchRetention).
Private Function StampMailboxTreeLimited(fld As Outlook.folder, ByVal remaining As Long, ByRef skipped As String) As Long
    Dim doneHere As Long, got As Long
    If remaining <= 0 Then Exit Function

    On Error GoTo node_err
    If fld.DefaultItemType = olMailItem Then
        Dim target As Long
        If remaining < BATCH_SIZE_MANUAL Then target = remaining Else target = BATCH_SIZE_MANUAL
        doneHere = StampFolderBatchRetention(fld, target)
        remaining = remaining - doneHere
    End If

    Dim subf As Outlook.folder
    For Each subf In fld.Folders
        If remaining <= 0 Then Exit For
        got = StampMailboxTreeLimited(subf, remaining, skipped)
        StampMailboxTreeLimited = StampMailboxTreeLimited + got
        remaining = remaining - got
    Next

    StampMailboxTreeLimited = StampMailboxTreeLimited + doneHere
    Exit Function

node_err:
    Dim pathText As String
    On Error Resume Next
    pathText = fld.FolderPath
    If Len(pathText) = 0 Then pathText = "(unknown folder)"
    skipped = skipped & "- " & pathText & " (access error)" & vbCrLf
    Err.Clear
    On Error GoTo 0
End Function

' Per-folder batcher used by both manual and mailbox-wide runs
Private Function StampFolderBatchRetention(fld As Outlook.folder, ByVal limit As Long) As Long
    Dim itms As Outlook.Items: Set itms = fld.Items
    On Error Resume Next: itms.Sort "[ReceivedTime]", True: On Error GoTo 0

    Dim i As Long, done As Long
    For i = itms.Count To 1 Step -1
        Dim it As Object
        On Error Resume Next: Set it = itms.Item(i)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextItem2
        On Error GoTo 0

        If TypeOf it Is Outlook.MailItem Then
            Dim mi As Outlook.MailItem: Set mi = it
            If Left$(mi.MessageClass, 8) = "IPM.Note" Then
                If Not (SKIP_ALREADY_STAMPED And IsAlreadyStampedValue(mi)) Then
                    If StampOneItemRetention(mi) Then
                        done = done + 1
                        If (done Mod YIELD_EVERY) = 0 Then DoEvents
                        If done >= limit Then Exit For
                    End If
                End If
            End If
        End If
NextItem2:
    Next
    StampFolderBatchRetention = done
End Function

' =================================================================================================
'                               CORE: compute & write Retention text
' =================================================================================================

' Returns True if the "Retention" user prop actually changed (we only Save on change).
' Leaves the value BLANK when no duration can be determined (no "None").
Private Function StampOneItemRetention(mi As Outlook.MailItem) As Boolean
    On Error GoTo EH

    Dim pa As Outlook.PropertyAccessor
    Dim upVal As Outlook.UserProperty, upHex As Outlook.UserProperty
    Dim valText As String, changed As Boolean
    Dim days As Long, haveDays As Boolean
    Dim expDate As Date, haveExp As Boolean

    Set pa = mi.PropertyAccessor

    ' 1) Retention period (days) if present
    On Error Resume Next
    days = CLng(pa.GetProperty(PR_RETENTION_PERIOD))
    If Err.Number = 0 And days > 0 Then haveDays = True
    Err.Clear: On Error GoTo EH

    ' 2) Else derive from explicit/computed expiration date
    If Not haveDays Then
        On Error Resume Next
        expDate = pa.GetProperty(PR_RETENTION_DATE)
        If Err.Number = 0 Then
            haveExp = (expDate > #1/1/1900# And expDate < #1/1/4501#)
        Else
            haveExp = False
        End If
        Err.Clear
        If Not haveExp Then
            expDate = mi.RetentionExpirationDate
            If Err.Number = 0 Then
                haveExp = (expDate > #1/1/1900# And expDate < #1/1/4501#)
            Else
                haveExp = False
            End If
            Err.Clear
        End If
        On Error GoTo EH

        If haveExp Then
            Dim baseDate As Date: baseDate = PickBestBaseDate(mi)
            If baseDate > #1/1/1900# And expDate > baseDate Then
                days = CLng(expDate - baseDate)
                haveDays = (days > 0)
            End If
        End If
    End If

    ' 3) Render compact text
    If haveDays Then
        valText = DaysToFriendly(days)
    Else
        valText = ""    ' unknown ? keep blank
    End If

    ' 4) Write only if changed; also stash TagHex (useful if you map policy tags later)
    Set upVal = EnsureUserProp(mi, "Retention")
    If CStr(upVal.Value) <> valText Then
        upVal.Value = valText
        changed = True
    End If

    On Error Resume Next
    Dim tagBin As Variant: tagBin = pa.GetProperty(PR_POLICY_TAG)
    If Not IsEmpty(tagBin) Then
        Set upHex = EnsureUserProp(mi, "Retention (TagHex)")
        Dim hexStr As String: hexStr = ByteArrayToHex(tagBin)
        If CStr(upHex.Value) <> hexStr Then upHex.Value = hexStr: changed = True
    End If
    Err.Clear: On Error GoTo EH

    If changed Then mi.Save
    StampOneItemRetention = changed
    Exit Function

EH:
    Err.Clear
    StampOneItemRetention = False
End Function

Private Function EnsureUserProp(mi As Outlook.MailItem, name As String) As Outlook.UserProperty
    Dim up As Outlook.UserProperty
    On Error Resume Next
    Set up = mi.UserProperties.Find(name)
    If up Is Nothing Then Set up = mi.UserProperties.Add(name, olText)
    On Error GoTo 0
    Set EnsureUserProp = up
End Function

Private Function IsAlreadyStampedValue(mi As Outlook.MailItem) As Boolean
    Dim up As Outlook.UserProperty
    On Error Resume Next
    Set up = mi.UserProperties.Find("Retention")
    If Not up Is Nothing Then
        If Len(Trim$(CStr(up.Value))) > 0 Then IsAlreadyStampedValue = True
    End If
    On Error GoTo 0
End Function

Private Function PickBestBaseDate(mi As Outlook.MailItem) As Date
    On Error Resume Next
    If mi.ReceivedTime > #1/1/1900# Then
        PickBestBaseDate = mi.ReceivedTime
    ElseIf mi.SentOn > #1/1/1900# Then
        PickBestBaseDate = mi.SentOn
    Else
        PickBestBaseDate = mi.CreationTime
    End If
    On Error GoTo 0
End Function

Private Function DaysToFriendly(ByVal days As Long) As String
    If days <= 0 Then DaysToFriendly = "": Exit Function
    If days < 60 Then
        DaysToFriendly = CStr(days) & IIf(days = 1, " Day", " Days"): Exit Function
    End If
    Dim months As Long: months = CLng(Int((days / 30#) + 0.5))
    If days < 730 Then
        If months < 1 Then months = 1
        DaysToFriendly = CStr(months) & IIf(months = 1, " Mo", " Mos"): Exit Function
    End If
    Dim years As Long: years = CLng(Int((days / 365#) + 0.5))
    If years < 1 Then years = 1
    DaysToFriendly = CStr(years) & IIf(years = 1, " Yr", " Yrs")
End Function

Private Function ByteArrayToHex(v As Variant) As String
    On Error GoTo fail
    Dim b() As Byte, i As Long
    b = v
    Dim hexChars As String: hexChars = "0123456789ABCDEF"
    Dim res() As Byte: ReDim res(1 To (UBound(b) - LBound(b) + 1) * 2)
    Dim hi As Byte, lo As Byte, idx As Long
    For i = LBound(b) To UBound(b)
        hi = (b(i) \ 16) And &HF
        lo = b(i) And &HF
        idx = (i - LBound(b)) * 2
        res(idx + 1) = Asc(Mid$(hexChars, hi + 1, 1))
        res(idx + 2) = Asc(Mid$(hexChars, lo + 1, 1))
    Next
    ByteArrayToHex = StrConv(res, vbUnicode)
    Exit Function
fail:
    ByteArrayToHex = ""
End Function

Private Function SafeFolderPath(fld As Outlook.folder) As String
    On Error Resume Next
    SafeFolderPath = fld.FolderPath
    If Len(SafeFolderPath) = 0 Then SafeFolderPath = "(unknown folder)"
End Function

' =================================================================================================
'                                 FIELD / VIEW UTILITIES (SAFE)
' =================================================================================================

' Register the "Retention" user field on ALL mail folders under the CURRENT mailbox
Public Sub Retention_RegisterField_OnCurrentMailbox()
    If Application.ActiveExplorer Is Nothing Then
        MsgBox "Open a folder in the mailbox first.", vbExclamation: Exit Sub
    End If

    Dim root As Outlook.folder
    On Error Resume Next
    Set root = Application.ActiveExplorer.CurrentFolder.Store.GetRootFolder
    If Err.Number <> 0 Or root Is Nothing Then
        MsgBox "Could not get current mailbox root.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim countFolders As Long, skipped As String
    countFolders = Retention_RegisterField_Tree_Safe(root, skipped)

    Dim msg As String
    msg = "Registered 'Retention' on " & countFolders & " folder(s) in this mailbox."
    If Len(skipped) > 0 Then msg = msg & vbCrLf & vbCrLf & "Skipped:" & vbCrLf & skipped
    MsgBox msg, vbInformation
End Sub

Private Function Retention_RegisterField_Tree_Safe(fld As Outlook.folder, ByRef skipped As String) As Long
    Dim touched As Long

    On Error GoTo node_err
    If fld.DefaultItemType = olMailItem Then
        If Not Retention_AddField_Safe(fld, "Retention") Then
            skipped = skipped & "- " & SafeFolderPath(fld) & " (no write)" & vbCrLf
        Else
            touched = touched + 1
        End If
        Call Retention_AddField_Safe(fld, "Retention (TagHex)") ' best-effort
    End If

    Dim subf As Outlook.folder
    On Error Resume Next
    For Each subf In fld.Folders
        If Err.Number <> 0 Then
            skipped = skipped & "- " & SafeFolderPath(fld) & " (cannot enumerate children)" & vbCrLf
            Err.Clear
            Exit For
        End If
        On Error GoTo child_err
        touched = touched + Retention_RegisterField_Tree_Safe(subf, skipped)
        On Error GoTo 0
    Next
    Retention_RegisterField_Tree_Safe = touched
    Exit Function

node_err:
    skipped = skipped & "- " & SafeFolderPath(fld) & " (access error)" & vbCrLf
    Err.Clear
    Retention_RegisterField_Tree_Safe = touched
    Exit Function

child_err:
    skipped = skipped & "- " & SafeFolderPath(subf) & " (child access error)" & vbCrLf
    Err.Clear
    Resume Next
End Function

Private Function Retention_AddField_Safe(fld As Outlook.folder, ByVal fieldName As String) As Boolean
    On Error GoTo fail
    Dim udfs As Outlook.UserDefinedProperties
    Set udfs = fld.UserDefinedProperties
    If udfs Is Nothing Then GoTo fail

    Dim udp As Outlook.UserDefinedProperty
    Set udp = udfs.Find(fieldName)
    If udp Is Nothing Then Set udp = udfs.Add(fieldName, olText)

    Retention_AddField_Safe = True
    Exit Function
fail:
    Err.Clear
    Retention_AddField_Safe = False
End Function

' Copy your CURRENT view (with the Retention column) to ALL mail folders in CURRENT mailbox
Public Sub Retention_ApplyActiveView_ToCurrentMailbox()
    If Application.ActiveExplorer Is Nothing Then
        MsgBox "Open a mail folder first.", vbExclamation: Exit Sub
    End If

    Dim srcView As Outlook.View
    Set srcView = Application.ActiveExplorer.CurrentView
    On Error Resume Next: srcView.Save: On Error GoTo 0

    Dim root As Outlook.folder
    On Error Resume Next
    Set root = Application.ActiveExplorer.CurrentFolder.Store.GetRootFolder
    If Err.Number <> 0 Or root Is Nothing Then
        MsgBox "Could not get current mailbox root.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim countFolders As Long
    countFolders = Retention_CopyViewToTree(root, srcView, "Retention View")
    MsgBox "Applied your current view to " & countFolders & " mail folder(s) in this mailbox.", vbInformation
End Sub

Private Function Retention_CopyViewToTree(fld As Outlook.folder, srcView As Outlook.View, ByVal viewName As String) As Long
    Dim applied As Long
    If fld.DefaultItemType = olMailItem Then
        On Error Resume Next
        Dim existing As Outlook.View
        Set existing = fld.Views(viewName)
        If Not existing Is Nothing Then existing.Delete  ' avoid duplicates

        Dim v As Outlook.View
        Set v = srcView.Copy(viewName, fld)
        v.Save
        v.Apply
        applied = applied + 1
        On Error GoTo 0
    End If

    Dim subf As Outlook.folder
    For Each subf In fld.Folders
        applied = applied + Retention_CopyViewToTree(subf, srcView, viewName)
    Next
    Retention_CopyViewToTree = applied
End Function

