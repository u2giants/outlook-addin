Option Explicit

' ================================================================
'  Attachment & Inline Image Toolkit (single module)
'  - RemoveInlineImages_SelectedEmails: inline-only (your working version)
'  - StripAttachments_SelectedEmails   : non-inline attachments (save/no-save)
'  - StripEverything_SelectedEmails    : inline first, then non-inline
' ================================================================


' ---------------------------
' PUBLIC ENTRY POINTS
' ---------------------------

' Inline images only (uses CID detection + HTML/VML scrub)
Public Sub RemoveInlineImages_SelectedEmails()
    Dim sel As Selection
    Set sel = Application.ActiveExplorer.Selection
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "Select one or more emails first.", vbExclamation
        Exit Sub
    End If

    Dim i As Long, done As Long, skipped As Long
    ScreenUpdating False
    For i = 1 To sel.Count
        If TypeName(sel(i)) = "MailItem" Then
            If RemoveInlineFromMailItem(sel(i)) Then
                done = done + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next i
    ScreenUpdating True
    MsgBox "Inline images removed from: " & done & vbCrLf & "Skipped/Failed: " & skipped, vbInformation
End Sub

' Non-inline attachments only; optional save-to-folder; keeps inline images by default.
Public Sub StripAttachments_SelectedEmails()
    Dim sel As Selection
    Dim resp As VbMsgBoxResult
    Dim savePath As String
    Dim skipInline As Boolean
    Dim addNote As Boolean

    Set sel = Application.ActiveExplorer.Selection
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "Select one or more emails first.", vbExclamation
        Exit Sub
    End If

    skipInline = True    ' leave inline images alone (recommended)
    addNote = True       ' add a small banner listing removed attachments

    resp = MsgBox("Save copies of attachments before removing?" & vbCrLf & _
                  "Yes = pick a folder and save; No = remove without saving.", _
                  vbYesNoCancel + vbQuestion, "Strip Attachments")
    If resp = vbCancel Then Exit Sub
    If resp = vbYes Then
        savePath = PickFolderPath("Choose where to save the stripped attachments")
        If Len(savePath) = 0 Then Exit Sub
    Else
        savePath = vbNullString
    End If

    ScreenUpdating False
    Dim i As Long, processed As Long, failed As Long
    For i = 1 To sel.Count
        On Error Resume Next
        If TypeName(sel(i)) = "MailItem" Then
            If ProcessMailItem(sel(i), savePath, skipInline, addNote) Then
                processed = processed + 1
            Else
                failed = failed + 1
            End If
        End If
        On Error GoTo 0
    Next i
    ScreenUpdating True

    Dim msg As String
    msg = "Processed: " & processed & vbCrLf & "Skipped/Failed: " & failed
    If Len(savePath) > 0 Then msg = msg & vbCrLf & "Saved to: " & savePath
    MsgBox msg, vbInformation, "Strip Attachments"
End Sub

' One-click convenience: remove inline images first, then strip non-inline attachments.
Public Sub StripEverything_SelectedEmails()
    Dim sel As Selection
    Set sel = Application.ActiveExplorer.Selection
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "Select one or more emails first.", vbExclamation
        Exit Sub
    End If

    ' Ask once about saving attachments
    Dim resp As VbMsgBoxResult
    Dim savePath As String
    resp = MsgBox("Save copies of (non-inline) attachments before removing?" & vbCrLf & _
                  "Yes = pick a folder and save; No = remove without saving.", _
                  vbYesNoCancel + vbQuestion, "Strip Everything")
    If resp = vbCancel Then Exit Sub
    If resp = vbYes Then
        savePath = PickFolderPath("Choose where to save the stripped attachments")
        If Len(savePath) = 0 Then Exit Sub
    Else
        savePath = vbNullString
    End If

    ScreenUpdating False

    Dim i As Long, inlineDone As Long, inlineSkipped As Long
    Dim attachDone As Long, attachFail As Long

    ' Pass 1: inline images
    For i = 1 To sel.Count
        If TypeName(sel(i)) = "MailItem" Then
            If RemoveInlineFromMailItem(sel(i)) Then
                inlineDone = inlineDone + 1
            Else
                inlineSkipped = inlineSkipped + 1
            End If
        End If
    Next i

    ' Pass 2: non-inline attachments (keep inline; most CIDs already gone)
    For i = 1 To sel.Count
        If TypeName(sel(i)) = "MailItem" Then
            If ProcessMailItem(sel(i), savePath, True, True) Then
                attachDone = attachDone + 1
            Else
                attachFail = attachFail + 1
            End If
        End If
    Next i

    ScreenUpdating True

    Dim msg As String
    msg = "Inline images removed from: " & inlineDone & "  (Skipped/Failed: " & inlineSkipped & ")" & vbCrLf & _
          "Attachments stripped from: " & attachDone & "  (Skipped/Failed: " & attachFail & ")"
    If Len(savePath) > 0 Then msg = msg & vbCrLf & "Saved copies to: " & savePath
    MsgBox msg, vbInformation, "Strip Everything"
End Sub


' ---------------------------
' INLINE IMAGE ENGINE
' ---------------------------

' Core: scrub HTML references to cid: images (and Outlook VML) then delete matching inline attachments.
Private Function RemoveInlineFromMailItem(ByVal mi As MailItem) As Boolean
    On Error GoTo fail

    If mi.BodyFormat <> olFormatHTML Then
        ' Non-HTML: best-effort—remove inline-ish attachments only.
        RemoveInlineAttachments mi, Nothing
        mi.Save
        RemoveInlineFromMailItem = True
        Exit Function
    End If

    Dim html As String
    html = mi.HTMLBody

    ' Collect CIDs referenced by <img> or VML <v:imagedata>.
    Dim cids As Object: Set cids = CreateObject("Scripting.Dictionary")
    cids.CompareMode = 1 ' TextCompare

    RxCollectCids html, "src\s*=\s*['""]cid:([^'""]+)['""]", cids
    RxCollectCids html, "(?:src|o:href)\s*=\s*['""]cid:([^'""]+)['""]", cids

    ' Remove <img ... src="cid:..."> tags.
    html = RxReplace(html, "<img\b[^>]*\bsrc\s*=\s*['""]cid:[^'""]+['""][^>]*>", "")

    ' Remove Outlook/Word VML image nodes that point at cid:...
    html = RxReplace(html, "<v:imagedata\b[^>]*\b(?:src|o:href)\s*=\s*['""]cid:[^'""]+['""][^>]*/?>", "")

    ' Clean up now-empty VML shapes.
    html = RxReplace(html, "<v:shape\b[^>]*>\s*</v:shape>", "")

    ' Optional: compress multiple blank lines created by removals.
    html = RxReplace(html, "(?:\r?\n\s*){3,}", vbCrLf & vbCrLf)

    mi.HTMLBody = html

    ' Now remove the actual inline attachments whose CIDs were referenced.
    RemoveInlineAttachments mi, cids

    mi.Save
    RemoveInlineFromMailItem = True
    Exit Function

fail:
    RemoveInlineFromMailItem = False
End Function

' Deletes attachments that are "inline" and (if a CID list is supplied) whose CID was referenced in HTML.
Private Sub RemoveInlineAttachments(ByVal mi As MailItem, ByVal cids As Object)
    On Error Resume Next
    Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

    Dim i As Long
    For i = mi.Attachments.Count To 1 Step -1
        Dim att As Attachment
        Set att = mi.Attachments(i)

        Dim cid As String
        Dim hidden As Variant
        cid = att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
        hidden = att.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN)

        Dim isInline As Boolean
        isInline = False

        ' Consider inline if it has a CID, is flagged hidden, or is not by-value.
        If Len(cid) > 0 Then isInline = True
        If VarType(hidden) = vbBoolean And hidden = True Then isInline = True
        If att.Type <> olByValue Then isInline = True

        If isInline Then
            If cids Is Nothing Then
                ' Non-HTML path: delete all inline-ish attachments.
                mi.Attachments.Remove i
            Else
                ' HTML path: only delete if its CID was actually referenced.
                If Len(cid) > 0 And cids.Exists(LCase$(cid)) Then
                    mi.Attachments.Remove i
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' ---- Regex helpers (late-bound; no reference needed) ----

Private Function RxReplace(ByVal s As String, ByVal pattern As String, Optional ByVal repl As String = "") As String
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = pattern
    rx.Global = True
    rx.Multiline = True
    rx.IgnoreCase = True
    RxReplace = rx.Replace(s, repl)
End Function

Private Sub RxCollectCids(ByVal s As String, ByVal pattern As String, ByRef dict As Object)
    Dim rx As Object, m As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = pattern
    rx.Global = True
    rx.Multiline = True
    rx.IgnoreCase = True
    For Each m In rx.Execute(s)
        If m.SubMatches.Count > 0 Then
            dict(LCase$(CStr(m.SubMatches(0)))) = True
        End If
    Next
End Sub


' ---------------------------
' NON-INLINE ATTACHMENT ENGINE
' ---------------------------

' Core worker: optionally saves, then removes attachments (skipping inline).
Private Function ProcessMailItem(ByVal mi As MailItem, ByVal savePath As String, _
                                 ByVal skipInline As Boolean, ByVal addNote As Boolean) As Boolean
    On Error GoTo fail

    Dim removedInfo As String
    Dim att As Attachment
    Dim idx As Long
    Dim savedCount As Long, removedCount As Long

    For idx = mi.Attachments.Count To 1 Step -1
        Set att = mi.Attachments.Item(idx)
        If skipInline And IsInlineAttachment(mi, att) Then
            ' leave inline in place
        Else
            Dim thisName As String: thisName = att.fileName
            Dim thisSize As Long: thisSize = att.Size  ' bytes
            If Len(savePath) > 0 Then
                Dim target As String
                target = UniquePath(savePath, SanitizeFileName(thisName))
                att.SaveAsFile target
                savedCount = savedCount + 1
            End If
            mi.Attachments.Remove idx
            removedCount = removedCount + 1
            removedInfo = removedInfo & "• " & thisName & " (" & FormatKB(thisSize) & ")" & vbCrLf
        End If
    Next idx

    If removedCount > 0 Then
        If addNote Then InsertRemovalNote mi, removedInfo
        mi.Save
    End If

    ProcessMailItem = True
    Exit Function

fail:
    ProcessMailItem = False
End Function

' Try to detect inline attachments: content-id referenced in HTML or hidden flag.
Private Function IsInlineAttachment(ByVal mi As MailItem, ByVal att As Attachment) As Boolean
    On Error Resume Next
    Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

    Dim cid As String
    Dim hidden As Variant
    cid = att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
    hidden = att.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN)

    ' Inline if HTML body references cid: or if attachment is flagged hidden
    If Len(cid) > 0 Then
        If InStr(1, LCase$(mi.HTMLBody), "cid:" & LCase$(cid), vbTextCompare) > 0 Then
            IsInlineAttachment = True
            Exit Function
        End If
    End If
    If VarType(hidden) = vbBoolean And hidden = True Then
        IsInlineAttachment = True
        Exit Function
    End If

    ' Treat embedded OLE/Item as inline-ish to avoid breaking forms/meeting items
    On Error GoTo 0
    On Error Resume Next
    If att.Type <> olByValue Then
        IsInlineAttachment = True
    End If
    On Error GoTo 0
End Function

' Insert a small banner at the top of the message indicating what was removed.
Private Sub InsertRemovalNote(ByVal mi As MailItem, ByVal removedList As String)
    On Error Resume Next
    If mi.BodyFormat = olFormatHTML Then
        Dim noteHtml As String
        noteHtml = "<div style=""border:1px solid #d0d0d0;padding:8px;margin:0 0 8px 0;font-family:Segoe UI,Arial,sans-serif;font-size:12px;background:#fffbe6"">" & _
                   "<b>Attachments removed by maintenance macro</b><br/>" & _
                   "<div style=""white-space:pre-wrap"">" & HtmlEncode(removedList) & "</div>" & _
                   "<div style=""color:#666"">Date: " & HtmlEncode(Format(Now, "yyyy-mm-dd HH:nn")) & "</div></div>"
        mi.HTMLBody = noteHtml & mi.HTMLBody
    Else
        Dim noteTxt As String
        noteTxt = "=== Attachments removed by maintenance macro ===" & vbCrLf & _
                  removedList & _
                  "Date: " & Format(Now, "yyyy-mm-dd HH:nn") & vbCrLf & vbCrLf
        mi.Body = noteTxt & mi.Body
    End If
End Sub


' ---------------------------
' UTILITIES
' ---------------------------

' Pick a folder using Office FileDialog, with Shell fallback if Office library missing.
Private Function PickFolderPath(ByVal prompt As String) As String
    On Error GoTo shellFallback
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = prompt
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolderPath = .SelectedItems(1)
        Else
            PickFolderPath = vbNullString
        End If
    End With
    Exit Function
shellFallback:
    Dim sh As Object, f As Object
    Set sh = CreateObject("Shell.Application")
    Set f = sh.BrowseForFolder(0, prompt, &H1)
    If Not f Is Nothing Then
        PickFolderPath = f.Self.path
    Else
        PickFolderPath = vbNullString
    End If
End Function

' Ensure filename is safe for Windows, preserve extension.
Private Function SanitizeFileName(ByVal nameIn As String) As String
    Dim base As String, ext As String, pos As Long
    pos = InStrRev(nameIn, ".")
    If pos > 0 Then
        base = Left$(nameIn, pos - 1)
        ext = Mid$(nameIn, pos)
    Else
        base = nameIn
        ext = ""
    End If
    base = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(base, _
           "\", "_"), "/", "_"), ":", "-"), "*", "_"), "?", "_"), """", "_"), "<", "("), ">", ")"), "|", "_"), vbTab, " ")
    SanitizeFileName = Trim$(base) & ext
End Function

' Generate a unique path if a file already exists (adds " (n)" suffix).
Private Function UniquePath(ByVal folder As String, ByVal fileName As String) As String
    Dim p As String, base As String, ext As String, n As Long, pos As Long
    p = folder
    If Right$(p, 1) <> "\" Then p = p & "\"
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        base = Left$(fileName, pos - 1)
        ext = Mid$(fileName, pos)
    Else
        base = fileName
        ext = ""
    End If
    Dim candidate As String
    candidate = p & base & ext
    n = 1
    Do While FileExists(candidate)
        candidate = p & base & " (" & n & ")" & ext
        n = n + 1
    Loop
    UniquePath = candidate
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path, vbNormal)) > 0)
    On Error GoTo 0
End Function

' Format bytes to KB with one decimal place.
Private Function FormatKB(ByVal bytes As Long) As String
    FormatKB = Format$(bytes / 1024#, "0.0") & " KB"
End Function

' Minimal HTML escaping for injected text.
Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEncode = s
End Function

' Simple screen flicker control.
Private Sub ScreenUpdating(ByVal enable As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = enable  ' available in recent Outlook builds
    On Error GoTo 0
End Sub
