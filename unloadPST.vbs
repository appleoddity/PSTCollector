' ============================================================================
' unloadPST.vbs - Removes PST files from Outlook before source deletion
' ============================================================================
'
' SYNOPSIS:
'   Detaches all PST data stores from a running Outlook instance.
'
' DESCRIPTION:
'   This script connects to a running Outlook application and removes all
'   PST (Personal Storage Table) files from the user's profile. This is
'   necessary BEFORE using PSTCollector's REMOVE mode to delete source
'   PST files, otherwise users will see errors in Outlook.
'
'   Preserves:
'   - SharePoint Lists (virtual PST stores)
'   - Internet Calendar Subscriptions
'
' USAGE:
'   Deploy as a login script or scheduled task in the USER context:
'   cscript \\server\PSTCollector\unloadPST.vbs
'
' REQUIREMENTS:
'   - Must run while Outlook is open
'   - Must run in the user's context (not SYSTEM)
'   - Safe to run multiple times
'
' AUTHOR: Appleoddity
' ============================================================================

On Error Resume Next

Dim objOutlook 'As Outlook.Application
Dim Stores     'As Outlook.Stores
Dim objFolder  'As Outlook.Folder
Dim i          'As Integer

' Connect to running Outlook instance
Set objOutlook = GetObject(, "Outlook.Application")

If Err.Number = 0 Then
    Set Stores = objOutlook.Session.Stores
    
    ' Iterate backwards through stores (removing changes indices)
    For i = Stores.Count To 0 Step -1
        ' ExchangeStoreType 3 = PST file
        If Stores(i).ExchangeStoreType = 3 Then
            ' Skip built-in virtual stores
            If (InStr(UCase(Stores(i).DisplayName), "SHAREPOINT LISTS") = 0) And _
               (InStr(UCase(Stores(i).DisplayName), "INTERNET CALENDAR SUBSCRIPTIONS") = 0) Then
                Set objFolder = Stores(i).GetRootFolder
                objOutlook.Session.RemoveStore objFolder
            End If
        End If
    Next
End If
     