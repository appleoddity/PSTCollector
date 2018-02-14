On Error Resume Next

Dim objOutlook 'As Outlook.Application
Dim Stores     'As Outlook.Stores
Dim objFolder  'As Outlook.Folder
Dim i          'As Integer

Set objOutlook = GetObject(, "Outlook.Application")
if Err.Number = 0 Then
    Set Stores = objOutlook.Session.Stores
 
    For i =  Stores.Count to 0 step -1
        If Stores(i).ExchangeStoreType = 3 Then
          If ((instr(ucase(Stores(i).DisplayName),"SHAREPOINT LISTS") = 0) and (instr(ucase(Stores(i).DisplayName),"INTERNET CALENDAR SUBSCRIPTIONS") = 0)) then
           Set objFolder = Stores(i).GetRootFolder
           objOutlook.Session.RemoveStore objFolder
          End if
        End if
    Next
End If
     