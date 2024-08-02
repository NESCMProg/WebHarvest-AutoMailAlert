Function GetLocalFile(wb As Workbook) As String
    ' Set default return
    GetLocalFile = wb.FullName
    
    Const HKEY_CURRENT_USER = &H80000001

    Dim strValue As String
    
    Dim objReg As Object: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Dim strRegPath As String: strRegPath = "Software\SyncEngines\Providers\OneDrive\"
    Dim arrSubKeys() As Variant
    objReg.EnumKey HKEY_CURRENT_USER, strRegPath, arrSubKeys
    
    Dim varKey As Variant
    For Each varKey In arrSubKeys
        ' check if this key has a value named "UrlNamespace", and save the value to strValue
        objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "UrlNamespace", strValue
    
        ' If the namespace is in FullName, then we know we have a URL and need to get the path on disk
        If InStr(wb.FullName, strValue) > 0 Then
            Dim strTemp As String
            Dim strCID As String
            Dim strMountpoint As String
            
            ' Get the mount point for OneDrive
            objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "MountPoint", strMountpoint
            
            ' Get the CID
            objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "CID", strCID
            
            ' Add a slash, if the CID returned something
            If strCID <> vbNullString Then
                strCID = "/" & strCID
            End If

            ' strip off the namespace and CID
            strTemp = Right(wb.FullName, Len(wb.FullName) - Len(strValue & strCID))
            
            ' replace all forward slashes with backslashes
            GetLocalFile = strMountpoint & Replace(strTemp, "/", "\")
            Exit Function
        End If
    Next
End Function