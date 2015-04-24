' Get Active Directory User Information
' Version 2003
' Created May 2003 by Ralph Montgomery - Firsthealth of the Carolinas (rmonty@myself.com)
' May be freely distributed to give back to the scripting community, please acknowledge
' the work where you can. I would appreciate it. Many items here were culled from MSDN, newsgroups
' the Windows 2000 Scripting Guide from MS and just many hours of work. If you recognize a routine
' that I have not acknowledged, please let me know and I will fix it for ya.
' Revision history:
'    Initial rollout after debugging and documentation 06-11-2003
'    09/21/03 Added HTML display alternative
'         Added display of last logged in workstation from SMS
'    11/11/03 fixed password expiry info so display correctly
'
' Caveats: The Terminal Service information can only be pulled by a WinXp workstation with the
'    Active Directory Users and Computers MMC console from a Server 2003 CD. Sorry, MS wants it
'    that way I guess. Otherwise it will always be no.
'
' Usage: ADUser <CR>
 
'Must do: Either add your SMS site server and SMS site Name under the Const or
'     remark out the calling line: FindMachineByUser(strGetUserName)
 
' Constants
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_PROPERTY_DELETE = 4
Const ADS_UF_ACCOUNTDISABLE = 2
Const ADS_UF_PASSWD_NOTREQD = &h00020
Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &h0080
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
Const ADS_UF_PASSWORD_EXPIRED = &h80000
Const ADS_UF_PASSWD_CANT_CHANGE = &h0040
Const CHANGE_PASSWORD_GUID = "{ab721a53-1e2f-11d0-9819-00aa0040529b}"
Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const SEC_IN_MIN = 60
Const SEC_IN_DAY = 86400
Const MIN_IN_DAY = 1440
Const ADS_SCOPE_SUBTREE = 2
 
'Must do: Either add your SMS site server and SMS site Name under this Const or
'     remark out the calling line: FindMachineByUser(strGetUserName)
Const cSMSmachine = "sms-ss-mrh" ' name of system where SMS lives
Const cSMSsite = "FHC" ' name of SMS site
 
'*********************Initialize the variable farm in one spot*******************************************
Public strGetUserName
Dim objUserName, objUserDomain, objGroup, objUser, strGroupList, WshShell, strMessage, strTitle, dtStart
Dim objDomain, strDomain, strUserName, strOS, strVer, strSortedGroups, arrMemberOf, strUserList, strCheckName
Dim strMsgNoUser, strUserMail, strExchange, sQ, strNoDomain, blnIsActive, strCN, strOU, strRootDSE
 
Dim objChangePwdTrue, objChangePwd, objUserProfile, objNet, strIsAccountLocked, strMailNickname, strRetry
Dim objPwdExpiresTrue, objFlags, oPwdExpire , dtmPwdLastChanged, strUserName2, strValueList, major, minor, ver
Dim objAcctDisabled, intPwdExpired, objPwdExpiredTrue, strTSProfile, strDisplayDelegates
 
Dim strGivenName, strInitials, strSn, strDisplayName, strPhysDelOfficeName, strTelephoneNumber, strGetUserNam
Dim strMail, strWwwHomePage, intUAC, intBadPwd,    strNetworkAddress, strAllowDialin, dtmLastLogin, strLogonName
Dim strWhenCreated,    strWhenChanged,    strPwdExpires, strValue, strUserMustChgPwd, strPwdNeverExpires, strPwdLastChanged
Dim strPwdExpired, strPwdAge, strAccountDisabled, strDisplayDescription, strDisplayOtherTelephone, strDisplayUrl
Dim strOtherTelephone, strUrl, strPwdCanChange, strPwdMinLength, strDisplayDepartment, strAccountExpires
 
Dim strTSHomeDir, strTSHomeDrive, strTSProfilePath, strTSConnectPrinters, strTSConnectDrives, strTSDefaultToMainPrinter
Dim strTSInitialProgram, strTSWorkingDir, strTSEnableRemoteControl, strTSBrokenConnAction, strTSMaxConnectTime
Dim strTSMaxDisconnectionTime, strTSMaxIdleTime, strTSReconnectionAction, strTSAllowLogon
 
Dim intMaxPwdAge, intMaxPwdAgeSeconds, intMinPwdAgeSeconds, intLockOutObservationWindowSeconds, blnChangePwdEnabled
Dim intLockoutDurationSeconds, intUserFlags,intMinPwdLength, intPwdHistoryLength, intPwdProperties, intLockoutThreshold
Dim    intMaxPwdAgeDays, intMinPwdAgeDays, intLockOutObservationWindowMinutes, intLockoutDurationMinutes
 
Dim strProfilePath, strScriptPath, strHomeDirectory, strHomeDrive, blnMsNPAllowDialin, strVPNAllow, strDLList, strSortedDLList
Dim ldapconnectstring, Ouser, strSearch, strDN, dtmNextFailedLogin, dtmLastFailedLogin, strPwdRequired
Dim arrDC(), intSize, strLastLoggedInWorkstation, objDocument, strPwdBGColor, strAcctBGColor, strLoginsBGColor, strMsgDisplay
Dim strPwdExpBGColor, strDelegateCount
Dim strMostRecentIP
 
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objNet = WScript.CreateObject("WScript.Network")' create network object for vars
Set objRootDSE = GetObject("LDAP://rootDSE")' bind to the rootDSE for portability
 
strADsConfPath = "LDAP://" & objRootDSE.Get("configurationNamingContext")' bind to configuration to get Domain Controllers later
strRootDSE = objRootDSE.Get("defaultNamingContext")' bind to the defaultContext for portability
strVer = "Ver 2003"' vanity
sQ = Chr(34)
strDomain = UCase(objNet.UserDomain)' pull user domain from environment variable
strUserName = UCase(objNet.UserName)' pull user name from environment variable
strOS = WshShell.ExpandEnvironmentStrings("%OS%")' pull OS from environment variable to use for other subs...
intSize = 0
strDelegateCount = 0
 
'SysTest() ' sub routine to check for Script Version/ADSI installed
GetUserName()' sub routine to get input for userID (sAMAccountName)
 
' this section added by John Ciccantelli
While strDN=""
  CheckForUser()' sub routine to check for user Existance & bind to if found
  If strDN = "" Then
    ReCheckUser()
  End If
Wend
' end of section added by John C.
GetUserAccount(strDN)
GetLastLogon()' sub routine to get absolute last login date from all Domain Controllers dynamically
' You must remark this next line out if NOT using SMS!!!!!!
' FindMachineByUser(strGetUserName) ' sub routine to query SMS for the last workstation logged into - remark out if not using SMS!
 
'strMsgDisplay = "To Display/Print Account information in" & vbCrLf & " Internet Explorer, press Yes, else press No"
'rtn = MsgBox(strMsgDisplay,vbYesNo,"Use HTML display output?")
'If rtn = 7 Then
'DisplayUser()' sub routine to Display gathered user Information in a popup box
'Elseif rtn = 6 Then
DisplayUserIE()' sub routine to Display gathered user Information in an Internet Explorer Window
'Else
'WScript.quit
'End if
 
'********************* Initial and only dialog box necessary *****************************
'********************* Looks for the sAMAccountName to bind to *****************************
Sub GetUserName()
  strMessage = "Enter the User Login ID (sAMAccountName) to search:" & vbCrLf & vbCrLf
  ' "Default is: " & strUserName & vbCrLf & vbCrLf
  strMessage = strMessage & "You may also search for a user by first or last name. "
  strMessage = strMessage & "(Searching will take a little bit longer)" & vbCrLf & vbCrLf & "or click Cancel to quit"
  strTitle = "Enter User Login ID"
  
  'get resource domain name, domain default via input box
  strGetUserName= UCase(InputBox(strMessage, strTitle, strUserName))
  
  ' Evaluate the user input.
  If strGetUserName = "" Then
    Cancelled()
  ElseIf Len(strGetUserName) < 1 Then
    strMessage = "Input name less than 1 character! Please Input at least 1!"
    strGetUserName= UCase(InputBox(strMessage, strTitle, strUserName))
  Else
    strGetUserName = strGetUserName
  End If
  
End Sub 'GetUserName
 
'********************* 'Attempt to bind to the sAMAccount Name provided search if not***************************
Sub CheckForUser()
  
  WScript.Echo "Searching for user " & strGetUserName & "..."
  
  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Provider = ("ADsDSOObject")
  objConnection.Open
  
  Set objCommand = CreateObject("ADODB.Command")
  
  objCommand.ActiveConnection = objConnection
  
  objCommand.CommandText = _
  "<LDAP://" & strRootDSE & ">;(&(objectCategory=User)" & _
  "(samAccountName=" & strGetUserName & "));distinguishedName,sAMAccountName,name;subtree"
  
  Set objRecordSet = objCommand.Execute
  
  If objRecordset.RecordCount = 0 Then
    dtStart = TimeValue(Now())
    strMessage = "Login ID: " & strGetUserName & " not found: " & vbCrLf & "This may take a few seconds. . ."
    WshShell.Popup strMessage,2,"Searching . . ."
    strMessage = ""
    Set objectRecordSet = Nothing
    objConnection.close
    Set objConnection = Nothing
  Else
    strDN = objRecordset.Fields("distinguishedName")
    Set objectRecordSet = Nothing
    objConnection.close
    Set objConnection = Nothing
  End If
  
End Sub ' CheckForUser
 
Sub Check4User()
  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Provider = ("ADsDSOObject")
  objConnection.Open
  
  Set objCommand = CreateObject("ADODB.Command")
  
  objCommand.ActiveConnection = objConnection
  
  objCommand.CommandText = _
  "<LDAP://" & strRootDSE & ">;(&(anr=" & strGetUserName & ")(|(objectCategory=organizationalPerson)(objectCategory=group)));ADsPath,name,distinguishedName,displayName,objectCategory;subtree"
  
  objCommand.Properties("Page Size") = 64
  objCommand.Properties("Timeout") = 30 'seconds
  
  Set objRecordSet = objCommand.Execute
  
  If objRecordset.RecordCount <> 1 Then
    dtStart = TimeValue(Now())
    strMessage = "Name not found: " & strGetUserName & vbCrLf & "This may take a few seconds. . ."
    WshShell.Popup strMessage,2,"Searching . . ."
    strMessage = ""
    Set objectRecordSet = Nothing
    objConnection.close
    Set objConnection = Nothing
  Else
    strDN = objRecordset.Fields("distinguishedName")
    Set objectRecordSet = Nothing
    objConnection.close
    Set objConnection = Nothing
  End If
  
End Sub ' Check4User
 
'********************* Recheck for user - uses Display name as the search key *****************************
Sub ReCheckUser()
  
  ldapconnectstring = "<LDAP://" & strRootDSE & ">"
  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open
  
  'strSearch = ldapconnectstring & ";(&(objectCategory=User)(CN=" & strGetUserName & "*));adspath;subtree"
  strSearch = ldapconnectstring & ";(&(anr=" & strGetUserName & ")(|(objectCategory=organizationalPerson)(objectCategory=group)));ADsPath,name,distinguishedName,displayName,objectCategory;subtree"
  Set objRecordSet = objConnection.Execute(strSearch)
  
  Do While Not objRecordset.EOF
    Set oUser = GetObject(objRecordSet("adspath"))
    strUserList = (strUserList & " " & oUser.givenName & " " & ouser.SN) & " - " & Mid(Replace(oUser.Name, "\,",","), 4) & vbCrLf
    If Err < 0 Then
      MsgBox "Error Occurred"
    End If
    objRecordSet.MoveNext
  Loop
  
  strMsgNoUser = "Your search found the following User Login IDs: " & vbCrLf & vbCrLf & strUserList & vbCrLf & _
  "Search completed in " & Second(TimeValue(Now()) - dtStart) & " second(s) or less." & vbCrLf & vbCrLf & _
  "Enter the User Login ID below, or cancel to exit"
  
  strRetry = InputBox(strMsgNoUser,"Search Reults . . .", strGetUserName)
  strUserList = ""
  strMsgNoUser = ""
  If strRetry = "" Then
    Set objectRecordSet = Nothing
    objConnection.close
    Set objConnection = Nothing
    Cancelled()
  Else
    Set objectRecordSet = Nothing
    objConnection.close
    Set objConnection = Nothing
    strGetUserName = strRetry
  End If
  strGetUserName = strRetry
End Sub ' ReCheckUser
 
'********************* 'Get Selected User Account Information *****************************
Sub GetUserAccount(strDN)
  On Error Resume Next
  If InStr(1,strDN,"/") Then strDN=Replace(strDN,"/","\/")
  Set objDomainNT = GetObject("WinNT://" & strDomain & "")    ' Use NT Provider for Domain Policy items
  Set objUser = GetObject("LDAP://" & strDN & "")                ' LDAP for User Info
  Set objAdS = GetObject("LDAP://" & strRootDSE & "")            ' LDAP for AD domain items
  
  With objDomainNT
    intMaxPwdAge =                             .Get("MaxPasswordAge")    'get NT value for MaxPasswordAge
    intMaxPwdAge =                             (intMaxPwdAge/SEC_IN_DAY) ' maximum password age in days
    intMaxPwdAgeSeconds =                     .Get("MaxPasswordAge")
    intMinPwdAgeSeconds =                     .Get("MinPasswordAge")
    intLockOutObservationWindowSeconds =     .Get("LockoutObservationInterval")
    intLockoutDurationSeconds =             .Get("AutoUnlockInterval")
  End With 'objDomainNT
  
  
  
  With objAdS
    intMinPwdLength =                         .Get("minPwdLength")
    intPwdHistoryLength =                     .Get("pwdHistoryLength")
    intPwdProperties =                         .Get("pwdProperties")
    intLockoutThreshold =                     .Get("lockoutThreshold")
    
    intMaxPwdAgeDays =                         ((intMaxPwdAgeSeconds/SEC_IN_MIN)/MIN_IN_DAY) & " days"
    intMinPwdAgeDays =                         ((intMinPwdAgeSeconds/SEC_IN_MIN)/MIN_IN_DAY) & " days"
    intLockOutObservationWindowMinutes =     (intLockOutObservationWindowSeconds/SEC_IN_MIN) & " minutes"
    
    If intLockoutDurationSeconds <> -1 Then
      intLockoutDurationMinutes =         (intLockOutDurationSeconds/SEC_IN_MIN) & " minutes"
    Else
      intLockoutDurationMinutes =         "Administrator must manually unlock locked accounts"
    End If
  End With ' objAdS
  
  With objUser
    '.GetInfo
    strGivenName =                 .Get("givenName")
    'MsgBox(strDN & VbCrLf & strGivenName)
    strInitials =                 .Get("initials")
    strSn =                     .Get("sn")
    strDisplayName =             .Get("displayName")
    strPhysDelOfficeName =         .Get("physicalDeliveryOfficeName")
    strTelephoneNumber =         .Get("telephoneNumber")
    strMail =                     .Get("mail")
    strWwwHomePage =             .Get("wWWHomePage")
    strAccountExpires =         .Get("accountExpires")
    strLogonName =                 .Get("sAMAccountName")
    strWhenCreated =             .Get("whenCreated")
    strWhenCreated = DateValue(strWhenCreated) & " at " & TimeValue(strWhenCreated - GREENWHICH_MEAN_TIME)
    strWhenChanged =             .Get("whenChanged")
    strWhenChanged = DateValue(strWhenChanged) & " at " & TimeValue(strWhenChanged - GREENWHICH_MEAN_TIME)
    strHomeDrive =                 .Get("homeDrive") & "\"
    strUserMail =                 .Get("mail")
    
    Set dtmGetLockout =         .Get("lockoutTime")
    
    If dtmGetLockout.HighPart = 0 And dtmGetLockout.LowPart = 0 Then
      strIsAccountLocked = "No"
      strAcctBGColor = "#00CC00"
    Else
      strIsAccountLocked = "Yes"
      strAcctBGColor = "#FF0000"
    End If
    
    strMailNickname =             .Get("mailNickname")
    If strMailNickname = "" Then
      strExchange = "No"
    Else
      strExchange = "Yes"
    End If
    
    strScriptPath =             .Get("scriptPath")
    If strScriptPath = "" Then
      strScriptPath = "No logon script defined"
    Else
      strScriptPath = strScriptPath
    End If
    
    strProfilePath =             .Get("profilePath")
    If strProfilePath = "" Then
      strProfilePath = "No profile path specified"
    Else
      strProfilePath = strProfilePath
    End If
    
    strHomeDirectory =             .Get("homeDirectory")
    If strHomeDirectory = "" Then
      strHomeDirectory = "No Home Directory specified"
    Else
      strHomeDirectory = strHomeDirectory
    End If
    
    blnMsNPAllowDialin =         .Get("msNPAllowDialin")
    If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
      strVPNAllow = "Control access through Remote Access Policy"
      Err.Clear
    Else
      If blnMsNPAllowDialin = True Then
        strVPNAllow = "Yes"
      Else
        strVPNAllow = "No"
      End If
    End If
    
    intUAC =                     .Get("userAccountControl")
    If intUAC And ADS_UF_DONT_EXPIRE_PASSWD Then
      strPwdExpires = "Password does not expire"
      strPwdNeverExpires = "Yes"
      strPwdExpBGColor = "#FF0000"'set the HTML view to Red
    Else
      dtmPwdLastChanged =         .PasswordLastChanged
      strPwdNeverExpires = "No"            
      strPwdExpires = DateValue(dtmPwdLastChanged + intMaxPwdAge) & " at " & TimeValue(dtmPwdLastChanged)
      strPwdExpBGColor = "#00CC00"'set the HTML view to Green
    End If
    
    dtmPwdLastChanged =         .PasswordLastChanged
    If dtmPwdLastChanged = "" Then
      strPwdAge = "" & vbTab
      strPwdLastChanged = "No record available"
      strPwdExpired = "Unknown"
      strPwdBGColor = "#FF0000" ' set the HTML view to Red
    Else
      strPwdLastChanged = DateValue(dtmPwdLastChanged) & " at " & TimeValue(dtmPwdLastChanged)
      strPwdAge = Int(Now - dtmPwdLastChanged) & " days"
      If intMaxPwdAgeDays >= strPwdAge Then
        strPwdExpired = "No"
        strPwdBGColor = "#00CC00"'set the HTML view to Green
      Else
        strPwdExpired = "Yes"
        strPwdBGColor = "#FF0000" ' set the HTML view to Red
        strPwdExpBGColor = "#FF0000"'set the HTML view to Red
      End If
      
    End If
    
    Set objSD =                 .Get("nTSecurityDescriptor")
    Set objDACL =             objSD.DiscretionaryAcl
    
    For Each Ace In objDACL
      If ((Ace.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT) And (LCase(Ace.ObjectType) = CHANGE_PASSWORD_GUID)) Then
        blnChangePwdEnabled = True
      End If
    Next
    
    If blnChangePwdEnabled Then
      strPwdCanChange = "No"
    Else
      strPwdCanChange = "Yes"
    End If
    
    'Terminal Services Info
    strTSHomeDir =                 .TerminalServicesHomeDirectory
    strTSHomeDrive =             .TerminalServicesHomeDrive
    strTSInitialProgram =         .TerminalServicesInitialProgram
    strTSWorkingDir =             .TerminalServicesWorkDirectory
    strTSBrokenConnAction =     .BrokenConnectionAction
    strTSMaxConnectTime =        .MaxConnectionTime
    strTSMaxDisconnectionTime = .MaxDisconnectionTime
    strTSMaxIdleTime =             .MaxIdleTime
    strTSReconnectionAction =     .ReconnectionAction
    strTSProfilePath =             .TerminalServicesProfilePath
    If strTSProfilePath = "" Then
      strTSProfilePath = "No profile path specified"
    Else
      strTSProfilePath = strTSProfilePath
    End If
    
    strTSAllowLogon =             .allowLogon
    If strTSAllowLogon = 1 Then
      strTSAllowLogon = "Yes"
    Else
      strTSAllowLogon = "No"
    End If
    
    strTSConnectPrinters =         .ConnectClientPrintersAtLogon
    If strTSConnectPrinters = 1 Then
      strTSConnectPrinters = "Yes"
    Else
      strTSConnectPrinters = "No"
    End If
    
    strTSConnectDrives =         .ConnectClientDrivesAtLogon
    If strTSConnectDrives = 1 Then
      strTSConnectDrives = "Yes"
    Else
      strTSConnectDrives = "No"
    End If
    
    strTSDefaultToMainPrinter = .DefaultToMainPrinter
    If strTSDefaultToMainPrinter = 1 Then
      strTSDefaultToMainPrinter = "Yes"
    Else
      strTSDefaultToMainPrinter = "No"
    End If
    
    strTSEnableRemoteControl =     .EnableRemoteControl
    If strTSEnableRemoteControl = 1 Then
      strTSEnableRemoteControl = "Yes"
    Else
      strTSEnableRemoteControl = "No"
    End If
    
    strDescription =             .GetEx("description")
    strDepartment =             .GetEx("department")
    strOtherTelephone =         .GetEx("otherTelephone")
    strUrl =                     .GetEx("url")
    arrMemberOf =                 .GetEx("memberOf")
    strDelegates =                .GetEx("publicDelegates")
    
    If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
      strGroupList = "The memberOf attribute is not set."
    Else
 
      ' *** Ammended Group Search ***
 
      ' Dictionary for loop prevention if we bump into circular nesting
      Dim objTemp : Set objTemp = CreateObject("Scripting.Dictionary")
 
      strGroupList = GetAllGroups(arrMemberOf, objTemp, "")
 
      Set objTemp = Nothing
 
      ' *** End ***
 
      ' Convert strgrouplist to Array
      arrGroupList = Split(strGroupList,",")
      'Sort the durn thing
      Quicksort arrGroupList, LBound(arrGroupList), UBound(arrGroupList)
      ' Now concatenate arrGroupList into a variable for display
      strSortedGroups = Join(arrGroupList, ", ")
      strSortedGroups = Mid(strSortedGroups, 4) ' cause the sort function is funky...
    End If
    
    For Each strValue In strDepartment
      strDisplayDepartment = strDisplayDepartment & strValue
    Next ' strDepartment Value
    
    For Each strValue In strDescription
      strDisplayDescription = strDisplayDescription & strValue
    Next ' strDecription Value
    
    For Each strValue In strOtherTelephone
      strDisplayOtherTelephone = strDisplayOtherTelephone & strValue
    Next ' strOtherTelephone Value
    
    For Each strValue In strUrl
      strDisplayUrl = strDisplayUrl & strValue
    Next ' strUrl value
    
    For Each strValue In strDelegates
      strDelegateCount = strDelegateCount + 1
      strValue = Mid(strValue,4)
      intLeft = InStr(strValue,",")
      strValue = Left(strValue, intLeft) & " "
      strValue = Replace(strValueList,"\,","")
      strValueList = strValueList & strValue
    Next ' strDelegate Value
    strDisplayDelegates = strValueList
    
    ' create dictionary for user account information
    Set objHash = CreateObject("Scripting.Dictionary")
    objHash.Add "ADS_UF_PASSWD_NOTREQD", &h00020
    objHash.Add "ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED", &h0080
    
    If intUAC And ADS_UF_ACCOUNTDISABLE Then
      strAccountDisabled = "Yes"
      strAcctBGColor = "#FF0000" ' set the HTML view to Red
    Else
      strAccountDisabled = "No"
      strAcctBGColor = "#00CC00" ' set the HTML view to Green
    End If
    
    For Each Key In objHash.Keys
      
      If objHash(Key) = ADS_UF_PASSWD_NOTREQD And intUAC Then
        strPwdRequired = "Yes"
      Else
        strPwdRequired = "No"
      End If
      
    Next 'objHash.keys
    
  End With ' objUser
  
End Sub 'GetUserAccount
 
Function GetAllGroups(arrGroups, objTemp, strGroupList)
 
  Dim strGroupDN
  For Each strGroupDN in arrGroups
 
    ' Make sure we haven't looked at this group before.
    If Not objTemp.Exists(strGroupDN) Then
 
      ' Connect to the group
      Dim objGroup : Set objGroup = GetObject("LDAP://" & strGroupDN)
      strGroupList = strGroupList & objGroup.get("name") & ","
 
      On Error Resume Next
      strGroupList = GetAllGroups(objGroup.GetEx("memberOf"), objTemp, strGroupList)
      On Error Goto 0
      Set objGroup = Nothing
    End If
  Next
 
  GetAllGroups = strGroupList
End Function
 
' ******************Sorts the items in the array (between the two values you pass in)*********************
Sub Quicksort(strValues(), ByVal min, ByVal max)
  
  Dim strMediumValue, high, low, i
  
  'If the list has only 1 item, it's sorted.
  If min >= max Then Exit Sub
  
  ' Pick a dividing item randomly.
  i = min + Int(Rnd(max - min + 1))
  strMediumValue = strValues(i)
  
  ' Swap the dividing item to the front of the list.
  strValues(i) = strValues(min)
  
  ' Separate the list into sublists.
  low = min
  high = max
  Do
    ' Look down from high for a value < strMediumValue.
    Do While strValues(high) >= strMediumValue
      high = high - 1
      If high <= low Then Exit Do
    Loop
    
    If high <= low Then
      'The list is separated.
      strValues(low) = strMediumValue
      Exit Do
    End If
    
    'Swap the low and high strValues.
    strValues(low) = strValues(high)
    
    'Look up from low for a value >= strMediumValue.
    low = low + 1
    Do While strValues(low) < strMediumValue
      low = low + 1
      If low >= high Then Exit Do
    Loop
    
    If low >= high Then
      'The list is separated.
      low = high
      strValues(high) = strMediumValue
      Exit Do
    End If
    
    'Swap the low and high strValues.
    strValues(high) = strValues(low)
  Loop 'Loop until the list is separated.
  
  'Recursively sort the sublists.
  Quicksort strValues, min, low - 1
  Quicksort strValues, low + 1, max
  
End Sub 'Quicksort
 
'********************* If user selects cancel at any dialog box *****************************
Sub Cancelled()
  strMessage = "Cancelled by user: " & strUserName
  strTitle = "Operation Cancelled"
  MsgBox strMessage,vbOKOnly,strTitle
  WScript.quit
End Sub 'Cancelled
 
'********************* 'Test for minimum system software needed to run *****************************
Sub SysTest()    
  On Error Resume Next
  ' Alan Kaplan for VISN 6 - many thanks for the SysTest routines
  ' akaplan@msdinc.com www.msdinc.com
  ' WSH version tested
  Major = (ScriptEngineMinorVersion())
  Minor = (ScriptEngineMinorVersion())/10
  Ver = major + minor
  'Need version 5.5
  If Err.number Or ver = 5.6 Then
    strMessage = "You must load Version 5.5 (or later) of Windows Script Host" & vbCrLf &_
    vbCrLf & "Located at: \\filer-mrh\software\wmi\scr56en.exe" & vbCrLf
    WScript.Quit
  End If
  
  'Test for ADSI
  Err.clear
  key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{E92B03AB-B707-11d2-9CBD-0000F87A369E}\version"
  key2 = WshShell.RegRead (key)
  If Err <> 0 Then
    If strOS = "Windows_NT" Then
      strMessage = "ADSI 5.2 must be installed on local workstation to continue" & vbCrLf &_
      vbCrLf & "Located at: \\filer-mrh\software\wmi\adsi5.2.exe" & vbCrLf
      
      WshShell.Popup strMessage,0,"Workstation Setup Error",vbCritical
      WScript.Quit
    Else ' Must be Windows 9x
      strMessage = "You appear to be running Windows 9x. If this is true, then" & vbCrLf
      strMessage = strMessage & "ADSI 5.2 AND WMI must be installed on local workstation to continue" & vbCrLf &_
      vbCrLf & "Located at: \\filer-mrh\software\wmi\adsi5.2.exe and dsclient.exe" & vbCrLf
      WshShell.Popup strMessage,0,"Workstation Setup Error",vbCritical
      WScript.Quit
    End If
  End If
  
End Sub 'SysTest
 
'********************* Get the absolute last login/failed login date from Domain Controllers*********************************
Sub GetLastLogon()
  On Error Resume Next
  Set objConnection = CreateObject("ADODB.Connection")
  Set objCommand = CreateObject("ADODB.Command")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  Set objCommand.ActiveConnection = objConnection
  objCommand.CommandText = _
  "SELECT distinguishedName FROM " _
  & "'" & "" & strADsConfPath & "" & "'" _
  & "WHERE objectClass='nTDSDSA'"
  objCommand.Properties("Page Size") = 1000
  objCommand.Properties("Timeout") = 30
  objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
  objCommand.Properties("Cache Results") = False
  Set objRecordSet = objCommand.Execute
  objRecordSet.MoveFirst
  
  Do Until objRecordSet.EOF
    strDCLeft = Mid(objRecordSet.Fields("distinguishedName").Value,21)
    strDCRight = InStr(strDCLeft,",")
    strDC = Left(strDCLeft,(strDCRight -1))
    ReDim Preserve arrDC(intSize)
    arrDC(intSize) = strDC
    intSize = intSize + 1
    objRecordSet.MoveNext
  Loop ' objRecordSet
  
  For Each strDC In arrDC
    Set objUser = GetObject("LDAP://" & strDC & "/" & strDN & "")
    
    With objUser
      dtmLastFailedLogin =         .LastFailedLogin
      If dtmNextFailedLogin > dtmLastFailedLogin Then
        dtmLastFailedLogin = dtmNextFailedLogin
      Else
        dtmLastFailedLogin = dtmLastFailedLogin
      End If
      
      intBadPwd =                 .Get("badPwdCount")
      If intNextBadPwd > intBadPwd Then
        intBadPwd = intNextBadPwd
      Else
        intBadPwd = intBadPwd
      End If
      
      If intBadPwd = 0 Then
        strLoginsBGColor = "#00CC00" ' set the HTML view to Green
      Else
        strLoginsBGColor = "#FF0000" ' set the HTML view to Red
      End If
      
      dtmLastLogin =                .LastLogin
      If dtmNextLogin > dtmLastLogin Then
        dtmLastLogin = dtmNextLogin
      Else
        dtmLastLogin = dtmLastLogin
      End If
    End With ' objUser
    dtmNextLogin = dtmLastLogin
    dtmNextFailedLogin = dtmLastFailedLogin
    intNextBadPwd = intBadPwd
  Next ' arrDC
  
  Set objectRecordSet = Nothing
  objConnection.close
  Set objConnection = Nothing
  
End Sub 'GetDomainControllers
 
 
 
'********************* Get the last logged in workstation name from SMS*********************************
 
Function FindMachineByUser(strGetUserName)
  On Error Resume Next
  
  Dim WinMgmt, SystemSet, strTime
  Dim objEnumerator, instance, strQuery
  Dim intMostRecentTime
  Dim i
  
  i=0
  FindMachineByUser = ""
  intMostRecentTime=""
  WinMgmt = "winmgmts:{impersonationLevel=impersonate}" & "!//" & cSMSmachine & "\root\sms\site_" & cSMSsite
  
  If Err <> 0 Then
    strLastLoggedInWorkstation = "Information Not available"
  Else
    
    Set SystemSet = GetObject(winmgmt)
    
    strQuery = _
    "SELECT Name, IPAddresses, AgentTime " & _
    "from sms_r_system where LastLogonUserName = '" & strGetUserName & "'"
    
    Set objEnumerator = SystemSet.ExecQuery(strQuery)
    For Each instance In objEnumerator
      strTime = instance.AgentTime(0)
      If strTime > intMostRecentTime Then
        intMostRecentTime=strTime
        If instance.IPAddresses(0) = "0.0.0.0" Then
          strMostRecentIP = instance.IPAddresses(1)
        Else
          strMostRecentIP = instance.IPAddresses(0)
        End If
        FindMachineByUser = instance.Name(0)
      End If
    Next
    If FindMachineByUser = "" Then
      strLastLoggedInWorkstation = "Information Not available"
    Else
      strLastLoggedInWorkstation = FindMachineByUser
    End If
  End If
End Function ' Get last logged in workstation from SMS
 
'********************************Create Internet Explorer Window to display the text in***************************
 
 
'*******************Kills the script if the IE Window is closed********************************
Sub IE_Quit()
  WScript.Quit
End Sub 'IE_Quit
 
 
'********************* Display USer Information in a Popup box*********************************
Sub DisplayUser()
  
  ' Set strMessage box variables to null
  strMessage =""
  ' Get rid of that annoying escape character for display purposes
  strDN = Replace(strDN,"\,",",")
  
  'popup user information: each line broken up for better reading
  strMessage = strMessage & "Logon Name: " & strLogonName & vbTab & "Display Name: " & strDisplayName & vbCrLf & _
  "Description: " & strDisplayDescription & vbCrLf & _
  "Department: " & strDisplayDepartment & vbTab & vbTab & "Telephone: " & strTelephoneNumber & vbCrLf & vbCrLf & _
  "Account Created: " & strWhenCreated & " GMT (-5 hours)" & vbTab & "Account changed: " & strWhenChanged & " GMT (-5 hours)" & vbCrLf & _
  "Distinguished Name: " & strDN & vbCrLf & _
  "Last logged in Workstation: " & strLastLoggedInWorkstation & vbTab & vbTab & "Last IPAddress: " & strMostRecentIP & vbCrLf & vbCrLf
  
  strMessage = strMessage & "Account Locked Out: " & strIsAccountLocked & vbTab & _
  "Account Disabled: " & strAccountDisabled & vbCrLf & _
  "Bad Logins: " & intBadPwd & vbTab & vbTab & "Attempts Left: " & (intLockoutThreshold - intBadPwd) & vbTab & "Max Attempts: " & intLockoutThreshold & vbCrLf & _
  "Last failed login: " & vbTab & vbTab & dtmLastFailedLogin & vbCrLf & _
  "Last Successful login: " & vbTab & dtmLastLogin & vbCrLf & vbCrLf
  
  strMessage = strMessage & "Password Last Changed: " & vbTab & strPwdLastChanged & vbCrLf & _
  "Password Expires: " & vbTab & vbTab& strPwdExpires & vbTab & vbCrLf & _
  "Password Age: " & strPwdAge & vbTab & "Password Expired: " & strPwdExpired & vbCrLf & vbCrLf
  
  strMessage = strMessage & "User can change Pwd: " & strPwdCanChange & vbTab & "Password Never Expires: " & _
  strPwdNeverExpires & vbCrLf & "Password Min Length: " & intMinPwdLength & vbTab & _
  "Passwords Kept In History: " & intPwdHistoryLength & " password(s)" & vbCrLf & _
  "Lockout Time: " & intLockoutDurationMinutes & vbTab & "AutoUnlock: " & intLockOutObservationWindowMinutes & vbCrLf & vbCrLf
  
  strMessage = strMessage & "Home Directory: " & strHomeDirectory & vbTab & vbTab & "Home Drive: " & strHomeDrive & vbCrLf & _
  "Roaming Profile Path: " & strProfilePath & vbCrLf & "Logon Script: " & strScriptPath & vbCrLf & vbCrLf
  
  strMessage = strMessage & "TS Profile Path: " & strTSProfilePath & vbCrLf & _
  "Allow TS Logon: " & strTSAllowLogon & vbTab & "Enable Remote Control: " & strTSEnableRemoteControl & vbCrLf & _
  "Connect Client Drives: " & strTSConnectDrives & vbTab & "Auto Create Printers: " & strTSConnectPrinters & vbCrLf & vbCrLf
  
  If strExchange = "Yes" Then    
    strMessage = strMessage & "Exchange Account: " & strExchange & vbTab & "External email address: " & strUserMail & vbCrLf & _
    "Exchange Alias: " & strMailNickname & vbTab & "Assigned Delegates: " & strDisplayDelegates & vbCrLf & _    
    "Allow VPN Access: " & strVPNAllow & vbCrLf & vbCrLf
    
  Else
    strMessage = strMessage & "Exchange Account: " & strExchange & vbTab & "Allow VPN Access: " & strVPNAllow & vbCrLf & vbCrLf
    
  End If
  
  strMessage = strMessage & "Group Membership: (Includes Distribution List Membership)" & vbCrLf & vbCrLf & _
  strSortedGroups
  
  ' Display User Information!
  strTitleMessage = " User Info for: " & strDisplayName & " in " & strDomain & " " & strVer
  WshShell.Popup strMessage,0,strTitleMessage
  
End Sub ' Display User
 
'********************* Display USer Information in a IE Window *********************************
Sub DisplayUserIE()
  Set objExplorer = WScript.CreateObject("InternetExplorer.Application", "IE_")
  objExplorer.Navigate "about:" & strDisplayName
  objExplorer.ToolBar = 0
  objExplorer.StatusBar = 0
  objExplorer.Width = 800
  objExplorer.Height = 600
  objExplorer.Left = 0
  objExplorer.Top = 0
  objExplorer.Visible = 1
  variable = "0"
  
  Set objDocument = objExplorer.Document
  
  objDocument.Open
  
  ' Set strPercent variable
  strPercent = "%"
  sQ = Chr(34)
  strBGColor = "#00CC00"
  ' Get rid of that annoying escape character for display purposes
  strDN = Replace(strDN,"\,",",")
  ' pre-defined HTML code- only have to change it ONCE to fix all
  
  sHTMLC1 = "<tr><td width=15" & strPercent & " bgcolor=" & sQ & strBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC1a = "<tr><td width=65" & strPercent & " style=" & sQ & "height:10px;" & sQ & " colspan=4>"
  sHTMLC1b = "<tr><td width=65" & strPercent & " bgcolor=" & sQ & strBGColor & sQ & " align=" & sQ & "right" & sQ & " colspan=4><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC1c = "<tr><td width=15" & strPercent & " bgcolor=" & sQ & strPwdBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC1d = "<tr><td width=15" & strPercent & " bgcolor=" & sQ & strAcctBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC1e = "<tr><td width=15" & strPercent & " bgcolor=" & sQ & strLoginsBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC1Close= "</strong></font></td>"
  sHTMLC2 = "<td width=20" & strPercent & "><font face=" & sQ & "Verdana" & sQ & " size=2"& sQ & ">"
  sHTMLC2a = "<td colspan=4><font face=" & sQ & "Verdana" & sQ & " size=2"& ">"
  sHTMLC2Close= "</strong></font></td>"
  sHTMLC3 = "<td width=15" & strPercent & " bgcolor=" & sQ & strBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC3b = "<td width=15" & strPercent & " bgcolor=" & sQ & strPwdBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC3c = "<td width=15" & strPercent & " bgcolor=" & sQ & strAcctBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC3d = "<td width=15" & strPercent & " bgcolor=" & sQ & strPwdExpBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC3e = "<td width=15" & strPercent & " bgcolor=" & sQ & strLoginsBGColor & sQ & " align=" & sQ & "right" & sQ & "><font face="& sQ & "Verdana" & sQ & "size=2 color="& sQ & "#FFFFFF"& sQ & "><strong>"
  sHTMLC3Close= "</strong></font></td>"
  sHTMLC4 = "<td><font face=" & sQ & "Verdana" & sQ & " size=2" & sQ & ">"
  sHTMLC4Close= "</strong></font></td>"
  
  'Display user information in HTML: each line broken up for better reading
  'objDocument.WriteLn "<marquee width=85" & strPercent & ">Active Directory Information for " & strDisplayName & ".</marquee>"
  objDocument.WriteLn "<html><head><meta name=" & sQ & "GENERATOR" & sQ & "content=" & sQ & "Ralph Montgomery, rmonty@myself.com" & sQ & "><title>Active Directory Information for: " & strDisplayName & "</title></head><body>"
  
  objDocument.WriteLn "<script language=" & sQ & "JavaScript1.2" & sQ & ">"
  objDocument.WriteLn "top.window.moveTo(0,0);"
  objDocument.Writeln "if (document.all) {"
  objDocument.WriteLn "top.window.resizeTo(screen.availWidth,screen.availHeight);"
  objDocument.WriteLn "}"
  objDocument.WriteLn "else if (document.layers||document.getElementById) {"
  objDocument.WriteLn "if (top.window.outerHeight<screen.availHeight||top.window.outerWidth<screen.availWidth){"
  objDocument.WriteLn    "top.window.outerHeight = screen.availHeight;"
  objDocument.WriteLn "top.window.outerWidth = screen.availWidth;"
  objDocument.WriteLn "}"
  objDocument.WriteLn "}"
  objDocument.WriteLn "</script>"
  
  objDocument.WriteLn "<Table border =0 Width = 65" & strPercent & "><Caption><strong>User Information for: </strong>" & strDisplayName & "</Caption>"
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  objDocument.WriteLn sHTMLC1 & "Distinguished Name:" & sHTMLC1Close & sHTMLC2a & strDN & sHTMLC2Close
  objDocument.WriteLn sHTMLC1 & "Acct Created:" & sHTMLC1Close & sHTMLC2a & strWhenCreated & " GMT" & sHTMLC2Close
  objDocument.WriteLn sHTMLC1 & "Acct changed:" & sHTMLC1Close & sHTMLC2a & strWhenChanged & " GMT" & sHTMLC2Close
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  objDocument.WriteLn sHTMLC1 & "Logon Name: " & sHTMLC1Close & sHTMLC2 & strLogonName & sHTMLC2Close & sHTMLC3 & "Description: " & sHTMLC3Close & sHTMLC4 & strDisplayDescription & sHTMLC4Close
  objDocument.WriteLn sHTMLC1 & "Department: " & sHTMLC1Close & sHTMLC2 & strDisplayDepartment & sHTMLC2Close & sHTMLC3 & "Telephone: " & sHTMLC3Close & sHTMLC4 & strTelephoneNumber & sHTMLC4Close
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  objDocument.WriteLn sHTMLC1d & "Acct Locked:" & sHTMLC1Close & sHTMLC2 & strIsAccountLocked & sHTMLC2Close & sHTMLC3c & "Account Disabled:" & stHTMLC3Close & sHTMLC4 & strAccountDisabled & sHTMLC4Close
  objDocument.WriteLn sHTMLC1e & "Bad Logins:" & sHTMLC1Close & sHTMLC2 & intBadPwd & sHTMLC2Close & sHTMLC3e & "Max/Attempts Left:" & sHTML3Close & sHTMLC4 & intLockoutThreshold & "/" & (intLockoutThreshold - intBadPwd) & sHTMLC4Close
  objDocument.WriteLn sHTMLC1 & "Last failed login:" & sHTMLC1Close & sHTMLC2 & dtmLastFailedLogin & sHTMLC2Close & sHTMLC3 & "Last Successful login:" & sHTMLC3Close & sHTMLC4 & dtmLastLogin & sHTMLC4Close
  
  'objDocument.WriteLn sHTMLC1 & "Last Workstation:" & sHTMLC1Close & sHTMLC2 & strLastLoggedInWorkstation & sHTMLC2Close &sHTMLC3 & "Last IP Address:" & sHTMLc3Close & sHTMLC4 & strMostRecentIP & sHTMLC4Close
  'objDocument.WriteLn sHTMLC1 & "Last Workstation:" & sHTMLC1Close & sHTMLC2 & strLastLoggedInWorkstation & sHTMLC2Close
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  objDocument.WriteLn sHTMLC1 & "Pwd Changed:" & sHTMLC1Close & sHTMLC2 & strPwdLastChanged & sHTMLC2Close & sHTMLC3 & "Pwd Age:" & sHTMLC3Close & sHTMLC4 & strPwdAge & sHTML4Close
  objDocument.WriteLn sHTMLC1 & "User change Pwd:" & sHTMLC1Close & sHTMLC2 & strPwdCanChange & sHTMLC2Close & sHTMLC3 & "Pwd Never Expires:" & sHTMLC3Close & sHTMLC4 & strPwdNeverExpires & sHTMLC4Close
  objDocument.WriteLn sHTMLC1 & "Min Pwd Length:" & sHTMLC1Close & sHTMLC2 & intMinPwdLength & sHTMLC2Close & sHTMLC3 & "Min Pwd History:" & sHTMLC3Close & sHTMLC4 & intPwdHistoryLength & " pwd(s)" & sHTMLC4Close
  objDocument.WriteLn sHTMLC1 & "Lockout Time:" & sHTMLC1Close & sHTMLC2 & intLockoutDurationMinutes & sHTMLC2Close & sHTMLC3 & "AutoUnlock:" & sHTMLC3Close & sHTMLC4 & intLockOutObservationWindowMinutes & sHTMLC4Close
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  objDocument.WriteLn sHTMLC1 & "Home Directory:" & sHTMLC1Close & sHTMLC2 & strHomeDirectory & sHTMLC2Close & sHTMLC3 & "Home Drive:" & sHTMLC3Close & sHTMLC4 & strHomeDrive & sHTMLC4Close
  objDocument.WriteLn sHTMLC1 & "Roaming Profile:" & sHTMLC1Close & sHTMLC2 & strProfilePath & sHTMLC2Close & sHTMLC3 & "Logon Script:" & sHTMLC3Close & sHTMLC4 & strScriptPath & sHTMLC4Close
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  objDocument.WriteLn sHTMLC1 & "Allow TS Logon:" & sHTMLC1Close & sHTMLC2 & strTSAllowLogon & sHTMLC2Close & sHTMLC3 & "Remote Control:" & sHTML3Close & sHTMLC4 & strTSEnableRemoteControl & sHTMLC4Close
  objDocument.WriteLn sHTMLC1 & "Connect Client Drives: " & sHTMLC1Close & sHTMLC2 & strTSConnectDrives & sHTMLC2Close & sHTMLC3 & "Auto Create Printers:" & sHTML3Close & sHTMLC4 & strTSConnectPrinters & sHTMLC4Close
  'objDocument.WriteLn sHTMLC1 & "TS Profile:" & sHTMLC1Close & sHTMLC2a & strTSProfilePath & sHTMLC2Close    
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  If strExchange = "Yes" Then    
    objDocument.WriteLn sHTMLC1 & "Exchange Account: " & sHTMLC1Close & sHTMLC2 & strExchange & sHTMLC2Close & sHTMLC3 & "Allow VPN Access:"& sHTMLC3Close & sHTMLC4 & strVPNAllow & sHTMLC4Close
    objDocument.WriteLn sHTMLC1 & "Exchange Alias:" & sHTMLC1Close & sHTMLC2 & strMailNickname & sHTMLC2Close & sHTMLC3 & "Assigned Delegates:" & sHTMLC3Close & sHTMLC4 & strDisplayDelegates & sHTMLC4Close    
    objDocument.WriteLn sHTMLC1 & "Ext email address:" & sHTMLC1Close & sHTMLC2 & strUserMail & sHTMLC2Close
    
  Else
    objDocument.WriteLn sHTMLC1 & "Exchange Account:" & sHTMLC1Close & sHTMLC2 & strExchange & sHTMLC2Close & sHTMLC3 & "Allow VPN Access:"& sHTMLC3Close & sHTMLC4 & strVPNAllow & sHTMLC4Close
    
  End If
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  objDocument.WriteLn sHTMLC1 & "Group Membership: " & sHTMLC1Close & sHTMLC2a & strSortedGroups & sHTMLC2Close
  objDocument.WriteLn sHTMLC1a & "<HR>" & sHTMLC1Close
  
  objDocument.WriteLn "</table></body></html>"
  
End Sub ' Display User in IE Window