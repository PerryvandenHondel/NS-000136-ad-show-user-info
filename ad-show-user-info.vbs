


Option Explicit


Const 	ADS_UF_ACCOUNTDISABLE = 					2
Const 	ADS_UF_PASSWD_NOTREQD = 					&h00020
Const 	ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 	&h0080
Const 	ADS_UF_DONT_EXPIRE_PASSWD = 				&h10000
Const 	ADS_UF_PASSWORD_EXPIRED = 					&h80000
Const 	ADS_UF_PASSWD_CANT_CHANGE = 				&h0040
Const 	CHANGE_PASSWORD_GUID = 						"{ab721a53-1e2f-11d0-9819-00aa0040529b}"
Const 	ADS_ACETYPE_ACCESS_DENIED_OBJECT = 			&H6
Const 	SEC_IN_MIN = 								60
Const 	SEC_IN_DAY = 								86400
Const 	MIN_IN_DAY = 								1440


Dim		strDn
Dim		objUser


Function EncloseWithDQ(ByVal s)
	''
	''	Returns an enclosed string s with double quotes around it.
	''	Check for exising quotes before adding adding.
	''
	''	s > "s"
	''
	
	If Left(s, 1) <> Chr(34) Then
		s = Chr(34) & s
	End If
	
	If Right(s, 1) <> Chr(34) Then
		s = s & Chr(34)
	End If

	EncloseWithDQ = s
End Function '' of Function EncloseWithDQ


Function RemoveEnclosedDQ(ByVal s)
	''
	''	Removes the enclosed Double Quotes around a string
	''
	''	"s" > s
	''
	
	If Left(s, 1) = Chr(34) Then
		s = Right(s, Len(s) - 1)
	End If
	
	If Right(s, 1) = Chr(34) Then
		s = Left(s, Len(s) - 1)
	End If

	RemoveEnclosedDQ = s
End Function '' of Function EncloseWithDQ


Function DsQueryGetDn(ByVal strRootDse, ByVal strCn)
	''
	''	Use the DSQUERY.EXE command to find a DN of a CN in a specific AD set by strRootDse
	''
	''		strRootDse: DC=prod,DC=ns,DC=nl
	''		strCn: 		ZZZ_NAME_OF_GROUP
	''
	''		Returns: 	The DN of blank if not found.
	''
	
	Dim		c			''	Command
	Dim		r			''	Result
	Dim		objShell
	Dim		objExec
	Dim		strOutput
	
	If InStr(strCn, "CN=") > 0 Then
		'' When the strCN already contains a Distinguished Name (DN), result = strCn
		r = strCn
	Else
		'' No, we must search for the DN based on the CN
	
		c = "dsquery.exe "
		c = c & "* "
		c = c & strRootDse & " "
		c = c & "-filter (CN=" & strCn & ")"

		Set objShell = CreateObject("WScript.Shell")
		Set objExec = objShell.Exec(c)
		
		Do
			strOutput = objExec.Stdout.ReadLine()
		Loop While Not objExec.Stdout.atEndOfStream

		Set objExec = Nothing
		Set objShell = Nothing
		If Len(strOutput) > 0 Then
			r = strOutput  '' BEWARE: r contains now " around the string, see "CN=name,OU=name,DC=domain,DC=nl"
		Else
			WScript.Echo "ERROR Could not find the Distinguished Name for " & strCn & " in " & strRootDse
			r = ""
		End If
	End If
	DsQueryGetDn = RemoveEnclosedDQ(r)
End Function '' DsQueryGetDn



'Set objDomainNt = GetObject


Dim 	intUac
Dim		objDomainNT
Dim		intMaxPwdAge
Dim		intMaxPwdAgeSeconds
Dim		intMinPwdAgeSeconds
Dim		intLockOutObservationWindowSeconds
Dim		intLockoutDurationSeconds

Set objDomainNT = GetObject("WinNT://PROD") 
With objDomainNT
    intMaxPwdAge =                             .Get("MaxPasswordAge")    'get NT value for MaxPasswordAge
    intMaxPwdAge =                             (intMaxPwdAge/SEC_IN_DAY) ' maximum password age in days
    intMaxPwdAgeSeconds =                     .Get("MaxPasswordAge")
    intMinPwdAgeSeconds =                     .Get("MinPasswordAge")
    intLockOutObservationWindowSeconds =     .Get("LockoutObservationInterval")
    intLockoutDurationSeconds =             .Get("AutoUnlockInterval")
 End With 'objDomainNT
 Set objDomainNT = Nothing
 
 WScript.Echo intMaxPwdAge





strDn = DsQueryGetDn("DC=prod,DC=ns,DC=nl", "Perry.vandenHondel")
WScript.Echo strDn

On Error Resume Next
Set objUser = GetObject("LDAP://" & strDn)
If Err.Number = 0 Then
	WScript.Echo "Connected to " & strDn

	WScript.Echo objUser.Get("displayName")
	WScript.Echo objUser.Get("mail")
	
	
	WScript.Echo objUser.PasswordLastChanged
	
	intUac = objUser.Get("userAccountControl")
	If intUac And ADS_UF_DONT_EXPIRE_PASSWD Then
		
		WScript.Echo "Password does not exipre"
		
	Else
		WScript.Echo "Password expires"
	End If
	
	
	Set objUser = Nothing
Else
	WScript.Echo "ERROR: Could not connect to user object for " & strDn & ", code " & Err.Number
End If


WScript.Quit(0)