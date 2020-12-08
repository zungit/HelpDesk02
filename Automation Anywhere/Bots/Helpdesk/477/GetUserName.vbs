UserName = getUserName() 
 
Function getUserName() 
  Set wshShell = CreateObject( "WScript.Shell" )
    getUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
End Function
