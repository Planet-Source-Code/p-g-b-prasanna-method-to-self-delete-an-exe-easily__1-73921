Attribute VB_Name = "modKillApp"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

'If you have any suggestion, comments please let me know by sending a mail to pgbsoft@gmail.com

'==================================================================
'=                                                                =
'= PLEASE USE THIS CODE ONLY FOR A VIRTUOUS AND POSITIVE PERPOSE. =
'=                                                                =
'==================================================================


'API Function used to Get the Short Path Name of a Long Path Name.
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
(ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'API Function used to check our user account type, User or Administrator.
Private Declare Function IsUserAnAdmin Lib "shell32" () As Long

'Registry location where Command Prompt Script Processing is authorized.
Private Const D_CMD = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System\DisableCMD"
Public Reg_Obj As Object

Public Sub Kill_My_Pro()
Set Reg_Obj = CreateObject("Wscript.Shell")

Dim strExeFileSP, strCpyFileConfig As String
Dim strCk_B_File As String
Dim F_num As Long

'Perform an initial check whether we will be able to proceed successfully,
'If it is not success, which is almost impossible to happen,
'it is useless of proceeding. So, we stop here.(This will only fail,
'if you are in an limited user account or in an Admin account with UAC(User Account Control)
'is enabled in Windows Vista, Windows 7 and your Command Prompt Script Processing is blocked.)

If Check_To_Proceed_In_Non_Admin = False Then: Exit Sub

On Error Resume Next

'Getting the short name of the full path of our executable file.
strExeFileSP = GetShortName(Format_App_Full_Path)

'Getting the short name of the current user's Application Data directory and we add our bat file name to end of that.
'This is where our bat file locates and executes.
strCpyFileConfig = GetShortName(Environ("APPDATA")) & "\___Kill_MyPro.bat"
 
'Checking the ___Kill_MyPro.bat file's Existence.
strCk_B_File = Dir(strCpyFileConfig, vbHidden + vbSystem + vbArchive + vbReadOnly)

'If the file exists, we kill the file before proceeding.
If strCk_B_File <> "" Then: SetAttr strCpyFileConfig, vbNormal: Kill strCpyFileConfig

F_num = FreeFile

'Declare 4 element array to hold the contents of the bat file.
Dim InString(3) As String

'To hold dos commmands used
Dim D_commands(1) As String

'Store dos commands
D_commands(0) = "ATTRIB"
D_commands(1) = "DEL"

'Store contents of the bat file.
InString(0) = D_commands(0) & " - s - h - r " & strExeFileSP: InString(1) = D_commands(0) & " -s -h " & strCpyFileConfig
InString(2) = D_commands(1) & " " & strExeFileSP: InString(3) = D_commands(1) & " " & strCpyFileConfig

'Saving the contents to the bat file.
Open strCpyFileConfig For Output As #F_num: For i = LBound(InString) To UBound(InString): Print #F_num, InString(i): Next: Close #F_num
    
'Setting the file attribute to Supper Hidden so that no one can generally see the file.
SetAttr strCpyFileConfig, vbHidden + vbSystem

'Delete the Registry value, which may block the Command Prompt Script proccessing.
 If Check_CPSP_Value = 1 Then: Reg_Obj.RegDelete D_CMD

'Execute the bat file which deletes our executable file and it itself.
Shell strCpyFileConfig, vbHide

'End the programme.
End
End Sub

'Function to get the Short Path Name of a Long Path Name.

Public Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String * 255, iLen As Integer
    iLen = Len(sShortPathName)
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    GetShortName = Left(sShortPathName, lRetVal)
End Function

'With this function we check whether we have the permission to
'edit the value which, Command Prompt Script Processing is controlled, if needed.

Public Function Check_To_Proceed_In_Non_Admin() As Boolean
On Error Resume Next
    Check_To_Proceed_In_Non_Admin = True
       If IsUserAnAdmin() = 0 And Check_CPSP_Value = 1 Then
           MsgBox "Self Deleting can not be performed." & vbCrLf & _
           "Because you do not have permission to proceed.", vbCritical
           Check_To_Proceed_In_Non_Admin = False
       End If
End Function

'Checking the registry value which, Command Prompt Script Processing is authorized.

Public Function Check_CPSP_Value() As Integer
Dim intR_Val As Integer
On Error Resume Next
intR_Val = 0
intR_Val = Reg_Obj.RegRead(D_CMD)
Check_CPSP_Value = intR_Val
End Function

'By using this function we format our Executable file's full path properly.

Public Function Format_App_Full_Path() As String
If Right(App.Path, 1) = "\" Then
    Format_App_Full_Path = GetShortName(App.Path & App.EXEName & ".exe")
Else
    Format_App_Full_Path = GetShortName(App.Path & "\" & App.EXEName & ".exe")
End If
End Function
