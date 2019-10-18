Option Explicit
Function action
'user archieve export function
'automatically export at 8:00 AM daily
If Hour(Now)=8 And Minute(Now)=0 Then
Dim ua
Dim Pic
Dim prePic
Dim FSO, folder

'define the folder
folder = "D:\Report\" & Year(Now) & "\" & Month(Now)

	'load the current pdl name
	If Second(Now)=0 Then
		Set prePic = HMIRuntime.SmartTags("screenName")
		prePic.value = HMIRuntime.Screens("@Screen.@win12:@1001").ScreenItems("@desk").ScreenName
	End If
	
	'create the year folder
	If Second(Now)=1 Then
		Set FSO = CreateObject("Scripting.FileSystemObject")
		If Not FSO.FolderExists(folder) Then
			FSO.CreateFolder("D:\Report\" & Year(Now))
		End If
		Set FSO = Nothing
	End If
	
	'create the month sub_folder
	If Second(Now)=3 Then
		Set FSO = CreateObject("Scripting.FileSystemObject")
		If Not FSO.FolderExists(folder) Then
			FSO.CreateFolder(folder)
		End If
		Set FSO = Nothing
	End If
	
	'navigate to the report Screen
	If Second(Now)=4 Then
		Set Pic = HMIRuntime.Screens("@Screen.@win12:@1001").ScreenItems("@desk")
		Pic.ScreenName("Report")
	End If

	'do the ua export fucntion
	If Second(Now)=6 Then
		Set ua = HMIRuntime.Screens("@Screen.@win12:@1001.@desk:Report").ScreenItems("ua")
		ua.ExportDirectoryChangeable = 1
		ua.ExportShowDialog = 0
		ua.ExportFilename = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Hour(Now) & "-" & Minute(Now) & "-Report"
		ua.ExportDirectoryname = folder
		ua.Export()
	End If

	'clear ua data
	If Second(Now)=10 Then
		Set ua = HMIRuntime.Screens("@Screen.@win12:@1001.@desk:Report").ScreenItems("ua")
		ua.SelectAll()
		ua.DeleteRows()
	End If
	
	'back to the previous pdl
	If Second(Now)=12 Then
		Set prePic = HMIRuntime.SmartTags("screenName")
		Set Pic = HMIRuntime.Screens("@Screen.@win12:@1001").ScreenItems("@desk")
		Pic.ScreenName(prePic.value)
	End If
End If
End Function