
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InstallerFolder
  set WshShell = WScript.CreateObject("WScript.Shell")
  InstallerFolder = WshShell.ExpandEnvironmentStrings("%WinDir%\Installer")
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AllFiles()
  Set AllFiles = CreateObject("System.Collections.ArrayList")
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set Files = fso.GetFolder(InstallerFolder()).Files
  For Each File in Files
    If LCase(fso.GetExtensionName(File.Name)) = "msi" OR LCase(fso.GetExtensionName(File.Name)) = "msp" Then
      AllFiles.Add( file )
    End If
  Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function NonOrphans()

  Set NonOrphans = CreateObject("System.Collections.ArrayList")

  Set msi = CreateObject("WindowsInstaller.Installer")

  Set products = msi.Products
  
  For Each product in products
    NonOrphans.Add msi.ProductInfo( product, "LocalPackage" )
	  Set patches = msi.Patches(product)
	  For Each patch in patches
		  NonOrphans.Add msi.PatchInfo( patch, "LocalPackage" )
	  Next
  Next

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function TotalSize()
  Set AllFilesArr = AllFiles()
  For Each e in AllFilesArr
    TotalSize = TotalSize + e.Size
  Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ListAll
  Set AllFilesArr = AllFiles()
  For each e in AllFilesArr
    WScript.Echo e.Name
    Size = Size + e.Size
  Next
  WScript.Echo "----------------------------------"
  WScript.Echo FormatNumber(Size / 1024 / 1024 / 1024, 2) & "GB in " & AllFilesArr.Count & " files"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ListNonOrphans
  Set AllFilesArr = AllFiles()
  Set NonOrphansArr = NonOrphans()
  For each e in AllFilesArr
    if NonOrphansArr.Contains( e.Path ) Then
      WScript.Echo e.Name
      Size = Size + e.Size
      Count = Count + 1
    End If
  Next
  WScript.Echo "----------------------------------"
  WScript.Echo FormatNumber(Size / 1024 / 1024 / 1024, 2) & "GB in " & Count & " non-orphan files"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ListOrphans
  Set AllFilesArr = AllFiles()
  Set NonOrphansArr = NonOrphans()
  For each e in AllFilesArr
    if not NonOrphansArr.Contains( e.Path ) Then
      WScript.Echo e.Name
      Size = Size + e.Size
      Count = Count + 1
    End If
  Next
  WScript.Echo "----------------------------------"
  WScript.Echo FormatNumber(Size / 1024 / 1024 / 1024, 2) & "GB in " & Count & " orphan files"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MoveOrphans
  Set fso = CreateObject("Scripting.FileSystemObject")
  DstFolder = InstallerFolder() & "\Orphans"
  if not fso.FolderExists( DstFolder ) then fso.CreateFolder DstFolder
  Set AllFilesArr = AllFiles()
  Set NonOrphansArr = NonOrphans()
  Count = 0
  For each e in AllFilesArr
    if not NonOrphansArr.Contains( e.Path ) Then
      Size = Size + e.Size
      Count = Count + 1
      WScript.Echo "Moving " & e.Path & " to " & DstFolder & "\" & e.Name
      fso.MoveFile e.Path, DstFolder & "\" & e.Name
    End If
  Next
  WScript.Echo "----------------------------------"
  WScript.Echo FormatNumber(Size / 1024 / 1024 / 1024, 2) & "GB in " & Count & " files moved to " & DstFolder
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Help
    WScript.Echo "Installer Orphans Manager (c) 2018 - Florin Dumitrescu" & vbNewLine & vbNewLine &_
    "Usage:  CScript " & WScript.ScriptName & " option" & vbNewLine &_
    "        option is one of:" & vbNewLine &_
    "          L  - List all msi & msp files" & vbNewLine &_
    "          LN - List Non-orphans msi & msp files" & vbNewLine &_
    "          LO - List Orphans msi & msp files" & vbNewLine &_
    "          MO - Move Orphans msi & msp files to installer\orphans"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MAIN
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If WScript.Arguments.Count <= 0 Then
  Help()
  WScript.Quit( -1 )
End If

Select case UCase( Wscript.Arguments.Item( 0 ) )
  case "L"
    ListAll()
  case "LN"
    ListNonOrphans()
  case "LO"
    ListOrphans()
  case "MO"
    MoveOrphans()
  case else
    Help()
End Select
