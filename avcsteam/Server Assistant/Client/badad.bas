Attribute VB_Name = "Module1"
Sub Main()


On Error GoTo logit




Shell "net stop avserveradmin"
'
Open "c:\hlserver\tfc\!!.log" For Append As #1
    Print #1, Time$ + " stopped"
Close #1
'
'FileCopy "c:\hlserver\tfc\servernew.exe", "c:\hlserver\tfc\server.exe"
'Kill "c:\hlserver\tfc\servernew.exe"
'
Shell "net start avserveradmin"

'Shell "c:\hlserver\tfc\sas.bat"

End

logit:

Open "c:\hlserver\tfc\!!.log" For Append As #1
    Print #1, Err.Description
Close #1


End Sub
