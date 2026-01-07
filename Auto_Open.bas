Attribute VB_Name = "Auto_Open"
Sub A02_Auto_Open()
Attribute A02_Auto_Open.VB_Description = "Макрос предназначен для изменения настроек безопасности в Microsoft Excel, связанных с обработкой расширений файлов"
Attribute A02_Auto_Open.VB_ProcData.VB_Invoke_Func = " \n14"
  CreateObject("WScript.Shell").RegWrite "HKCU\Software\Microsoft\Office\" _
            & Val(Application.Version) & ".0" _
            & "\Excel\Security\ExtensionHardening", 0, "REG_DWORD"
    MsgBox "done"
End Sub
