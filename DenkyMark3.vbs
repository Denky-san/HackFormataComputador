'parece que você não é tão bobo igual eu pensei, use esse código para assustar seus amigos também!

Dim senha_correta
senha_correta = "denkyquebrandocodigos"

Dim tentativas_restantes
tentativas_restantes = 3 ' número de tentativas permitidas

Function ReiniciarComputador()
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    For i = 1 To 3
        MsgBox "ERROR", vbOKOnly + vbCritical, "ALERTA"
    Next
	
	'comandos de listagem no CMD para assustar a pessoa
	WshShell.Run "cmd.exe /c schtasks /query", 1, True
	WScript.Sleep 1000
	WshShell.Run "cmd.exe /c tasklist", 1, True
	WScript.Sleep 1000
	WshShell.Run "cmd.exe /c schtasks /query", 1, True
	WScript.Sleep 1000
	WshShell.Run "cmd.exe /c tasklist", 1, True
	WScript.Sleep 1000
	
    WshShell.Run "cmd.exe /c shutdown /r /t 0", 0, True
End Function


Do Until tentativas_restantes < 0
    Dim senha_digitada
    senha_digitada = InputBox("Seu PC sera FORMATADO em breve. Para CANCELAR, insira a senha correta.", "ALERTA", "")
    
    If senha_digitada = senha_correta Then
        MsgBox "Senha correta. O processo de formatacao foi cancelado.", 64, "ALERTA"
        Exit Do
	ElseIf senha_digitada = "" Then
		tentativas_restantes = tentativas_restantes - 1
		If tentativas_restantes = 0 Then
			ReiniciarComputador()
        Else
           MsgBox "Boa tentativa, mas voce nao vai fugir. Tentativas restantes: "  & tentativas_restantes, vbOKOnly + vbExclamation, "ALERTA"
    End If
    Else
        tentativas_restantes = tentativas_restantes - 1
        If tentativas_restantes = 0 Then
			ReiniciarComputador()
        Else
            MsgBox "Senha incorreta. Tente novamente. Tentativas restantes: " & tentativas_restantes, vbOKOnly + vbExclamation, "ALERTA"
        End If
    End If
Loop
