Attribute VB_Name = "Botão"
Rem Link útil, como executar script Python a partir de botão no Excel: https://pythonandvba.com/blog/how-to-execute-a-python-script-from-excel-using-vba/
Rem Link útil, como passar argumento para o script Python: https://stackoverflow.com/questions/63873954/excel-vba-pass-arguments-to-python-script
Sub RunPythonScript()
    
    Rem Armazena o endereço da célula atual. Deverá ser a célula com o ID da NCMR a ser gerada.
    Dim aba As String
    Dim endereco_ID As String
    aba = ActiveSheet.Name
    endereco_ID = ActiveCell.Address
    
    Rem Armazena o conteúdo da célula atual. Deverá ser a célula com o ID da NCMR a ser gerada.
    Dim codigo As String
    codigo = ActiveCell.Value
    
    Rem Armazena o resultado da MsgBox
    Dim result As Integer
    result = MsgBox("A célula selecionada é " + endereco_ID + "." + vbCrLf _
                    + "O seu conteúdo é: " + vbCrLf + codigo + vbCrLf + vbCrLf _
                    + "Essa é a célula com o código desejado?", _
                    1)
    
    Rem Seleciona que fazer conforme o resultado da MsgBox
    Select Case result
    Case 1
        Dim objShell As Object
        Set objShell = VBA.CreateObject("Wscript.Shell")
        
        Rem Este é o path da pasta
        Rem Dim folder_path As String
        
        Rem A função GetLocalPath está no módulo GetLocalPathModule e foi copiada do link: https://stackoverflow.com/a/73577057/12287457
        Rem folder_path = "C:\Users\" & Environ("Username") & "\OneDrive - Wabtec Corporation\Teste\GitHub\pptx-automatico\"
        folder_path = GetLocalPath(ThisWorkbook.path)
        
        Rem Este é o path do arquivo Python.exe instalado na máquina ou no .venv. O Chr(34) é uma aspa dupla ("), necessária para passar os argumentos.
        Dim PythonExePath As String
        PythonExePath = "\.venv\Scripts\python.exe"
        PythonExePathFull = Chr(34) & folder_path & PythonExePath & Chr(34)
        
        Rem Este é o path do script Python a ser executado. O Chr(34) é uma aspa dupla ("), necessária para passar os argumentos.
        Dim PythonScriptPath As String
        PythonScriptPath = "\pptx-automatico.py"
        PythonScriptPathFull = Chr(34) & folder_path & PythonScriptPath & Chr(34)
        
        Dim ExePath As String
        ExePath = "\pptx-automatico.exe"
        ExePathFull = Chr(34) & folder_path & ExePath & Chr(34)
        
        Rem Estes são os argumentos a ser passados para o script Python
        Args = aba & " " & endereco_ID
        
        Rem Range("$A$2").Value = ExePathFull & " " & Args
        objShell.Run ExePathFull & " " & Args
        Rem objShell.Run PythonExePathFull & " " & PythonScriptPathFull & " " & Args
    Case 2
        MsgBox ("Você clicou em Cancelar. Selecione a célula correta.")
    End Select
End Sub

