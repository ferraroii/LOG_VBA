VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

       Dim ztext As String
        Dim zpasta As String
        Dim arquivo As String

        Public Property Let pasta(value)
            zpasta = value
        End Property
        
        Public Property Let text(value)
            ztext = value
        End Property
        
        Private Property Get arq() As String
            arq = arquivo
        End Property

        Function criar_log()
            Set fs = CreateObject("Scripting.FileSystemObject")
            arquivo = zpasta & "\log" & Format(Now, "_ddmmyyyy_hhmmss") & ".txt"
            Set a = fs.CreateTextFile(arquivo, True)
            a.WriteLine Now & " - Log criado"
            a.Close
            criar_log = arquivo
        End Function

        Sub escrever_log()
            Dim FSO As New FileSystemObject
            Set FSO = CreateObject("Scripting.FileSystemObject")
            Set FileToWrite = FSO.OpenTextFile(arq, ForAppending)
            FileToWrite.WriteLine Now & " - " & ztext
            FileToWrite.Close
        End Sub
