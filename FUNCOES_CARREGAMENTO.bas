Attribute VB_Name = "FUNCOES_CARREGAMENTO"
Option Explicit

Function barra_de_Progresso(linha As Long, tam_Execucao, nome_Da_Barra As String, mensagem_Form As String)
        
    Dim current_Progress As Double
    Dim current_Percentage As Double
    Dim bar_Width As Long
    
    current_Progress = linha / tam_Execucao
    bar_Width = Progresso.Borda.Width * current_Progress
    current_Percentage = Round(current_Progress * 100, 0)
    Progresso.Barra.Width = bar_Width
    Progresso.Texto.Caption = current_Percentage & "% Completo"
    Progresso.Caption = "" & nome_Da_Barra
    Progresso.Mensagem.Caption = "" & mensagem_Form
    DoEvents
End Function
Sub iniciando_BarraProgresso()
    With Progresso
        .Barra.Width = 0
        .Texto.Caption = "0% Completo"
        .Show vbModeless
    End With
End Sub
Function termina_BarraProgresso()
    Unload Progresso
End Function
