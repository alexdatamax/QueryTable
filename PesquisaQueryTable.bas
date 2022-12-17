Attribute VB_Name = "Módulo1"
Option Explicit


Public Sub pesquisaDados()

    Dim pesqTb As QueryTable
    Dim url As String
    
    Range("A1").CurrentRegion.ClearContents
    
    url = "Url;" & "https://pt.wikipedia.org/wiki/Lista_de_munic%C3%ADpios_do_Brasil_por_popula%C3%A7%C3%A3o"
    
    
    Set pesqTb = Planilha1.QueryTables.Add(url, Range("A1"))
    
    With pesqTb
    
        .RefreshOnFileOpen = False
        .Name = "Tabela_Cidades"
        .WebFormatting = xlWebFormattingNone
        .WebTables = "1"
        .Refresh
        
    End With
    
End Sub
