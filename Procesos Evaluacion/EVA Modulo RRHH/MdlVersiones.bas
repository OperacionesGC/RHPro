Attribute VB_Name = "MdlVersiones"
Option Explicit


Public Function ValidarV(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que determina si el proceso esta en condiciones de ejecutarse.
' Autor      : Leticia
' Fecha      : 10/05/2012
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

Select Case TipoProceso
Case 69: 'Formulario GDD

    If Version >= "1.19" Then
        'Tabla nueva

        '/* este scrip fue probado en sql server*/
        'CREATE TABLE [dbo].[evaroleta] (
        '    [evacabnro] [int] NOT NULL ,
        '    [evatevnro] [int] NOT NULL ,
        '    [evaetanro] [int] NOT NULL ,
        ')
        'GO
              
        'ALTER TABLE [dbo].[evaroleta] ADD
        '    CONSTRAINT [FK_evaroleta1] FOREIGN KEY
        '    (
        '        [evacabnro]
        '    ) REFERENCES [evacab] (
        '        [evacabnro]
        '    )
        
                
        'ALTER TABLE [dbo].[evaroleta] ADD
        '    CONSTRAINT [FK_evaroleta2] FOREIGN KEY
        '    (
        '        [evatevnro]
        '    ) REFERENCES [evatipevalua] (
        '        [evatevnro]
        '    )
        
        Texto = "Revisar que exista tabla evaroleta    -y su estructura sea correcta."
        
        StrSql = "SELECT * FROM evaroleta WHERE evacabnro = 1 "
        OpenRecordset StrSql, rs
                
        V = True
    End If
    
Case Else:
    Texto = "version correcta"
    V = True
End Select


  ValidarV = V
  
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function

