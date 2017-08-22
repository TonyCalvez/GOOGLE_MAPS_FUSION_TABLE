Attribute VB_Name = "Module1"
Sub Export_Vers_Carto()

'Programme réalisé par Tony CALVEZ - Pas responsable des utilisations

'Dernière ligne renseignée de la feuille de calculs
Dim DerniereLigne As Integer
DerniereLigne = Range("A1").SpecialCells(xlCellTypeLastCell).Row

For n = 1 To DerniereLigne

    'Supprime les mots clefs suivant
    Range("C" & n) = Replace(Range("C" & n).Value, "LOTISSEMENT", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "LIEU", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "DIT", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "LIEU-DIT", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "LD", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "TER", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "LOT", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "BIS", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "TER", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "QUATER", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "ZA", "")
    Range("C" & n) = Replace(Range("C" & n).Value, "IMPASSE", "")
    'Supprime les chiffres de 1 à 10
    For i = 0 To 10
    Range("C" & n) = Replace(Range("C" & n).Value, i, "")
    Next

Range("A" & n) = Range("C" & n) & " " & Range("D" & n)
    
    If Range("B" & n) = "Branchement individuel neuf en soutirage" Or Range("B" & n) = "Branchement collectif neuf" Then
    Range("C" & n) = "small_red"
    
    ElseIf Range("B" & n) = "Modification de branchement" Then
    Range("C" & n) = "small_blue"
    
    Else
    Range("C" & n) = "small_green"
    
    End If
    
Next

Range("A1") = "Location"
Range("C1") = "Icon"

Range("D:D").Delete


ActiveWorkbook.SaveAs Filename:="C:\Users\" & Environ("USERNAME") & "\Desktop\" & Format(Date, "yyyymmdd") & "ExportFusion.xlsx", FileFormat:=xlOpenXMLWorkbook


Shell """" & Right(Environ(2), Len(Environ(2)) - 8) & "\FirefoxEDF\FirefoxPortable.exe " & """" & "fusiontables.google.com/DataSource?dsrcid=implicit""""", _
vbMaximizedFocus

ActiveWorkbook.Close SaveChanges:=False
ThisWorkbook.Close SaveChanges:=False

End Sub
