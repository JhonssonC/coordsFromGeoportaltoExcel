Attribute VB_Name = "COORDSGEOPORTAL"

Private Function Col_Letter(lngCol As Variant)
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function


Private Function requestWeb(STR As Variant)
Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    Dim arrResult() As Variant
    Dim myxml As String
    
    objHTTP.Open "GET", "" & STR, False

    objHTTP.SetRequestHeader "Content-Type", "text/json;charset=utf-8"
    objHTTP.Send
    requestWeb = objHTTP.responseText
    
End Function


Public Sub reqFromGeoportal(destinoX As Range, destinoY As Range)

    Dim clsJSON As clsJSParse
    Set clsJSON = New clsJSParse
    Dim strVar As String
    
    cta = Range(Sheets("VAR").Range("B2") & ActiveCell.Row)
    initUrl = Sheets("VAR").Range("B1")
  
    result = requestWeb(initUrl & "query?f=json&returnIdsOnly=true&where=UPPER(CODIGOCLIENTE)%20LIKE%20%27%25" & cta & "%25%27&returnGeometry=false&spatialRel=esriSpatialRelIntersects&outSR=102100")
     

    strVar = result
    
    clsJSON.LoadString strVar

    objectId = clsJSON.Value(2)
    
    
    If "" & objectId <> "" Then
    

        result = requestWeb(initUrl & "query?f=json&where=&returnGeometry=true&spatialRel=esriSpatialRelIntersects&objectIds=" & objectId & "&outFields=OBJECTID%2CMIOID%2CALIMENTADORID%2CCODIGOEMPRESA%2CPROVINCIA%2CCANTON%2CPARROQUIA%2CFASECONEXION%2CSUBTIPO%2CCODIGOCLIENTE%2CMEDIDOR%2CCOORD_X%2CCOORD_Y&outSR=102100")
        
        strVar = result
    

        clsJSON.LoadString strVar
        
        
        out_x = Replace(clsJSON.Value(clsJSON.NumElements - 3), ".", ",")
        out_y = Replace(clsJSON.Value(clsJSON.NumElements - 2), ".", ",")
        
        
        If out_x <> "" And out_y <> "" Then

         
            destinoX = "'" & out_x
            destinoY = "'" & out_y
    
        End If
        
    End If


End Sub


Sub COORDS_GEOPORTAL()

    If Selection.Cells.Rows.Count > 1 Then
        
        RF = Selection.Cells(Selection.Cells.Rows.Count, 1).Row
        RI = Selection.Cells(1, 1).Row
        
        COLU = Col_Letter(Selection.Cells.Column)
        
        Dim celda As Range
    
        For Each celda In Range(COLU & RI & ":" & COLU & RF).SpecialCells(xlCellTypeVisible)
            celda.Select
            If ActiveCell <> "" Then
                reqFromGeoportal Range(Sheets("VAR").Range("B3") & ActiveCell.Row), Range(Sheets("VAR").Range("B4") & ActiveCell.Row)
            End If
        Next
    Else
        If ActiveCell <> "" Then
            reqFromGeoportal Range(Sheets("VAR").Range("B3") & ActiveCell.Row), Range(Sheets("VAR").Range("B4") & ActiveCell.Row)
        End If
    End If

End Sub

