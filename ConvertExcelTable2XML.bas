Option Explicit
    
Sub Gen_Click()

' External variables:
'
'   Range("doz_databegin")
'
'   This is the top-left of the table to convert to XML
'
'   The Excel table
'
'       The first row is the header which indicate the XML structure
'
'           use period '.' delimited to separate XML element and XML attribute
'           for example:
'                user.id     |   user.name   |  accessright.id   | accessright.enabled
'              --------------+---------------+-------------------+----------------------
'
'       The second row+ are attribute values
'
'   Range("doz_xml_root")
'   Range("doz_temp_data")
'   Range("doz_xml_header")
'   Range("doz_xml_footer")



    Dim objDom As DOMDocument
    Dim objXMLRootelement As IXMLDOMElement
    Dim objXMLelement As IXMLDOMElement
    Dim objXMLattr As IXMLDOMAttribute
    
    Range("doz_databegin").Select
    
    Dim col_start, row_start, col_end, row_end, col_count, row_count, col_iter, row_iter
    Dim ch$, struc$, attr$, val$, prev_val$, prev_struc$
    Dim c, r, i, j, k

    col_start = Selection.Column
    row_start = Selection.Row + 1
    
    col_iter = 0
    row_iter = 0
    
    col_end = Selection.End(xlToRight).Column
    col_count = col_end - col_start + 1
    
    row_end = Selection.End(xlDown).Row
    row_count = row_end - row_start + 1
    
    
    Dim prev_row_val(10) As String
    
    ' struct start
    
    Dim structure(5) As String
    Dim structure_attr_count(5) As Integer
    Dim structure_attr_count_iter(5) As Integer
    Dim struc_count, struc_idx, struc_attr_counter
    Dim header_structure(10) As String
    Dim header_attr(10) As String
    
    struc_count = 0
    struc_idx = 0
    struc_attr_counter = 0
    
    For c = 0 To col_count - 1
    
        '     ActiveCell.Offset(2, 0).Range("A1").Select
        ch$ = Range("doz_databegin").Offset(0, c).Value
        struc$ = Left(ch$, InStr(ch$, ".") - 1)
        attr$ = Mid(ch$, InStr(ch$, ".") + 1)
        header_structure(c) = struc$
        header_attr(c) = attr$
        
        If (c = 0) Then
            structure(struc_idx) = struc$
            prev_struc$ = struc$
        End If
        
        If (prev_struc$ = struc$) Then
            struc_attr_counter = struc_attr_counter + 1
        Else
            structure_attr_count(struc_idx) = struc_attr_counter
            prev_struc$ = struc$
            struc_idx = struc_idx + 1
            structure(struc_idx) = struc$
            struc_attr_counter = 0
        End If
        
        ' init prev_val
        prev_row_val(c) = Range("doz_databegin").Offset(1, c).Value

    Next c
    
    If (prev_struc$ = struc$) Then
        structure_attr_count(struc_idx) = struc_attr_counter + 1
        struc_attr_counter = 0
    End If
    struc_count = struc_idx + 1
    
    ' struct end

    ' row data processing

' ------------------------------------------------------
    ' process first row
    
    Set objDom = New DOMDocument
    Set objXMLRootelement = objDom.createElement(Range("doz_xml_root").Value)
    objDom.appendChild objXMLRootelement

    Dim cursor(5) As IXMLDOMElement
    
    c = 0
    Set cursor(0) = objXMLRootelement
    
    For i = 0 To struc_count - 1
    
        Set objXMLelement = objDom.createElement(structure(i))
        Set cursor(i + 1) = objXMLelement
        
        For j = 0 To structure_attr_count(i) - 1
        
            Set objXMLattr = objDom.createAttribute(header_attr(c))
            objXMLattr.NodeValue = Range("doz_databegin").Offset(1, c).Value
            objXMLelement.setAttributeNode objXMLattr

            c = c + 1
        Next j ' attr
        
        cursor(i).appendChild objXMLelement
    
    Next i ' struc
    


    Dim changed_val_idx
    changed_val_idx = -1
    
    Dim si, sj

    For r = 1 To row_count
    
        c = 0
        
        For si = 0 To struc_count - 1
            
            For sj = 0 To structure_attr_count(si) - 1
                val$ = Range("doz_databegin").Offset(r, c).Value
                If (prev_row_val(c) <> val$) Then
                    changed_val_idx = c
                    GoTo GOT_IDX
                End If
                c = c + 1
            Next sj ' attr
        Next si ' struc
GOT_IDX:
        If (changed_val_idx >= 0) Then
        
            c = c - sj
            For i = si To struc_count - 1
            
                Set objXMLelement = objDom.createElement(structure(i))
                Set cursor(i + 1) = objXMLelement
    
                For j = 0 To structure_attr_count(i) - 1
                    val$ = Range("doz_databegin").Offset(r, c).Value
                    prev_row_val(c) = val$
                    
                    Set objXMLattr = objDom.createAttribute(header_attr(c))
                    objXMLattr.NodeValue = Range("doz_databegin").Offset(r, c).Value
                    objXMLelement.setAttributeNode objXMLattr
                    
                    c = c + 1
                Next j ' attr
                
                cursor(i).appendChild objXMLelement
                
            Next i ' struc
            
        End If
            
    Next r
        
    
    
Debug.Print cursor(0).XML

Range("doz_temp_data").Value = Range("doz_xml_header").Value & vbCrLf & cursor(0).XML & vbCrLf & Range("doz_xml_footer").Value

'
End Sub

