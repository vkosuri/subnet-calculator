Public Function pow(var) As Variant
    Dim mytemp As Variant
    mytemp = 1
    For i = 1 To var
        mytemp = mytemp * 2 + temp
    Next i
    pow = mytemp
End Function

Public Function dec2bin(var) As Variant

    Do
        tempArr = " " & (var Mod 2) & tempArr
        var = Int(var / 2)
    Loop While var > 0
    Do While Len(tempArr) <> 16
        tempArr = " " & 0 & tempArr
    Loop
    dec2bin = tempArr
End Function

Sub calc()
    Dim ipAdd As Variant
    Dim nonSubnetBits As Variant
    Dim defaultBits As Variant
    Dim remainBits As Variant
    Dim temp As Variant
    Dim subnetworks As Variant
    Dim calcSubnet(7) As Integer
      
    Dim temp1 As Variant
    Dim temp2 As Variant
    Dim temp3 As String
    Dim calcTemp(3) As String
    Dim arr1 As Variant
            
    calcTemp(0) = 0
    calcTemp(1) = 0
    calcTemp(2) = 0
    calcTemp(3) = 0
    
    calcSubnet(0) = 128
    calcSubnet(1) = 64
    calcSubnet(2) = 32
    calcSubnet(3) = 16
    calcSubnet(4) = 8
    calcSubnet(5) = 4
    calcSubnet(6) = 2
    calcSubnet(7) = 1
    
    reqBits = Worksheets("Sheet1").Cells(13, 4)
    
    ipAdd = Split(Worksheets("Sheet1").Cells(12, 4), ".")
    'UBound(Arr) - LBound(Arr) + 1 len of the array'UBound(Arr) - LBound(Arr) + 1 len of the array
    lenOfArr = (UBound(ipAdd) - LBound(ipAdd) + 1)
    
    If (lenOfArr <= 3) And (Worksheets("Sheet1").Cells(13, 4) <= 30) Then
        MsgBox "U Must Enter a valid IP Address", vbYesNo
        Worksheets("Sheet1").Cells(12, 4).Value = "0.0.0.0"
        Worksheets("Sheet1").Cells(13, 4).Value = "0"
    Else:
        reqBits = Worksheets("Sheet1").Cells(13, 4)
        nonSubnetBits = reqBits
        defaultBits = Worksheets("Sheet1").Cells(17, 2)
        sunbnetworks = Worksheets("Sheet1").Cells(17, 6)
        reqBits = Worksheets("Sheet1").Cells(13, 4)
    
        Worksheets("Sheet1").Cells(17, 3).Value = reqBits
        If ipAdd(0) >= 1 And ipAdd(0) <= 127 Then
            Worksheets("Sheet1").Cells(20, 4).Value = "Class A"
            Worksheets("Sheet1").Cells(21, 4).Value = "255.0.0.0"
            Worksheets("Sheet1").Cells(17, 2).Value = 8
        ElseIf ipAdd(0) >= 128 And ipAdd(0) <= 191 Then
            Worksheets("Sheet1").Cells(20, 4).Value = "Class B"
            Worksheets("Sheet1").Cells(21, 4).Value = "255.255.0.0"
            Worksheets("Sheet1").Cells(17, 2).Value = 16
        ElseIf ipAdd(0) >= 192 And ipAdd(0) <= 223 Then
            Worksheets("Sheet1").Cells(20, 4).Value = "Class C"
            Worksheets("Sheet1").Cells(21, 4).Value = "255.255.255.0"
            Worksheets("Sheet1").Cells(17, 2).Value = 24
        End If
    
        ' Interesting Octent
        If nonSubnetBits < 8 And nonSubnetBits >= 0 Then
            Worksheets("Sheet1").Cells(17, 4).Value = 1
        ElseIf nonSubnetBits < 16 And nonSubnetBits >= 8 Then
            Worksheets("Sheet1").Cells(17, 4).Value = 2
        ElseIf nonSubnetBits < 24 And nonSubnetBits >= 16 Then
            Worksheets("Sheet1").Cells(17, 4).Value = 3
        ElseIf nonSubnetBits <= 32 And nonSubnetBits >= 24 Then
            Worksheets("Sheet1").Cells(17, 4).Value = 4
        End If
    
        
        ' Increment how much
        If nonSubnetBits < 8 And nonSubnetBits >= 0 Then
            Worksheets("Sheet1").Cells(17, 5).Value = pow(2 - nonSubnetBits)
        ElseIf nonSubnetBits < 16 And nonSubnetBits >= 8 Then
            Worksheets("Sheet1").Cells(17, 5).Value = pow(16 - nonSubnetBits)
        ElseIf nonSubnetBits < 24 And nonSubnetBits >= 16 Then
            Worksheets("Sheet1").Cells(17, 5).Value = pow(24 - nonSubnetBits)
        ElseIf nonSubnetBits < 32 And nonSubnetBits >= 24 Then
            Worksheets("Sheet1").Cells(17, 5).Value = pow(32 - nonSubnetBits)
        Else: Worksheets("Sheet1").Cells(17, 4).Value = "N / A"
        End If
    
        ' Number of subnet bits
        nonSubnetBits = Worksheets("Sheet1").Cells(17, 3)
        defaultBits = Worksheets("Sheet1").Cells(17, 2)
        remainBits = nonSubnetBits - defaultBits
        If remainBits >= 0 Then
            Worksheets("Sheet1").Cells(17, 6).Value = remainBits
        Else: Worksheets("Sheet1").Cells(17, 6).Value = "N / A"
        End If
    
        ' Subnetworks
        sunbnetworks = Worksheets("Sheet1").Cells(17, 6)
        If sunbnetworks >= 0 And sunbnetworks <= 32 Then
            Worksheets("Sheet1").Cells(17, 7).Value = pow(sunbnetworks)
        Else: Worksheets("Sheet1").Cells(17, 7).Value = "N / A"
        End If
    
        ' Host bits
        Worksheets("Sheet1").Cells(17, 8).Value = 32 - nonSubnetBits
        If nonSubnetBits < 32 Then
            temp = pow(Worksheets("Sheet1").Cells(17, 8))
            Worksheets("Sheet1").Cells(17, 9).Value = (temp - 2)
        Else: Worksheets("Sheet1").Cells(17, 9).Value = "N / A"
        End If
        
          
        ' Subnet mask
        temp1 = Worksheets("Sheet1").Cells(17, 3)
        For i = 0 To temp1 - 1
            If j = 8 Then
                j = 0
            End If
            temp2 = calcSubnet(j) + temp2
            calcTemp(k) = temp2
            If temp2 = 255 And j <= 8 Then
                temp2 = 0
                k = k + 1
            End If
            j = j + 1
        Next i
    
        For i = 0 To 3
            If i <= 2 Then
                temp3 = temp3 & (calcTemp(i) & ".")
            Else
                temp3 = temp3 & calcTemp(i)
            End If
        Next i
        Worksheets("Sheet1").Cells(22, 4).Value = temp3
    
                
        ipAdd = Split(Worksheets("Sheet1").Cells(12, 4), ".")
        reqValue = ipAdd((Worksheets("Sheet1").Cells(17, 4)) - 1)
        arrBits = dec2bin(reqValue)
        arr = Split(arrBits, " ")
        
        ' Network Address
        intOct = Worksheets("Sheet1").Cells(17, 4)
        tot = 0
        If intOct = 2 Then
            diff = Worksheets("Sheet1").Cells(17, 8) - 16
        ElseIf intOct = 3 Then
            diff = Worksheets("Sheet1").Cells(17, 8) - 8
        ElseIf intOct = 4 Then
            diff = Worksheets("Sheet1").Cells(17, 8)
        End If
        
        For i = 0 To 8
            If arr(i) = 1 And i <= (8 - diff) Then
                tot = calcSubnet(i - 1) + tot
            End If
        Next
        
        If intOct = 2 Then
            Worksheets("Sheet1").Cells(23, 4).Value = ((ipAdd(0) & ".") & (tot & ".")) & ("0.0")
        ElseIf intOct = 3 Then
            Worksheets("Sheet1").Cells(23, 4).Value = (ipAdd(0) & "." & ((ipAdd(1) & ".") & (tot & "."))) & ("0")
        ElseIf intOct = 4 Then
            Worksheets("Sheet1").Cells(23, 4).Value = (ipAdd(0) & "." & ((ipAdd(1) & ".") & (ipAdd(2) & "."))) & (tot)
        End If
        
        ' Broadcast address
        arr3 = arr
               
        j = 8
        For i = 0 To diff - 1
            arr3(j) = 1
            j = j - 1
        Next
        
        For i = 0 To 8
            If arr3(i) = 1 Then
                tot1 = calcSubnet(i - 1) + tot1
            End If
        Next
        
        If intOct = 1 Then
            Worksheets("Sheet1").Cells(24, 4).Value = (ipAdd(0) & ".255.255.255")
        ElseIf intOct = 2 Then
            Worksheets("Sheet1").Cells(24, 4).Value = ((ipAdd(0) & ".") & (tot1 & ".")) & "255.255"
        ElseIf intOct = 3 Then
            Worksheets("Sheet1").Cells(24, 4).Value = ((ipAdd(0) & ".") & (ipAdd(1) & ".") & tot1) & ".255"
        ElseIf intOct = 4 Then
            Worksheets("Sheet1").Cells(24, 4).Value = ((ipAdd(0) & ".") & (ipAdd(1) & ".") & (ipAdd(2) & ".")) & (tot1)
        End If
        
    End If
      
End Sub
