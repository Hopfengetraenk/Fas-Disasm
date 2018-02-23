Attribute VB_Name = "CRC_Adler32"
Public Function ADLER32$(data) ' As StringReader)
   Dim mydata() As Byte
   With data
'            Dim a
            
            Dim L&, H&
            H = 0: L = 1
'            a = GetTickCount
' taken out for performance reason
'               .EOS = False
'               .DisableAutoMove = False
'               Do Until .EOS
'                 'The largest prime less than 2^16
'                  l = (.int8 + l) Mod 65521 '&HFFF1
'                  H = (H + l) Mod 65521 '&HFFF1
'                  If (l And 8) Then myDoEvents
'               Loop
'
'            Debug.Print "a: ", GetTickCount - a 'Benchmark: 20203

 '           a = GetTickCount
               
               Dim StrCharPos&, tmpBuff() As Byte
               tmpBuff = UToA(data)
'               tmpBuff = .mvardata
               For StrCharPos = 0 To UBound(tmpBuff)
                  'The largest prime less than 2^16
                  L = (tmpBuff(StrCharPos) + L) Mod &HFFF1&  '65521
                  H = (H + L) Mod &HFFF1& '65521
                  
                  'If 0 = (StrCharPos Mod &H8000) Then myDoEvents

               Next
'            Debug.Print "b: ", GetTickCount - a 'Benchmark: 5969

      ADLER32 = H16(H) & H16(L)
   End With
End Function
