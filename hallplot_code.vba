'this code was developed collaboratively, and some samples were written by PETEX.
'to see my contribution, see where it says "'Author: Akmal Aulia"

Option Explicit
'Petroleum Experts Ltd - Open Server VBA Example

' These lines declare global variables
Dim X As String
Dim Server As Object
Dim Connected As Integer
Dim lErr As Long
Dim Command As String
Dim AppName As String
Dim OSString As String
Dim typeFluide As Long

Sub OpenFile_cka_20()
'Author: Akmal Aulia

Dim Path As String

Path = Application.GetOpenFilename()

Worksheets("CKA-20").Cells(2, 3).Value = Path

VBA.Shell "Explorer.exe " & Range("C2"), vbMaximizedFocus
    
End Sub

Sub OpenFile_cka_21()

Dim Path As String

Path = Application.GetOpenFilename()

Worksheets("CKA-21").Cells(2, 3).Value = Path

VBA.Shell "Explorer.exe " & Range("C2"), vbMaximizedFocus
    
End Sub

Sub Macro()

Dim NodePressure As Double
Dim LiquidRate As Double

Call Connect

NodePressure = Cells(5, 4).Value
LiquidRate = Cells(5, 3).Value

'set value to First Node Pressure
DoSet ("PROSPER.ANL.GRD.Pres"), NodePressure
'set value to Liquid Rate
DoSet ("PROSPER.ANL.GRD.Rate"), LiquidRate

Disconnect
    
End Sub

Sub Calculate_cka20()
'Author: Akmal Aulia
Dim NumCalcs As Integer
Dim MD As Variant
Dim BottomPress As Variant
Dim NodePressure As Double
Dim LiquidRate As Double
Dim i As Integer, j As Integer, last_row As Integer
Dim cum_vol As Double, cum_press As Double

Call Connect

'Set rRng = Range("E14")
'If IsEmpty(rRng.Value) Then
If IsEmpty(Cells(14, 5)) Then
    MsgBox "Nothing to compute."
Else

    If IsEmpty(Cells(15, 5)) Then
        last_row = 14
    Else
        last_row = Cells(14, 5).End(xlDown).Row
    End If
    
    'initiate cums
    cum_vol = 0
    cum_press = 0
    
    For i = 14 To last_row
        
        'gather input data
        NodePressure = Cells(i, 10).Value
        LiquidRate = Cells(i, 6).Value
        
        'set value to First Node Pressure
        DoSet ("PROSPER.ANL.GRD.Pres"), NodePressure
        
        'set value to Liquid Rate
        DoSet ("PROSPER.ANL.GRD.Rate"), LiquidRate
        
        'Perform gradient traverse calculation
        DoCmd ("PROSPER.ANL.GRD.CALC")
        
        'store total count of calculated result to a variable
        NumCalcs = DoGet("PROSPER.OUT.GRD.Results[0][0][0].Pres.COUNT") - 1

        'print out results in excel
        Cells(i, 13).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].MSD[" & CStr(NumCalcs) & "]")
        
        If LiquidRate < 0.0001 Then
            Cells(i, 14).Value = 0
        Else
            Cells(i, 14).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].PRES[" & CStr(NumCalcs) & "]")
        End If
        
        'print cumulatives
        Cells(i, 19) = cum_vol + Cells(i, 6) 'cumulative volume
        Cells(i, 20) = cum_press + (Cells(i, 14) - 1700)
        
        'update cumulatives
        cum_vol = Cells(i, 19)
        cum_press = Cells(i, 20)
        
    Next i
    
    MsgBox "Extraction from PROSPER completed for CKA-20."
End If

 
End Sub

Sub Calculate_cka20_v2()
'Author: Akmal Aulia
Dim NumCalcs As Integer
Dim MD As Variant
Dim BottomPress As Variant
Dim NodePressure As Double
Dim LiquidRate As Double
Dim i As Integer, j As Integer, last_row As Integer
Dim cum_vol As Double, cum_press As Double

Call Connect


If IsEmpty(Cells(14, 5)) Then
    MsgBox "Nothing to compute."
Else

    If IsEmpty(Cells(15, 5)) Then
        last_row = 14
    Else
        last_row = Cells(14, 5).End(xlDown).Row
    End If
    
    'initiate cums
    cum_vol = 0
    cum_press = 0
    
    For i = 14 To last_row
        
        'gather input data
        NodePressure = Cells(i, 8).Value
        LiquidRate = Cells(i, 6).Value
        
        'set value to First Node Pressure
        DoSet ("PROSPER.ANL.GRD.Pres"), NodePressure
        
        'set value to Liquid Rate
        DoSet ("PROSPER.ANL.GRD.Rate"), LiquidRate
        
        'Perform gradient traverse calculation
        DoCmd ("PROSPER.ANL.GRD.CALC")
        
        'store total count of calculated result to a variable
        NumCalcs = DoGet("PROSPER.OUT.GRD.Results[0][0][0].Pres.COUNT") - 1

        'print out results in excel
        Cells(i, 10).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].MSD[" & CStr(NumCalcs) & "]")
        
        If LiquidRate < 0.0001 Then
            Cells(i, 11).Value = 0
        Else
            Cells(i, 11).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].PRES[" & CStr(NumCalcs) & "]")
        End If
        
        'print cumulatives
        Cells(i, 13) = cum_vol + Cells(i, 6) 'cumulative volume
        Cells(i, 14) = cum_press + (Cells(i, 11) - 1700)
        
        'update cumulatives
        cum_vol = Cells(i, 13)
        cum_press = Cells(i, 14)
        
    Next i
    
    MsgBox "Extraction from PROSPER completed for CKA-20."
End If

 
End Sub


Sub Calculate_cka21()
'Author: Akmal Aulia
Dim NumCalcs As Integer
Dim MD As Variant
Dim BottomPress As Variant
Dim NodePressure As Double
Dim LiquidRate As Double
Dim i As Integer, j As Integer, last_row As Integer
Dim cum_vol As Double, cum_press As Double

Call Connect

'Set rRng = Range("E14")
'If IsEmpty(rRng.Value) Then
If IsEmpty(Cells(14, 5)) Then
    MsgBox "Nothing to compute."
Else
    If IsEmpty(Cells(15, 5)) Then
        last_row = 14
    Else
        last_row = Cells(14, 5).End(xlDown).Row
    End If
    
    For i = 14 To last_row
        
        'gather input data
        NodePressure = Cells(i, 11).Value
        LiquidRate = Cells(i, 7).Value
        
        'set value to First Node Pressure
        DoSet ("PROSPER.ANL.GRD.Pres"), NodePressure
        
        'set value to Liquid Rate
        DoSet ("PROSPER.ANL.GRD.Rate"), LiquidRate
        
        
        
        'Perform gradient traverse calculation
        DoCmd ("PROSPER.ANL.GRD.CALC")
        
        'store total count of calculated result to a variable
        NumCalcs = DoGet("PROSPER.OUT.GRD.Results[0][0][0].Pres.COUNT") - 1

        'print out results in excel
        Cells(i, 16).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].MSD[" & CStr(NumCalcs) & "]")
        


        If LiquidRate < 0.0001 Then
            Cells(i, 17).Value = 0
        Else
            Cells(i, 17).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].PRES[" & CStr(NumCalcs) & "]")
        End If
        
       
        
        'print cumulatives
        Cells(i, 22) = cum_vol + Cells(i, 7) 'cumulative volume
        Cells(i, 23) = cum_press + (Cells(i, 17) - 1700) 'cumulative pressure
        
        'update cumulatives
        cum_vol = Cells(i, 22)
        cum_press = Cells(i, 23)
        
    Next i
    
    MsgBox "Extraction from PROSPER completed for CKA-21."
End If

 
End Sub

Sub Calculate_cka21_v2()
'Author: Akmal Aulia
Dim NumCalcs As Integer
Dim MD As Variant
Dim BottomPress As Variant
Dim NodePressure As Double
Dim LiquidRate As Double
Dim i As Integer, j As Integer, last_row As Integer
Dim cum_vol As Double, cum_press As Double

Call Connect


If IsEmpty(Cells(14, 5)) Then
    MsgBox "Nothing to compute."
Else

    If IsEmpty(Cells(15, 5)) Then
        last_row = 14
    Else
        last_row = Cells(14, 5).End(xlDown).Row
    End If
    
    'initiate cums
    cum_vol = 0
    cum_press = 0
    
    For i = 14 To last_row
        
        'gather input data
        NodePressure = Cells(i, 8).Value
        LiquidRate = Cells(i, 6).Value
        
        'set value to First Node Pressure
        DoSet ("PROSPER.ANL.GRD.Pres"), NodePressure
        
        'set value to Liquid Rate
        DoSet ("PROSPER.ANL.GRD.Rate"), LiquidRate
        
        'Perform gradient traverse calculation
        DoCmd ("PROSPER.ANL.GRD.CALC")
        
        'store total count of calculated result to a variable
        NumCalcs = DoGet("PROSPER.OUT.GRD.Results[0][0][0].Pres.COUNT") - 1

        'print out results in excel
        Cells(i, 10).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].MSD[" & CStr(NumCalcs) & "]")
        
        If LiquidRate < 0.0001 Then
            Cells(i, 11).Value = 0
        Else
            Cells(i, 11).Value = DoGet("PROSPER.OUT.GRD.Results[0][0][0].PRES[" & CStr(NumCalcs) & "]")
        End If
        
        'print cumulatives
        Cells(i, 13) = cum_vol + Cells(i, 6) 'cumulative volume
        Cells(i, 14) = cum_press + (Cells(i, 11) - 1700)
        
        'update cumulatives
        cum_vol = Cells(i, 13)
        cum_press = Cells(i, 14)
        
    Next i
    
    MsgBox "Extraction from PROSPER completed for CKA-21."
End If

 
End Sub

Sub Connect() 'This utility creates the OpenServer object which allows comunication between Excel and IPM tools
    
    If Connected = 0 Then
        Set Server = CreateObject("PX32.OpenServer.1")
        Connected = 1
    End If

End Sub
Sub Disconnect()
    
    If Connected = 1 Then
       Set Server = Nothing
       Connected = 0
    End If

End Sub

' This utility function extracts the application name from the tag string
Function GetAppName(Strval As String) As String
   Dim Pos
   Pos = InStr(Strval, ".")
   If Pos < 2 Then
        MsgBox "Badly formed tag string"
        End
   End If
   GetAppName = Left(Strval, Pos - 1)
  
End Function
' Perform a command, then check for errors
Sub DoCmd(Cmd As String)
    Dim lErr As Long
    lErr = Server.DoCommand(Cmd)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Set Server = Nothing
        End
    End If
End Sub
'Set a value, then check for errors
Sub DoSet(Sv As String, Val)
    Dim lErr As Long
    lErr = Server.SetValue(Sv, Val)
    AppName = GetAppName(Sv)
    lErr = Server.GetLastError(AppName)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Set Server = Nothing
        End
    End If
End Sub
' Get a value, then check for errors
Function DoGet(Gv As String) As String
    Dim lErr As Long
    DoGet = Server.GetValue(Gv)
    AppName = GetAppName(Gv)
    lErr = Server.GetLastError(AppName)
    If lErr > 0 Then
        MsgBox Server.GetLastErrorMessage(AppName)
        Set Server = Nothing
        End
    End If
End Function
' Perform a command, then wait for the command to exit
' Then check for errors
Sub DoSlowCmd(Cmd As String)
    Dim starttime As Single
    Dim endtime As Single
    Dim CurrentTime As Single
    Dim lErr As Long
    Dim bLoop As Boolean
    Dim step As Single
        
    step = 0.001
    AppName = GetAppName(Cmd)
    lErr = Server.DoCommandAsync(Cmd)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Disconnect
        End
    End If
    While Server.IsBusy(AppName) > 0
        If step < 2 Then
            step = step * 2
        End If
        starttime = Timer
        endtime = starttime + step
        Do
            CurrentTime = Timer
            'DoEvents
            bLoop = True
            Rem Check first for the case where we have gone over midnight
            Rem and the number of seconds will go back to zero
            If CurrentTime < starttime Then
                bLoop = False
            Rem Now check for the 2 second pause finishing
            ElseIf CurrentTime > endtime Then
                bLoop = False
            End If
        Loop While bLoop
    Wend
    AppName = GetAppName(Cmd)
    lErr = Server.GetLastError(AppName)
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        Disconnect
        End
    End If
End Sub


' Perform a function in GAP, then retrieve return value
' Finally, check for errors
Function DoGAPFunc(Gv As String) As String
    DoSlowCmd Gv
    DoGAPFunc = DoGet("GAP.LASTCMDRET")
    lErr = Server.GetLastError("GAP")
    If lErr > 0 Then
        MsgBox Server.GetErrorDescription(lErr)
        End
    End If
End Function

Sub Pulling()
'Author: Akmal Aulia
'declare variables
Dim sqlStr2 As String
Dim i As Integer
Dim wnam(130) As String

'declare connection variables
Dim oConn2 As ADODB.Connection
Dim rs2 As ADODB.Recordset
Dim sConnString2 As String

''declare worksheet variable
'Dim wdat As Worksheet, wc As Worksheet
'
''set worksheets
'Set wdat = Sheets("WMI Input")
'Set wc = Sheets("CO2 Real Time")

'create connection string
sConnString2 = "Provider=******;data source=******;Initial Catalog=******;User Id=******;Password=******"

'create the connection and recordset objects
Set oConn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset

'Open connection and execute
oConn2.Open sConnString2

'NOTE: date in yyyy-mm-dd
'sqlStr2 = "SELECT GHV FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
'sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK20 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) = '2020-01-18';"
'sqlStr2 = "SELECT  cast(timestamp as date) AS dates,VOL_INJ_CK20,VOL_INJ_CK21,INJ_BP_CK20,INJ_BP_CK21 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) between '2019-01-01' and '2019-05-01'"

Dim last_date As String, last_row As Integer
'last_date = "2019-01-01"
last_date = Cells(10, 6)
MsgBox "Request date: " & last_date

'Set rRng = Range("G6")
If IsEmpty(Cells(14, 5)) Then
    sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK20,VOL_INJ_CK21,INJ_BP_CK20,INJ_BP_CK21 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) >='" & last_date & "';"
    Set rs2 = oConn2.Execute(sqlStr2)
    Cells(14, 5).CopyFromRecordset rs2
Else
    'grab date from last row
    last_date = Cells(Rows.Count, 5).End(xlUp)
    
    'grab last row index
    last_row = Cells(14, 5).End(xlDown).Row
    MsgBox "Current last row = " & last_row
    
    sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK20,VOL_INJ_CK21,INJ_BP_CK20,INJ_BP_CK21 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) >'" & last_date & "';"
    Set rs2 = oConn2.Execute(sqlStr2)
    Cells(last_row + 1, 5).CopyFromRecordset rs2 'copy value to the empty row
    
End If

'convert pressure units from [bar] to [psig]

If IsEmpty(Cells(14, 5)) Then
    MsgBox "No data for this request date!"
Else
    last_row = Cells(14, 5).End(xlDown).Row
    For i = 14 To last_row
        Cells(i, 10) = Cells(i, 8) * 14.5 'pressure, CKA-20
        Cells(i, 11) = Cells(i, 9) * 14.5 'pressure, CKA-21
        
        'handle values below zero
        If Cells(i, 10) < 0 Then
            Cells(i, 10) = 0
        End If
        
        If Cells(i, 11) < 0 Then
            Cells(i, 11) = 0
        End If
    Next i
End If


End Sub

Sub copy_to_container()
'Author: Akmal Aulia
Dim first_row As Integer, maxit As Integer, j As Integer
Dim i As Integer, cnt As Integer, num_tab As Integer, shift_row As Integer


first_row = 4 'location of column names
maxit = 5000 'any sufficiently large number


num_tab = 14
shift_row = 4

For j = 1 To num_tab

    If j = 1 Then
    
        For i = 1 To maxit
        
            If Cells(first_row + i, 1) = "NA" Then
                'MsgBox "Row number " & first_row + i & " is empty."
                Exit For
            End If
            
            'copy to containers
            If Cells(first_row + i, 1) <= Cells(3, 5) Then
                'to 2020 containers
                Cells(first_row + i, 5) = Cells(first_row + i, 1)
                Cells(first_row + i, 6) = Cells(first_row + i, 2)
                Cells(first_row + i, 7) = Cells(first_row + i, 3)
            End If
    
        Next i
    
    End If
    
    
    If j > 1 Then
    
        cnt = 0
        For i = 1 To maxit
    
            If Cells(first_row + i, 1) = "NA" Then
                'MsgBox "Row number " & first_row + i & " is empty."
                Exit For
            End If
    
            'copy to containers
            If Cells(first_row + i, 1) <= Cells(3, 5 + ((j - 1) * shift_row)) And Cells(first_row + i, 1) > Cells(3, 5 + ((j - 2) * shift_row)) Then
                cnt = cnt + 1
                Cells(first_row + cnt, 5 + ((j - 1) * shift_row)) = Cells(first_row + i, 1)
                Cells(first_row + cnt, 6 + ((j - 1) * shift_row)) = Cells(first_row + i, 2)
                Cells(first_row + cnt, 7 + ((j - 1) * shift_row)) = Cells(first_row + i, 3)
            End If
    
        Next i
    
    End If

Next j

End Sub


Sub Pulling_cka20()
'Author: Akmal Aulia
'declare variables
Dim sqlStr2 As String
Dim i As Integer
Dim wnam(130) As String

'declare connection variables
Dim oConn2 As ADODB.Connection
Dim rs2 As ADODB.Recordset
Dim sConnString2 As String

''declare worksheet variable
'Dim wdat As Worksheet, wc As Worksheet
'
''set worksheets
'Set wdat = Sheets("WMI Input")
'Set wc = Sheets("CO2 Real Time")

'create connection string
sConnString2 = "Provider=******;data source=******;Initial Catalog=******;User Id=******;Password=******"

'create the connection and recordset objects
Set oConn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset

'Open connection and execute
oConn2.Open sConnString2

'NOTE: date in yyyy-mm-dd
'sqlStr2 = "SELECT GHV FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
'sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK20 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) = '2020-01-18';"
'sqlStr2 = "SELECT  cast(timestamp as date) AS dates,VOL_INJ_CK20,VOL_INJ_CK21,INJ_BP_CK20,INJ_BP_CK21 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) between '2019-01-01' and '2019-05-01'"

Dim last_date As String, last_row As Integer
'last_date = "2019-01-01"
last_date = Cells(10, 6)
MsgBox "Request date: " & last_date



'Set rRng = Range("G6")
If IsEmpty(Cells(14, 5)) Then
    sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK20,INJ_BP_CK20 FROM DataMining.dbo.vIPRIZ_WaterInj where VOL_INJ_CK20 > 0 and cast(timestamp as date) >='" & last_date & "';"
    Set rs2 = oConn2.Execute(sqlStr2)
    Cells(14, 5).CopyFromRecordset rs2
Else
    'grab date from last row
    last_date = Cells(Rows.Count, 5).End(xlUp)
    
    'grab last row index
    last_row = Cells(14, 5).End(xlDown).Row
    MsgBox "Current last row = " & last_row
    
    sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK20,INJ_BP_CK20 FROM DataMining.dbo.vIPRIZ_WaterInj where VOL_INJ_CK20 > 0 and cast(timestamp as date) >'" & last_date & "';"
    Set rs2 = oConn2.Execute(sqlStr2)
    Cells(last_row + 1, 5).CopyFromRecordset rs2 'copy value to the empty row
    
End If

'convert pressure units from [bar] to [psig]

If IsEmpty(Cells(14, 5)) Then
    MsgBox "No data for this request date!"
Else
    last_row = Cells(14, 5).End(xlDown).Row
    For i = 14 To last_row
        Cells(i, 8) = Cells(i, 7) * 14.5 'pressure, CKA-20
        'Cells(i, 11) = Cells(i, 9) * 14.5 'pressure, CKA-21
        
        'handle values below zero
        If Cells(i, 8) < 0 Then
            Cells(i, 8) = 0
        End If
        
        'If Cells(i, 11) < 0 Then
        '    Cells(i, 11) = 0
        'End If
    Next i
End If


End Sub

Sub Pulling_cka21()
'Author: Akmal Aulia
'declare variables
Dim sqlStr2 As String
Dim i As Integer
Dim wnam(130) As String

'declare connection variables
Dim oConn2 As ADODB.Connection
Dim rs2 As ADODB.Recordset
Dim sConnString2 As String

''declare worksheet variable
'Dim wdat As Worksheet, wc As Worksheet
'
''set worksheets
'Set wdat = Sheets("WMI Input")
'Set wc = Sheets("CO2 Real Time")

'create connection string
sConnString2 = "Provider=******;data source=******;Initial Catalog=******;User Id=******;Password=******"

'create the connection and recordset objects
Set oConn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset

'Open connection and execute
oConn2.Open sConnString2

'NOTE: date in yyyy-mm-dd
'sqlStr2 = "SELECT GHV FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
'sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK20 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) = '2020-01-18';"
'sqlStr2 = "SELECT  cast(timestamp as date) AS dates,VOL_INJ_CK20,VOL_INJ_CK21,INJ_BP_CK20,INJ_BP_CK21 FROM DataMining.dbo.vIPRIZ_WaterInj where cast(timestamp as date) between '2019-01-01' and '2019-05-01'"

Dim last_date As String, last_row As Integer
'last_date = "2019-01-01"
last_date = Cells(10, 6)
MsgBox "Request date: " & last_date



'Set rRng = Range("G6")
If IsEmpty(Cells(14, 5)) Then
    sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK21,INJ_BP_CK21 FROM DataMining.dbo.vIPRIZ_WaterInj where VOL_INJ_CK21 > 0 and cast(timestamp as date) >='" & last_date & "';"
    Set rs2 = oConn2.Execute(sqlStr2)
    Cells(14, 5).CopyFromRecordset rs2
Else
    'grab date from last row
    last_date = Cells(Rows.Count, 5).End(xlUp)
    
    'grab last row index
    last_row = Cells(14, 5).End(xlDown).Row
    MsgBox "Current last row = " & last_row
    
    sqlStr2 = "SELECT cast(timestamp as date) AS dates, VOL_INJ_CK21,INJ_BP_CK21 FROM DataMining.dbo.vIPRIZ_WaterInj where VOL_INJ_CK21 > 0 and cast(timestamp as date) >'" & last_date & "';"
    Set rs2 = oConn2.Execute(sqlStr2)
    Cells(last_row + 1, 5).CopyFromRecordset rs2 'copy value to the empty row
    
End If

'convert pressure units from [bar] to [psig]

If IsEmpty(Cells(14, 5)) Then
    MsgBox "No data for this requested date."
Else
    last_row = Cells(14, 5).End(xlDown).Row
    For i = 14 To last_row
        Cells(i, 8) = Cells(i, 7) * 14.5 'pressure, CKA-20
        'Cells(i, 11) = Cells(i, 9) * 14.5 'pressure, CKA-21
        
        'handle values below zero
        If Cells(i, 8) < 0 Then
            Cells(i, 8) = 0
        End If
        
        'If Cells(i, 11) < 0 Then
        '    Cells(i, 11) = 0
        'End If
    Next i
End If


End Sub


Sub Clearing()
'Author: Akmal Aulia
Dim i As Integer, j As Integer, last_row As Integer
Dim last_col As Integer

last_col = 23

If IsEmpty(Cells(14, 5)) Then
    MsgBox "Nothing to delete."
Else
    last_row = Cells(14, 5).End(xlDown).Row
    For i = 14 To last_row
        For j = 5 To last_col
            Cells(i, j).ClearContents
        Next j
    Next i
    
    MsgBox "Data deleted."
        
End If

End Sub



Sub delete_rows()
'Author: Akmal Aulia
    'Dim i As Integer, str_arg As String
    'Dim cum_rows As Integer, last_row As Integer
    
    'i = 14
    
    
    Dim i As Integer, str_arg As String
    Dim nflag As Integer 'nflag = 0 if empty, 1 otherwise
    Dim cum_rows As Integer
    
    i = 14
    
    If IsEmpty(Cells(14, 5)) Then
        nflag = 0
    Else
        nflag = 1
    End If
    
    
    
    Do While nflag = 1
        If Cells(i, 6) < 0.001 Then
            'remove row
            str_arg = "E" & i & ":W" & i
            Range(str_arg).Select
            Selection.Delete Shift:=xlUp
            cum_rows = cum_rows + 1
            
        Else
            i = i + 1
            If IsEmpty(Cells(i, 6)) Then
                nflag = 0
            End If
        End If

    Loop
    
    


    
    
    MsgBox cum_rows & " rows were removed."


End Sub
