Attribute VB_Name = "CONSOLIDATION"
Option Explicit

' Summary:
' Handles uploading of data to the SQL server.
'
' Compatibility:
' 32-bit and 64-bit.
'
' Dependencies:
' StringBuilder class
'
' Remarks:


' Summary:
' Entry point procedure to control uploading of the data to SQL server from the workbook.
'
' Parameters:
' None.
'
' Returns:
' Nothing.
'
' Remarks:
' This method relies on a range name (Settings.ConsolidationUploadConfig) containing the configuration for upload.
' It loops through each entry and calls the UploadConfigEntry method to transform the source data and send it to the appropriate stored procedure.
'
' For performance reasons, we upload in batches. We can specify the batch size just by adjusting the number. Primitive, but it works!
' I haven't tested what batch size is the most performant, but 1500 works nicely enough.
'
Public Sub UploadData()

'10        On Error GoTo UploadData_Error

'10        Call Refresh_01_TableTypes

ActiveWorkbook.Connections("Connection_UploadConfig").Refresh

20        ActiveWorkbook.Connections("Connection_types_Koro_live").Refresh
          ActiveWorkbook.Connections("Connection_types_Koro_live_nkey").Refresh
          ActiveWorkbook.Connections("Connection_types_ip_live").Refresh
          ActiveWorkbook.Connections("Connection_types_central_forecast").Refresh
        

          'Debug.Print "Started upload process. Time is: "; Now
          '
          Dim cnSQLServer As ADODB.Connection
30        Set cnSQLServer = New ADODB.Connection
'40        cnSQLServer.Open "Provider=SQLOLEDB;Data Source=EUBASCIRDB01;Initial Catalog=D2C_PLAN;Integrated Security=SSPI"
40        cnSQLServer.Open "Provider=SQLOLEDB;Data Source=biseappdb.eu.sony.com;Initial Catalog=D2C_PLAN;Integrated Security=SSPI"

50        If [Table_UploadConfig].Rows.count = 0 Then Exit Sub

          Dim uploadConfigEntry As Range
          Dim batches As Integer

          Dim batchSize As Integer
60        batchSize = 500

          Dim van_e_error As Boolean
          Dim error_szoveg As String
70        van_e_error = False
80        For Each uploadConfigEntry In [Table_UploadConfig].Rows

              Dim uploadTable As Range
              Dim typeTable As Range
              Dim i As Integer
              Dim totalUploaded As Long
              Dim offSet As Integer

              Dim startCmd As ADODB.Command
              Dim endCmd As ADODB.Command

90            Set uploadTable = UploadConfig.Evaluate(uploadConfigEntry.Cells(, 1).value)
100           Set typeTable = UploadConfig.Evaluate(uploadConfigEntry.Cells(, 2).value)

110           batches = Ceil(uploadTable.Rows.count / batchSize)

120           totalUploaded = 1

              'Start cmd to execute before anything else?
130           If Not IsEmpty(uploadConfigEntry.Cells(, 5)) Then

140               Set startCmd = New ADODB.Command
150               With startCmd
160                   .CommandType = adCmdStoredProc
170                   .ActiveConnection = cnSQLServer
180                   .CommandText = (uploadConfigEntry.Cells(, 5).value)
190                   .Execute
200               End With


210           Else
                  'Debug.Print "first"
220               cnSQLServer.BeginTrans

230           End If

              'Upload each batch
240           For i = 1 To batches

                  Dim batchRange As Range
                  Dim endOfBatchRange As Long

                  'Range may end somewhere different if we are at the end of the upload process
250               If i = batches Then
260                   endOfBatchRange = uploadTable.Rows.count
270               Else
280                   endOfBatchRange = (totalUploaded + batchSize) - 1
290               End If

                  'Set upload range for this batch.
300               Set batchRange = uploadTable.Worksheet.Range("" & uploadTable.Cells(totalUploaded, 1).Address() & ":" & uploadTable.Cells(endOfBatchRange, uploadTable.Columns.count).Address())

                  'Have to do this AFTER we make up a range.
310               totalUploaded = totalUploaded + batchSize

                  'Debug.Print "About to upload batch " & i
                  '''''
                  'Debug.Print GetUploadDataSQLString(batchRange, typeTable)

                  Dim errorTest
320               errorTest = 1

                  'On Error GoTo errorcmd

330               WriteToServer cnSQLServer, uploadConfigEntry.Cells(, 3).value, uploadConfigEntry.Cells(, 4).value, _
                                GetUploadDataSQLString(batchRange, typeTable)



                  'batchRange.Select

                  'Debug.Print "Uploaded batch " & i

340           Next i

              'End cmd to execute after anything else?
350           If Not IsEmpty(uploadConfigEntry.Cells(, 6)) And errorTest = 1 Then

                  'If Not IsEmpty(uploadConfigEntry.Cells(, 6)) Then

360               Set endCmd = New ADODB.Command
370               With endCmd: .CommandTimeout = 500
380                   .CommandType = adCmdStoredProc
390                   .ActiveConnection = cnSQLServer
400                   .CommandText = (uploadConfigEntry.Cells(, 6).value)
410                   .Execute
420               End With

430           Else
                  'Debug.Print "last"
440               cnSQLServer.CommitTrans

450           End If


460       Next uploadConfigEntry



          'Debug.Print "Finished upload process. Time is: "; Now

470       cnSQLServer.Close


480       On Error GoTo 0
490       Exit Sub

UploadData_Error:

'500       m_name = "CONSOLIDATION"
'510       c_name = "UploadData"
'520       error_line_number = Erl
'530       Call Handling_Those_Errors(m_name, c_name, error_line_number)

End Sub
' Summary:
' Gets the ceiling integer of a double.
'
' Parameters:
' [num] In: Double to take ceiling of
'
' Returns:
' Ceiling of the given number.
'
' Remarks:
' e.g. Ceil(6.4) = 7
Public Function Ceil(ByVal num As Double) As Integer
      'Trickery
10        Ceil = (Int(num / 1) - (num - Int(num / 1) > 0))
End Function

' Summary:
' Handles writing a series of select statements (containing the data to send) to an open SQL server connection, using table-valued parameters.
'
' Parameters:
' [con] In: a reference to an open ADODB connection to a DB server.
' [tableTypeName] In: a string containing the name of a table type on the SQL server that will be used for the TVP process.
' [storedProcName] In: a string containing the name of a stored procedure on the SQL server that will handle the update/insert using TVP and the
' tableTypeName table type.
' [dataSelectStatement] In: a string containing a series of SELECT statements which contain the data to be updated/inserted and passed to the stored
' procedure.
' [timeout] In: an integer specifying the timeout in seconds for the command's execution. Default is 10 minutes (600 seconds).
'
' Returns:
' Nothing.
'
' Remarks:
' Normally, the SQL commands should be fully parameterised using the command object to avoid SQL injection attacks, however, this is not
' simple to do with traditional ADO when using adCmdTexts and DECLARE statements as used below.
Private Sub WriteToServer(ByRef con As ADODB.Connection, ByRef tableTypeName As String, ByRef storedProcName As String, ByRef dataSelectStatement As String, Optional ByVal timeout As Integer = 600)
10        If con Is Nothing Then Err.Raise vbObjectError + 128, "WriteToServer", "Database connection cannot be null."

          Dim cmd As ADODB.Command
20        Set cmd = New ADODB.Command

30        Set cmd.ActiveConnection = con
40        cmd.CommandType = adCmdText
50        cmd.Prepared = True

60        cmd.CommandText = "DECLARE @ReturnMessage nvarchar(255);" _
                          + "DECLARE @ConsolidationTableType " & tableTypeName & ";" _
                          + "INSERT INTO @ConsolidationTableType " _
                          + "EXEC ('" _
                          + dataSelectStatement _
                          + "') " _
                          + "EXEC " & storedProcName & " @ConsolidationTableType, @ReturnMessage OUTPUT;" _
                          + "SELECT @ReturnMessage as ReturnMessage"

70        cmd.CommandTimeout = timeout

          'Debug.Print cmd.CommandText

80        cmd.Execute Options:=adExecuteNoRecords
End Sub

' Summary:
' Creates a string containing a series of SELECT statements (one for each data row) containing the data to be later passed to a stored procedure in the
' database.
'
' Parameters:
' [sourceData] In: a reference to an Excel range which contains the data to be used. This range should exclude any header rows.
' [sourceDataTypes] In: a reference to an Excel range which specifies the data types for each field in the sourceData range.
' These are specified using SQL types, e.g. VARCHAR
'
' Returns:
' A string containing the concatenated SELECT statements with the data from sourceData.
'
' Remarks:
' This method uses the StringBuilder class to append the strings, since this is much faster than using traditional string concatenation.
'
' The data type for each field is determined from the sourceDataTypes parameter, which is normally a 2D range, therefore, the column which is
' used to determine the field type is controlled by the DATATYPE_ID_COL constant.
'
' There is currently no escaping performed for the source data, therefore, any row data with quotes or single quotes will probably cause mangled data
' (due to mismatched quotes) to be returned. This could then cause problems if the string is sent to a stored procedure.
'
' NB. This method currently ignores the last two columns of the sourceData range as these are not required to be uploaded, change this as appropriate.

'2019-11-12 - due to slow down on some machines (maybe due to missing 2017 x64 C++ redistributable?), rewritten to use simple strings. Matt White

Private Function GetUploadDataSQLString(ByRef sourceData As Excel.Range, ByRef sourceDataTypes As Excel.Range) As String

          Const DT_FLOAT As Integer = 5   ' SQL's internal data type ID for FLOAT
          Const DT_VARCHAR As Integer = 130    ' SQL's internal data type ID for VARCHAR
          Const DATATYPE_ID_COL As Integer = 3    ' column number of sourceDataTypes that contains the actual field type

10        Dim data As Variant: data = sourceData
20        Dim dataTypes As Variant: dataTypes = sourceDataTypes

          Dim i As Long
          Dim j As Long
          Dim strSQLString As String
            strSQLString = ""
          
'30        Dim sb As StringBuilder: Set sb = New StringBuilder

40        For i = LBound(data, 1) To UBound(data, 1)
'50            sb.Append " SELECT "
50            strSQLString = strSQLString & " SELECT "

              ' ignore the last two columns of the source data as these are not required to be uploaded.
60            For j = LBound(data, 2) To UBound(data, 2)
                  ' Float or integer field values should be left unquoted ('')
70                If CInt(dataTypes(j, DATATYPE_ID_COL)) = DT_FLOAT Then

                      ' if the field type is float, and the cell is empty, then force it to be zero.
80                    If Len(CStr(data(i, j))) = 0 Then
90                         'sb.Append "0,"
                            strSQLString = strSQLString & "0,"
100                   Else
110                        'sb.Append Replace(CStr(data(i, j)), ",", ".") & ","
                            strSQLString = strSQLString & Replace(CStr(data(i, j)), ",", ".") & ","
120                   End If
                      ' Otherwise strings should be quoted using single quotes (e.g. ''Hello world'').
130               ElseIf CInt(dataTypes(j, DATATYPE_ID_COL)) = DT_VARCHAR Then
140                        'sb.Append "''" & Left(CStr(data(i, j)), 50) & "'',"
                            strSQLString = strSQLString & "''" & Replace(Left(CStr(data(i, j)), 50), "'", "") & "'',"  'replaces quotes within the words
150               End If
160           Next j

170                        'sb.Length = sb.Length - 1    'remove the last comma from the "row"
                            strSQLString = Left(strSQLString, Len(strSQLString) - 1) 'remove the last comma from the "row"
180       Next i

190       'GetUploadDataSQLString = sb.ToString
           GetUploadDataSQLString = strSQLString
End Function

'Sub C_1_DataGL()
'End Sub
'
'Sub C_1_DataFB()
'End Sub
'
'Sub C_2_DB()
'End Sub
'
'Sub C_3_DB()
'End Sub
'
'Sub C_6_DB()
'End Sub
'
'Sub C_6a_DB()
'End Sub
'
'Sub C_6b_DB()
'End Sub
'
'Sub C_9_DB()
'End Sub
'
'Sub C_8_DB()
'End Sub



'-------------END-------------
Public Sub UploadData_cf()

'10        On Error GoTo UploadData_Error

'10        Call Refresh_01_TableTypes

ActiveWorkbook.Connections("Connection_UploadConfig").Refresh

20        ActiveWorkbook.Connections("Connection_types_Koro_live").Refresh
          ActiveWorkbook.Connections("Connection_types_Koro_live_nkey").Refresh
          ActiveWorkbook.Connections("Connection_types_ip_live").Refresh
          ActiveWorkbook.Connections("Connection_types_central_forecast").Refresh
        

          'Debug.Print "Started upload process. Time is: "; Now
          '
          Dim cnSQLServer As ADODB.Connection
30        Set cnSQLServer = New ADODB.Connection
'40        cnSQLServer.Open "Provider=SQLOLEDB;Data Source=EUBASCIRDB01;Initial Catalog=D2C_PLAN;Integrated Security=SSPI"
40        cnSQLServer.Open "Provider=SQLOLEDB;Data Source=biseappdb.eu.sony.com;Initial Catalog=D2C_PLAN;Integrated Security=SSPI"

50        If [Table_UploadConfig].Rows.count = 0 Then Exit Sub

          Dim uploadConfigEntry As Range
          Dim batches As Integer

          Dim batchSize As Integer
60        batchSize = 500

          Dim van_e_error As Boolean
          Dim error_szoveg As String
70        van_e_error = False

              Dim uploadTable As Range
              Dim typeTable As Range
              Dim i As Integer
              Dim totalUploaded As Long
              Dim offSet As Integer

              Dim startCmd As ADODB.Command
              Dim endCmd As ADODB.Command
              
            With Worksheets("raw_extract_input_query")
            Set uploadTable = .Range("A2:I2", .Range("B" & .Rows.count).End(xlUp))
            End With
                
            With Worksheets("UploadConfig")
            Set typeTable = .Range("AR2:AU2", .Range("AR" & .Rows.count).End(xlUp))
            End With
                
90
100

110           batches = Ceil(uploadTable.Rows.count / batchSize)

120           totalUploaded = 1

              'Start cmd to execute before anything else?
130           If Not IsEmpty(UploadConfig.Cells(6, 5)) Then

140               Set startCmd = New ADODB.Command
150               With startCmd
160                   .CommandType = adCmdStoredProc
170                   .ActiveConnection = cnSQLServer
180                   .CommandText = UploadConfig.Cells(6, 5).value
190                   .Execute
200               End With


210           Else
                  'Debug.Print "first"
220               cnSQLServer.BeginTrans

230           End If

              'Upload each batch
240           For i = 1 To batches

                  Dim batchRange As Range
                  Dim endOfBatchRange As Long

                  'Range may end somewhere different if we are at the end of the upload process
250               If i = batches Then
260                   endOfBatchRange = uploadTable.Rows.count
270               Else
280                   endOfBatchRange = (totalUploaded + batchSize) - 1
290               End If

                  'Set upload range for this batch.
300               Set batchRange = uploadTable.Worksheet.Range("" & uploadTable.Cells(totalUploaded, 1).Address() & ":" & uploadTable.Cells(endOfBatchRange, uploadTable.Columns.count).Address())

                  'Have to do this AFTER we make up a range.
310               totalUploaded = totalUploaded + batchSize

                  'Debug.Print "About to upload batch " & i
                  '''''
                  'Debug.Print GetUploadDataSQLString(batchRange, typeTable)

                  Dim errorTest
320               errorTest = 1

                  'On Error GoTo errorcmd

330               WriteToServer cnSQLServer, UploadConfig.Cells(6, 3).value, UploadConfig.Cells(6, 4).value, _
                                GetUploadDataSQLString(batchRange, typeTable)



                  'batchRange.Select

                  'Debug.Print "Uploaded batch " & i

340           Next i

              'End cmd to execute after anything else?
350           If Not IsEmpty(UploadConfig.Cells(6, 6)) And errorTest = 1 Then

                  'If Not IsEmpty(uploadConfigEntry.Cells(, 6)) Then

360               Set endCmd = New ADODB.Command
370               With endCmd: .CommandTimeout = 500
380                   .CommandType = adCmdStoredProc
390                   .ActiveConnection = cnSQLServer
400                   .CommandText = UploadConfig.Cells(6, 6).value
410                   .Execute
420               End With

430           Else
                  'Debug.Print "last"
440               cnSQLServer.CommitTrans

450           End If


460



          'Debug.Print "Finished upload process. Time is: "; Now

470       cnSQLServer.Close


480       On Error GoTo 0
490       Exit Sub

UploadData_Error:

'500       m_name = "CONSOLIDATION"
'510       c_name = "UploadData"
'520       error_line_number = Erl
'530       Call Handling_Those_Errors(m_name, c_name, error_line_number)

End Sub


