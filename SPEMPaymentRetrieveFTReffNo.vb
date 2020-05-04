Imports System.Data

Public Class SPEMPaymentRetrieveFTReffNo
    ''' <summary>
    ''' It's like enumeration method.
    ''' </summary>
    ''' <remarks>0 = No Data, 1 = Have a Data, 2 = Failed or Error</remarks>
    Private _statusCode As Integer
    Public Property statusCode() As Integer
        Get
            Return _statusCode
        End Get
        Set(value As Integer)
            _statusCode = value
        End Set
    End Property

    Private _statusMessage As String
    Public Property statusMessage() As String
        Get
            Return _statusMessage
        End Get
        Set(value As String)
            _statusMessage = value
        End Set
    End Property

    Private _statusProcessing As Boolean
    Public Property statusProcessing() As String
        Get
            Return _statusProcessing
        End Get
        Set(value As String)
            _statusProcessing = value
        End Set
    End Property

    Private _informToAdministrator As Boolean
    Public Property informToAdministrator() As String
        Get
            Return _informToAdministrator
        End Get
        Set(value As String)
            _informToAdministrator = value
        End Set
    End Property

    Public Function execute() As Boolean
        Try
            informToAdministrator = False
            If (statusProcessing = False) Then
                retrieveFTStatusAndReffNo()

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            statusCode = 2
            statusMessage = ex.Message.ToString()
            statusProcessing = False

            Return False
        End Try
    End Function
    Private Sub retrieveFTStatusAndReffNo()
        statusProcessing = True

        Dim CUtility As New Utility()
        Dim CDatabase As New Database()

        Try
            Dim userID As String = CDatabase.getFieldValue("MasterWindowServiceConfiguration", "ServiceID = '" & AbstractClass.serviceID() & "' AND configurationName", "SPEMSVCUID", "value").Trim()
            Dim deptID As String = CDatabase.getDeptID(userID) 'Copied this function from Data.vb

            Dim linkedserver As String = AbstractClass.linkedServer()
            Dim libraryFT As String = AbstractClass.ftLib()
            Dim libraryMQ As String = AbstractClass.midasLibMQ()
            Dim isUpdated As Integer = 0

            Dim query As String = "declare @CountFRNTID int,@count1 int,@FRNT varchar(50),@countPer100 int,@Query varchar(max),@FRNTList varchar(max)                           " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "select @count1 = 1,@countPer100 = 0                                                                                                                " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'BULKMidasTemp_" & userID & "')                                             " & vbCrLf & _
                            "BEGIN                                                                                                                                              " & vbCrLf & _
                            "   DROP TABLE BULKMidasTemp_" & userID & "                                                                                                         " & vbCrLf & _
                            "End                                                                                                                                                " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "create table BULKMidasTemp_" & userID & "(                                                                                                         " & vbCrLf & _
                            "ReferenceNo varchar(20),                                                                                                                           " & vbCrLf & _
                            "FRNTID		varchar(20),                                                                                                                            " & vbCrLf & _
                            "MidasReffNo varchar(20),                                                                                                                           " & vbCrLf & _
                            "AUIN		varchar(1)                                                                                                                              " & vbCrLf & _
                            ")                                                                                                                                                  " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "--get reference has been approved                                                                                                                  " & vbCrLf & _
                            "select ROW_NUMBER()over (order by d.FTReffNo) as [_No], d.FTReffNo [FRNTMidas]	INTO #ListFRNT											            " & vbCrLf & _
                            "from TrxSPEMPayment h with (nolock) INNER join TrxSPEMPaymentDetails d on h.TransNo = d.TransNo                                                    " & vbCrLf & _
                            "and convert(varchar(10),ISNULL(d.ValueDateAfterEdited,d.ValueDate),120)=convert(varchar(10),GETDATE(),120)                                         " & vbCrLf & _
                            "where h.Flag='0' and ISNULL(h.ApprovedBy,'')<>'' and h.TransferType in ('FFR','FHT') and UPPER(ISNULL(d.MidasFTStatus, '')) <> 'AUT'	            " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "-- count total trx per 100                                                                                                                         " & vbCrLf & _
                            "set @CountFRNTID = (select count(FRNTMidas) from #ListFRNT)                                                                                        " & vbCrLf & _
                            "if @CountFRNTID > 0                                                                                                                                " & vbCrLf & _
                            "begin                                                                                                                                              " & vbCrLf & _
                            "    select @countPer100 = ceiling(_No/100.) from #ListFRNT                                                                                         " & vbCrLf & _
                            "End                                                                                                                                                " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "if @countPer100 > 0                                                                                                                                " & vbCrLf & _
                            "begin                                                                                                                                              " & vbCrLf & _
                            "    while @countPer100 >= @count1                                                                                                                  " & vbCrLf & _
                            "    begin                                                                                                                                          " & vbCrLf & _
                            "        select top 1 @FRNTList = STUFF((select ',''''' + FRNTMidas + '''''' [text()]                                                               " & vbCrLf & _
                            "					                    from #ListFRNT                                                                                              " & vbCrLf & _
                            "					                    where [_No] > case when @count1 = 1 then 0                                                                  " & vbCrLf & _
                            "									                    when @count1 > 1 then ((@count1-1)*100)                                                     " & vbCrLf & _
                            "                                                       End                                                                                         " & vbCrLf & _
                            "					                    and [_No] <= (@count1*100)                                                                                  " & vbCrLf & _
                            "			                    for xml path ('')), 1, 1, '')                                                                                       " & vbCrLf & _
                            "        from #ListFRNT                                                                                                                             " & vbCrLf & _
                            "        where [_No] > case when @count1 = 1 then 0                                                                                                 " & vbCrLf & _
                            "	                    when @count1 > 1 then ((@count1-1)*100)                                                                                     " & vbCrLf & _
                            "                        End                                                                                                                        " & vbCrLf & _
                            "        and [_No] <= (@count1*100)                                                                                                                 " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "        set @Query = 'insert into BULKMidasTemp_" & userID & "                                                                                     " & vbCrLf & _
                            "	                    select RLPR,FRNT,PREF,AUIN                                                                                                  " & vbCrLf & _
                            "	                    from openquery(" & linkedserver & ",''SELECT PREF,RLPR,FRNT,AUIN FROM " & libraryFT & ".OTPAYDD where FRNT in (' + @FRNTList + ') and RECI<>''''C'''' for read only'')' " & vbCrLf & _
                            "        exec(@Query)                                                                                                                               " & vbCrLf & _
                            "        -- check new repair queue FBM                                                                                                              " & vbCrLf & _
                            "        set @Query = 'insert into BULKMidasTemp_" & userID & "                                                                                     " & vbCrLf & _
                            "                        select ''''[RLPR],MSWFOTID,''''[MidasReffNo],''''[AUIN]                                                                    " & vbCrLf & _
                            "                        from openquery(" & linkedserver & ",''SELECT MSWFOTID FROM " & libraryMQ & ".T_MSRPRQ where MSWFOTID in (' + @FRNTList + ') for read only'')'  " & vbCrLf & _
                            "        exec(@Query)                                                                                                                               " & vbCrLf & _
                            "        -- check WIP                                                                                                                               " & vbCrLf & _
                            "        set @Query = 'insert into BULKMidasTemp_" & userID & "                                                                                     " & vbCrLf & _
                            "                        select ''''[RLPR],MSWFOTID,''WIP''[MidasReffNo],''''[AUIN]                                                                 " & vbCrLf & _
                            "                        from openquery(" & linkedserver & ",''SELECT MSWFOTID FROM " & libraryMQ & ".T_MSWIP where MSWFOTID in (' + @FRNTList + ') and MSWSTAT not in (''''AUTHORISED'''',''''REJECTED'''') for read only'')' " & vbCrLf & _
                            "        exec(@Query)                                                                                                                               " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "        set @count1 = @count1 + 1                                                                                                                  " & vbCrLf & _
                            "    End                                                                                                                                            " & vbCrLf & _
                            "End                                                                                                                                                " & vbCrLf & _
                            "                                                                                                                                                   " & vbCrLf & _
                            "-------------------  =================== SKN & RTGS ======================== ------------------------------------------                            " & vbCrLf & _
                            "select d.FTReffNo ReferenceNo, '00' AS MidasRefNo, NULL AS ERPRefNo, d.FTReffNo RLPR, d.MidasFTStatus StatusID,''[AUIN]                            " & vbCrLf & _
                            "from TrxSPEMPayment h INNER join TrxSPEMPaymentDetails d on h.TransNo = d.TransNo                                                                  " & vbCrLf & _
                            "and convert(varchar(10),ISNULL(d.ValueDateAfterEdited,d.ValueDate),120)=convert(varchar(10),GETDATE(),120)                                         " & vbCrLf & _
                            "WHERE h.Flag = 0 and DATEDIFF(MINUTE,h.DateApproved,GETDATE()) > 1 and h.TransferType in ('FFR','FHT') and UPPER(ISNULL(d.MidasFTStatus, '')) <> 'AUT'      " & vbCrLf & _
                            "AND (ISNULL(h.ApprovedBy,'')<>'') and d.FTReffNo not in (select FRNTID from BULKMidasTemp_" & userID & ")                                          " & vbCrLf & _
                            "union all                                                                                                                                          " & vbCrLf & _
                            "select a.ReferenceNo, MidasReffNo MidasRefNo, FRNTID ERPRefNo, b.FTReffNo RLPR,'' StatusID, AUIN from BULKMidasTemp_" & userID & " a               " & vbCrLf & _
                            "left join (select h.GCMSReffNo, d.* from TrxSPEMPayment h INNER join TrxSPEMPaymentDetails d on h.TransNo = d.TransNo                              " & vbCrLf & _
                            "and convert(varchar(10),ISNULL(d.ValueDateAfterEdited,d.ValueDate),120)=convert(varchar(10),GETDATE(),120) ) b on a.FRNTID=b.FTReffNo              " & vbCrLf & _
                            "drop table BULKMidasTemp_" & userID & "                                                                                                            " & vbCrLf & _
                            "drop table #ListFRNT"

            Dim dt As DataTable = CDatabase.getDataTable(query)
            If (dt.Rows.Count > 0) Then
                For i = 0 To dt.Rows.Count - 1
                    'Checking Get MIDAS+ Reference No
                    If (Not String.IsNullOrEmpty(dt.Rows(i)("MidasRefNo").ToString())) Then
                        Dim _RLPR As String = dt.Rows(i)("RLPR").ToString()
                        Dim _MidasRefNo As String = dt.Rows(i)("MidasRefNo").ToString()
                        Dim _AUIN As String = dt.Rows(i)("AUIN").ToString()
                        'Dim _RefNo As String = dt.Rows(i)("ReferenceNo").ToString()

                        Dim _MidasStatusID As String = dt.Rows(i)("StatusID").ToString()
                        Dim _StatusID As String = String.Empty

                        If _MidasRefNo = "WIP" Then
                            _StatusID = "WIP"
                        ElseIf _MidasRefNo = "00" Then
                            _StatusID = "DEL"
                        Else
                            If _AUIN.Trim.ToUpper = "Y" Then
                                _StatusID = "AUT" 'Authorized
                            Else
                                _StatusID = "WAT"
                            End If
                        End If

                        If _StatusID <> CDatabase.getFieldValue("TrxSPEMPaymentDetails", "FTReffNo", _RLPR, "MidasFTStatus") Then
                            Dim queryUpdate As String = "UPDATE TrxSPEMPaymentDetails SET FTReffNo=FTReffNo"

                            If _MidasRefNo <> CDatabase.getFieldValue("TrxSPEMPaymentDetails", "FTReffNo", _RLPR, "MidasPReffNo") And _MidasRefNo <> "WIP" Then
                                queryUpdate += ",MidasRefNo = '" & IIf(_MidasRefNo = "00", "", _MidasRefNo) & "' "
                            End If

                            If _StatusID.Trim <> "" Then
                                queryUpdate += ",MidasFTStatus = '" & _StatusID & "' ,MidasFTModifiedBy = '" & userID & "' ,MidasFTDateModified = GETDATE() ,MidasFTModifiedByDept = '" & deptID & "' "
                            End If

                            queryUpdate += " WHERE FTReffNo = '" & _MidasRefNo & "'"

                            Dim _errUpdate As String = String.Empty
                            CDatabase.execute(queryUpdate, 0, _errUpdate, True)

                            If (Not String.IsNullOrEmpty(_errUpdate)) Then
                                Throw New Exception(_errUpdate)
                                Exit For
                            End If
                            isUpdated = isUpdated + 1
                        End If
                    Else
                        'Checking for ERP Reference No in MIDAS+ (Repair)
                        If (Not String.IsNullOrEmpty(dt.Rows(i)("ERPRefNo").ToString())) Then
                            Dim _RLPR As String = dt.Rows(i)("RLPR").ToString()
                            Dim _ERPRefNo As String = dt.Rows(i)("ERPRefNo").ToString()

                            Dim _MidasStatusID As String = dt.Rows(i)("StatusID").ToString()
                            Dim _StatusID As String = String.Empty

                            _StatusID = "WRP"
                            If _StatusID <> CDatabase.getFieldValue("TrxSPEMPaymentDetails", "FTReffNo", _RLPR, "MidasFTStatus") Then
                                Dim queryUpdate As String = "UPDATE TrxSPEMPaymentDetails SET MidasFTStatus = '" & _StatusID & "' " & _
                                                        ",MidasFTModifiedBy = '" & userID & "' ,MidasFTDateModified = GETDATE() ,MidasFTModifiedByDept = '" & deptID & "' " & _
                                                        "WHERE FTReffNo = '" & _RLPR & "'"
                                Dim _errUpdate As String = String.Empty
                                CDatabase.execute(queryUpdate, 0, _errUpdate, False)
                                'execute(queryUpdate, 0, _errUpdate, False)

                                If (Not String.IsNullOrEmpty(_errUpdate)) Then
                                    Throw New Exception(_errUpdate)
                                    Exit For
                                End If
                                isUpdated = isUpdated + 1
                            End If
                        End If

                    End If
                Next

            End If

            'Cek failed upload trx
            query = ""

            'If we found just 1 processing, it will be return message that we have converting and saving SPEM files
            If (isUpdated > 0) Then
                statusCode = 1
            Else
                statusCode = 0
            End If

            statusProcessing = False
        Catch ex As Exception
            statusCode = 2
            statusMessage = ex.Message.ToString()
            statusProcessing = False

            Throw New Exception(statusMessage)
        End Try
    End Sub
    

    'Private Sub retrieveFTStatusAndReffNo()
    '    statusProcessing = True

    '    Dim CUtility As New Utility()
    '    Dim CDatabase As New Database()

    '    Try
    '        Dim userID As String = CDatabase.getFieldValue("MasterWindowServiceConfiguration", "ServiceID = '" & AbstractClass.serviceID() & "' AND configurationName", "SPEMSVCUID", "value").Trim()
    '        Dim deptID As String = CDatabase.getDeptID(userID) 'Copied this function from Data.vb
    '        Dim actionStatus As Boolean = False

    '        Dim query As String = "WITH cteMidasFT AS ( " & vbCrLf & _
    '                              " SELECT SPD.TransNo, SPD.SEQ, SPD.FTReffNo, DD.PREF AS MidasPReffNo, " & vbCrLf & _
    '                              "        CASE WHEN(UPPER(ISNULL(DD.RECI, '')) = 'C') THEN " & vbCrLf & _
    '                              "         'DEL' " & vbCrLf & _
    '                              "        ELSE " & vbCrLf & _
    '                              "         CASE WHEN(UPPER(ISNULL(DD.AUIN, '')) = 'Y') THEN 'AUT' ELSE 'WAT' END " & vbCrLf & _
    '                              "        END AS MidasFTStatus " & vbCrLf & _
    '                              " FROM TrxSPEMPayment SP LEFT JOIN TrxSPEMPaymentDetails SPD ON SP.TransNo = SPD.TransNo " & vbCrLf & _
    '                              "      INNER JOIN " & AbstractClass.ftLinkedServers() & ".OTPAYDD DD ON SP.GCMSReffNo = DD.RLPR COLLATE SQL_Latin1_General_CP1_CI_AS AND " & vbCrLf & _
    '                              "                                                         SPD.FTReffNo = DD.FRNT COLLATE SQL_Latin1_General_CP1_CI_AS " & vbCrLf & _
    '                              "WHERE SP.TransferType IN ('FFR', 'FHT') AND  ISNULL(SP.ApprovedBy, '') <> '' AND UPPER(ISNULL(SPD.MidasFTStatus, '')) <> 'AUT' " & vbCrLf & _
    '                              "UNION " & vbCrLf & _
    '                              "SELECT SPD.TransNo, SPD.SEQ, SPD.FTReffNo, NULL AS MidasPReffNo, 'WRP' AS MidasFTStatus " & vbCrLf & _
    '                              "FROM TrxSPEMPayment SP LEFT JOIN TrxSPEMPaymentDetails SPD ON SP.TransNo = SPD.TransNo " & vbCrLf & _
    '                              "      INNER JOIN " & AbstractClass.ftLinkedServers() & ".FTIOPAYPD PD ON SP.GCMSReffNo = PD.SRLPR COLLATE SQL_Latin1_General_CP1_CI_AS AND " & vbCrLf & _
    '                              "                                                           SPD.FTReffNo = PD.SFOTRANID COLLATE SQL_Latin1_General_CP1_CI_AS " & vbCrLf & _
    '                              "WHERE SP.TransferType IN ('FFR', 'FHT') AND ISNULL(SP.ApprovedBy, '') <> '' AND UPPER(ISNULL(SPD.MidasFTStatus, '')) <> 'AUT' " & vbCrLf & _
    '                              ") " & vbCrLf & _
    '                              "UPDATE USPD " & vbCrLf & _
    '                              "   SET USPD.MidasFTStatus = MF.MidasFTStatus, USPD.MidasPReffNo = MF.MidasPReffNo, USPD.MidasFTModifiedBy = '" & userID & "', " & vbCrLf & _
    '                              "       USPD.MidasFTDateModified = GETDATE(), USPD.MidasFTModifiedByDept = '" & deptID & "' " & vbCrLf & _
    '                              "FROM TrxSPEMPaymentDetails USPD INNER JOIN cteMidasFT MF ON USPD.TransNo = MF.TransNo AND USPD.SEQ = MF.SEQ"

    '        Dim errMessage As String = String.Empty
    '        CDatabase.execute(query, 0, errMessage, False)
    '        If (Not String.IsNullOrEmpty(errMessage)) Then
    '            Throw New Exception(errMessage)
    '        Else
    '            actionStatus = True
    '        End If

    '        'If we found just 1 processing, it will be return message that we have converting and saving SPEM files
    '        If (actionStatus = True) Then
    '            statusCode = 1
    '        Else
    '            statusCode = 0
    '        End If

    '        statusProcessing = False
    '    Catch ex As Exception
    '        statusCode = 2
    '        statusMessage = ex.Message.ToString()
    '        statusProcessing = False

    '        Throw New Exception(statusMessage)
    '    End Try
    'End Sub
End Class
