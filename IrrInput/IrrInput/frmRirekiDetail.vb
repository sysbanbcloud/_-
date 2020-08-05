Public Class frmRirekiDetail
    Const mstrfrmId As String = "A03"

    Private mstrKibanKbn As String = ""      '機器区分（1:スリット機 2:巻き替え機 左記以外:その他上工程など）

    Private mstrId As String
    ''' <summary>
    ''' フォーム間のデータ受け渡し用プロパティ
    ''' 前画面で選択した行のIR実績ID    
    ''' </summary>
    ''' <returns></returns>
    Public Property rstrId As String
        Get
            rstrId = mstrId
        End Get
        Set(value As String)
            mstrId = value
        End Set
    End Property

    Private mintMode As String
    ''' <summary>
    ''' フォーム間のデータ受け渡し用プロパティ
    ''' 前画面で押下したボタン（1:修正、9:照会）    
    ''' </summary>
    ''' <returns></returns>
    Public Property rintMode As String
        Get
            rintMode = mintMode
        End Get
        Set(value As String)
            mintMode = value
        End Set
    End Property

    Private Sub frmIrrJissekiInput_Load(sender As Object, e As EventArgs) Handles Me.Load
        lblTitle.Text = mstrfrmId & ":作成履歴詳細"
        lblKiban.Text = ""      '機番非表示

        fncSetList()

        If mintMode = 1 Then
            '修正
            grdList.ReadOnly = False
            btnUpd.Visible = True
            btnDel.Visible = True
            'btnRowDel.Visible = True
            grdList.Columns("合否cbo").Visible = True
            grdList.Columns("合否txt").Visible = False
        Else
            '照会
            grdList.ReadOnly = True
            btnUpd.Visible = False
            btnDel.Visible = False
            'btnRowDel.Visible = False
            grdList.Columns("合否cbo").Visible = False
            grdList.Columns("合否txt").Visible = True
            grdList.AllowUserToAddRows = False
            grdList.AllowUserToDeleteRows = False
        End If

        grdList.ClearSelection()

        '機器区分取得（1:スリット機 2:巻き替え機 左記以外:その他上工程など）
        Dim strSql As String = String.Empty
        strSql = "select 機器区分 from " & gcTblKibanMst
        strSql &= " where 機番 = " & gfncMakeSqlValue(lblKiban2.Text, 0)
        mstrKibanKbn = gfncGetSqlValue(strSql)

    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Owner.Show()
        Me.Close()
    End Sub

    Private Sub btnUpd_Click(sender As Object, e As EventArgs) Handles btnUpd.Click
        '入力チェック
        If Not fncCheck() Then
            Exit Sub
        End If

        If MessageBox.Show("この内容で登録しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Dim strSql As String = String.Empty
        Dim strErr As String = String.Empty
        Dim strRet As String = String.Empty
        Dim intId As Integer = 0

        Try
            If Not gfncDbConnect(strErr) Then
                MessageBox.Show(gcErrDb010, "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If Not gfncDBBeginTrans(strErr) Then
                MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor

            ' IR実績明細TBLに登録　グリッドの行数分
            For intRow As Integer = 0 To grdList.RowCount - 1
                '明細行入力ナシは登録しない（最終行など）
                If grdList("第一給紙", intRow).Value = "" And grdList("第二給紙", intRow).Value = "" And grdList("仕上りm", intRow).Value = "" And grdList("継目回数", intRow).Value = "" And
                   grdList("継目位置１", intRow).Value = "" And grdList("継目位置２", intRow).Value = "" And grdList("継目位置３", intRow).Value = "" And grdList("継目位置４", intRow).Value = "" And grdList("継目位置５", intRow).Value = "" And
                   grdList("合否cbo", intRow).Value = "" Then
                    Exit For
                End If

                If grdList("SEQ", intRow).Value <> "" Then
                    '修正行
                    strSql = "UPDATE " & gcTblIrMeisai & " SET "
                    If grdList("第一給紙", intRow).Value = "" Then
                        strSql &= "第一給紙=NULL"
                    Else
                        strSql &= "第一給紙=" & gfncMakeSqlValue(grdList("第一給紙", intRow).Value, 0)
                    End If
                    If grdList("第二給紙", intRow).Value = "" Then
                        strSql &= ",第二給紙=NULL"
                    Else
                        strSql &= ",第二給紙=" & gfncMakeSqlValue(grdList("第二給紙", intRow).Value, 0)
                    End If
                    If grdList("製造ロットNo", intRow).Value = "" Then
                        strSql &= ",ロットNo=NULL"
                    Else
                        strSql &= ",ロットNo=" & gfncMakeSqlValue(grdList("製造ロットNo", intRow).Value, 0)
                    End If
                    If grdList("仕上りm", intRow).Value = "" Then
                        strSql &= ",仕上りm=0"
                    Else
                        strSql &= ",仕上りm=" & gfncMakeSqlValue(grdList("仕上りm", intRow).Value.Replace(",", ""), 1)
                    End If
                    strSql &= ",継目回数=" & gfncMakeSqlValue(grdList("継目回数", intRow).Value, 1)
                    If grdList("継目位置１", intRow).Value = "" Then
                        strSql &= ",継目位置１=NULL"
                    Else
                        strSql &= ",継目位置１=" & gfncMakeSqlValue(grdList("継目位置１", intRow).Value, 1)
                    End If
                    If grdList("継目位置２", intRow).Value = "" Then
                        strSql &= ",継目位置２=NULL"
                    Else
                        strSql &= ",継目位置２=" & gfncMakeSqlValue(grdList("継目位置２", intRow).Value, 1)
                    End If
                    If grdList("継目位置３", intRow).Value = "" Then
                        strSql &= ",継目位置３=NULL"
                    Else
                        strSql &= ",継目位置３=" & gfncMakeSqlValue(grdList("継目位置３", intRow).Value, 1)
                    End If
                    If grdList("継目位置４", intRow).Value = "" Then
                        strSql &= ",継目位置４=NULL"
                    Else
                        strSql &= ",継目位置４=" & gfncMakeSqlValue(grdList("継目位置４", intRow).Value, 1)
                    End If
                    If grdList("継目位置５", intRow).Value = "" Then
                        strSql &= ",継目位置５=NULL"
                    Else
                        strSql &= ",継目位置５=" & gfncMakeSqlValue(grdList("継目位置５", intRow).Value, 1)
                    End If
                    If grdList("合否cbo", intRow).Value = "合" Then
                        strSql &= ",合否=0"
                    Else
                        strSql &= ",合否=1"
                    End If
                    strSql &= " WHERE IR実績ID=" & gfncMakeSqlValue(mstrId, 1)
                    strSql &= " AND SEQ=" & gfncMakeSqlValue(grdList("SEQ", intRow).Value, 1)
                Else
                    '新規行
                    'SEQ取得
                    Dim strSql2 As String
                    Dim strSeq As String
                    strSql2 = "select max(SEQ+1) as maxSEQ from " & gcTblIrMeisai & " where IR実績ID=" & gfncMakeSqlValue(mstrId, 1)
                    strSeq = gfncGetSqlValueTrn(strSql2)
                    If strSeq = "" Then
                        strSeq = "1"
                    End If

                    '登録
                    strSql = "INSERT INTO " & gcTblIrMeisai & "("
                        strSql &= " IR実績ID"
                        strSql &= ",SEQ"
                        strSql &= ",第一給紙"
                        strSql &= ",第二給紙"
                        strSql &= ",ロットNo"
                        strSql &= ",仕上りm"
                        strSql &= ",継目回数"
                        strSql &= ",継目位置１"
                        strSql &= ",継目位置２"
                        strSql &= ",継目位置３"
                        strSql &= ",継目位置４"
                        strSql &= ",継目位置５"
                        strSql &= ",合否"
                        strSql &= " )VALUES( "
                        strSql &= gfncMakeSqlValue(mstrId, 1)
                    strSql &= "," & gfncMakeSqlValue(strSeq, 1)
                    If grdList("第一給紙", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("第一給紙", intRow).Value, 0)
                    End If
                    If grdList("第二給紙", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("第二給紙", intRow).Value, 0)
                    End If
                    If grdList("製造ロットNo", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("製造ロットNo", intRow).Value, 0)
                    End If
                    If grdList("仕上りm", intRow).Value = "" Then
                        strSql &= ",0"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("仕上りm", intRow).Value, 1)
                    End If
                    strSql &= "," & gfncMakeSqlValue(grdList("継目回数", intRow).Value, 1)
                    If grdList("継目位置１", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("継目位置１", intRow).Value, 1)
                    End If
                    If grdList("継目位置２", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("継目位置２", intRow).Value, 1)
                    End If
                    If grdList("継目位置３", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("継目位置３", intRow).Value, 1)
                    End If
                    If grdList("継目位置４", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("継目位置４", intRow).Value, 1)
                    End If
                    If grdList("継目位置５", intRow).Value = "" Then
                        strSql &= ",NULL"
                    Else
                        strSql &= "," & gfncMakeSqlValue(grdList("継目位置５", intRow).Value, 1)
                    End If
                    If grdList("合否cbo", intRow).Value = "合" Then
                        strSql &= ",0"
                    Else
                        strSql &= ",1"
                    End If
                    strSql &= ")"

                    strSeq = CStr(CInt(strSeq) + 1)
                End If

                If Not gfncDBExecute(strErr, strSql, strRet) Then
                    If Not gfncDBRollback(strErr) Then
                        MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If
                    MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
            Next

            ' データ保存 コミット
            If Not gfncDBCommitTrans(strErr) Then
                MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            fncSetList()    '再表示
            grdList.ClearSelection()
            MessageBox.Show("登録しました", "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            If Not gfncDBRollback(strErr) Then
                MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql & Environment.NewLine & ex.ToString, "IR実績詳細データ修正", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        Finally
            If Not gfncDBClose(strErr) Then

            End If
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        Dim strSql As String = String.Empty
        Dim strErr As String = String.Empty
        Dim strRet As String = String.Empty
        Dim intId As Integer = 0

        ' 選択行取得
        Dim intCurrentRow = CInt(grdList.CurrentRow.Index)

        Dim intRowCnt As Integer = grdList.Rows.GetRowCount(DataGridViewElementStates.Selected)
        If intRowCnt <= 0 Then
            MessageBox.Show("削除する行を選択してください", "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If MessageBox.Show("選択行を削除してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        ' 削除する行がDB登録済みか、登録中か判定
        strSql = "SELECT COUNT(IR実績ID) as 件数 FROM " & gcTblIrMeisai
        strSql &= " WHERE IR実績ID=" & gfncMakeSqlValue(mstrId, 1)
        strSql &= " AND SEQ=" & gfncMakeSqlValue(grdList("No", intCurrentRow).Value, 1)
        If gfncGetSqlValue(strSql) = "0" Then
            '---登録中の行を削除
            If grdList.CurrentCellAddress.Y = grdList.RowCount - 1 Then
                For intCol As Integer = 0 To grdList.ColumnCount - 1
                    grdList(intCol, grdList.CurrentCellAddress.Y).Value = ""
                Next
            Else
                grdList.Rows.RemoveAt(grdList.CurrentCellAddress.Y)
            End If

        Else
            '---登録済の行を削除
            Try
                If Not gfncDbConnect(strErr) Then
                    MessageBox.Show(gcErrDb010, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
                If Not gfncDBBeginTrans(strErr) Then
                    MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor

                ' 選択行削除
                strSql = "DELETE FROM " & gcTblIrMeisai
                strSql &= " WHERE IR実績ID=" & gfncMakeSqlValue(mstrId, 1)
                strSql &= " AND SEQ=" & gfncMakeSqlValue(grdList("No", intCurrentRow).Value, 1)
                If Not gfncDBExecute(strErr, strSql, strRet) Then
                    If Not gfncDBRollback(strErr) Then
                        MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If
                    MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                ' コミット
                If Not gfncDBCommitTrans(strErr) Then
                    MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                'fncSetList()    '再表示
                grdList.Rows.RemoveAt(grdList.CurrentCellAddress.Y)

            Catch ex As Exception
                If Not gfncDBRollback(strErr) Then
                    MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
                MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql & Environment.NewLine & ex.ToString, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            Finally
                If Not gfncDBClose(strErr) Then

                End If
                Me.Cursor = Cursors.Default
            End Try
        End If

        ' 番号振り直し
        For intRow As Integer = 0 To grdList.RowCount - 1
            grdList("No", intRow).Value = intRow + 1
        Next
        grdList.ClearSelection()
        MessageBox.Show("削除しました", "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    'Private Sub btnRowDel_Click(sender As Object, e As EventArgs) Handles btnRowDel.Click
    '    ' グリッド選択行削除
    '    If grdList.CurrentCellAddress.Y = grdList.RowCount - 1 Then
    '        For intCol As Integer = 0 To grdList.ColumnCount - 1
    '            grdList(intCol, grdList.CurrentCellAddress.Y).Value = ""
    '        Next
    '    Else
    '        grdList.Rows.RemoveAt(grdList.CurrentCellAddress.Y)
    '    End If

    '    ' 番号振り直し
    '    For intRow As Integer = 0 To grdList.RowCount - 2
    '        grdList("No", intRow).Value = intRow + 1
    '        'grdList("製造ロットNo", intRow).Value = txtLotNo.Text & (intRow + 1).ToString("00")
    '    Next
    'End Sub

    '一覧表示
    Private Sub fncSetList()
        Dim strSql As String = String.Empty
        Dim dt As New DataTable

        Try
            ' データ取得
            strSql = "SELECT"
            strSql &= " IR実績ID"
            strSql &= ",SEQ"
            strSql &= ",第一給紙"
            strSql &= ",第二給紙"
            strSql &= ",ロットNo"
            strSql &= ",仕上りm"
            strSql &= ",継目回数"
            strSql &= ",継目位置１"
            strSql &= ",継目位置２"
            strSql &= ",継目位置３"
            strSql &= ",継目位置４"
            strSql &= ",継目位置５"
            strSql &= ",合否"
            strSql &= " FROM " & gcTblIrMeisai
            strSql &= " WHERE 1 = 1"
            strSql &= " and IR実績ID = " & gfncMakeSqlValue(mstrId, 1)
            strSql &= " Order by SEQ"
            If Not gfncbol_RecordGet(strSql, dt) Then
                Exit Sub
            End If

            ' データをグリッドに表示
            ' 描画を止める
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            grdList.SuspendLayout()
            grdList.RowCount = 1

            For lLoop = 0 To grdList.ColumnCount - 1
                grdList.Columns(lLoop).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            Next

            For Each dr As DataRow In dt.Rows
                Dim item As New DataGridViewRow
                item.CreateCells(grdList)
                Dim intRow As Integer = grdList.Rows.Add(item)

                grdList("No", intRow).Value = intRow + 1
                grdList("SEQ", intRow).Value = dr("SEQ").ToString
                grdList("第一給紙", intRow).Value = dr("第一給紙").ToString
                grdList("第二給紙", intRow).Value = dr("第二給紙").ToString
                grdList("製造ロットNo", intRow).Value = dr("ロットNo").ToString
                If gfncNVL(dr("仕上りm"), "") = "" Then
                    grdList("仕上りm", intRow).Value = ""
                Else
                    grdList("仕上りm", intRow).Value = CInt(dr("仕上りm")).ToString("#,0")
                End If
                grdList("継目回数", intRow).Value = gfncNVL(dr("継目回数"), "")
                If gfncNVL(dr("継目位置１"), "") = "" Then
                    grdList("継目位置１", intRow).Value = ""
                Else
                    grdList("継目位置１", intRow).Value = CInt(dr("継目位置１")).ToString("#,0")
                End If
                If gfncNVL(dr("継目位置２"), "") = "" Then
                    grdList("継目位置２", intRow).Value = ""
                Else
                    grdList("継目位置２", intRow).Value = CInt(dr("継目位置２")).ToString("#,0")
                End If
                If gfncNVL(dr("継目位置３"), "") = "" Then
                    grdList("継目位置３", intRow).Value = ""
                Else
                    grdList("継目位置３", intRow).Value = CInt(dr("継目位置３")).ToString("#,0")
                End If
                If gfncNVL(dr("継目位置４"), "") = "" Then
                    grdList("継目位置４", intRow).Value = ""
                Else
                    grdList("継目位置４", intRow).Value = CInt(dr("継目位置４")).ToString("#,0")
                End If
                If gfncNVL(dr("継目位置５"), "") = "" Then
                    grdList("継目位置５", intRow).Value = ""
                Else
                    grdList("継目位置５", intRow).Value = CInt(dr("継目位置５")).ToString("#,0")
                End If
                If dr("合否").ToString = "0" Then
                    grdList("合否cbo", intRow).Value = "合"
                    grdList("合否txt", intRow).Value = "合"
                Else
                    grdList("合否cbo", intRow).Value = "否"
                    grdList("合否txt", intRow).Value = "否"
                End If
            Next

            If mintMode = 9 Then    '照会
                ' 列幅は自動とし、再描画
                For lLoop = 0 To grdList.ColumnCount - 1
                    grdList.Columns(lLoop).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Next
                grdList.ResumeLayout(True)
            End If

            Exit Sub

        Catch ex As Exception
            gsubExceptionProc(ex.Message, mstrfrmId, Reflection.MethodBase.GetCurrentMethod.Name)
            Exit Sub
        Finally
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Function fncCheck() As Boolean
        'If grdList.RowCount = 0 Then
        '    MessageBox.Show("明細を一行以上入力してください", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '    Return False
        'End If
        Dim intVal As Integer

        Dim intTsugime = 0
        For intRow As Integer = 0 To grdList.RowCount - 1
            '明細行入力ナシは登録しない（最終行など）
            If grdList("第一給紙", intRow).Value = "" And grdList("第二給紙", intRow).Value = "" And grdList("仕上りm", intRow).Value = "" And grdList("継目回数", intRow).Value = "" And
               grdList("継目位置１", intRow).Value = "" And grdList("継目位置２", intRow).Value = "" And grdList("継目位置３", intRow).Value = "" And grdList("継目位置４", intRow).Value = "" And grdList("継目位置５", intRow).Value = "" And
               grdList("合否cbo", intRow).Value = "" Then
                Exit For
            End If

            If CStr(grdList("No", intRow).Value) <> "" Then
                If grdList("第一給紙", intRow).Value = "" Then
                    MessageBox.Show(String.Format(gcErrInp001, "（" & intRow + 1 & "行目）第一給紙"), "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
                If grdList("製造ロットNo", intRow).Value = "" Then
                    MessageBox.Show(String.Format(gcErrInp001, "（" & intRow + 1 & "行目）製造ロットNo"), "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
                If Not Integer.TryParse(grdList("仕上りm", intRow).Value.Replace(",", ""), intVal) Then
                    MessageBox.Show("（" & intRow + 1 & "行目）仕上りmは数字で入力してください", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
                If grdList("継目回数", intRow).Value = "" Then
                    MessageBox.Show(String.Format(gcErrInp001, "継目回数"), "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
                If Not Integer.TryParse(grdList("継目回数", intRow).Value, intVal) Then
                    MessageBox.Show("（" & intRow + 1 & "行目）継目回数は数字で入力してください", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
                If CInt(grdList("継目回数", intRow).Value) < 0 Or CInt(grdList("継目回数", intRow).Value) > 5 Then
                    MessageBox.Show("（" & intRow + 1 & "行目）継目回数は0～5の間で入力してください", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
                If grdList("継目回数", intRow).Value = "0" Then
                    '継目回数=0
                    If grdList("継目位置１", intRow).Value <> "" Or grdList("継目位置２", intRow).Value <> "" Or grdList("継目位置３", intRow).Value <> "" Or grdList("継目位置４", intRow).Value <> "" Or grdList("継目位置５", intRow).Value <> "" Then
                        MessageBox.Show("継目回数が0の場合、継目位置の入力はできません", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return False
                    End If
                Else
                    '継目回数≠0
                    Dim bolFull As Boolean = False  '継目位置の入力チェックが継目回数に達したかどうかのフラグ
                    For intI As Integer = 1 To 5    '継目回数分ループ
                        If bolFull = False Then
                            If grdList("継目位置" & StrConv(intI, VbStrConv.Wide), intRow).Value = "" Then
                                MessageBox.Show("（" & intRow + 1 & "行目）継目回数分の継目位置を入力してください", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Return False
                            End If
                            If intI = CInt(grdList("継目回数", intRow).Value) Then
                                bolFull = True
                            End If
                            If Not Integer.TryParse(grdList("継目位置" & StrConv(intI, VbStrConv.Wide), intRow).Value, intVal) Then
                                MessageBox.Show("（" & intRow + 1 & "行目）継目位置は数字で入力してください", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Return False
                            End If
                        Else
                            If grdList("継目位置" & StrConv(intI, VbStrConv.Wide), intRow).Value <> "" Then
                                MessageBox.Show("（" & intRow + 1 & "行目）継目回数以上の継目位置は入力できません", "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Return False
                            End If
                        End If
                    Next
                End If

                If grdList("合否cbo", intRow).Value = "" Then
                    MessageBox.Show(String.Format(gcErrInp007, "（" & intRow + 1 & "行目）合否"), "入力チェック", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Private Sub grdList_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdList.CellClick
        If rintMode = 9 Then Exit Sub     '照会モードは処理なし

        Dim strBasicNo As String = String.Empty     'ロットNo基本部
        Dim intSerialNo As Integer = 0              'ロットNo連番部

        If e.RowIndex < 0 Then Exit Sub      ' タイトル行は抜ける

        grdList("No", e.RowIndex).Value = e.RowIndex + 1

        '1行目（初回）
        If e.RowIndex = 0 Then
            If grdList("製造ロットNo", 0).Value = "" Then
                If mstrKibanKbn = "1" Or mstrKibanKbn = "2" Then
                    grdList("製造ロットNo", 0).Value = lblLotNo.Text & "0101"
                Else
                    grdList("製造ロットNo", 0).Value = lblLotNo.Text & "01"
                End If
            End If
            Exit Sub
        End If

        '2行目以降（新規行に対応）
        If grdList("製造ロットNo", e.RowIndex - 1).Value = "" Then Exit Sub
        If grdList("製造ロットNo", e.RowIndex - 1).Value.ToString.Length <> 9 And grdList("製造ロットNo", e.RowIndex - 1).Value.ToString.Length <> 8 Then Exit Sub

        If mstrKibanKbn = "1" Or mstrKibanKbn = "2" Then
            strBasicNo = grdList("製造ロットNo", e.RowIndex - 1).Value.ToString.Substring(0, 7)
            Try
                intSerialNo = CInt(grdList("製造ロットNo", e.RowIndex - 1).Value.ToString.Substring(7, 2))
            Catch ex As Exception
            End Try
        Else
            strBasicNo = grdList("製造ロットNo", e.RowIndex - 1).Value.ToString.Substring(0, 6)
            Try
                intSerialNo = CInt(grdList("製造ロットNo", e.RowIndex - 1).Value.ToString.Substring(6, 2))
            Catch ex As Exception
            End Try
        End If

        If e.RowIndex = grdList.RowCount - 1 Then
            If grdList("製造ロットNo", e.RowIndex).Value = "" Then
                grdList("製造ロットNo", e.RowIndex).Value = strBasicNo & (intSerialNo + 1).ToString("D2")
            End If
        End If
    End Sub


    'セルのフォーマット
    Private Sub grdList_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles grdList.CellFormatting
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim intVal As Integer

        'セルの列を確認
        If dgv.Columns(e.ColumnIndex).Name = "仕上りm" AndAlso TypeOf e.Value Is String Then
            Dim str As String = e.Value.ToString()
            If Integer.TryParse(str, intVal) Then
                e.Value = CInt(str).ToString("#,0")
            End If
            'フォーマットの必要がないことを知らせる
            e.FormattingApplied = True
            Exit Sub
        End If
        If dgv.Columns(e.ColumnIndex).Name = "継目位置１" AndAlso TypeOf e.Value Is String Then
            Dim str As String = e.Value.ToString()
            If Integer.TryParse(str, intVal) Then
                e.Value = CInt(str).ToString("#,0")
            End If
            'フォーマットの必要がないことを知らせる
            e.FormattingApplied = True
            Exit Sub
        End If
        If dgv.Columns(e.ColumnIndex).Name = "継目位置２" AndAlso TypeOf e.Value Is String Then
            Dim str As String = e.Value.ToString()
            If Integer.TryParse(str, intVal) Then
                e.Value = CInt(str).ToString("#,0")
            End If
            'フォーマットの必要がないことを知らせる
            e.FormattingApplied = True
            Exit Sub
        End If
        If dgv.Columns(e.ColumnIndex).Name = "継目位置３" AndAlso TypeOf e.Value Is String Then
            Dim str As String = e.Value.ToString()
            If Integer.TryParse(str, intVal) Then
                e.Value = CInt(str).ToString("#,0")
            End If
            'フォーマットの必要がないことを知らせる
            e.FormattingApplied = True
            Exit Sub
        End If
        If dgv.Columns(e.ColumnIndex).Name = "継目位置４" AndAlso TypeOf e.Value Is String Then
            Dim str As String = e.Value.ToString()
            If Integer.TryParse(str, intVal) Then
                e.Value = CInt(str).ToString("#,0")
            End If
            'フォーマットの必要がないことを知らせる
            e.FormattingApplied = True
            Exit Sub
        End If
        If dgv.Columns(e.ColumnIndex).Name = "継目位置５" AndAlso TypeOf e.Value Is String Then
            Dim str As String = e.Value.ToString()
            If Integer.TryParse(str, intVal) Then
                e.Value = CInt(str).ToString("#,0")
            End If
            'フォーマットの必要がないことを知らせる
            e.FormattingApplied = True
            Exit Sub
        End If
    End Sub


End Class
