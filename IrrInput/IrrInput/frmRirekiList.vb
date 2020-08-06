Public Class frmRirekiList
    Const mstrfrmId As String = "A02"

    Private Sub frmIrrJissekiInput_Load(sender As Object, e As EventArgs) Handles Me.Load
        lblTitle.Text = mstrfrmId & ":作成履歴"
        lblKiban.Text = ""      '機番非表示

        ' 画面初期化
        subFrmClear()

    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Owner.Show()
        Me.Close()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        fncSetList()
    End Sub

    Private Sub btnDetail_Click(sender As Object, e As EventArgs) Handles btnDetail.Click
        If grdList.RowCount = 0 Then Exit Sub

        Dim objfrmRirekiDetail As New frmRirekiDetail
        Try
            ' 詳細画面に値セット
            objfrmRirekiDetail.lblJyutyuNo.Text = grdList.CurrentRow.Cells(4).Value      '受注書No
            objfrmRirekiDetail.lblShiyouNo.Text = grdList.CurrentRow.Cells(5).Value      '仕様書No
            objfrmRirekiDetail.lblHinName.Text = grdList.CurrentRow.Cells(6).Value       '製品名
            objfrmRirekiDetail.lblKakouDate.Text = grdList.CurrentRow.Cells(1).Value     '加工日
            objfrmRirekiDetail.lblKiban2.Text = grdList.CurrentRow.Cells(2).Value        '機番
            objfrmRirekiDetail.lblLotNo.Text = grdList.CurrentRow.Cells(11).Value        'ロットNo（基本）

            ' 詳細画面に値を渡す
            objfrmRirekiDetail.rstrId = grdList.CurrentRow.Cells(10).Value   'IR実績ID
            objfrmRirekiDetail.rintMode = 9     '照会モード

            objfrmRirekiDetail.Show(Me)
            Me.Hide()
        Catch ex As Exception
            gsubExceptionProc(ex.Message, mstrfrmId, Reflection.MethodBase.GetCurrentMethod.Name)
        Finally
        End Try
    End Sub

    Private Sub btnMod_Click(sender As Object, e As EventArgs) Handles btnMod.Click
        If grdList.RowCount = 0 Then Exit Sub

        Dim objfrmRirekiDetail As New frmRirekiDetail
        Try
            ' 詳細画面に値セット
            objfrmRirekiDetail.lblJyutyuNo.Text = grdList.CurrentRow.Cells(4).Value      '受注書No
            objfrmRirekiDetail.lblShiyouNo.Text = grdList.CurrentRow.Cells(5).Value      '仕様書No
            objfrmRirekiDetail.lblHinName.Text = grdList.CurrentRow.Cells(6).Value       '製品名
            objfrmRirekiDetail.lblKakouDate.Text = grdList.CurrentRow.Cells(1).Value     '加工日
            objfrmRirekiDetail.lblKiban2.Text = grdList.CurrentRow.Cells(2).Value        '機番
            objfrmRirekiDetail.lblLotNo.Text = grdList.CurrentRow.Cells(11).Value        'ロットNo（基本）

            ' 詳細画面に値を渡す
            objfrmRirekiDetail.rstrId = grdList.CurrentRow.Cells(10).Value   'IR実績ID
            objfrmRirekiDetail.rintMode = 1     '修正モード

            objfrmRirekiDetail.Show(Me)
            Me.Hide()
        Catch ex As Exception
            gsubExceptionProc(ex.Message, mstrfrmId, Reflection.MethodBase.GetCurrentMethod.Name)
        Finally
        End Try
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        Dim intRowCnt As Integer = grdList.Rows.GetRowCount(DataGridViewElementStates.Selected)
        If intRowCnt <= 0 Then
            MessageBox.Show("削除する行を選択してください", "IR実績データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If MessageBox.Show("選択行を削除してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Dim strSql As String = String.Empty
        Dim strErr As String = String.Empty
        Dim strRet As String = String.Empty
        Dim intId As Integer = 0

        Try
            If Not gfncDbConnect(strErr) Then
                MessageBox.Show(gcErrDb010, "IR実績データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If Not gfncDBBeginTrans(strErr) Then
                MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor

            ' IR実績基本削除
            Dim intCurrentRow = CInt(grdList.CurrentRow.Index)
            strSql = "DELETE FROM " & gcTblIrKihon
            strSql &= " WHERE IR実績ID=" & gfncMakeSqlValue(grdList("IR実績ID", intCurrentRow).Value, 1)
            If Not gfncDBExecute(strErr, strSql, strRet) Then
                If Not gfncDBRollback(strErr) Then
                    MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
                MessageBox.Show(gcErrDb011 & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strErr & Environment.NewLine & Reflection.MethodBase.GetCurrentMethod.Name & Environment.NewLine & strSql, "IR実績詳細データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            ' IR実績明細削除
            strSql = "DELETE FROM " & gcTblIrMeisai
            strSql &= " WHERE IR実績ID=" & gfncMakeSqlValue(grdList("IR実績ID", intCurrentRow).Value, 1)
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

            fncSetList()    '再表示
            grdList.ClearSelection()
            MessageBox.Show("削除しました", "IR実績データ削除", MessageBoxButtons.OK, MessageBoxIcon.Information)

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
    End Sub

    '一覧表示
    Private Sub fncSetList()
        Dim strSql As String = String.Empty
        Dim dt As New DataTable

        Try
            ' データ取得
            strSql = "SELECT"
            strSql &= " I.IR実績ID as IR実績ID"
            strSql &= ",I.注文No as 注文No"
            strSql &= ",I.仕様書No as 仕様書No"
            strSql &= ",I.製品名 as 製品名"
            strSql &= ",I.加工日 as 加工日"
            strSql &= ",I.機番 as 機番"
            strSql &= ",I.ロットNo as ロットNo"
            strSql &= ",I.備考 as 備考"
            strSql &= ",I.印刷枚数 as 印刷枚数"
            strSql &= ",I.入力日 as 入力日"
            strSql &= ",I.入力者 as 入力者"
            strSql &= ",I.修正日 as 修正日"
            strSql &= ",I.修正者 as 修正者"
            strSql &= ",I.実績反映 as 実績反映"
            strSql &= ",S.社員名 as 社員名"
            strSql &= ",(select sum(仕上りm) from " & gcTblIrMeisai & " M where M.IR実績ID = I.IR実績ID) as 仕上り総m数"
            strSql &= ",(select TOP 1 ロットNo from " & gcTblIrMeisai & " M where M.IR実績ID = I.IR実績ID order by SEQ) as 明細ロットNo"
            strSql &= " FROM " & gcTblIrKihon & " I"
            strSql &= " INNER JOIN " & gcTblSyainMst & " S ON I.入力者 = S.社員コード"
            strSql &= " WHERE 1 = 1"
            If txtJyutyuNo.Text.Trim <> "" Then
                strSql &= " and 注文No = " & gfncMakeSqlValue(txtJyutyuNo.Text, 1)
            End If
            If txtShiyouNo.Text.Trim <> "" Then
                strSql &= " and 仕様書No = " & gfncMakeSqlValue(txtShiyouNo.Text, 0)
            End If
            If cboKiban.SelectedValue <> "" Then
                strSql &= " and 機番 = " & gfncMakeSqlValue(cboKiban.SelectedValue, 0)
            End If
            If txtLotNo.Text.Trim <> "" Then
                strSql &= " and ロットNo = " & gfncMakeSqlValue(txtLotNo.Text, 0)
            End If
            If gfncNVL(ndtp_from.Value) <> "" Then
                strSql &= " and 加工日 >= " & gfncMakeSqlValue(Date.Parse(ndtp_from.Value).ToString("yyyy/MM/dd"), 0)
            End If
            If gfncNVL(ndtp_to.Value) <> "" Then
                strSql &= " and 加工日 <=" & gfncMakeSqlValue(Date.Parse(ndtp_to.Value).ToString(" yyyy/MM/dd"), 0)
            End If
            strSql &= " Order by IR実績ID"
            If Not gfncbol_RecordGet(strSql, dt) Then
                Exit Sub
            End If
            If dt.Rows.Count > gcintRowMax Then
                MessageBox.Show("一覧表示対象件数が" & gcintRowMax & "件を超えています" & Environment.NewLine & "検索条件を見直してください", "表示リミット超え", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                grdList.RowCount = 0
                Exit Sub
            End If

            lblCnt.Text = dt.Rows.Count.ToString("#,0") & "件"

            ' データをグリッドに表示
            ' 描画を止める
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            grdList.SuspendLayout()
            grdList.RowCount = 0

            For lLoop = 0 To grdList.ColumnCount - 1
                grdList.Columns(lLoop).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            Next

            For Each dr As DataRow In dt.Rows
                Dim item As New DataGridViewRow
                item.CreateCells(grdList)
                Dim intRow As Integer = grdList.Rows.Add(item)

                grdList("No", intRow).Value = intRow + 1
                grdList("加工日", intRow).Value = gfncFormatDateTime(dr("加工日"), 0)
                grdList("機番", intRow).Value = dr("機番").ToString
                grdList("ロットNo", intRow).Value = dr("明細ロットNo").ToString
                grdList("受注書No", intRow).Value = dr("注文No").ToString
                grdList("仕様書No", intRow).Value = dr("仕様書No").ToString
                grdList("製品名", intRow).Value = dr("製品名").ToString
                grdList("仕上り総m数", intRow).Value = CInt(gfncNVL(dr("仕上り総m数").ToString, "0")).ToString("#,0")
                grdList("入力日", intRow).Value = gfncFormatDateTime(dr("入力日"), 0)
                grdList("入力者", intRow).Value = dr("社員名").ToString
                grdList("IR実績ID", intRow).Value = dr("IR実績ID").ToString
                grdList("ロットNo基本", intRow).Value = dr("ロットNo").ToString
            Next

            ' 列幅は自動とし、再描画
            For lLoop = 0 To grdList.ColumnCount - 1
                grdList.Columns(lLoop).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next
            grdList.ResumeLayout(True)

            Exit Sub

        Catch ex As Exception
            gsubExceptionProc(ex.Message, mstrfrmId, Reflection.MethodBase.GetCurrentMethod.Name)
            Exit Sub
        Finally
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End Try
    End Sub


    ' コントロール初期化
    Private Sub subFrmClear()
        txtJyutyuNo.Text = ""
        txtShiyouNo.Text = ""
        ndtp_from.Value = DBNull.Value
        ndtp_to.Text = Now
        gsubCboKiban(cboKiban, True, gstrKiban, 0)  ' 機番コンボ値セット
        txtLotNo.Text = ""
        grdList.Rows.Clear()

        txtJyutyuNo.Focus()
    End Sub


End Class
