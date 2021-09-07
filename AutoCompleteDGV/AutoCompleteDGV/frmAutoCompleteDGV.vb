#Region "ABOUT"
' / --------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com/
' /
' / Purpose: AutoComplete TextBox in dataGridView.
' / Microsoft Visual Basic .NET (2010) + MS Access 2007
' /
' / This is open source code under @Copyleft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------
#End Region

Imports System.Data.OleDb

Public Class frmAutoCompleteDGV
    Dim Conn As OleDb.OleDbConnection
    Dim DA As New System.Data.OleDb.OleDbDataAdapter()
    Dim DS As New System.Data.DataSet()
    'Dim DR As OleDbDataReader
    'Dim DT As New DataTable
    'Dim Cmd As New System.Data.OleDb.OleDbCommand
    'Dim strSQL As String

    '// Connect MS Access DataBase
    Function ConnectDataBase() As OleDb.OleDbConnection
        Return New OleDb.OleDbConnection( _
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
            MyPath(Application.StartupPath) & "data\" & "Stock.accdb;Persist Security Info=True")
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    ' / AppPath = C:\My Project\bin\debug
    ' / Replace "\bin\debug" with "\"
    ' / Return : C:\My Project\
    Function MyPath(ByVal AppPath As String) As String
        '/ MessageBox.Show(AppPath);
        MyPath = AppPath.ToLower.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '/ Return Value
        '// If not found folder then put the \ (BackSlash) at the end.
        If Microsoft.VisualBasic.Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function

    Private Sub frmAutoCompleteDGV_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Conn = ConnectDataBase()
        Dim colProduct As New DataGridViewTextBoxColumn
        colProduct.Name = "ProductName"
        colProduct.HeaderText = "Product Name"
        dgvData.Columns.Add(colProduct)
        dgvData.Columns(0).Width = 200
        With dgvData
            .RowHeadersVisible = True
            .AllowUserToAddRows = True
            .AllowUserToDeleteRows = True
            .AllowUserToResizeRows = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .ReadOnly = False
            .Font = New Font("Century Gothic", 10)
            '// Even-Odd Color
            .AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue
            '// Header Styles.
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.Black
                .Font = New Font("Century Gothic", 11, FontStyle.Bold)
                .WrapMode = DataGridViewTriState.False
            End With
        End With

    End Sub

    Private Sub dgvData_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvData.EditingControlShowing
        '// Check Column Name.
        If dgvData.Columns(dgvData.CurrentCell.ColumnIndex).Name = "ProductName" Then
            Dim TB As TextBox = e.Control
            If IsNothing(TB) = False Then
                TB.AutoCompleteMode = AutoCompleteMode.Suggest
                TB.AutoCompleteSource = AutoCompleteSource.CustomSource
                If Conn.State = ConnectionState.Closed Then Conn.Open()
                Dim AutoComplete As New AutoCompleteStringCollection()
                DA = New OleDbDataAdapter("SELECT Product.ProductName FROM Product", Conn)
                DS = New DataSet
                DA.Fill(DS)
                For i = 0 To DS.Tables(0).Rows.Count - 1
                    AutoComplete.Add(DS.Tables(0).Rows(i)("ProductName"))
                Next
                TB.AutoCompleteCustomSource = AutoComplete
                DS.Dispose()
                DA.Dispose()
                Conn.Close()
            End If
        End If

    End Sub

    Private Sub frmAutoCompleteDGV_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Conn.State = ConnectionState.Open Then Conn.Close()
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub
End Class
