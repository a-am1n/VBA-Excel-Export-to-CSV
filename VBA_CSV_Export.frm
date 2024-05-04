VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBA_CSV_Export 
   Caption         =   "Экспорт CSV"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5460
   OleObjectBlob   =   "VBA_CSV_Export.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VBA_CSV_Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim csv_path As String  ' Адрес будущего файла CSV
Dim csv_name As String  ' Имя будущего файла CSV
Dim csv_sep As String   ' Разделить будущего файла CSV

Sub BT_cancel_Click()
'    Кнопка отмены
    Unload Me
End Sub
Sub BT_export_Click()
    On Error GoTo ErrMess
    
    If Len(csv_path) = 0 Or csv_path = "(необходимо указать)" Then
        MsgBox _
            prompt:="Необходимо указать папку, где будет сохранён экспортированный файл!", _
            Buttons:=vbOKOnly + vbCritical, _
            Title:="Ошибка!"
        Exit Sub
    Else
    End If
    
    If Len(csv_name) = 0 Then
        MsgBox _
            prompt:="Необходимо указать имя будущего файла!", _
            Buttons:=vbOKOnly + vbCritical, _
            Title:="Ошибка!"
        Exit Sub
    Else
    End If
    
    If sep_t_z.Value = True Then
        csv_sep = ";"
    ElseIf sep_z.Value = True Then
        csv_sep = ","
    End If
    
    Dim txtStream As Object

    Set txtStream = CreateObject("ADODB.Stream")
    With txtStream
        .Charset = "UTF-8"  ' Кодировка
        .Type = 2
        .Open
    End With
    
    Dim arrRange() As Variant
    arrRange = Selection
       
    Dim cell As Long
    Dim row As Long
    
        
    For row = 1 To UBound(arrRange, 1)
        For cell = 1 To UBound(arrRange, 2)
            If cell <> UBound(arrRange, 2) Then
                txtStream.WriteText (arrRange(row, cell) & csv_sep)
            Else
                txtStream.WriteText (arrRange(row, cell))
            End If
        Next cell
        txtStream.WriteText (vbNewLine)
    Next row
    
    txtStream.SaveToFile csv_path & "\" & csv_name & ".csv", 2
    txtStream.Close
    Set txtStream = Nothing
    
    MsgBox "Файл готов!"
    BT_cancel_Click
    Exit Sub
    
ErrMess:
    MsgBox "Возникла ошибка при экспорте!", vbCritical
    Err.Clear
    Call BT_cancel_Click
End Sub

Sub file_name_Change()
    csv_name = file_name.Text
End Sub

Private Sub Image1_Click()
    Call BT_export_Click
End Sub
Private Sub Image2_Click()
    Call BT_cancel_Click
End Sub

Private Sub Image3_Click()
    Call select_folder_Click
End Sub

Sub select_folder_Click()
    Dim od As FileDialog
    
    Set od = Application.FileDialog(msoFileDialogFolderPicker)
        With od
        .Title = "Выбрать папку для сохранения CSV"
        If .Show = 0 Then Exit Sub
        .ButtonName = "Выбрать папку"
        csv_path = .SelectedItems(1)
        file_path = csv_path & "\"
    End With
End Sub

Sub UserForm_Initialize()
'    Подгружаем папка назначения по-умолчанию (папка самого файла), если файл не сохранён, то выводим, что необходимо указать

    If Selection.Count < 2 Then
        MsgBox _
            prompt:="Сначало выделите диапазон данных для экспортирования!", _
            Buttons:=vbCritical + vbOKOnly, _
            Title:="Ошибка!"
        End
    End If
    Image1.Picture = Application.CommandBars.GetImageMso("TableExportTableToSharePointList", 16, 16)
    Image2.Picture = Application.CommandBars.GetImageMso("WindowClose", 16, 16)
    Image3.Picture = Application.CommandBars.GetImageMso("SaveSentItemRecentlyUsedFolder", 16, 16)
    If Len(ActiveWorkbook.Path) = 0 Then
        file_path = "(необходимо указать)"
    Else
        file_path = ActiveWorkbook.Path & "\"
        csv_path = file_path
    End If
    
    system_sep.Value = True
    
    csv_sep = Application.International(xlColumnSeparator)
    
    export_range = ActiveSheet.Name & "!" & Selection.Address(False, False)
    
End Sub


