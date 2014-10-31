VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Изменение категорий"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8865
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Dim oApp As New Outlook.Application
    Dim oExp As Outlook.Explorer
    Dim oSel As Outlook.Selection   ' You need a selection object for getting the selection.
    Dim oItem As Object             ' You don't know the type yet.
    

    Dim oAppointItem As Outlook.AppointmentItem
    Dim oContactItem As Outlook.ContactItem
    Dim oMailItem As Outlook.MailItem
    Dim oJournalItem As Outlook.JournalItem
    Dim oNoteItem As Outlook.NoteItem
    Dim oTaskItem As Outlook.TaskItem








Private Sub btnAddToSubject_Click()

    Dim strMessageClass As String
    

    Set oExp = oApp.ActiveExplorer  ' Get the ActiveExplorer.
    Set oSel = oExp.Selection       ' Get the selection.
    
    For i = 1 To oSel.Count         ' Loop through all the currently .selected items
        Set oItem = oSel.Item(i)    ' Get a selected item.
       
        ' You need the message class to determine the type.
        strMessageClass = oItem.MessageClass
        
        If (strMessageClass = "IPM.Note") Then          ' Mail Entry.
            Set oMailItem = oItem
            If Len(txtAddToSubject.text) > 0 Then 'если пользователь просил что-то добавлять в сабжект
                'если в сабжекте отсутствует строка которую просил добавить пользователь
                If InStr(1, oMailItem.Subject, txtAddToSubject, vbTextCompare) = 0 Then
            
                    oMailItem.Subject = txtAddToSubject.text & " " & oMailItem.Subject
                    'если не сделать сэйв то сохранить только первое изменение - очень неочевидная штука
                    oMailItem.Save
                End If
            End If
            
        End If
        
    Next i

End Sub






Private Sub btnAddToSubject_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'по нажатию ескэйп независимо от того какой контрол активный форма должна закрыться
    If KeyAscii = 27 Then Unload UserForm1

End Sub





Private Sub btnApply_Click()

    Dim strMessageClass As String
    Dim objMailItem As Object
    

    Set oExp = oApp.ActiveExplorer  ' Get the ActiveExplorer.
    Set oSel = oExp.Selection       ' Get the selection.
    
    For i = 1 To oSel.Count         ' Loop through all the currently .selected items
        Set oItem = oSel.Item(i)    ' Get a selected item.
       ' DisplayInfo oItem           ' Display information about it.
    
        ' You need the message class to determine the type.
        strMessageClass = oItem.MessageClass
        
        'If (strMessageClass = "IPM.Note") Then          ' Mail Entry.
        Set objMailItem = oItem
        
        On Error Resume Next
        objMailItem.Categories = txtCategInput.text
        'если не сделать сэйв то сохранить только первое изменение - очень неочевидная штука
        objMailItem.Save
        On Error GoTo 0
            
        'End If
        
    Next i

End Sub








Private Sub btnApply_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'по нажатию ескэйп независимо от того какой контрол активный форма должна закрыться
    If KeyAscii = 27 Then Unload UserForm1
End Sub






Private Sub btnCancel_Click()
Unload UserForm1
End Sub






Private Sub btnCancel_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'по нажатию ескэйп независимо от того какой контрол активный форма должна закрыться
    If KeyAscii = 27 Then Unload UserForm1
End Sub













Private Sub btnMove_Click()
    
    'Dim strMessageClass As String
    Dim objItem As Object
    
    Dim strFolderEntyID As String
    strFolderEntryID = Interaction.GetSetting("outlook toys", "outlook categorizer", "ArchiveFolderEntryID", "")
    If Len(strFolderEntryID) = 0 Then
        SelectFolder
        strFolderEntryID = Interaction.GetSetting("outlook toys", "outlook categorizer", "ArchiveFolderEntryID", "")
    End If



    On Error Resume Next
    Dim objFolder As Outlook.Folder
    Set objFolder = Application.GetNamespace("MAPI").GetFolderFromID(strFolderEntryID)
    If Err.Number <> 0 Then Set objFolder = Nothing
    On Error GoTo 0
    
    If objFolder Is Nothing Then
        SelectFolder
        strFolderEntryID = Interaction.GetSetting("outlook toys", "outlook categorizer", "ArchiveFolderEntryID", "")
        If Len(strFolderEntryID) = 0 Then
            MsgBox "Archive folder not selected", vbOKOnly
            Exit Sub
        End If
        
        On Error Resume Next
        Set objFolder = Application.GetNamespace("MAPI").GetFolderFromID(strFolderEntryID)
        If Err.Number <> 0 Then
            MsgBox "Can't find archive folder :(", vbOKOnly
            Exit Sub
        End If
        On Error GoTo 0
               
    End If
    
    
    
    If objFolder Is Nothing Then
            MsgBox "Can't find archive folder :( :(", vbOKOnly
            Exit Sub
    End If

    Set oExp = oApp.ActiveExplorer  ' Get the ActiveExplorer.
    Set oSel = oExp.Selection       ' Get the selection.
    
    For i = 1 To oSel.Count         ' Loop through all the currently .selected items
        Set oItem = oSel.Item(i)    ' Get a selected item.
       ' DisplayInfo oItem           ' Display information about it.
    
        ' You need the message class to determine the type.
     '   strMessageClass = oItem.MessageClass
        
        'If (strMessageClass = "IPM.Note") Then          ' Mail Entry.
        On Error Resume Next
        Set objItem = oItem
        'MsgBox oMailItem.Subject
        'перемещаем в архив
        objItem.Move objFolder
            
        'End If
        On Error GoTo 0
    Next i
    
    Unload UserForm1

End Sub






Private Sub btnMove_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'по нажатию ескэйп независимо от того какой контрол активный форма должна закрыться
    If KeyAscii = 27 Then Unload UserForm1
End Sub




Private Sub btnMove2_Click()
    'делаем тоже самое что и по нажатию первой кнопки "Переместить в архив и закрыть"
    btnMove_Click
End Sub




Private Sub btnMove2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'по нажатию ескэйп независимо от того какой контрол активный форма должна закрыться
    If KeyAscii = 27 Then Unload UserForm1
End Sub






Private Sub btnSelectArchFolder_Click()

    'open standard outlook folder dialog and save selected folder to registry
    SelectFolder
       
End Sub





Private Sub txtAddToSubject_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'по нажатию ескэйп независимо от того какой контрол активный форма должна закрыться
    If KeyAscii = 27 Then Unload UserForm1

End Sub

Private Sub txtCategInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'по нажатию ескэйп независимо от того какой контрол активный форма должна закрыться
    If KeyAscii = 27 Then Unload UserForm1
End Sub







Private Sub UserForm_Initialize()

    Dim strMessageClass As String
    
    Dim strCurrentCategories As String
    strCurrentCategories = ""
    


    Set oExp = oApp.ActiveExplorer  ' Get the ActiveExplorer.
    Set oSel = oExp.Selection       ' Get the selection.
    
    For i = 1 To oSel.Count         ' Loop through all the currently .selected items
        Set oItem = oSel.Item(i)    ' Get a selected item.
       ' DisplayInfo oItem           ' Display information about it.
    
        ' You need the message class to determine the type.
        strMessageClass = oItem.MessageClass
        
        If (strMessageClass = "IPM.Note") Then          ' Mail Entry.
            Set oMailItem = oItem
            'MsgBox oMailItem.Subject
            'oMailItem.Categories = "zzz_test, fuermb"
            Dim strCurrentCateg As String
            
            'получили строку типа " REFWorkProcedure; RZ23; FUSTANDIN"
            'описывающую категории для каждого конкретного письма
            strCurrentCateg = oMailItem.Categories
            
            'нормализовали строку- убрали  пробелы после точки с запятой
            strCurrentCateg = Replace(oMailItem.Categories, " ", "")
            
            'нормализовали строку - добавили в конец точку с запятой
            '                   (только если строка непустая - т.е. категории у письма есть)
            If Len(strCurrentCateg) > 0 Then
                strCurrentCateg = strCurrentCateg & ";"
            End If
            
            'добавили нормализованную строку к единой строке описывающей все категории выделенных писем
            strCurrentCategories = strCurrentCategories & strCurrentCateg
        End If
        
    Next i
    
    
    'хотим алфавитно отсортировать эту строку
    Dim arrCurrentCategories() As String
    arrCurrentCategories = Split(strCurrentCategories, ";")
    
    If LBound(arrCurrentCategories) < UBound(arrCurrentCategories) Then ' без этой проверки падает сортировка
        SortStringArray arrCurrentCategories, LBound(arrCurrentCategories), UBound(arrCurrentCategories)
    End If
    
    
    
    
    'дедублицируем категории в строке (т.е. удаляем повторяющиеся)
    '     путём использования категорий как ключей в коллекции
    Dim cltTemp As New Collection, varTmp As Variant
    For Each varTmp In arrCurrentCategories
        On Error Resume Next
        'здесь будет ошибка при добавлении дублицированной категории, но мы её игнорим
        If Len(varTmp) > 0 Then 'пустые категории отбрасываем - это побочный эффект сплита
            cltTemp.Add varTmp, varTmp
        End If
        On Error GoTo 0
    Next
    
    
    
    
    'формируем итоговую строку описывающую все категории
    Dim varTmp1 As Variant
    
    'обнуляем итоговую строку
    strCurrentCategories = ""
    For Each varTmp1 In cltTemp
        'strCurrentCategories = strCurrentCategories & varTmp1 & "~~~"
        'UserForm1.txtCategories.Text = UserForm1.txtCategories.Text & varTmp1 & "; "
        UserForm1.txtCategInput.text = UserForm1.txtCategInput.text & varTmp1 & "; "
    Next
    
    
    
    'отображаем количество выбранных писем
    lblItemCount.Caption = "items:" & oSel.Count
    
    'draw actual archive folder name on buttons
    OnArchiveFolderChange
    
    
    'strCurrentCategories = Join(arrCurrentCategories, "~~~")
    '
    
    
    
    
    'отображаем категории на форме
    'UserForm1.txtCategories.Text = strCurrentCategories
    'UserForm1.txtCategInput.Text = strCurrentCategories

End Sub




Public Sub SortStringArray(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long
  
  If inHi <= inLo Then
    Err.Raise 40915, , "Попытка отсортировать пустой массив"
  End If

  tmpLow = inLow
  tmpHi = inHi
  
  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then SortStringArray vArray, inLow, tmpHi
  If (tmpLow < inHi) Then SortStringArray vArray, tmpLow, inHi

End Sub




'when user change Archive folder we need to refresh captions on some buttons
Private Sub OnArchiveFolderChange()


Dim strFolderEntyID As String
strFolderEntryID = Interaction.GetSetting("outlook toys", "outlook categorizer", "ArchiveFolderEntryID", "")

On Error Resume Next
Dim strFolderName As String
strFolderName = Application.GetNamespace("MAPI").GetFolderFromID(strFolderEntryID)
If Err.Number <> 0 Then strFolderName = ""
On Error GoTo 0

If Len(strFolderName) = 0 Then
    'Archive folder not setted
    btnMove.Caption = "Move to <ARCHIVE> and close"
    btnMove2.Caption = "Move to <ARCHIVE> and close"
Else
    btnMove.Caption = "Move to \" & strFolderName & " and close"
    btnMove2.Caption = "Move to \" & strFolderName & " and close"
End If


'btnMove.Caption = "Move to

End Sub





Private Sub SelectFolder()
    Dim objFolder As Outlook.Folder
    Dim strSelectedFolder As String
    Set objFolder = Application.GetNamespace("MAPI").PickFolder
           
    If objFolder Is Nothing Then
        'folder have not been selected
        Interaction.SaveSetting "outlook toys", "Outlook Categorizer", "ArchiveFolderEntryID", ""
    Else
        Interaction.SaveSetting "outlook toys", "Outlook Categorizer", "ArchiveFolderEntryID", objFolder.EntryID
    End If
    
    'change captions
    OnArchiveFolderChange

End Sub
