Attribute VB_Name = "systemButton"
'created by nandur93 12/07/2017 http://nandur93.blogspot.com/VBA
'library ini untuk tutup form menggunakan X merah pojok kiri atas
'this library as to ask user some action if RED X button clicked
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'when X button clicked
    If CloseMode = 0 Then 'ketika X di klik
        'defined as integer
        Dim xClose As Integer 'dim variablenya sebagai integer
        'ask user 2 question before quit between YES or NO
        xClose = MsgBox("Tutup Form?", vbYesNo + vbQuestion, "Keluar") 'tanyakan mau keluar apa tidak
            'if YES button clicked
            If xClose = vbYes Then 'jika klik IYA
            '//put your logic here
            '...
            'here for example, my logic is Unload Me (close the userform)
                Unload Me 'keluar dari aplikasi
                    'else or if NO button clicked
                    Else
                '//put your logic here
                '...
                'here for example, my logic is CANCEL the action (don't close the userform)
                Cancel = True
            End If
    End If
End Sub
