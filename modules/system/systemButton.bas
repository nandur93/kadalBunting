'Attribute VB_Name = "systemButton"
'created by nandur93 12/07/2017 https://nandur93.com/VBA
'update 18/04/2019
'+fix code
'+add simple tutorial
'update 01/05/2019
'+fix indent
'+fix simple tutorial
'library ini untuk tutup form menggunakan X merah pojok kiri atas
'this library is to ask user some action if RED X button clicked
'copy and paste this code to end of your UserForm module
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'when X button clicked
    If CloseMode = 0 Then                        'ketika X di klik
        'defined as integer
        Dim xClose As Integer                    'dim variablenya sebagai integer
        'ask user two question before quit between YES or NO
        xClose = MsgBox("Tutup Form?", vbYesNo + vbQuestion, "Keluar") 'tanyakan mau keluar apa tidak
        'if YES button clicked
        If xClose = vbYes Then                   'jika klik IYA
            '//put your logic here
            '...
            'here for example, my logic is Unload Me (close the userform)
            Unload Me                            'keluar dari aplikasi
            'else or if NO button clicked
        Else
            '//put your logic here
            '...
            'here for example, my logic is CANCEL the action (don't close the userform)
            Cancel = True
        End If
    End If
End Sub
