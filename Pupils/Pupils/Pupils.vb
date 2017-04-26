Imports System.IO

Public Class Pupils
    Private Structure Pupil
        Public PupilId As String
        Public FirstName As String
        Public Surname As String
        Public HomeAdd As String                  'Creating the structure that will hold the  data.
        Public EmergencyNo As String
    End Structure


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim PupilData As New Pupil
        Dim sw As New System.IO.StreamWriter("Pupils.txt", True)
        PupilData.PupilId = LSet(txtPupilID.Text, 50)
        PupilData.FirstName = LSet(txtFirstName.Text, 50)
        PupilData.Surname = LSet(txtSurname.Text, 50)
        PupilData.HomeAdd = LSet(txtHomeAdd.Text, 50)                      'Filling the structure with data.
        PupilData.EmergencyNo = LSet(txtEmergencyNo.Text, 50)

        sw.WriteLine(PupilData.PupilId & PupilData.FirstName & PupilData.Surname & PupilData.HomeAdd & PupilData.EmergencyNo)
        sw.Close()                                                                  'Always need to close afterwards
        MsgBox("File Saved!")

    End Sub

    Private Sub Pupils_Load() Handles MyBase.Load
        If Dir$("Pupils.txt") = "" Then
            Dim sw As New StreamWriter("Pupils.txt", True)    'This makes sure there is actually a database to enter/read data. If not, it creates a new blank one.
            sw.WriteLine("")
            sw.Close()
            MsgBox("A new file has been created", vbExclamation, "Warning!")
        End If
    End Sub

    Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click
        Dim CountGot As Integer
        CountGot = 0
        Dim PupilCount As Integer
        PupilCount = 0


        '                ". . . ." Indicates missing or broken code.


        Dim PupilData() As String = File.ReadAllLines("Pupils.txt")
        For i = 0 To UBound(PupilData)
           
            CountGot = 0
            If Trim(Mid(PupilData(i), 1, 50)) = txtPupilID.Text And Not txtPupilID.Text = "" Then CountGot = CountGot + 1 'Counting how many attributes follow the search
            If Trim(Mid(PupilData(i), 51, 50)) = txtFirstName.Text And Not txtFirstName.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(PupilData(i), 101, 50)) = txtSurname.Text And Not txtSurname.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(PupilData(i), 151, 50)) = txtHomeAdd.Text And Not txtHomeAdd.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(PupilData(i), 201, 50)) = txtEmergencyNo.Text And Not txtEmergencyNo.Text = "" Then CountGot = CountGot + 1
            If CountGot > 0 Then PupilCount = PupilCount + 1
        Next i
        MsgBox("There were: " & PupilCount & " Pupils Found")

    End Sub
End Class
