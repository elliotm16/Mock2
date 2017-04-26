Imports System.IO

Public Class StaffDetails

    ' Structure to store Staff data
    Private Structure Staff

        Public StaffID As String ' Used to uniquely identify a member of staff
        Public FirstName As String
        Public Surname As String
        Public EmailAddress As String
        Public PhoneNumber As String

    End Structure

    Private Sub StaffDetails_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        ' If the textfile doesn't already exist
        If Dir$("staffdetails.txt") = "" Then

            ' Create the new textfile with this name
            Dim sw As New StreamWriter("staffdetails.txt", True)

            ' Write '0' to the textfile
            sw.WriteLine("0")

            ' Must be closed after use
            sw.Close()

            ' Output that a new textfile has been created
            MsgBox("A new file has been created", vbExclamation, "Warning!")

        End If

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim StaffData As New Staff

        Dim sw As New System.IO.StreamWriter("staffdetails.txt", True)

        GenerateID()

        If Validation() = True Then

            StaffData.StaffID = LSet(txtStaffID.Text, 4)
            StaffData.FirstName = LSet(txtFirstName.Text, 20)
            StaffData.Surname = LSet(txtSurname.Text, 20)
            StaffData.EmailAddress = LSet(txtEmailAddress.Text, 30)
            StaffData.PhoneNumber = LSet(txtPhoneNumber.Text, 11)

            sw.WriteLine(StaffData.StaffID & StaffData.FirstName & StaffData.Surname & StaffData.EmailAddress & StaffData.PhoneNumber)

            sw.Close()

            MsgBox("File Saved!")

        End If

    End Sub

    Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click

        Dim StaffFound As Integer
        StaffFound = 0

        Dim StaffsData() As String = File.ReadAllLines("staffdetails.txt")

        ' If a Staff ID isn't been searched for
        If txtStaffID.Text = "" Then

            For i = 0 To UBound(StaffsData)



            Next

        Else



        End If

    End Sub

    Public Function Validation()

        If txtFirstName.Text = "" Or txtSurname.Text = "" Or txtEmailAddress.Text = "" Or txtPhoneNumber.Text = "" Then

            ' Presence check
            MsgBox("Please enter a Name, Email Address and Phone Number.")

            Return False

        ElseIf (txtFirstName.Text).Length > 20 Or (txtSurname.Text).Length > 20 Then

            ' Length check
            MsgBox("First Name and Surname must be a maximum of 20 characters in length.")

            Return False

        ElseIf txtEmailAddress.Text.Contains("@") = False Then

            ' Format check
            MsgBox("Email Address must contain the '@' symbol.")

            Return False

        ElseIf IsNumeric(txtPhoneNumber.Text) = False Then

            ' Type check
            MsgBox("Phone Number must be numeric.")

            Return False

        Else

            Return True

        End If

    End Function

    Public Sub GenerateID()

        Dim StaffsData() As String = File.ReadAllLines("staffdetails.txt")

        Dim CurrentStaffID
        CurrentStaffID = 0

        For i = 0 To UBound(StaffsData)

            If Trim(Mid(StaffsData(i), 1, 50)) = CurrentStaffID Then

                CurrentStaffID = CurrentStaffID + 1

                txtStaffID.Text = CurrentStaffID

            End If

        Next

    End Sub

End Class