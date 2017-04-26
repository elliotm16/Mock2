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

        Dim StaffsData() As String = File.ReadAllLines("staffdetails.txt")

        ' Call the function which generates the Staff ID
        GenerateID(StaffsData)

        ' If the validation checks have been passed
        If Validation() = True Then

            Dim sw As New System.IO.StreamWriter("staffdetails.txt", True)

            StaffData.StaffID = LSet(txtStaffID.Text, 4)
            StaffData.FirstName = LSet(txtFirstName.Text, 20)
            StaffData.Surname = LSet(txtSurname.Text, 20)
            StaffData.EmailAddress = LSet(txtEmailAddress.Text, 30)
            StaffData.PhoneNumber = LSet(txtPhoneNumber.Text, 11)

            ' Write the data in the structure to the textfile
            sw.WriteLine(StaffData.StaffID & StaffData.FirstName & StaffData.Surname & StaffData.EmailAddress & StaffData.PhoneNumber)

            ' Must be closed after use
            sw.Close()

            ' Output that the file has been saved
            MsgBox("File Saved!")

            ClearTextboxes()

        End If

    End Sub

    Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click

        Dim StaffFound As Integer
        StaffFound = 0

        Dim OneFound As Integer

        Dim StaffsData() As String = File.ReadAllLines("staffdetails.txt")

        ' If a Staff ID isn't been searched for
        If txtStaffID.Text = "" Then

            For i = 0 To UBound(StaffsData)

                If ((Trim(Mid(StaffsData(i), 5, 20)) = txtFirstName.Text) Or txtFirstName.Text = "") And ((Trim(Mid(StaffsData(i), 25, 20)) = txtSurname.Text) Or txtSurname.Text = "") And ((Trim(Mid(StaffsData(i), 45, 20)) = txtEmailAddress.Text) Or txtEmailAddress.Text = "") And ((Trim(Mid(StaffsData(i), 75, 11)) = txtPhoneNumber.Text) Or txtPhoneNumber.Text = "") Then

                    StaffFound = StaffFound + 1

                End If

            Next

            If StaffFound = 1 Then

                MsgBox("One member of staff was found.")

                txtStaffID.Text = (Trim(Mid(StaffsData(OneFound), 1, 4)))
                txtFirstName.Text = (Trim(Mid(StaffsData(OneFound), 5, 20)))
                txtSurname.Text = (Trim(Mid(StaffsData(OneFound), 25, 20)))
                txtEmailAddress.Text = (Trim(Mid(StaffsData(OneFound), 45, 30)))
                txtPhoneNumber.Text = (Trim(Mid(StaffsData(OneFound), 75, 11)))

            ElseIf StaffFound > 0 Then

                MsgBox(StaffFound & " members of staff were found.")

            Else

                MsgBox("No members of staff were found.")

            End If

        Else

            For i = 0 To UBound(StaffsData)

                If Trim(Mid(StaffsData(i), 1, 4)) = txtStaffID.Text Then

                    StaffFound = 0

                    MsgBox("A member of staff with this Staff ID has been found.")

                    txtFirstName.Text = (Trim(Mid(StaffsData(i), 5, 20)))
                    txtSurname.Text = (Trim(Mid(StaffsData(i), 25, 20)))
                    txtEmailAddress.Text = (Trim(Mid(StaffsData(i), 45, 30)))
                    txtPhoneNumber.Text = (Trim(Mid(StaffsData(i), 75, 11)))

                    Exit Sub

                End If

            Next

            If StaffFound = 0 Then

                MsgBox("A member of staff with this Staff ID has not been found.")

            End If

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

    Public Sub GenerateID(StaffsData)

        Dim CurrentStaffID
        CurrentStaffID = 0

        For i = 0 To UBound(StaffsData)

            If Trim(Mid(StaffsData(i), 1, 4)) = CurrentStaffID Then

                CurrentStaffID = CurrentStaffID + 1

                txtStaffID.Text = CurrentStaffID

            End If

        Next

    End Sub

    Public Sub ClearTextboxes()

        txtStaffID.Text = ""
        txtFirstName.Text = ""
        txtSurname.Text = ""
        txtEmailAddress.Text = ""
        txtPhoneNumber.Text = ""

    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click

        ClearTextboxes()

    End Sub

End Class