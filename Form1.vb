Imports System.Data.OleDb

Public Class Form1

    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\imafo\Documents\School\ICT3611\Assign6_Database\Rugby.accdb;Persist Security Info=False"
    Dim connection As New OleDbConnection(connectionString)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            connection.Open()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

        LoadTeams()
        TabOne()
        TabTwo()
        TabThree()
        TabFour()
        TabFive()
    End Sub

    Private Sub LoadTeams()
        Dim query As String = "SELECT Team FROM Teams"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        lbTeams5.Items.Clear()

        While reader.Read()
            Dim teamName As String = reader("Team").ToString()
            lbTeams5.Items.Add(teamName)
        End While

        reader.Close()
    End Sub

    Private Sub lbTeams5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbTeams5.SelectedIndexChanged
        If lbTeams5.SelectedIndex <> -1 Then
            Dim selectedTeamName As String = lbTeams5.SelectedItem.ToString()
            DisplayPlayers(selectedTeamName)
        End If
    End Sub

    Private Sub DisplayPlayers(teamName As String)
        Dim query As String = "SELECT Player, AVG(Points * 1.0 / Games) AS AveragePoints " &
                              "FROM Players WHERE Team = @TeamName " &
                              "GROUP BY Player " &
                              "ORDER BY AveragePoints DESC"

        Dim command As New OleDbCommand(query, connection)
        command.Parameters.AddWithValue("@TeamName", teamName)
        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            gDGVTeams5.DataSource = dataTable
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub TabOne()
        'Tab Number One Queries
        Dim query As String = "Select * FROM Teams"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        ListBox1.Items.Clear()

        While reader.Read()
            Dim teamName As String = reader("Team").ToString()
            ListBox1.Items.Add(teamName)
        End While

        reader.Close()
    End Sub

    Private Sub TabTwo()
        'Tab Number Two Queries
        Dim query As String = "SELECT Team, Stadium, AVG(Points * 1.0 / Games) AS AveragePoints FROM Teams GROUP BY Team, Stadium"
        Dim command As New OleDbCommand(query, connection)
        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            gDGV2.DataSource = dataTable
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub TabThree()
        ' Tab Number Three Queries
        Dim query As String = "SELECT Player FROM Players WHERE Points = (SELECT MAX(Points) FROM Players)"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        lbHighScore3.Items.Clear()

        While reader.Read()
            Dim playerName As String = reader("Player").ToString()
            lbHighScore3.Items.Add(playerName)
        End While

        reader.Close()
    End Sub

    Private Sub TabFour()
        ' Tab Number Four Queries
        Dim query As String = "SELECT Player FROM Players WHERE (Points * 1.0 / Games) = (SELECT MAX(Points * 1.0 / Games) FROM Players)"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        lbHighAve4.Items.Clear()

        While reader.Read()
            Dim playerName As String = reader("Player").ToString()
            lbHighAve4.Items.Add(playerName)
        End While

        reader.Close()
    End Sub

    Private Sub TabFive()

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex <> -1 Then
            Dim selectedTeamName As String = ListBox1.SelectedItem.ToString()
            Dim query As String = "SELECT Stadium FROM Teams WHERE Team = @TeamName"
            Dim command As New OleDbCommand(query, connection)
            command.Parameters.AddWithValue("@TeamName", selectedTeamName)

            Try
                Dim stadium As Object = command.ExecuteScalar()
                If stadium IsNot Nothing AndAlso Not IsDBNull(stadium) Then
                    txtNo1HomeStadium.Text = stadium.ToString()
                Else
                    txtNo1HomeStadium.Text = "No information"
                End If
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        Else
            txtNo1HomeStadium.Text = ""
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs)
        If connection.State = ConnectionState.Open Then
            connection.Close()
        End If
    End Sub

    Private Sub Guna2PictureBox1_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox1.Click
        Application.Exit()
    End Sub

    Private Sub Guna2PictureBox2_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox2.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub gDGV2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles gDGV2.CellContentClick

    End Sub
End Class
