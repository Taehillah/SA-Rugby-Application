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
        LoadAveragesAbove3()
    End Sub

    Private Sub LoadTeams()
        Dim query As String = "SELECT Team FROM Teams"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        lbTeams5.Items.Clear()
        glbSARUs13.Items.Clear() 'just added

        While reader.Read()
            Dim teamName As String = reader("Team").ToString()
            lbTeams5.Items.Add(teamName)
            glbSARUs13.Items.Add(teamName) 'just added
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
                              "ORDER BY AVG(Points * 1.0 / Games) DESC"

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

    Private Sub LoadAveragesAbove3()
        Dim query As String = "SELECT DISTINCT AVG(Points * 1.0 / Games) AS AveragePoints " &
                              "FROM Players " &
                              "GROUP BY Player " &
                              "HAVING AVG(Points * 1.0 / Games) > 3.0 " &
                              "ORDER BY AVG(Points * 1.0 / Games) DESC"

        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        lbAve3of6.Items.Clear()

        While reader.Read()
            Dim averagePoints As Double = CDbl(reader("AveragePoints"))
            lbAve3of6.Items.Add(averagePoints.ToString("0.00"))
        End While

        reader.Close()
    End Sub

    Private Sub lbAve3of6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbAve3of6.SelectedIndexChanged
        If lbAve3of6.SelectedIndex <> -1 Then
            Dim selectedAverage As Double = CDbl(lbAve3of6.SelectedItem)
            DisplayPlayersWithAverage(selectedAverage)
        End If
    End Sub

    Private Sub DisplayPlayersWithAverage(average As Double)
        Dim query As String = "SELECT Player, Team " &
                              "FROM Players " &
                              "WHERE (Points * 1.0 / Games) > @Average " &
                              "ORDER BY Player"

        Dim command As New OleDbCommand(query, connection)
        command.Parameters.AddWithValue("@Average", average)
        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            gDGVTeam6.DataSource = dataTable
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub gCBtn7_Click(sender As Object, e As EventArgs) Handles gCBtn7.Click
        Dim averagePoints As Double = CalculateAveragePointsInCurrieCup()
        gtxtPts7.Text = averagePoints.ToString("0.00")
    End Sub

    Private Function CalculateAveragePointsInCurrieCup() As Double
        Dim query As String = "SELECT AVG(Points * 1.0 / Games) AS AveragePoints " &
                              "FROM Players " &
                              "WHERE Team IN (SELECT Team FROM Teams WHERE League = 'Currie Cup')"

        Dim command As New OleDbCommand(query, connection)

        Try
            Dim result As Object = command.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CDbl(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

        Return 0.0

    End Function
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs)
        If connection.State = ConnectionState.Open Then
            connection.Close()
        End If
    End Sub


    Private Sub gCBtnTotal8_Click(sender As Object, e As EventArgs) Handles gCBtnTotal8.Click
        Dim totalPoints As Integer = CalculateTotalPointsInCurrieCup()
        gtxtTotalPts8.Text = totalPoints.ToString()
    End Sub

    Private Function CalculateTotalPointsInCurrieCup() As Integer
        Dim query As String = "SELECT SUM(Points) AS TotalPoints " &
                              "FROM Teams " &
                              "WHERE League = 'Currie Cup'"

        Dim command As New OleDbCommand(query, connection)

        Try
            Dim result As Object = command.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CInt(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

        Return 0 ' Return 0 if there's an error or no result
    End Function

    Private Sub gCBtnPlayers9_Click(sender As Object, e As EventArgs) Handles gCBtnPlayers9.Click
        Dim playerCount As Integer = CountCurrieCupPlayers()
        gtxtPlayers9.Text = playerCount.ToString()
    End Sub

    Private Function CountCurrieCupPlayers() As Integer
        Dim query As String = "SELECT COUNT(*) AS PlayerCount " &
                              "FROM Players P " &
                              "WHERE EXISTS (SELECT 1 FROM Teams T WHERE T.Team = P.Team AND T.League = 'Currie Cup')"

        Dim command As New OleDbCommand(query, connection)

        Try
            Dim result As Object = command.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CInt(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

        Return 0 'If there are errors, the function returns 0
    End Function

    Private Sub gCBtnBloem10_Click(sender As Object, e As EventArgs) Handles gCBtnBloem10.Click
        Dim playerCount As Integer = CountBloemfonteinPlayers()
        gtxtBloem10.Text = playerCount.ToString()
    End Sub

    Private Function CountBloemfonteinPlayers() As Integer
        Dim query As String = "SELECT COUNT(*) AS PlayerCount " &
                              "FROM Players P " &
                              "WHERE EXISTS (SELECT 1 FROM Teams T WHERE T.Team = P.Team AND T.Location = 'Bloemfontein')"

        Dim command As New OleDbCommand(query, connection)

        Try
            Dim result As Object = command.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CInt(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

        Return 0 ' Return 0 if there's an error or no result
    End Function

    Private Sub RadioButton_CheckedChanged11(sender As Object, e As EventArgs) Handles gRBtnCurrieCup11.CheckedChanged, gRBtnSARugby11.CheckedChanged
        If gRBtnCurrieCup11.Checked Then
            DisplayTeamsByLeague("Currie Cup")
        ElseIf gRBtnSARugby11.Checked Then
            DisplayTeamsByLeague("SA Rugby")
        End If
    End Sub

    Private Sub DisplayTeamsByLeague(league As String)
        Dim query As String = "SELECT Team, AVG(Points * 1.0 / Games) AS AveragePoints " &
                              "FROM Teams " &
                              "WHERE League = @League " &
                              "GROUP BY Team " &
                              "ORDER BY AVG(Points * 1.0 / Games) DESC"

        Dim command As New OleDbCommand(query, connection)
        command.Parameters.AddWithValue("@League", league)
        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            gDGVTeams11.DataSource = dataTable
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub RadioButton_CheckedChanged12(sender As Object, e As EventArgs) Handles gRBtnCurrieCup12.CheckedChanged, gRBtnSARugby12.CheckedChanged
        If gRBtnCurrieCup12.Checked Then
            DisplayPlayers("Currie Cup", 40)
        ElseIf gRBtnSARugby12.Checked Then
            DisplayPlayers("SA Rugby", 40)
        End If
    End Sub

    Private Sub DisplayPlayers(league As String, minPoints As Integer)
        Dim query As String = "SELECT Players.Player, Players.Points, Teams.Stadium, AVG(Players.Points * 1.0 / Players.Games) AS AveragePoints " &
                              "FROM Players " &
                              "INNER JOIN Teams ON Players.Team = Teams.Team " &
                              "WHERE Teams.League = @League AND Players.Points > @MinPoints " &
                              "GROUP BY Players.Player, Players.Points, Teams.Stadium " &
         "ORDER BY AVG(Players.Points * 1.0 / Players.Games) DESC"

        Dim command As New OleDbCommand(query, connection)
        command.Parameters.AddWithValue("@League", league)
        command.Parameters.AddWithValue("@MinPoints", minPoints)
        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            gDGVForty12.DataSource = dataTable
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub



    Private Sub glbSARUs13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles glbSARUs13.SelectedIndexChanged
        If glbSARUs13.SelectedIndex <> -1 Then
            Dim selectedTeamName As String = glbSARUs13.SelectedItem.ToString()
            DisplayPlayersAboveTeamAverage(selectedTeamName)
        End If
    End Sub

    Private Sub DisplayPlayersAboveTeamAverage(teamName As String)
        ' First, get the team's points average
        Dim teamAverageQuery As String = "SELECT AVG(Points * 1.0 / Games) AS TeamAverage FROM Players WHERE Team = @TeamName"
        Dim teamAverageCommand As New OleDbCommand(teamAverageQuery, connection)
        teamAverageCommand.Parameters.AddWithValue("@TeamName", teamName)

        Dim teamAverage As Double = 0.0

        Try
            Dim result As Object = teamAverageCommand.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                teamAverage = CDbl(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

        ' Then, get players whose points average is greater than the team's points average
        Dim query As String = "SELECT Player, AVG(Points * 1.0 / Games) AS PlayerAverage " &
                              "FROM Players WHERE Team = @TeamName " &
                              "GROUP BY Player " &
                              "HAVING AVG(Points * 1.0 / Games) > @TeamAverage " &
                              "ORDER BY AVG(Points * 1.0 / Games) DESC"

        Dim command As New OleDbCommand(query, connection)
        command.Parameters.AddWithValue("@TeamName", teamName)
        command.Parameters.AddWithValue("@TeamAverage", teamAverage)

        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            glbPoints13.Items.Clear()

            For Each row As DataRow In dataTable.Rows
                Dim playerName As String = row("Player").ToString()
                glbPoints13.Items.Add(playerName)
            Next
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    'Tab 15
    Private Sub gCBtnDisplay15_Click(sender As Object, e As EventArgs) Handles gCBtnDisplay15.Click
        Dim highestAverage As Double = GetHighestCurrieCupAverage()
        gtxtHighAvg15.Text = highestAverage.ToString("0.00")
    End Sub

    Private Function GetHighestCurrieCupAverage() As Double
        Dim query As String = "SELECT MAX(AveragePoints) AS HighestAverage " &
                          "FROM (SELECT AVG(Points * 1.0 / Games) AS AveragePoints " &
                          "      FROM Players " &
                          "      WHERE Team IN (SELECT Team FROM Teams WHERE League = 'Currie Cup') " &
                          "      GROUP BY Player) AS Subquery"

        Dim command As New OleDbCommand(query, connection)

        Try
            Dim result As Object = command.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CDbl(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

        Return 0.0
    End Function

    Private Sub Guna2PictureBox1_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox1.Click
        Application.Exit()
    End Sub


    Private Sub Guna2PictureBox2_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox2.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub gDGV2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles gDGV2.CellContentClick

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub gtxtTotalPts8_TextChanged(sender As Object, e As EventArgs) Handles gtxtTotalPts8.TextChanged

    End Sub
End Class
