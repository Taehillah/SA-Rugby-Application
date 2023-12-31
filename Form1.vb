﻿Imports System.Data.OleDb

Public Class Form1
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\imafo\Documents\School\ICT3611\Assign6_Database\Rugby.accdb;Persist Security Info=False"
    Dim connection As New OleDbConnection(connectionString)
    'Forms load supports the subs to load data when the form opens
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
        gcbLeague14.Items.AddRange(New String() {"Currie Cup", "SA Rugby"})
        LoadAveragesAbove3()
    End Sub


    'LoadTeams()
    'The code below serves to load Teams data from the db to the lb 
    Private Sub LoadTeams()
        Dim query As String = "SELECT Team FROM Teams"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        'lb cleared on load
        lbTeams5.Items.Clear()
        glbSARUs13.Items.Clear() 'just added

        While reader.Read()
            Dim teamName As String = reader("Team").ToString()
            lbTeams5.Items.Add(teamName)
            glbSARUs13.Items.Add(teamName) 'just added
        End While
        reader.Close()

    End Sub





    'Tab 1
    Private Sub TabOne()
        'Tab Number One Queries
        Dim query As String = "Select * FROM Teams"
        Dim command As New OleDbCommand(query, connection)
        Dim reader As OleDbDataReader = command.ExecuteReader()

        glbHighAvg15.Items.Clear()

        While reader.Read()
            Dim teamName As String = reader("Team").ToString()
            glbHighAvg15.Items.Add(teamName)
        End While

        reader.Close()
    End Sub


    'Tab 2
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


    'Tab 3
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

    'Tab 4 
    'The tab is loaded when the form loads
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


    'Tab 5
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles glbHighAvg15.SelectedIndexChanged
        If glbHighAvg15.SelectedIndex <> -1 Then
            Dim selectedTeamName As String = glbHighAvg15.SelectedItem.ToString()
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
    Private Sub lbTeams5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbTeams5.SelectedIndexChanged
        If lbTeams5.SelectedIndex <> -1 Then
            Dim selectedTeamName As String = lbTeams5.SelectedItem.ToString()
            DisplayPlayers(selectedTeamName)
        End If
    End Sub


    'Tab6
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
        'The code below fills the datagridview with data
        Try
            adapter.Fill(dataTable)
            gDGVTeam6.DataSource = dataTable
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub


    'Tab 7 Code
    'The code below calculates the average points in Currie Cup using an event handler
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


    'Tab 8
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
        'The below code returns 0 if theres errors
        Return 0
    End Function


    'Tab9
    Private Sub gCBtnPlayers9_Click(sender As Object, e As EventArgs) Handles gCBtnPlayers9.Click
        Dim playerCount As Integer = CountCurrieCupPlayers()
        gtxtPlayers9.Text = playerCount.ToString()
    End Sub
    Private Function CountCurrieCupPlayers() As Integer
        Dim query As String = "SELECT COUNT(*) AS PlayerCount " &
                              "FROM Players P " &
                              "WHERE EXISTS (SELECT 1 FROM Teams T WHERE T.Team = P.Team AND T.League = 'Currie Cup')"

        Dim command As New OleDbCommand(query, connection)
        'I validate the database 
        Try
            Dim result As Object = command.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CInt(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
        'The below code returns 0 if theres errors
        Return 0
    End Function


    'Tab10
    'The entire code is enclosed within the btn event handler
    Private Sub gCBtnBloem10_Click(sender As Object, e As EventArgs) Handles gCBtnBloem10.Click
        Dim playerCount As Integer = CountBloemPlayers()
        gtxtBloem10.Text = playerCount.ToString()
    End Sub
    'The sql query within a count bloem players
    Private Function CountBloemPlayers() As Integer
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
        ' The code below returns 0 if there's an error or no result
        Return 0
    End Function


    'Tab11
    Private Sub RadioButton_CheckedChanged11(sender As Object, e As EventArgs) Handles gRBtnCurrieCup11.CheckedChanged, gRBtnSARugby11.CheckedChanged
        If gRBtnCurrieCup11.Checked Then
            DisplayTeamsByLeague("Currie Cup")
        ElseIf gRBtnSARugby11.Checked Then
            DisplayTeamsByLeague("SA Rugby")
        End If
    End Sub
    'The code below displays according to the sql query
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
        'This below code populates the datagridview
        Try
            adapter.Fill(dataTable)
            gDGVTeams11.DataSource = dataTable
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub


    'Tab12
    'The code below checks which radio button is checked
    Private Sub RadioButton_CheckedChanged12(sender As Object, e As EventArgs) Handles gRBtnCurrieCup12.CheckedChanged, gRBtnSARugby12.CheckedChanged
        If gRBtnCurrieCup12.Checked Then
            DisplayPlayers("Currie Cup", 40)
        ElseIf gRBtnSARugby12.Checked Then
            DisplayPlayers("SA Rugby", 40)
        End If
    End Sub
    'The code below displays according to the sql query
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


    'Tab13
    'This code selects the player with the team avg
    Private Sub glbSARUs13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles glbSARUs13.SelectedIndexChanged
        If glbSARUs13.SelectedIndex <> -1 Then
            Dim selectedTeamName As String = glbSARUs13.SelectedItem.ToString()
            DisplayPlayersAboveTeamAverage(selectedTeamName)
        End If
    End Sub
    'Function with a Query to collect points
    Private Sub DisplayPlayersAboveTeamAverage(teamName As String)
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

        'I get the players whose points average is greater than the team's points average
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


    'Tab14
    Private Sub gcbLeague14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gcbLeague14.SelectedIndexChanged
        'Confirms the selected league as a string
        Dim selectedLeague As String = gcbLeague14.SelectedItem.ToString()

        ' Here I disabled the text box if the user doesnt select the league
        If selectedLeague = "Currie Cup" OrElse selectedLeague = "SA Rugby" Then
            gtxtAvgPts14.Enabled = True
        Else
            gtxtAvgPts14.Enabled = False
        End If
    End Sub
    Private Sub gtxtAvgPts14_TextChanged(sender As Object, e As EventArgs) Handles gtxtAvgPts14.TextChanged
        ' Validates the input for average points
        Dim inputText As String = gtxtAvgPts14.Text
        'The code that Converts the text into a double
        If Double.TryParse(inputText, Nothing) Then
            Dim averagePoints As Double = Double.Parse(inputText)

            'This code validates the min and max
            If averagePoints < 0 Then
                MessageBox.Show("Average points cannot be less than 0.")
                gtxtAvgPts14.Text = "0"
            ElseIf averagePoints > 10 Then
                MessageBox.Show("Average points cannot be greater than 10.")
                gtxtAvgPts14.Text = "10"
            End If
        Else
            'I removed the error or validation msg before the user enters input
            gtxtAvgPts14.Text = ""
        End If
    End Sub
    'I enclosed the code for tab 14 inside the btn as an event handler
    Private Sub gbtnCalculate14_Click(sender As Object, e As EventArgs) Handles gbtnCalculate14.Click
        'The code below collects the League's name from the selection
        Dim selectedLeague As String = gcbLeague14.SelectedItem.ToString()

        'The code below collects the average points from the user
        Dim averagePoints As Double
        If Double.TryParse(gtxtAvgPts14.Text, averagePoints) Then

            Dim query As String = "SELECT Player FROM Players " &
                                  "WHERE Team IN (SELECT Team FROM Teams WHERE League = @League) " &
                                  "AND (Points * 1.0 / Games) > @AveragePoints " &
                                  "ORDER BY Player"

            Dim command As New OleDbCommand(query, connection)
            command.Parameters.AddWithValue("@League", selectedLeague)
            command.Parameters.AddWithValue("@AveragePoints", averagePoints)

            Dim adapter As New OleDbDataAdapter(command)
            Dim dataTable As New DataTable()

            Try
                adapter.Fill(dataTable)
                glbPlayers14.Items.Clear()

                For Each row As DataRow In dataTable.Rows
                    Dim playerName As String = row("Player").ToString()
                    glbPlayers14.Items.Add(playerName)
                Next
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        Else
            MessageBox.Show("Please enter a valid numeric value for average points.")
        End If
    End Sub


    'Tab15
    'The code below takes the player(s) with the highest average points
    'I enclosed the code inside the tab 15 button as an event handler
    Private Sub gCBtnDisplay15_Click(sender As Object, e As EventArgs) Handles gCBtnDisplay15.Click

        Dim query As String = "SELECT Player, MAX(Points) AS HighestPoints " &
                              "FROM Players " &
                              "WHERE Team IN (SELECT Team FROM Teams WHERE League = 'Currie Cup') " &
                              "GROUP BY Player " &
                              "HAVING MAX(Points) = (SELECT MAX(Points) FROM Players WHERE Team IN (SELECT Team FROM Teams WHERE League = 'Currie Cup'))"

        Dim command As New OleDbCommand(query, connection)
        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            glbPtsAvg15.Items.Clear()

            For Each row As DataRow In dataTable.Rows
                Dim playerName As String = row("Player").ToString()
                glbPtsAvg15.Items.Add(playerName)
            Next
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub


    'Tab16
    'The code below takes the player with the highest points in the SA Rugby league
    'I enclose the code inside the tab 16 button and an event handler
    Private Sub gCBtnHighPts16_Click(sender As Object, e As EventArgs) Handles gCBtnHighPts16.Click

        Dim query As String = "SELECT Player, MAX(Points) AS HighestPoints " &
                              "FROM Players " &
                              "WHERE Team IN (SELECT Team FROM Teams WHERE League = 'SA Rugby') " &
                              "GROUP BY Player " &
                              "HAVING MAX(Points) = (SELECT MAX(Points) FROM Players WHERE Team IN (SELECT Team FROM Teams WHERE League = 'SA Rugby'))"

        Dim command As New OleDbCommand(query, connection)
        Dim adapter As New OleDbDataAdapter(command)
        Dim dataTable As New DataTable()

        Try
            adapter.Fill(dataTable)
            glbPlayerHG16.Items.Clear()

            For Each row As DataRow In dataTable.Rows
                Dim playerName As String = row("Player").ToString()
                glbPlayerHG16.Items.Add(playerName)
            Next
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub


    'Exit Btn
    'This code is for the custom exit button I made
    'This is because I used the ellipse tool to remove the title bar
    Private Sub Guna2PictureBox1_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox1.Click
        Application.Exit()
    End Sub


    'Minimize Btn
    'This code is for minimizing the window from the custom btn I made
    Private Sub Guna2PictureBox2_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox2.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

End Class
