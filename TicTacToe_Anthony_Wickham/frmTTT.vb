' Program Name: Tic Tac Ultimate
' Programmer:   Anthony Wickham
' Date:         3/27/16
' Purpose:      User plays the classic game of Tic Tac Toe against the CPU.

Option Strict On

Public Class frmTTT
    ' Declare Class-wide variables

    ' Identifies if the CPU turn is complete
    Dim turncomplete As Boolean
    Dim gameover As String
    Dim turnnum As Integer
    Dim allspacesgone As Boolean


    ' Identifies difficulty
    Dim cpudifficulty As String
    Dim nomercy As Boolean

    ' Identifies the players preferred mark
    Dim PlayerMark As String
    Dim cpuMark As String

    ' Keeps track of wins
    Dim xwins As Integer
    Dim owins As Integer
    Dim cwins As Integer

    ' Identifies CPU color
    Dim cpucolor As Color
    Dim playercolor As Color
    Dim btnbackcolor As Color
    Dim btncolor As Color

    ' Identifies if buttons have been selected or not
    Dim btn1selected As Boolean
    Dim btn2selected As Boolean
    Dim btn3selected As Boolean
    Dim btn4selected As Boolean
    Dim btn5selected As Boolean
    Dim btn6selected As Boolean
    Dim btn7selected As Boolean
    Dim btn8selected As Boolean
    Dim btn9selected As Boolean






    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSel1.Click

        If btn1selected = False Then
            turncomplete = False
            btn1selected = True
            turnnum = turnnum + 1
            btnSel1.BackColor = btncolor
            btnSel1.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if a corner is selected
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If
    End Sub



    Private Sub frmTTT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PlayerMark = "X"
        cpuMark = "O"
        turnnum = 0
        nomercy = False
        cpudifficulty = "aggressive"
        cpucolor = Color.Red
        btncolor = Color.Cyan
        playercolor = Color.Black
        btnbackcolor = Color.DarkTurquoise
        btn1selected = False
        btn2selected = False
        btn3selected = False
        btn4selected = False
        btn5selected = False
        btn6selected = False
        btn7selected = False
        btn8selected = False
        btn9selected = False
        lblDifficulty.Text = "CPU Difficulty: Aggressive"

        ' Set default theme to pizza
        Me.BackgroundImage = picPizza.Image
        Me.BackgroundImageLayout = ImageLayout.Stretch
        cpucolor = Color.Black
        playercolor = Color.White
        btnbackcolor = Color.DarkRed
        btncolor = Color.Firebrick
        lblDifficulty.BackColor = Color.Firebrick
        lblDifficulty.ForeColor = Color.White
        lblWin.ForeColor = Color.White
        lblWin.BackColor = Color.Firebrick
        Label1.ForeColor = Color.White
        Label1.BackColor = Color.Firebrick
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Firebrick
        Label3.ForeColor = Color.White
        Label3.BackColor = Color.Firebrick
        btnSel1.BackColor = Color.Firebrick
        btnSel2.BackColor = Color.Firebrick
        btnSel3.BackColor = Color.Firebrick
        btnSel4.BackColor = Color.Firebrick
        btnSel5.BackColor = Color.Firebrick
        btnSel6.BackColor = Color.Firebrick
        btnSel7.BackColor = Color.Firebrick
        btnSel8.BackColor = Color.Firebrick
        btnSel9.BackColor = Color.Firebrick
        btnSel1.ForeColor = Color.White
        btnSel2.ForeColor = Color.White
        btnSel3.ForeColor = Color.White
        btnSel4.ForeColor = Color.White
        btnSel5.ForeColor = Color.White
        btnSel6.ForeColor = Color.White
        btnSel7.ForeColor = Color.White
        btnSel8.ForeColor = Color.White
        btnSel9.ForeColor = Color.White
        btnNew.BackColor = Color.Firebrick
        btnNew.ForeColor = Color.White
        btnClear.BackColor = Color.Firebrick
        btnClear.ForeColor = Color.White
        btnExit.BackColor = Color.Firebrick
        btnExit.ForeColor = Color.White
        Me.BackColor = Color.White
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub btnSel1_MouseHover(sender As Object, e As EventArgs) Handles btnSel1.MouseHover
        If btn1selected = False Then
            btnSel1.Text = PlayerMark
            btnSel1.BackColor = btnbackcolor
        End If
    End Sub

    Private Sub btnSel1_MouseLeave(sender As Object, e As EventArgs) Handles btnSel1.MouseLeave
        If btn1selected = False Then
            btnSel1.Text = ""
            btnSel1.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel2_MouseHover(sender As Object, e As EventArgs) Handles btnSel2.MouseHover
        If btn2selected = False Then
            btnSel2.Text = PlayerMark
            btnSel2.BackColor = btnbackcolor
        End If

    End Sub

    Private Sub btnSel3_MouseHover(sender As Object, e As EventArgs) Handles btnSel3.MouseHover
        If btn3selected = False Then
            btnSel3.Text = PlayerMark
            btnSel3.BackColor = btnbackcolor
        End If

    End Sub

    Private Sub btnSel4_MouseHover(sender As Object, e As EventArgs) Handles btnSel4.MouseHover
        If btn4selected = False Then
            btnSel4.Text = PlayerMark
            btnSel4.BackColor = btnbackcolor
        End If

    End Sub

    Private Sub btnSel5_MouseHover(sender As Object, e As EventArgs) Handles btnSel5.MouseHover
        If btn5selected = False Then
            btnSel5.Text = PlayerMark
            btnSel5.BackColor = btnbackcolor
        End If
    End Sub

    Private Sub btnSel6_MouseHover(sender As Object, e As EventArgs) Handles btnSel6.MouseHover
        If btn6selected = False Then
            btnSel6.Text = PlayerMark
            btnSel6.BackColor = btnbackcolor
        End If
    End Sub

    Private Sub btnSel7_MouseHover(sender As Object, e As EventArgs) Handles btnSel7.MouseHover
        If btn7selected = False Then
            btnSel7.Text = PlayerMark
            btnSel7.BackColor = btnbackcolor
        End If
    End Sub

    Private Sub btnSel8_MouseHover(sender As Object, e As EventArgs) Handles btnSel8.MouseHover
        If btn8selected = False Then
            btnSel8.Text = PlayerMark
            btnSel8.BackColor = btnbackcolor
        End If
    End Sub

    Private Sub btnSel9_MouseHover(sender As Object, e As EventArgs) Handles btnSel9.MouseHover
        If btn9selected = False Then
            btnSel9.Text = PlayerMark
            btnSel9.BackColor = btnbackcolor
        End If
    End Sub

    Private Sub btnSel2_MouseLeave(sender As Object, e As EventArgs) Handles btnSel2.MouseLeave
        If btn2selected = False Then
            btnSel2.Text = ""
            btnSel2.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel3_MouseLeave(sender As Object, e As EventArgs) Handles btnSel3.MouseLeave
        If btn3selected = False Then
            btnSel3.Text = ""
            btnSel3.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel4_MouseLeave(sender As Object, e As EventArgs) Handles btnSel4.MouseLeave
        If btn4selected = False Then
            btnSel4.Text = ""
            btnSel4.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel5_MouseLeave(sender As Object, e As EventArgs) Handles btnSel5.MouseLeave
        If btn5selected = False Then
            btnSel5.Text = ""
            btnSel5.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel6_MouseLeave(sender As Object, e As EventArgs) Handles btnSel6.MouseLeave
        If btn6selected = False Then
            btnSel6.Text = ""
            btnSel6.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel7_MouseLeave(sender As Object, e As EventArgs) Handles btnSel7.MouseLeave
        If btn7selected = False Then
            btnSel7.Text = ""
            btnSel7.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel8_MouseLeave(sender As Object, e As EventArgs) Handles btnSel8.MouseLeave
        If btn8selected = False Then
            btnSel8.Text = ""
            btnSel8.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel9_MouseLeave(sender As Object, e As EventArgs) Handles btnSel9.MouseLeave
        If btn9selected = False Then
            btnSel9.Text = ""
            btnSel9.BackColor = btncolor
        End If
    End Sub

    Private Sub btnSel3_Click(sender As Object, e As EventArgs) Handles btnSel3.Click
        If btn3selected = False Then
            turncomplete = False
            btn3selected = True
            turnnum = turnnum + 1
            btnSel3.BackColor = btncolor
            btnSel3.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if a corner is selected
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnSel7_Click(sender As Object, e As EventArgs) Handles btnSel7.Click
        If btn7selected = False Then
            turncomplete = False
            btn7selected = True
            turnnum = turnnum + 1
            btnSel7.BackColor = btncolor
            btnSel7.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if a corner is selected
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnSel9_Click(sender As Object, e As EventArgs) Handles btnSel9.Click
        If btn9selected = False Then
            turncomplete = False
            btn9selected = True
            turnnum = turnnum + 1
            btnSel9.BackColor = btncolor
            btnSel9.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if available on cpu's first turn
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnSel5_Click(sender As Object, e As EventArgs) Handles btnSel5.Click


        If btn5selected = False Then
            btn5selected = True
            turncomplete = False

            turnnum = turnnum + 1
            btnSel5.BackColor = btncolor
            btnSel5.Text = PlayerMark


            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If


            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if a corner is selected
                If ((btn1selected = False) Or (btn3selected = False) Or (btn7selected = False) Or (btn9selected = False)) And (nomercy = True) Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 4)
                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            End If
                        ElseIf rand = 2 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            End If
                        ElseIf rand = 3 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            End If
                        ElseIf rand = 4 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            End If
                        End If
                    End While

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnSel2_Click(sender As Object, e As EventArgs) Handles btnSel2.Click
        If btn2selected = False Then
            turncomplete = False
            btn2selected = True
            turnnum = turnnum + 1
            btnSel2.BackColor = btncolor
            btnSel2.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if available on cpu's first turn
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnSel4_Click(sender As Object, e As EventArgs) Handles btnSel4.Click
        If btn4selected = False Then
            turncomplete = False
            btn4selected = True
            turnnum = turnnum + 1
            btnSel4.BackColor = btncolor
            btnSel4.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if available on cpu's first turn
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnSel6_Click(sender As Object, e As EventArgs) Handles btnSel6.Click
        If btn6selected = False Then
            turncomplete = False
            btn6selected = True
            turnnum = turnnum + 1
            btnSel6.BackColor = btncolor
            btnSel6.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if available on cpu's first turn
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnSel8_Click(sender As Object, e As EventArgs) Handles btnSel8.Click
        If btn8selected = False Then
            turncomplete = False
            btn8selected = True
            turnnum = turnnum + 1
            btnSel8.BackColor = btncolor
            btnSel8.Text = PlayerMark
            playerwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
            If (btn1selected = True) And (btn2selected = True) And (btn3selected = True) And (btn4selected = True) And (btn5selected = True) And (btn6selected = True) And (btn7selected = True) And (btn8selected = True) And (btn9selected = True) Then
                allspacesgone = True
            End If
            If (cpudifficulty = "aggressive") And (allspacesgone <> True) Then

                ' Prioritizes a center selection for CPU if available on cpu's first turn
                If (btn5selected = False) And (turnnum = 1) And (nomercy = True) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True

                    'Prioritizes a win for CPU if the player leaves it open
                ElseIf (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                    ' Allows for a defensive selection if player is about to win
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True

                ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn4selected = False) Then
                    btnSel4.ForeColor = cpucolor
                    btnSel4.Text = cpuMark
                    btn4selected = True
                    turncomplete = True
                ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn8selected = False) Then
                    btnSel8.ForeColor = cpucolor
                    btnSel8.Text = cpuMark
                    btn8selected = True
                    turncomplete = True
                ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btn2selected = False) Then
                    btnSel2.ForeColor = cpucolor
                    btnSel2.Text = cpuMark
                    btn2selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn6selected = False) Then
                    btnSel6.ForeColor = cpucolor
                    btnSel6.Text = cpuMark
                    btn6selected = True
                    turncomplete = True
                ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn9selected = False) Then
                    btnSel9.ForeColor = cpucolor
                    btnSel9.Text = cpuMark
                    btn9selected = True
                    turncomplete = True
                ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btn1selected = False) Then
                    btnSel1.ForeColor = cpucolor
                    btnSel1.Text = cpuMark
                    btn1selected = True
                    turncomplete = True

                ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btn7selected = False) Then
                    btnSel7.ForeColor = cpucolor
                    btnSel7.Text = cpuMark
                    btn7selected = True
                    turncomplete = True
                ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn5selected = False) Then
                    btnSel5.ForeColor = cpucolor
                    btnSel5.Text = cpuMark
                    btn5selected = True
                    turncomplete = True
                ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btn3selected = False) Then
                    btnSel3.ForeColor = cpucolor
                    btnSel3.Text = cpuMark
                    btn3selected = True
                    turncomplete = True

                Else
                    If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                        While turncomplete = False
                            Dim rn As New Random
                            Dim rand As Integer
                            rand = rn.Next(1, 9)





                            If rand = 1 Then
                                If btn1selected = False Then
                                    btnSel1.ForeColor = cpucolor
                                    btnSel1.Text = cpuMark
                                    btn1selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 2 Then
                                If btn2selected = False Then
                                    btnSel2.ForeColor = cpucolor
                                    btnSel2.Text = cpuMark
                                    btn2selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 3 Then
                                If btn3selected = False Then
                                    btnSel3.ForeColor = cpucolor
                                    btnSel3.Text = cpuMark
                                    btn3selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 4 Then
                                If btn4selected = False Then
                                    btnSel4.ForeColor = cpucolor
                                    btnSel4.Text = cpuMark
                                    btn4selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 5 Then
                                If btn5selected = False Then
                                    btnSel5.ForeColor = cpucolor
                                    btnSel5.Text = cpuMark
                                    btn5selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 6 Then
                                If btn6selected = False Then
                                    btnSel6.ForeColor = cpucolor
                                    btnSel6.Text = cpuMark
                                    btn6selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 7 Then
                                If btn7selected = False Then
                                    btnSel7.ForeColor = cpucolor
                                    btnSel7.Text = cpuMark
                                    btn7selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 8 Then
                                If btn8selected = False Then
                                    btnSel8.ForeColor = cpucolor
                                    btnSel8.Text = cpuMark
                                    btn8selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If
                            If rand = 9 Then
                                If btn9selected = False Then
                                    btnSel9.ForeColor = cpucolor
                                    btnSel9.Text = cpuMark
                                    btn9selected = True
                                    turncomplete = True
                                Else
                                    turncomplete = False
                                End If
                            End If



                        End While
                    End If
                End If
            ElseIf (cpudifficulty = "random") And (allspacesgone <> True) Then
                If (btnSel1.Text <> "") Or (btnSel2.Text <> "") Or (btnSel3.Text <> "") Or (btnSel4.Text <> "") Or (btnSel5.Text <> "") Or (btnSel6.Text <> "") Or (btnSel7.Text <> "") Or (btnSel8.Text <> "") Or (btnSel9.Text <> "") Then
                    While turncomplete = False
                        Dim rn As New Random
                        Dim rand As Integer
                        rand = rn.Next(1, 9)





                        If rand = 1 Then
                            If btn1selected = False Then
                                btnSel1.ForeColor = cpucolor
                                btnSel1.Text = cpuMark
                                btn1selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 2 Then
                            If btn2selected = False Then
                                btnSel2.ForeColor = cpucolor
                                btnSel2.Text = cpuMark
                                btn2selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 3 Then
                            If btn3selected = False Then
                                btnSel3.ForeColor = cpucolor
                                btnSel3.Text = cpuMark
                                btn3selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 4 Then
                            If btn4selected = False Then
                                btnSel4.ForeColor = cpucolor
                                btnSel4.Text = cpuMark
                                btn4selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 5 Then
                            If btn5selected = False Then
                                btnSel5.ForeColor = cpucolor
                                btnSel5.Text = cpuMark
                                btn5selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 6 Then
                            If btn6selected = False Then
                                btnSel6.ForeColor = cpucolor
                                btnSel6.Text = cpuMark
                                btn6selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 7 Then
                            If btn7selected = False Then
                                btnSel7.ForeColor = cpucolor
                                btnSel7.Text = cpuMark
                                btn7selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 8 Then
                            If btn8selected = False Then
                                btnSel8.ForeColor = cpucolor
                                btnSel8.Text = cpuMark
                                btn8selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If
                        If rand = 9 Then
                            If btn9selected = False Then
                                btnSel9.ForeColor = cpucolor
                                btnSel9.Text = cpuMark
                                btn9selected = True
                                turncomplete = True
                            Else
                                turncomplete = False
                            End If
                        End If



                    End While
                End If
            End If
            cpuwinchk()
            tiechk()
            If gameover <> "" Then
                gameoverchk()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        lblWin.Visible = False
        btnSel1.Enabled = True
        btnSel2.Enabled = True
        btnSel3.Enabled = True
        btnSel4.Enabled = True
        btnSel5.Enabled = True
        btnSel6.Enabled = True
        btnSel7.Enabled = True
        btnSel8.Enabled = True
        btnSel9.Enabled = True
        btn1selected = False
        btn2selected = False
        btn3selected = False
        btn4selected = False
        btn5selected = False
        btn6selected = False
        btn7selected = False
        btn8selected = False
        btn9selected = False
        turnnum = 0
        btnSel1.Text = ""
        btnSel2.Text = ""
        btnSel3.Text = ""
        btnSel4.Text = ""
        btnSel5.Text = ""
        btnSel6.Text = ""
        btnSel7.Text = ""
        btnSel8.Text = ""
        btnSel9.Text = ""
        gameover = ""
        btnSel1.ForeColor = playercolor
        btnSel2.ForeColor = playercolor
        btnSel3.ForeColor = playercolor
        btnSel4.ForeColor = playercolor
        btnSel5.ForeColor = playercolor
        btnSel6.ForeColor = playercolor
        btnSel7.ForeColor = playercolor
        btnSel8.ForeColor = playercolor
        btnSel9.ForeColor = playercolor
        allspacesgone = False
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        xwins = 0
        owins = 0
        cwins = 0
        Label1.Text = "X Wins: "
        Label2.Text = "O Wins: "
        Label3.Text = "Cat Wins: "

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub NewGameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewGameToolStripMenuItem.Click
        lblWin.Visible = False
        btnSel1.Enabled = True
        btnSel2.Enabled = True
        btnSel3.Enabled = True
        btnSel4.Enabled = True
        btnSel5.Enabled = True
        btnSel6.Enabled = True
        btnSel7.Enabled = True
        btnSel8.Enabled = True
        btnSel9.Enabled = True
        btn1selected = False
        btn2selected = False
        btn3selected = False
        btn4selected = False
        btn5selected = False
        btn6selected = False
        btn7selected = False
        btn8selected = False
        btn9selected = False
        turnnum = 0
        btnSel1.Text = ""
        btnSel2.Text = ""
        btnSel3.Text = ""
        btnSel4.Text = ""
        btnSel5.Text = ""
        btnSel6.Text = ""
        btnSel7.Text = ""
        btnSel8.Text = ""
        btnSel9.Text = ""
        gameover = ""
        btnSel1.ForeColor = playercolor
        btnSel2.ForeColor = playercolor
        btnSel3.ForeColor = playercolor
        btnSel4.ForeColor = playercolor
        btnSel5.ForeColor = playercolor
        btnSel6.ForeColor = playercolor
        btnSel7.ForeColor = playercolor
        btnSel8.ForeColor = playercolor
        btnSel9.ForeColor = playercolor
        allspacesgone = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        xwins = 0
        owins = 0
        cwins = 0
        Label1.Text = "X Wins: "
        Label2.Text = "O Wins: "
        Label3.Text = "Cat Wins: "
    End Sub

    Private Sub RandomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RandomToolStripMenuItem.Click
        If MessageBox.Show(" Changing the difficulty will reset your score and current game progress.", "",
         MessageBoxButtons.OKCancel, MessageBoxIcon.Information) _
         = DialogResult.OK Then
            cpudifficulty = "random"
            lblWin.Visible = False
            btnSel1.Enabled = True
            btnSel2.Enabled = True
            btnSel3.Enabled = True
            btnSel4.Enabled = True
            btnSel5.Enabled = True
            btnSel6.Enabled = True
            btnSel7.Enabled = True
            btnSel8.Enabled = True
            btnSel9.Enabled = True
            btn1selected = False
            btn2selected = False
            btn3selected = False
            btn4selected = False
            btn5selected = False
            btn6selected = False
            btn7selected = False
            btn8selected = False
            btn9selected = False
            turnnum = 0
            btnSel1.Text = ""
            btnSel2.Text = ""
            btnSel3.Text = ""
            btnSel4.Text = ""
            btnSel5.Text = ""
            btnSel6.Text = ""
            btnSel7.Text = ""
            btnSel8.Text = ""
            btnSel9.Text = ""
            gameover = ""
            btnSel1.ForeColor = playercolor
            btnSel2.ForeColor = playercolor
            btnSel3.ForeColor = playercolor
            btnSel4.ForeColor = playercolor
            btnSel5.ForeColor = playercolor
            btnSel6.ForeColor = playercolor
            btnSel7.ForeColor = playercolor
            btnSel8.ForeColor = playercolor
            btnSel9.ForeColor = playercolor
            allspacesgone = False
            xwins = 0
            owins = 0
            cwins = 0
            Label1.Text = "X Wins: "
            Label2.Text = "O Wins: "
            Label3.Text = "Cat Wins: "
            lblDifficulty.Text = "CPU Difficulty: Easy"
        End If


    End Sub

    Private Sub AggressiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AggressiveToolStripMenuItem.Click
        If MessageBox.Show(" Changing the difficulty will reset your score and current game progress.", "",
         MessageBoxButtons.OKCancel, MessageBoxIcon.Information) _
         = DialogResult.OK Then
            cpudifficulty = "aggressive"
            lblWin.Visible = False
            btnSel1.Enabled = True
            btnSel2.Enabled = True
            btnSel3.Enabled = True
            btnSel4.Enabled = True
            btnSel5.Enabled = True
            btnSel6.Enabled = True
            btnSel7.Enabled = True
            btnSel8.Enabled = True
            btnSel9.Enabled = True
            btn1selected = False
            btn2selected = False
            btn3selected = False
            btn4selected = False
            btn5selected = False
            btn6selected = False
            btn7selected = False
            btn8selected = False
            btn9selected = False
            turnnum = 0
            btnSel1.Text = ""
            btnSel2.Text = ""
            btnSel3.Text = ""
            btnSel4.Text = ""
            btnSel5.Text = ""
            btnSel6.Text = ""
            btnSel7.Text = ""
            btnSel8.Text = ""
            btnSel9.Text = ""
            gameover = ""
            btnSel1.ForeColor = playercolor
            btnSel2.ForeColor = playercolor
            btnSel3.ForeColor = playercolor
            btnSel4.ForeColor = playercolor
            btnSel5.ForeColor = playercolor
            btnSel6.ForeColor = playercolor
            btnSel7.ForeColor = playercolor
            btnSel8.ForeColor = playercolor
            btnSel9.ForeColor = playercolor
            allspacesgone = False
            xwins = 0
            owins = 0
            cwins = 0
            Label1.Text = "X Wins: "
            Label2.Text = "O Wins: "
            Label3.Text = "Cat Wins: "
            lblDifficulty.Text = "CPU Difficulty: Aggressive"
        End If
    End Sub

    Private Sub NoviceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NoviceToolStripMenuItem.Click
        If MessageBox.Show(" Changing the difficulty will reset your score and current game progress.", "",
         MessageBoxButtons.OKCancel, MessageBoxIcon.Information) _
         = DialogResult.OK Then
            cpudifficulty = "aggressive"
            nomercy = True
            lblWin.Visible = False
            btnSel1.Enabled = True
            btnSel2.Enabled = True
            btnSel3.Enabled = True
            btnSel4.Enabled = True
            btnSel5.Enabled = True
            btnSel6.Enabled = True
            btnSel7.Enabled = True
            btnSel8.Enabled = True
            btnSel9.Enabled = True
            btn1selected = False
            btn2selected = False
            btn3selected = False
            btn4selected = False
            btn5selected = False
            btn6selected = False
            btn7selected = False
            btn8selected = False
            btn9selected = False
            turnnum = 0
            btnSel1.Text = ""
            btnSel2.Text = ""
            btnSel3.Text = ""
            btnSel4.Text = ""
            btnSel5.Text = ""
            btnSel6.Text = ""
            btnSel7.Text = ""
            btnSel8.Text = ""
            btnSel9.Text = ""
            gameover = ""
            btnSel1.ForeColor = playercolor
            btnSel2.ForeColor = playercolor
            btnSel3.ForeColor = playercolor
            btnSel4.ForeColor = playercolor
            btnSel5.ForeColor = playercolor
            btnSel6.ForeColor = playercolor
            btnSel7.ForeColor = playercolor
            btnSel8.ForeColor = playercolor
            btnSel9.ForeColor = playercolor
            allspacesgone = False
            xwins = 0
            owins = 0
            cwins = 0
            Label1.Text = "X Wins: "
            Label2.Text = "O Wins: "
            Label3.Text = "Cat Wins: "
            lblDifficulty.Text = "CPU Difficulty: No Mercy"
        End If
    End Sub

    Private Sub XToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles XToolStripMenuItem.Click
        If MessageBox.Show("   Changing your mark will reset your score and current game progress.", "",
         MessageBoxButtons.OKCancel, MessageBoxIcon.Information) _
         = DialogResult.OK Then
            PlayerMark = "X"
            cpuMark = "O"
            lblWin.Visible = False
            btnSel1.Enabled = True
            btnSel2.Enabled = True
            btnSel3.Enabled = True
            btnSel4.Enabled = True
            btnSel5.Enabled = True
            btnSel6.Enabled = True
            btnSel7.Enabled = True
            btnSel8.Enabled = True
            btnSel9.Enabled = True
            btn1selected = False
            btn2selected = False
            btn3selected = False
            btn4selected = False
            btn5selected = False
            btn6selected = False
            btn7selected = False
            btn8selected = False
            btn9selected = False
            turnnum = 0
            btnSel1.Text = ""
            btnSel2.Text = ""
            btnSel3.Text = ""
            btnSel4.Text = ""
            btnSel5.Text = ""
            btnSel6.Text = ""
            btnSel7.Text = ""
            btnSel8.Text = ""
            btnSel9.Text = ""
            gameover = ""
            btnSel1.ForeColor = playercolor
            btnSel2.ForeColor = playercolor
            btnSel3.ForeColor = playercolor
            btnSel4.ForeColor = playercolor
            btnSel5.ForeColor = playercolor
            btnSel6.ForeColor = playercolor
            btnSel7.ForeColor = playercolor
            btnSel8.ForeColor = playercolor
            btnSel9.ForeColor = playercolor
            allspacesgone = False
            xwins = 0
            owins = 0
            cwins = 0
            Label1.Text = "X Wins: "
            Label2.Text = "O Wins: "
            Label3.Text = "Cat Wins: "
        End If
    End Sub

    Private Sub OToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OToolStripMenuItem.Click
        If MessageBox.Show("   Changing your mark will reset your score and current game progress.", "",
         MessageBoxButtons.OKCancel, MessageBoxIcon.Information) _
         = DialogResult.OK Then
            PlayerMark = "O"
            cpuMark = "X"
            lblWin.Visible = False
            btnSel1.Enabled = True
            btnSel2.Enabled = True
            btnSel3.Enabled = True
            btnSel4.Enabled = True
            btnSel5.Enabled = True
            btnSel6.Enabled = True
            btnSel7.Enabled = True
            btnSel8.Enabled = True
            btnSel9.Enabled = True
            btn1selected = False
            btn2selected = False
            btn3selected = False
            btn4selected = False
            btn5selected = False
            btn6selected = False
            btn7selected = False
            btn8selected = False
            btn9selected = False
            turnnum = 0
            btnSel1.Text = ""
            btnSel2.Text = ""
            btnSel3.Text = ""
            btnSel4.Text = ""
            btnSel5.Text = ""
            btnSel6.Text = ""
            btnSel7.Text = ""
            btnSel8.Text = ""
            btnSel9.Text = ""
            gameover = ""
            btnSel1.ForeColor = playercolor
            btnSel2.ForeColor = playercolor
            btnSel3.ForeColor = playercolor
            btnSel4.ForeColor = playercolor
            btnSel5.ForeColor = playercolor
            btnSel6.ForeColor = playercolor
            btnSel7.ForeColor = playercolor
            btnSel8.ForeColor = playercolor
            btnSel9.ForeColor = playercolor
            allspacesgone = False
            xwins = 0
            owins = 0
            cwins = 0
            Label1.Text = "X Wins: "
            Label2.Text = "O Wins: "
            Label3.Text = "Cat Wins: "
        End If
    End Sub

    Private Sub DarkCyanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DarkCyanToolStripMenuItem.Click
        Me.BackgroundImage = Nothing
        cpucolor = Color.Red
        playercolor = Color.Black
        btnbackcolor = Color.DarkTurquoise
        btncolor = Color.Cyan
        lblDifficulty.BackColor = Color.Transparent
        lblDifficulty.ForeColor = Color.White
        lblWin.ForeColor = Color.White
        lblWin.BackColor = Color.Transparent
        Label1.ForeColor = Color.White
        Label1.BackColor = Color.Transparent
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Transparent
        Label3.ForeColor = Color.White
        Label3.BackColor = Color.Transparent
        btnSel1.BackColor = Color.Cyan
        btnSel2.BackColor = Color.Cyan
        btnSel3.BackColor = Color.Cyan
        btnSel4.BackColor = Color.Cyan
        btnSel5.BackColor = Color.Cyan
        btnSel6.BackColor = Color.Cyan
        btnSel7.BackColor = Color.Cyan
        btnSel8.BackColor = Color.Cyan
        btnSel9.BackColor = Color.Cyan
        btnNew.BackColor = Color.Black
        btnNew.ForeColor = Color.White
        btnClear.BackColor = Color.Black
        btnClear.ForeColor = Color.White
        btnExit.BackColor = Color.Black
        btnExit.ForeColor = Color.White
        Me.BackColor = Color.Black


    End Sub

    Private Sub CyanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CyanToolStripMenuItem.Click
        Me.BackgroundImage = Nothing
        cpucolor = Color.Red
        playercolor = Color.Black
        btnbackcolor = Color.DarkTurquoise
        btncolor = Color.Cyan
        lblDifficulty.BackColor = Color.Transparent
        lblDifficulty.ForeColor = Color.Black
        lblWin.ForeColor = Color.Black
        lblWin.BackColor = Color.Transparent
        Label1.ForeColor = Color.Black
        Label1.BackColor = Color.Transparent
        Label2.ForeColor = Color.Black
        Label2.BackColor = Color.Transparent
        Label3.ForeColor = Color.Black
        Label3.BackColor = Color.Transparent
        btnSel1.BackColor = Color.Cyan
        btnSel2.BackColor = Color.Cyan
        btnSel3.BackColor = Color.Cyan
        btnSel4.BackColor = Color.Cyan
        btnSel5.BackColor = Color.Cyan
        btnSel6.BackColor = Color.Cyan
        btnSel7.BackColor = Color.Cyan
        btnSel8.BackColor = Color.Cyan
        btnSel9.BackColor = Color.Cyan
        btnSel1.ForeColor = Color.Black
        btnSel2.ForeColor = Color.Black
        btnSel3.ForeColor = Color.Black
        btnSel4.ForeColor = Color.Black
        btnSel5.ForeColor = Color.Black
        btnSel6.ForeColor = Color.Black
        btnSel7.ForeColor = Color.Black
        btnSel8.ForeColor = Color.Black
        btnSel9.ForeColor = Color.Black
        btnNew.BackColor = Color.PaleTurquoise
        btnNew.ForeColor = Color.Black
        btnClear.BackColor = Color.PaleTurquoise
        btnClear.ForeColor = Color.Black
        btnExit.BackColor = Color.PaleTurquoise
        btnExit.ForeColor = Color.Black
        Me.BackColor = Color.White
    End Sub

    Private Sub SpaceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpaceToolStripMenuItem.Click
        Me.BackColor = Color.Black
        cpucolor = Color.Red
        playercolor = Color.Black
        btnbackcolor = Color.DarkViolet
        btncolor = Color.BlueViolet
        lblDifficulty.BackColor = Color.DarkViolet
        lblDifficulty.ForeColor = Color.White
        lblWin.ForeColor = Color.White
        lblWin.BackColor = Color.DarkViolet
        Label1.ForeColor = Color.White
        Label1.BackColor = Color.DarkViolet
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.DarkViolet
        Label3.ForeColor = Color.White
        Label3.BackColor = Color.DarkViolet
        btnNew.BackColor = Color.BlueViolet
        btnNew.ForeColor = Color.White
        btnClear.BackColor = Color.BlueViolet
        btnClear.ForeColor = Color.White
        btnExit.BackColor = Color.BlueViolet
        btnExit.ForeColor = Color.White
        btnSel1.BackColor = Color.BlueViolet
        btnSel2.BackColor = Color.BlueViolet
        btnSel3.BackColor = Color.BlueViolet
        btnSel4.BackColor = Color.BlueViolet
        btnSel5.BackColor = Color.BlueViolet
        btnSel6.BackColor = Color.BlueViolet
        btnSel7.BackColor = Color.BlueViolet
        btnSel8.BackColor = Color.BlueViolet
        btnSel9.BackColor = Color.BlueViolet
        btnSel1.ForeColor = Color.Black
        btnSel2.ForeColor = Color.Black
        btnSel3.ForeColor = Color.Black
        btnSel4.ForeColor = Color.Black
        btnSel5.ForeColor = Color.Black
        btnSel6.ForeColor = Color.Black
        btnSel7.ForeColor = Color.Black
        btnSel8.ForeColor = Color.Black
        btnSel9.ForeColor = Color.Black
        Me.BackgroundImage = picSpace.Image
        Me.BackgroundImageLayout = ImageLayout.Zoom
    End Sub

    Private Sub LavaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LavaToolStripMenuItem.Click
        Me.BackgroundImage = picLava.Image
        Me.BackgroundImageLayout = ImageLayout.Zoom
        cpucolor = Color.Black
        playercolor = Color.White
        btnbackcolor = Color.DarkRed
        btncolor = Color.Firebrick
        lblDifficulty.BackColor = Color.Firebrick
        lblDifficulty.ForeColor = Color.White
        lblWin.ForeColor = Color.White
        lblWin.BackColor = Color.Firebrick
        Label1.ForeColor = Color.White
        Label1.BackColor = Color.Firebrick
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Firebrick
        Label3.ForeColor = Color.White
        Label3.BackColor = Color.Firebrick
        btnSel1.BackColor = Color.Firebrick
        btnSel2.BackColor = Color.Firebrick
        btnSel3.BackColor = Color.Firebrick
        btnSel4.BackColor = Color.Firebrick
        btnSel5.BackColor = Color.Firebrick
        btnSel6.BackColor = Color.Firebrick
        btnSel7.BackColor = Color.Firebrick
        btnSel8.BackColor = Color.Firebrick
        btnSel9.BackColor = Color.Firebrick
        btnSel1.ForeColor = Color.White
        btnSel2.ForeColor = Color.White
        btnSel3.ForeColor = Color.White
        btnSel4.ForeColor = Color.White
        btnSel5.ForeColor = Color.White
        btnSel6.ForeColor = Color.White
        btnSel7.ForeColor = Color.White
        btnSel8.ForeColor = Color.White
        btnSel9.ForeColor = Color.White
        btnNew.BackColor = Color.Firebrick
        btnNew.ForeColor = Color.White
        btnClear.BackColor = Color.Firebrick
        btnClear.ForeColor = Color.White
        btnExit.BackColor = Color.Firebrick
        btnExit.ForeColor = Color.White
        Me.BackColor = Color.White
    End Sub

    Private Sub PizzaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PizzaToolStripMenuItem.Click
        Me.BackgroundImage = picPizza.Image
        Me.BackgroundImageLayout = ImageLayout.Stretch
        cpucolor = Color.Black
        playercolor = Color.White
        btnbackcolor = Color.DarkRed
        btncolor = Color.Firebrick
        lblDifficulty.BackColor = Color.Firebrick
        lblDifficulty.ForeColor = Color.White
        lblWin.ForeColor = Color.White
        lblWin.BackColor = Color.Firebrick
        Label1.ForeColor = Color.White
        Label1.BackColor = Color.Firebrick
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Firebrick
        Label3.ForeColor = Color.White
        Label3.BackColor = Color.Firebrick
        btnSel1.BackColor = Color.Firebrick
        btnSel2.BackColor = Color.Firebrick
        btnSel3.BackColor = Color.Firebrick
        btnSel4.BackColor = Color.Firebrick
        btnSel5.BackColor = Color.Firebrick
        btnSel6.BackColor = Color.Firebrick
        btnSel7.BackColor = Color.Firebrick
        btnSel8.BackColor = Color.Firebrick
        btnSel9.BackColor = Color.Firebrick
        btnSel1.ForeColor = Color.White
        btnSel2.ForeColor = Color.White
        btnSel3.ForeColor = Color.White
        btnSel4.ForeColor = Color.White
        btnSel5.ForeColor = Color.White
        btnSel6.ForeColor = Color.White
        btnSel7.ForeColor = Color.White
        btnSel8.ForeColor = Color.White
        btnSel9.ForeColor = Color.White
        btnNew.BackColor = Color.Firebrick
        btnNew.ForeColor = Color.White
        btnClear.BackColor = Color.Firebrick
        btnClear.ForeColor = Color.White
        btnExit.BackColor = Color.Firebrick
        btnExit.ForeColor = Color.White
        Me.BackColor = Color.White
    End Sub

    Private Sub NatureToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NatureToolStripMenuItem.Click
        Me.BackgroundImage = picNature.Image
        Me.BackgroundImageLayout = ImageLayout.Stretch
        cpucolor = Color.Red
        playercolor = Color.White
        btnbackcolor = Color.Green
        btncolor = Color.LightGreen
        lblDifficulty.BackColor = Color.Green
        lblDifficulty.ForeColor = Color.White
        lblWin.ForeColor = Color.White
        lblWin.BackColor = Color.Green
        Label1.ForeColor = Color.White
        Label1.BackColor = Color.Green
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Green
        Label3.ForeColor = Color.White
        Label3.BackColor = Color.Green
        btnSel1.BackColor = Color.LightGreen
        btnSel2.BackColor = Color.LightGreen
        btnSel3.BackColor = Color.LightGreen
        btnSel4.BackColor = Color.LightGreen
        btnSel5.BackColor = Color.LightGreen
        btnSel6.BackColor = Color.LightGreen
        btnSel7.BackColor = Color.LightGreen
        btnSel8.BackColor = Color.LightGreen
        btnSel9.BackColor = Color.LightGreen
        btnSel1.ForeColor = Color.White
        btnSel2.ForeColor = Color.White
        btnSel3.ForeColor = Color.White
        btnSel4.ForeColor = Color.White
        btnSel5.ForeColor = Color.White
        btnSel6.ForeColor = Color.White
        btnSel7.ForeColor = Color.White
        btnSel8.ForeColor = Color.White
        btnSel9.ForeColor = Color.White
        btnNew.BackColor = Color.Green
        btnNew.ForeColor = Color.White
        btnClear.BackColor = Color.Green
        btnClear.ForeColor = Color.White
        btnExit.BackColor = Color.Green
        btnExit.ForeColor = Color.White
        Me.BackColor = Color.White
    End Sub

    Private Sub OceanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OceanToolStripMenuItem.Click
        Me.BackgroundImage = picOcean.Image
        Me.BackgroundImageLayout = ImageLayout.Stretch
        cpucolor = Color.White
        playercolor = Color.Black
        btnbackcolor = Color.SkyBlue
        btncolor = Color.LightBlue
        lblDifficulty.BackColor = Color.LightBlue
        lblDifficulty.ForeColor = Color.Black
        lblWin.ForeColor = Color.Black
        lblWin.BackColor = Color.LightBlue
        Label1.ForeColor = Color.Black
        Label1.BackColor = Color.LightBlue
        Label2.ForeColor = Color.Black
        Label2.BackColor = Color.LightBlue
        Label3.ForeColor = Color.Black
        Label3.BackColor = Color.LightBlue
        btnSel1.BackColor = Color.LightBlue
        btnSel2.BackColor = Color.LightBlue
        btnSel3.BackColor = Color.LightBlue
        btnSel4.BackColor = Color.LightBlue
        btnSel5.BackColor = Color.LightBlue
        btnSel6.BackColor = Color.LightBlue
        btnSel7.BackColor = Color.LightBlue
        btnSel8.BackColor = Color.LightBlue
        btnSel9.BackColor = Color.LightBlue
        btnSel1.ForeColor = Color.Black
        btnSel2.ForeColor = Color.Black
        btnSel3.ForeColor = Color.Black
        btnSel4.ForeColor = Color.Black
        btnSel5.ForeColor = Color.Black
        btnSel6.ForeColor = Color.Black
        btnSel7.ForeColor = Color.Black
        btnSel8.ForeColor = Color.Black
        btnSel9.ForeColor = Color.Black
        btnNew.BackColor = Color.LightBlue
        btnNew.ForeColor = Color.Black
        btnClear.BackColor = Color.LightBlue
        btnClear.ForeColor = Color.Black
        btnExit.BackColor = Color.LightBlue
        btnExit.ForeColor = Color.Black
        Me.BackColor = Color.White
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        If MessageBox.Show("Program Name: TicTac v1.0" & vbCrLf & "Programmed by: Anthony Wickham" & vbCrLf & "Date of Last Revision: 3/29/16", "About TicTac",
         MessageBoxButtons.OK, MessageBoxIcon.Information) _
         = DialogResult.OK Then



        End If
    End Sub

    Private Sub HelpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem1.Click
        If MessageBox.Show("Try to get three of your marks in a row!" & vbCrLf & "If nobody gets three in a row, the cat wins!" & vbCrLf & vbCrLf & "Difficulties:" & vbCrLf & "Easy - The CPU will choose spaces at random." & vbCrLf & "Aggressive - The CPU will play offensively and defensively." & vbCrLf & "No Mercy - The CPU cannot be beaten.", "Help",
         MessageBoxButtons.OK, MessageBoxIcon.Information) _
         = DialogResult.OK Then



        End If
    End Sub

    Private Sub frmTTT_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Confirms closing on any close of the form

        If MessageBox.Show("          Are you sure you want to quit?", "Are you sure?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) <> DialogResult.Yes Then
            e.Cancel = True
        End If
    End Sub

    Public Sub gameoverchk()
        ' Does actions if the game was evaluated as over
        If gameover = "cpuwin" Then
            lblWin.Text = cpuMark & " Wins!"
            lblWin.Visible = True

            btnNew.Enabled = True
            btnNew.Focus()
            If cpuMark = "X" Then
                xwins = xwins + 1
                Label1.Text = "X Wins: " & xwins
            Else
                owins = owins + 1
                Label2.Text = "O Wins: " & owins
            End If

        End If
        If gameover = "playerwin" Then
            lblWin.Text = PlayerMark & " Wins!"
            lblWin.Visible = True

            btnNew.Enabled = True
            btnNew.Focus()
            If PlayerMark = "X" Then
                xwins = xwins + 1
                Label1.Text = "X Wins: " & xwins
            Else
                owins = owins + 1
                Label2.Text = "O Wins: " & owins
            End If
        End If
        If gameover = "tie" Then
            lblWin.Text = "Cat Wins!"
            lblWin.Visible = True

            btnNew.Enabled = True
            btnNew.Focus()
            cwins = cwins + 1
            Label3.Text = "Cat Wins: " & cwins
        End If

    End Sub

    Public Sub playerwinchk()
        ' Checks for Player win
        If (btnSel1.Text = PlayerMark) And (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel1.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btnSel2.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel2.Text = PlayerMark) And (btnSel3.Text = PlayerMark) And (btnSel1.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel4.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False

            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel4.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btnSel5.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False

            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel5.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btnSel4.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False

            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False

        ElseIf (btnSel7.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
        ElseIf (btnSel7.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btnSel8.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
        ElseIf (btnSel8.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btnSel7.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False

        ElseIf (btnSel1.Text = PlayerMark) And (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel1.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btnSel4.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel4.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btnSel1.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False

        ElseIf (btnSel2.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel2.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btnSel5.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel5.Text = PlayerMark) And (btnSel8.Text = PlayerMark) And (btnSel2.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel9.Enabled = False

        ElseIf (btnSel3.Text = PlayerMark) And (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel3.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btnSel6.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel6.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btnSel3.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False

        ElseIf (btnSel1.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel1.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btnSel5.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel5.Text = PlayerMark) And (btnSel9.Text = PlayerMark) And (btnSel1.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False

        ElseIf (btnSel3.Text = PlayerMark) And (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel3.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btnSel5.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel5.Text = PlayerMark) And (btnSel7.Text = PlayerMark) And (btnSel3.Text = PlayerMark) Then
            gameover = "playerwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False


        End If


    End Sub

    Public Sub cpuwinchk()
        ' Checks for CPU win
        If (btnSel1.Text = cpuMark) And (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) Then
            gameover = "cpuwin"

            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel1.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btnSel2.Text = cpuMark) Then
            gameover = "cpuwin"

            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel2.Text = cpuMark) And (btnSel3.Text = cpuMark) And (btnSel1.Text = cpuMark) Then
            gameover = "cpuwin"

            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel4.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False

            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel4.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btnSel5.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False

            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel5.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btnSel4.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False

            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False

        ElseIf (btnSel7.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False

        ElseIf (btnSel7.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btnSel8.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False

        ElseIf (btnSel8.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btnSel7.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False

        ElseIf (btnSel1.Text = cpuMark) And (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel1.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btnSel4.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel4.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btnSel1.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False

        ElseIf (btnSel2.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel2.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btnSel5.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel5.Text = cpuMark) And (btnSel8.Text = cpuMark) And (btnSel2.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel9.Enabled = False

        ElseIf (btnSel3.Text = cpuMark) And (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel3.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btnSel6.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel6.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btnSel3.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False

        ElseIf (btnSel1.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel1.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btnSel5.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
        ElseIf (btnSel5.Text = cpuMark) And (btnSel9.Text = cpuMark) And (btnSel1.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False

        ElseIf (btnSel3.Text = cpuMark) And (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel3.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btnSel5.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        ElseIf (btnSel5.Text = cpuMark) And (btnSel7.Text = cpuMark) And (btnSel3.Text = cpuMark) Then
            gameover = "cpuwin"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel4.Enabled = False
            btnSel6.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False

        End If
    End Sub

    Public Sub tiechk()
        If (btnSel1.Text <> "") And (btnSel2.Text <> "") And (btnSel3.Text <> "") And (btnSel4.Text <> "") And (btnSel5.Text <> "") And (btnSel6.Text <> "") And (btnSel7.Text <> "") And (btnSel8.Text <> "") And (btnSel9.Text <> "") Then
            gameover = "tie"
            btnSel1.Enabled = False
            btnSel2.Enabled = False
            btnSel3.Enabled = False
            btnSel4.Enabled = False
            btnSel5.Enabled = False
            btnSel6.Enabled = False
            btnSel7.Enabled = False
            btnSel8.Enabled = False
            btnSel9.Enabled = False
        End If
    End Sub

End Class
