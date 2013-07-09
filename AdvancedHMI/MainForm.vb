Public Class MainForm
    Public RunFlag As Integer
    Public CntQty As Integer

    '*******************************************************************************
    '* Stop polling when the form is not visible in order to reduce communications
    '* Copy this section of code to every new form created
    '*******************************************************************************
    Private Sub Form_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If components IsNot Nothing Then
            Dim drv As AdvancedHMIDrivers.IComComponent
            '*****************************
            '* Search for comm components
            '*****************************
            For i As Integer = 0 To components.Components.Count - 1
                If components.Components(i).GetType.GetInterface("AdvancedHMIDrivers.IComComponent") IsNot Nothing Then
                    drv = components.Components.Item(i)
                    '* Stop/Start polling based on form visibility
                    drv.DisableSubscriptions = Not Me.Visible
                End If
            Next
        End If
    End Sub

    '************************************************************
    '* This will guarantee that even hidden forms are all closed
    '* when the main application is closed
    '************************************************************
    'Private Sub DemoForm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    Environment.Exit(0)
    'End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim N272Flag As String
        Dim N274Flag As String

        N272Flag = DF1Comm1.ReadAny("N27:2")
        N274Flag = DF1Comm1.ReadAny("N27:4")

        If Qty1.Value = 0 Or Length1.Value = 0 Then
            RunFlag = 0
        End If

        Label27.Text = N272Flag
        Label28.Text = N274Flag
        Label29.Text = Timer1.Interval

        If RunFlag = 1 Then
            If CntQty = 20 And N272Flag = 0 And N274Flag = 0 And Qty20.Value <> 0 And Length20.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty20.Value)
                DF1Comm1.WriteData("F18:25", Length20.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty20.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 19 And N272Flag = 0 And N274Flag = 0 And Qty19.Value <> 0 And Length19.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty19.Value)
                DF1Comm1.WriteData("F18:25", Length19.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty19.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 18 And N272Flag = 0 And N274Flag = 0 And Qty18.Value <> 0 And Length18.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty18.Value)
                DF1Comm1.WriteData("F18:25", Length18.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty18.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 17 And N272Flag = 0 And N274Flag = 0 And Qty17.Value <> 0 And Length17.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty17.Value)
                DF1Comm1.WriteData("F18:25", Length17.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty17.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 16 And N272Flag = 0 And N274Flag = 0 And Qty16.Value <> 0 And Length16.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty16.Value)
                DF1Comm1.WriteData("F18:25", Length16.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty16.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 15 And N272Flag = 0 And N274Flag = 0 And Qty15.Value <> 0 And Length15.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty15.Value)
                DF1Comm1.WriteData("F18:25", Length15.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty15.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 14 And N272Flag = 0 And N274Flag = 0 And Qty14.Value <> 0 And Length14.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty14.Value)
                DF1Comm1.WriteData("F18:25", Length14.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty14.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 13 And N272Flag = 0 And N274Flag = 0 And Qty13.Value <> 0 And Length13.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty13.Value)
                DF1Comm1.WriteData("F18:25", Length13.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty13.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 12 And N272Flag = 0 And N274Flag = 0 And Qty12.Value <> 0 And Length12.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty12.Value)
                DF1Comm1.WriteData("F18:25", Length12.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty12.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 11 And N272Flag = 0 And N274Flag = 0 And Qty11.Value <> 0 And Length11.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty11.Value)
                DF1Comm1.WriteData("F18:25", Length11.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty11.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 10 And N272Flag = 0 And N274Flag = 0 And Qty10.Value <> 0 And Length10.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty10.Value)
                DF1Comm1.WriteData("F18:25", Length10.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty10.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 9 And N272Flag = 0 And N274Flag = 0 And Qty9.Value <> 0 And Length9.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty9.Value)
                DF1Comm1.WriteData("F18:25", Length9.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty9.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 8 And N272Flag = 0 And N274Flag = 0 And Qty8.Value <> 0 And Length8.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty8.Value)
                DF1Comm1.WriteData("F18:25", Length8.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty8.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 7 And N272Flag = 0 And N274Flag = 0 And Qty7.Value <> 0 And Length7.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty7.Value)
                DF1Comm1.WriteData("F18:25", Length7.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty7.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 6 And N272Flag = 0 And N274Flag = 0 And Qty6.Value <> 0 And Length6.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty6.Value)
                DF1Comm1.WriteData("F18:25", Length6.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty6.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 5 And N272Flag = 0 And N274Flag = 0 And Qty5.Value <> 0 And Length5.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty5.Value)
                DF1Comm1.WriteData("F18:25", Length5.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty5.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 4 And N272Flag = 0 And N274Flag = 0 And Qty4.Value <> 0 And Length4.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty4.Value)
                DF1Comm1.WriteData("F18:25", Length4.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty4.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 3 And N272Flag = 0 And N274Flag = 0 And Qty3.Value <> 0 And Length3.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty3.Value)
                DF1Comm1.WriteData("F18:25", Length3.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty3.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 2 And N272Flag = 0 And N274Flag = 0 And Qty2.Value <> 0 And Length2.Value <> 0 Then
                CntQty = CntQty + 1
                DF1Comm1.WriteData("N27:3", Qty2.Value)
                DF1Comm1.WriteData("F18:25", Length2.Value)
                'DF1Comm1.WriteData("B23/6", 1)
                DF1Comm1.WriteData("B23/3", 0)
                DF1Comm1.WriteData("B23/3", 1)
                Qty2.BackColor = Color.BlueViolet
                Timer1.Interval = 3000 + Length1.Value * 325
            End If
            If CntQty = 1 And Qty1.Value <> 0 And Length1.Value <> 0 Then
                CntQty = CntQty + 1
                Try
                    DF1Comm1.WriteData("N27:3", Qty1.Value)
                    DF1Comm1.WriteData("F18:25", Length1.Value)
                    'DF1Comm1.WriteData("B23/6", 1)
                    DF1Comm1.WriteData("B23/3", 0)
                    DF1Comm1.WriteData("B23/3", 1)
                    N272Flag = DF1Comm1.ReadAny("N27:2")
                    N274Flag = DF1Comm1.ReadAny("N27:4")
                    Qty1.BackColor = Color.BlueViolet
                    Timer1.Interval = 3000 + Length1.Value * 325
                Catch ex As Exception
                End Try

            End If
        End If
    End Sub



    Private Sub Btn_Run_Click(sender As Object, e As EventArgs) Handles Btn_Run.Click
        RunFlag = 1
        CntQty = 1
        DF1Comm1.WriteData("B23/6", 1)
        Timer1.Interval = 2000
    End Sub

    Private Sub Btn_Clear_Click(sender As Object, e As EventArgs) Handles Btn_Clear.Click
        Qty1.Value = 0
        Qty2.Value = 0
        Qty3.Value = 0
        Qty4.Value = 0
        Qty5.Value = 0
        Qty6.Value = 0
        Qty7.Value = 0
        Qty8.Value = 0
        Qty9.Value = 0
        Qty10.Value = 0
        Qty11.Value = 0
        Qty12.Value = 0
        Qty13.Value = 0
        Qty14.Value = 0
        Qty15.Value = 0
        Qty16.Value = 0
        Qty17.Value = 0
        Qty18.Value = 0
        Qty19.Value = 0
        Qty20.Value = 0

        Length1.Value = 0
        Length2.Value = 0
        Length3.Value = 0
        Length4.Value = 0
        Length5.Value = 0
        Length6.Value = 0
        Length7.Value = 0
        Length8.Value = 0
        Length9.Value = 0
        Length10.Value = 0
        Length11.Value = 0
        Length12.Value = 0
        Length13.Value = 0
        Length14.Value = 0
        Length15.Value = 0
        Length16.Value = 0
        Length17.Value = 0
        Length18.Value = 0
        Length19.Value = 0
        Length20.Value = 0
        Qty1.BackColor = Color.White
        Qty2.BackColor = Color.White
        Qty3.BackColor = Color.White
        Qty4.BackColor = Color.White
        Qty5.BackColor = Color.White
        Qty6.BackColor = Color.White
        Qty7.BackColor = Color.White
        Qty8.BackColor = Color.White
        Qty9.BackColor = Color.White
        Qty10.BackColor = Color.White
        Qty11.BackColor = Color.White
        Qty12.BackColor = Color.White
        Qty13.BackColor = Color.White
        Qty14.BackColor = Color.White
        Qty15.BackColor = Color.White
        Qty16.BackColor = Color.White
        Qty17.BackColor = Color.White
        Qty18.BackColor = Color.White
        Qty19.BackColor = Color.White
        Qty20.BackColor = Color.White


    End Sub

    Private Sub Btn_Stop_Click(sender As Object, e As EventArgs) Handles Btn_Stop.Click
        RunFlag = 0
        Qty1.BackColor = Color.White
        Qty2.BackColor = Color.White
        Qty3.BackColor = Color.White
        Qty4.BackColor = Color.White
        Qty5.BackColor = Color.White
        Qty6.BackColor = Color.White
        Qty7.BackColor = Color.White
        Qty8.BackColor = Color.White
        Qty9.BackColor = Color.White
        Qty10.BackColor = Color.White
        Qty11.BackColor = Color.White
        Qty12.BackColor = Color.White
        Qty13.BackColor = Color.White
        Qty14.BackColor = Color.White
        Qty15.BackColor = Color.White
        Qty16.BackColor = Color.White
        Qty17.BackColor = Color.White
        Qty18.BackColor = Color.White
        Qty19.BackColor = Color.White
        Qty20.BackColor = Color.White
    End Sub


 
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub
End Class
