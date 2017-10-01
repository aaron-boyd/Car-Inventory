Imports System.IO

Public Class Form1
    Dim arrVins(10000) As String
    Dim arrCheck(12) As CheckBox
    Dim p As Integer, x As Integer
    Dim pos As Integer, count As Integer, C As Integer

    Private Sub cmbMake_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMake.SelectedIndexChanged
        cmbModel.Text = ""
        cmbYear.Text = ""
        cmbModel.Items.Clear()
        cmbYear.Items.Clear()
        cmbHP.Items.Clear()
        cmbHP.Text = ""
        opt4cyl.Checked = False
        opt6cyl.Checked = False
        opt8cyl.Checked = False
        optFourDoor.Checked = False
        optTwoDoor.Checked = False
        optSixDoor.Checked = False
        chkLeather.Checked = False
        optLight.Checked = False
        optMedium.Checked = False
        optDark.Checked = False
        chkMoon.Checked = False
        chkPoly.Checked = False
        chkTurbo.Checked = False
        chkBack.Checked = False
        chkBluetooth.Checked = False
        chkVoice.Checked = False
        chkWifi.Checked = False
        opt10.Checked = False
        opt12.Checked = False
        opt15.Checked = False
        chkHeated.Checked = False
        chkWheel.Checked = False
        chkMassage.Checked = False
        chkWiper.Checked = False
        opt4cyl.Enabled = False
        opt6cyl.Enabled = False
        opt8cyl.Enabled = False
        optTwoDoor.Enabled = False
        optFourDoor.Enabled = False
        optSixDoor.Enabled = False
        opt10.Enabled = False
        opt12.Enabled = False
        opt15.Enabled = False
        optLight.Enabled = False
        optMedium.Enabled = False
        optDark.Enabled = False
        If cmbMake.Text = "Lamborghini" Then
            pic1.Image = Image.FromFile("D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\Lamborghini Logo.png")
            cmbModel.Items.Add("Huracan")
            cmbModel.Items.Add("Aventador")
            cmbModel.Items.Add("Murciélago")
        ElseIf cmbMake.Text = "Bentley" Then
            pic1.Image = Image.FromFile("D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\Bentley Logo.png")
            cmbModel.Items.Add("Bentayga")
            cmbModel.Items.Add("Flying Spur")
            cmbModel.Items.Add("Mulsanne")
        ElseIf cmbMake.Text = "Maserati" Then
            pic1.Image = Image.FromFile("D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\Maserati Logo.png")
            cmbModel.Items.Add("Grancabrio")
            cmbModel.Items.Add("Ghibli")
            cmbModel.Items.Add("Quattroporte")
        End If
        If count < 1 Then
            count = count + 1
            progbar1.Value = 12.5
        End If
        cmbModel.Enabled = True
        'lblX.Text = Str(x)
    End Sub

    Private Sub cmbModel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbModel.SelectedIndexChanged
        cmbYear.Items.Clear()
        cmbYear.Items.Add("2013")
        cmbYear.Items.Add("2014")
        cmbYear.Items.Add("2015")
        cmbYear.Items.Add("2016")
        'lblX.Text = Str(x)
        If count < 2 Then
            count = count + 1
            progbar1.Value = 25
            'MessageBox.Show(count)w
        End If
        cmbYear.Enabled = True
    End Sub

    Private Sub cmdVin_Click(sender As Object, e As EventArgs) Handles cmdVin.Click
        Dim Vin As String
        txtVin.Text = ""
        If cmbMake.Text = "Bentley" Then
            Vin = Vin + "BT"
            If cmbModel.Text = "Bentayga" Then
                Vin = Vin + "BT"
            ElseIf cmbModel.Text = "Flying Spur" Then
                Vin = Vin + "FS"
            ElseIf cmbModel.Text = "Mulsanne" Then
                Vin = Vin + "MS"
            End If
        ElseIf cmbMake.Text = "Lamborghini" Then
            Vin = Vin + "LB"
            If cmbModel.Text = "Huracan" Then
                Vin = Vin + "HN"
            ElseIf cmbModel.Text = "Aventador" Then
                Vin = Vin + "AV"
            ElseIf cmbModel.Text = "Murciélago" Then
                Vin = Vin + "MG"
            End If
        ElseIf cmbMake.Text = "Maserati" Then
            Vin = Vin + "MS"
            If cmbModel.Text = "Grancabrio" Then
                Vin = Vin + "GC"
            ElseIf cmbModel.Text = "Ghibli" Then
                Vin = Vin + "GB"
            ElseIf cmbModel.Text = "Quattroporte" Then
                Vin = Vin + "QP"
            End If
        End If
        If cmbYear.Text = "2013" Then
            Vin = Vin + "13"
        ElseIf cmbYear.Text = "2014" Then
            Vin = Vin + "14"
        ElseIf cmbYear.Text = "2015" Then
            Vin = Vin + "15"
        ElseIf cmbYear.Text = "2016" Then
            Vin = Vin + "16"
        End If
        If cmbHP.Text = "600" Then
            Vin = Vin + "6"
        ElseIf cmbHP.Text = "700" Then
            Vin = Vin + "7"
        ElseIf cmbHP.Text = "800" Then
            Vin = Vin + "8"
        ElseIf cmbHP.Text = "900" Then
            Vin = Vin + "9"
        ElseIf cmbHP.Text = "1000" Then
            Vin = Vin + "1"
        End If
        If opt4cyl.Checked = True Then
            Vin = Vin + "4"
        ElseIf opt6cyl.Checked = True Then
            Vin = Vin + "6"
        ElseIf opt8cyl.Checked = True Then
            Vin = Vin + "8"
        End If
        If optFourDoor.Checked = True Then
            Vin = Vin + "4"
        ElseIf optTwoDoor.Checked = True Then
            Vin = Vin + "2"
        ElseIf optSixDoor.Checked = True Then
            Vin = Vin + "6"
        End If
        If opt10.Checked = True Then
            Vin = Vin + "10"
        ElseIf opt12.Checked = True Then
            Vin = Vin + "12"
        ElseIf opt15.Checked = True Then
            Vin = Vin + "15"
        End If
        If optLight.Checked = True Then
            Vin = Vin + "L"
        ElseIf optMedium.Checked = True Then
            Vin = Vin + "M"
        ElseIf optDark.Checked = True Then
            Vin = Vin + "D"
        End If
        For p = 1 To 12
            If arrCheck(p).Checked = True Then
                Vin = Vin + "1"
            ElseIf arrCheck(p).Checked = False Then
                Vin = Vin + "0"
            End If
        Next
        txtVin.Text = Vin
        'lstVins.Items.Add(Vin)
        cmbMake.Text = ""
        cmbModel.Text = ""
        cmbYear.Text = ""
        cmbHP.Text = ""
        cmbModel.Items.Clear()
        cmbYear.Items.Clear()
        opt4cyl.Checked = False
        opt6cyl.Checked = False
        opt8cyl.Checked = False
        optLight.Checked = False
        optMedium.Checked = False
        optDark.Checked = False
        optFourDoor.Checked = False
        optTwoDoor.Checked = False
        optSixDoor.Checked = False
        chkLeather.Checked = False
        chkMoon.Checked = False
        chkPoly.Checked = False
        chkTurbo.Checked = False
        chkBack.Checked = False
        chkBluetooth.Checked = False
        chkVoice.Checked = False
        chkWifi.Checked = False
        chkHeated.Checked = False
        chkWheel.Checked = False
        chkMassage.Checked = False
        chkWiper.Checked = False
        opt10.Checked = False
        opt12.Checked = False
        opt15.Checked = False
        cmbModel.Enabled = False
        cmbYear.Enabled = False
        cmbHP.Enabled = False
        opt4cyl.Enabled = False
        opt6cyl.Enabled = False
        opt8cyl.Enabled = False
        optTwoDoor.Enabled = False
        optFourDoor.Enabled = False
        optSixDoor.Enabled = False
        opt10.Enabled = False
        opt12.Enabled = False
        opt15.Enabled = False
        optLight.Enabled = False
        optMedium.Enabled = False
        optDark.Enabled = False
        cmdVin.Enabled = False
        arrVins(x) = Vin
        lstVins.Items.Add(arrVins(x))
        x = x + 1
        'lblX.Text = Str(x)
        progbar1.Value = 0
        count = 0
        pic1.Image = Nothing
    End Sub

    Private Sub lstVins_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lstVins.MouseDoubleClick
        Dim vin As String
        vin = lstVins.Text
        txtVin.Text = vin
        cmbMake.Text = ""
        cmbModel.Text = ""
        cmbYear.Text = ""
        cmbModel.Text = ""
        cmbModel.Enabled = True
        cmbYear.Enabled = True
        cmbHP.Enabled = True
        cmbModel.Items.Clear()
        cmbYear.Items.Clear()
        opt4cyl.Checked = False
        opt6cyl.Checked = False
        opt8cyl.Checked = False
        optLight.Checked = False
        optMedium.Checked = False
        optDark.Checked = False
        optFourDoor.Checked = False
        optTwoDoor.Checked = False
        chkLeather.Checked = False
        chkMoon.Checked = False
        chkPoly.Checked = False
        chkTurbo.Checked = False
        chkBack.Checked = False
        chkBluetooth.Checked = False
        chkVoice.Checked = False
        chkWifi.Checked = False
        chkHeated.Checked = False
        chkWheel.Checked = False
        chkMassage.Checked = False
        chkWiper.Checked = False
        opt10.Checked = False
        opt12.Checked = False
        opt15.Checked = False
        cmdVin.Enabled = False
        opt4cyl.Enabled = True
        opt6cyl.Enabled = True
        opt8cyl.Enabled = True
        optTwoDoor.Enabled = True
        optFourDoor.Enabled = True
        optSixDoor.Enabled = True
        opt10.Enabled = True
        opt12.Enabled = True
        opt15.Enabled = True
        optLight.Enabled = True
        optMedium.Enabled = True
        optDark.Enabled = True

        If Mid(vin, 1, 2) = "BT" Then
            cmbMake.Text = "Bentley"
            cmbModel.Items.Clear()
            cmbModel.Items.Add("Bentayga")
            cmbModel.Items.Add("Flying Spur")
            cmbModel.Items.Add("Mulsanne")
        ElseIf Mid(vin, 1, 2) = "LB" Then
            cmbMake.Text = "Lamborghini"
            cmbModel.Items.Clear()
            cmbModel.Items.Add("Huracan")
            cmbModel.Items.Add("Aventador")
            cmbModel.Items.Add("Murciélago")
        ElseIf Mid(vin, 1, 2) = "MS" Then
            cmbMake.Text = "Maserati"
            cmbModel.Items.Clear()
            cmbModel.Items.Add("Grancabrio")
            cmbModel.Items.Add("Ghibli")
            cmbModel.Items.Add("Quattroporte")
        End If
        If Mid(vin, 3, 2) = "BT" Then
            cmbModel.Text = "Bentayga"
        ElseIf Mid(vin, 3, 2) = "FS" Then
            cmbModel.Text = "Flying Spur"
        ElseIf Mid(vin, 3, 2) = "MS" Then
            cmbModel.Text = "Mulsanne"
        ElseIf Mid(vin, 3, 2) = "HN" Then
            cmbModel.Text = "Huracan"
        ElseIf Mid(vin, 3, 2) = "AV" Then
            cmbModel.Text = "Aventador"
        ElseIf Mid(vin, 3, 2) = "MG" Then
            cmbModel.Text = "Murciélago"
        ElseIf Mid(vin, 3, 2) = "GC" Then
            cmbModel.Text = "Grancabrio"
        ElseIf Mid(vin, 3, 2) = "GB" Then
            cmbModel.Text = "Ghibli"
        ElseIf Mid(vin, 3, 2) = "QP" Then
            cmbModel.Text = "Quattroporte"
        End If
        If Mid(vin, 5, 2) = "13" Then
            cmbYear.Text = "2013"
        ElseIf Mid(vin, 5, 2) = "14" Then
            cmbYear.Text = "2014"
        ElseIf Mid(vin, 5, 2) = "15" Then
            cmbYear.Text = "2015"
        ElseIf Mid(vin, 5, 2) = "16" Then
            cmbYear.Text = "2016"
        End If
        If Mid(vin, 7, 1) = "6" Then
            cmbHP.Text = "600"
        ElseIf Mid(vin, 7, 1) = "7" Then
            cmbHP.Text = "700"
        ElseIf Mid(vin, 7, 1) = "8" Then
            cmbHP.Text = "800"
        ElseIf Mid(vin, 7, 1) = "9" Then
            cmbHP.Text = "900"
        ElseIf Mid(vin, 7, 1) = "1" Then
            cmbHP.Text = "1000"
        End If
        If Mid(vin, 8, 1) = "4" Then
            opt4cyl.Checked = True
        ElseIf Mid(vin, 8, 1) = "6" Then
            opt6cyl.Checked = True
        ElseIf Mid(vin, 8, 1) = "8" Then
            opt8cyl.Checked = True
        End If
        If Mid(vin, 9, 1) = "4" Then
            optFourDoor.Checked = True
        ElseIf Mid(vin, 9, 1) = "2" Then
            optTwoDoor.Checked = True
        ElseIf Mid(vin, 9, 1) = "6" Then
            optSixDoor.Checked = True
        End If
        If Mid(vin, 10, 2) = "10" Then
            opt10.Checked = True
        ElseIf Mid(vin, 10, 2) = "12" Then
            opt12.Checked = True
        ElseIf Mid(vin, 10, 2) = "15" Then
            opt15.Checked = True
        End If
        If Mid(vin, 12, 1) = "L" Then
            optLight.Checked = True
        ElseIf Mid(vin, 12, 1) = "M" Then
            optMedium.Checked = True
        ElseIf Mid(vin, 12, 1) = "D" Then
            optDark.Checked = True
        End If
        For p = 1 To 12
            If Mid(vin, (p + 12), 1) = "1" Then
                arrCheck(p).Checked = True
            End If
        Next
        cmbMake.Enabled = False
        cmdVin.Enabled = False
        If cmbMake.Text = "Bentley" Then
            pic1.Image = Image.FromFile("D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\Bentley Logo.png")
        ElseIf cmbMake.Text = "Lamborghini" Then
            pic1.Image = Image.FromFile("D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\Lamborghini Logo.png")
        ElseIf cmbMake.Text = "Maserati" Then
            pic1.Image = Image.FromFile("D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\Maserati Logo.png")
        End If
        progbar1.Value = 0
        count = 100
        C = lstVins.SelectedIndex
        cmdChange.Enabled = True
        'lblX.Text = Str(x)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        x = 0
        count = 0
        arrCheck(1) = chkMoon
        arrCheck(2) = chkLeather
        arrCheck(3) = chkPoly
        arrCheck(4) = chkTurbo
        arrCheck(5) = chkBack
        arrCheck(6) = chkBluetooth
        arrCheck(7) = chkWifi
        arrCheck(8) = chkVoice
        arrCheck(9) = chkHeated
        arrCheck(10) = chkWheel
        arrCheck(11) = chkMassage
        arrCheck(12) = chkWiper

        cmbModel.Enabled = False
        cmbYear.Enabled = False
        cmbHP.Enabled = False
        opt4cyl.Enabled = False
        opt6cyl.Enabled = False
        opt8cyl.Enabled = False
        optTwoDoor.Enabled = False
        optFourDoor.Enabled = False
        optSixDoor.Enabled = False
        opt10.Enabled = False
        cmdVin.Enabled = False
        opt12.Enabled = False
        opt15.Enabled = False
        optLight.Enabled = False
        optMedium.Enabled = False
        optDark.Enabled = False

    End Sub

    Private Sub cmdOpen_Click(sender As Object, e As EventArgs) Handles cmdOpen.Click
        Dim File As StreamReader, i As Integer
        File = New StreamReader("D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\CarVins.txt")
        Do While Not File.EndOfStream
            arrVins(x) = File.ReadLine
            lstVins.Items.Add(arrVins(x))
            x = x + 1
        Loop
        File.Close()
        cmdOpen.Enabled = False
        txtVin.Text = ""
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmbHorse.Click

        FileOpen(1, "D:\School\Computer Programming II\VB.Net Projects\Car Inventory\Car Inventory\Car Inventory\CarVins.txt", OpenMode.Output)
        For i = 0 To (x - 1)
            PrintLine(1, arrVins(i))
        Next i
        FileClose(1)
        'lblX.Text = Str(x)
    End Sub

    Private Sub cmbYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbYear.SelectedIndexChanged
        If count < 3 Then
            count = count + 1
            progbar1.Value = 37.5
        End If
        cmbHP.Items.Clear()
        cmbHP.Items.Add("600")
        cmbHP.Items.Add("700")
        cmbHP.Items.Add("800")
        cmbHP.Items.Add("900")
        cmbHP.Items.Add("1000")
        cmbHP.Enabled = True
    End Sub
    Private Sub cmbHP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHP.SelectedIndexChanged
        If count < 4 Then
            count = count + 1
            progbar1.Value = 50
        End If
        opt4cyl.Enabled = True
        opt6cyl.Enabled = True
        opt8cyl.Enabled = True
    End Sub

    Private Sub opt4cyl_CheckedChanged(sender As Object, e As EventArgs) Handles opt4cyl.CheckedChanged
        If count < 5 Then
            count = count + 1
            progbar1.Value = 62.5
        End If
        optFourDoor.Enabled = True
        optTwoDoor.Enabled = True
        optSixDoor.Enabled = True
    End Sub

    Private Sub opt6cyl_CheckedChanged(sender As Object, e As EventArgs) Handles opt6cyl.CheckedChanged
        If count < 5 Then
            count = count + 1
            progbar1.Value = 62.5
        End If
        optFourDoor.Enabled = True
        optTwoDoor.Enabled = True
        optSixDoor.Enabled = True
    End Sub

    Private Sub opt8cyl_CheckedChanged(sender As Object, e As EventArgs) Handles opt8cyl.CheckedChanged
        If count < 5 Then
            count = count + 1
            progbar1.Value = 62.5
        End If
        optFourDoor.Enabled = True
        optTwoDoor.Enabled = True
        optSixDoor.Enabled = True
    End Sub

    Private Sub optTwoDoor_CheckedChanged(sender As Object, e As EventArgs) Handles optTwoDoor.CheckedChanged
        If count < 6 Then
            count = count + 1
            progbar1.Value = 75
        End If
        opt10.Enabled = True
        opt12.Enabled = True
        opt15.Enabled = True
    End Sub

    Private Sub optFourDoor_CheckedChanged(sender As Object, e As EventArgs) Handles optFourDoor.CheckedChanged
        If count < 6 Then
            count = count + 1
            progbar1.Value = 75
        End If
        opt10.Enabled = True
        opt12.Enabled = True
        opt15.Enabled = True
    End Sub

    Private Sub optSixDoor_CheckedChanged(sender As Object, e As EventArgs) Handles optSixDoor.CheckedChanged
        If count < 6 Then
            count = count + 1
            progbar1.Value = 75
        End If
        opt10.Enabled = True
        opt12.Enabled = True
        opt15.Enabled = True
    End Sub

    Private Sub opt10_CheckedChanged(sender As Object, e As EventArgs) Handles opt10.CheckedChanged
        If count < 7 Then
            count = count + 1
            progbar1.Value = 87.5
        End If
        optLight.Enabled = True
        optMedium.Enabled = True
        optDark.Enabled = True
    End Sub

    Private Sub opt12_CheckedChanged(sender As Object, e As EventArgs) Handles opt12.CheckedChanged
        If count < 7 Then
            count = count + 1
            progbar1.Value = 87.5
        End If
        optLight.Enabled = True
        optMedium.Enabled = True
        optDark.Enabled = True
    End Sub

    Private Sub opt15_CheckedChanged(sender As Object, e As EventArgs) Handles opt15.CheckedChanged
        If count < 7 Then
            count = count + 1
            progbar1.Value = 87.5
        End If
        optLight.Enabled = True
        optMedium.Enabled = True
        optDark.Enabled = True
    End Sub

    Private Sub optLight_CheckedChanged(sender As Object, e As EventArgs) Handles optLight.CheckedChanged
        If count < 8 Then
            count = count + 1
            progbar1.Value = 100
        End If
        cmdVin.Enabled = True
    End Sub

    Private Sub optMedium_CheckedChanged(sender As Object, e As EventArgs) Handles optMedium.CheckedChanged
        If count < 8 Then
            count = count + 1
            progbar1.Value = 100
        End If
        cmdVin.Enabled = True
    End Sub

    Private Sub cmdChange_Click(sender As Object, e As EventArgs) Handles cmdChange.Click
        arrVins(C) = ""
        Dim Vin As String
        txtVin.Text = ""
        If cmbMake.Text = "Bentley" Then
            Vin = Vin + "BT"
            If cmbModel.Text = "Bentayga" Then
                Vin = Vin + "BT"
            ElseIf cmbModel.Text = "Flying Spur" Then
                Vin = Vin + "FS"
            ElseIf cmbModel.Text = "Mulsanne" Then
                Vin = Vin + "MS"
            End If
        ElseIf cmbMake.Text = "Lamborghini" Then
            Vin = Vin + "LB"
            If cmbModel.Text = "Huracan" Then
                Vin = Vin + "HN"
            ElseIf cmbModel.Text = "Aventador" Then
                Vin = Vin + "AV"
            ElseIf cmbModel.Text = "Murciélago" Then
                Vin = Vin + "MG"
            End If
        ElseIf cmbMake.Text = "Maserati" Then
            Vin = Vin + "MS"
            If cmbModel.Text = "Grancabrio" Then
                Vin = Vin + "GC"
            ElseIf cmbModel.Text = "Ghibli" Then
                Vin = Vin + "GB"
            ElseIf cmbModel.Text = "Quattroporte" Then
                Vin = Vin + "QP"
            End If
        End If
        If cmbYear.Text = "2013" Then
            Vin = Vin + "13"
        ElseIf cmbYear.Text = "2014" Then
            Vin = Vin + "14"
        ElseIf cmbYear.Text = "2015" Then
            Vin = Vin + "15"
        ElseIf cmbYear.Text = "2016" Then
            Vin = Vin + "16"
        End If
        If cmbHP.Text = "600" Then
            Vin = Vin + "6"
        ElseIf cmbHP.Text = "700" Then
            Vin = Vin + "7"
        ElseIf cmbHP.Text = "800" Then
            Vin = Vin + "8"
        ElseIf cmbHP.Text = "900" Then
            Vin = Vin + "9"
        ElseIf cmbHP.Text = "1000" Then
            Vin = Vin + "1"
        End If
        If opt4cyl.Checked = True Then
            Vin = Vin + "4"
        ElseIf opt6cyl.Checked = True Then
            Vin = Vin + "6"
        ElseIf opt8cyl.Checked = True Then
            Vin = Vin + "8"
        End If
        If optFourDoor.Checked = True Then
            Vin = Vin + "4"
        ElseIf optTwoDoor.Checked = True Then
            Vin = Vin + "2"
        ElseIf optSixDoor.Checked = True Then
            Vin = Vin + "6"
        End If
        If opt10.Checked = True Then
            Vin = Vin + "10"
        ElseIf opt12.Checked = True Then
            Vin = Vin + "12"
        ElseIf opt15.Checked = True Then
            Vin = Vin + "15"
        End If
        If optLight.Checked = True Then
            Vin = Vin + "L"
        ElseIf optMedium.Checked = True Then
            Vin = Vin + "M"
        ElseIf optDark.Checked = True Then
            Vin = Vin + "D"
        End If
        For p = 1 To 12
            If arrCheck(p).Checked = True Then
                Vin = Vin + "1"
            ElseIf arrCheck(p).Checked = False Then
                Vin = Vin + "0"
            End If
        Next
        txtVin.Text = Vin
        'lstVins.Items.Add(Vin)
        cmbMake.Text = ""
        cmbModel.Text = ""
        cmbYear.Text = ""
        cmbHP.Text = ""
        cmbModel.Items.Clear()
        cmbYear.Items.Clear()
        opt4cyl.Checked = False
        opt6cyl.Checked = False
        opt8cyl.Checked = False
        optLight.Checked = False
        optMedium.Checked = False
        optDark.Checked = False
        optFourDoor.Checked = False
        optTwoDoor.Checked = False
        optSixDoor.Checked = False
        chkLeather.Checked = False
        chkMoon.Checked = False
        chkPoly.Checked = False
        chkTurbo.Checked = False
        chkBack.Checked = False
        chkBluetooth.Checked = False
        chkVoice.Checked = False
        chkWifi.Checked = False
        chkHeated.Checked = False
        chkWheel.Checked = False
        chkMassage.Checked = False
        chkWiper.Checked = False
        opt10.Checked = False
        opt12.Checked = False
        opt15.Checked = False
        cmbModel.Enabled = False
        cmbYear.Enabled = False
        cmbHP.Enabled = False
        opt4cyl.Enabled = False
        opt6cyl.Enabled = False
        opt8cyl.Enabled = False
        optTwoDoor.Enabled = False
        optFourDoor.Enabled = False
        optSixDoor.Enabled = False
        opt10.Enabled = False
        opt12.Enabled = False
        opt15.Enabled = False
        optLight.Enabled = False
        optMedium.Enabled = False
        optDark.Enabled = False
        cmdVin.Enabled = False
        arrVins(C) = Vin
        lstVins.Items.Clear()
        progbar1.Value = 0
        count = 0
        pic1.Image = Nothing
        For i = 0 To (x - 1)
            lstVins.Items.Add(arrVins(i))
        Next i
        txtVin.Text = Vin
        cmdChange.Enabled = False
        cmbMake.Enabled = True
        cmdDele.Enabled = False
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Dim q As Integer, p As Integer, search As String, o As Integer
        Dim index As Integer, space As String
        space = " "
        index = 0
        search = UCase(txtSearch.Text)
        txtSearchBox.Text = ""
        If txtSearch.Text <> "" Then
            For q = 0 To x
                For p = 1 To Len(arrVins(q))
                    If Mid(arrVins(q), p, Len(search)) = search Then
                        txtSearchBox.Text = txtSearchBox.Text & Str(q) & space & arrVins(q) & ControlChars.NewLine
                        Exit For
                    End If
                Next p
            Next q
            While index < txtSearchBox.Text.LastIndexOf(search)
                txtSearchBox.Find(search, index, txtSearchBox.TextLength, RichTextBoxFinds.None)
                txtSearchBox.SelectionColor = Color.Green
                index = txtSearchBox.Text.IndexOf(search, index) + 1
            End While
        End If
    End Sub

    Private Sub lstVins_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstVins.SelectedIndexChanged
        pos = lstVins.SelectedIndex
        cmdDele.Enabled = True
    End Sub

    Private Sub cmdDele_Click(sender As Object, e As EventArgs) Handles cmdDele.Click
        lstVins.Items.Clear()
        For i = pos To x - 1
            arrVins(i) = arrVins(i + 1)
        Next i
        x = x - 2
        For i = 0 To x
            lstVins.Items.Add(arrVins(i))
        Next i
        x = x + 1
        cmdDele.Enabled = False
        cmbMake.Text = ""
        cmbModel.Text = ""
        cmbYear.Text = ""
        cmbHP.Text = ""
        cmbModel.Items.Clear()
        cmbYear.Items.Clear()
        opt4cyl.Checked = False
        opt6cyl.Checked = False
        opt8cyl.Checked = False
        optLight.Checked = False
        optMedium.Checked = False
        optDark.Checked = False
        optFourDoor.Checked = False
        optTwoDoor.Checked = False
        optSixDoor.Checked = False
        chkLeather.Checked = False
        chkMoon.Checked = False
        chkPoly.Checked = False
        chkTurbo.Checked = False
        chkBack.Checked = False
        chkBluetooth.Checked = False
        chkVoice.Checked = False
        chkWifi.Checked = False
        chkHeated.Checked = False
        chkWheel.Checked = False
        chkMassage.Checked = False
        chkWiper.Checked = False
        opt10.Checked = False
        opt12.Checked = False
        opt15.Checked = False
        cmbModel.Enabled = False
        cmbYear.Enabled = False
        cmbHP.Enabled = False
        opt4cyl.Enabled = False
        opt6cyl.Enabled = False
        opt8cyl.Enabled = False
        optTwoDoor.Enabled = False
        optFourDoor.Enabled = False
        optSixDoor.Enabled = False
        opt10.Enabled = False
        opt12.Enabled = False
        opt15.Enabled = False
        optLight.Enabled = False
        optMedium.Enabled = False
        optDark.Enabled = False
        cmdVin.Enabled = False
        pic1.Image = Nothing
    End Sub

    Private Sub optDark_CheckedChanged(sender As Object, e As EventArgs) Handles optDark.CheckedChanged
        If count < 8 Then
            count = count + 1
            progbar1.Value = 100
        End If
        cmdVin.Enabled = True
    End Sub
End Class
