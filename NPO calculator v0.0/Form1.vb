Public Class Form1

    Public Function LRDensity(D1 As Decimal, D2 As Decimal, T1 As Decimal, T2 As Decimal, T3 As Decimal)
        'Calculating density at given temperature by lineal regression
        Dim A As Decimal = (D2 - D1) / (T2 - T1)
        Dim B As Decimal = D1 - T1 * (D2 - D1) / (T2 - T1)
        Dim D3 As Decimal = A * T3 + B
        Return D3
    End Function
    Public Function D341(KV1 As Decimal, KV2 As Decimal, T1 As Decimal, T2 As Decimal, T3 As Decimal)
        'Calculating KV at given temperature as per ASTM D341
        Dim Z100 As Double = KV2 + 0.7 + Math.Exp(-1.47 - 1.84 * KV2 - 0.51 * KV2 * KV2)
        Dim Z40 As Double = KV1 + 0.7 + Math.Exp(-1.47 - 1.84 * KV1 - 0.51 * KV1 * KV1)
        Dim A As Double = (Math.Log10(273.15 + T1) * Math.Log10(Math.Log10(Z100)) - Math.Log10(273.15 + T2) * Math.Log10(Math.Log10(Z40))) / (Math.Log10(273.15 + T1) - Math.Log10(273.15 + T2))
        Dim B As Double = (A - Math.Log10(Math.Log10(Z40))) / Math.Log10(273.15 + T1)
        Dim Znew As Double = Math.Pow(10, Math.Pow(10, (A - B * Math.Log10(273.15 + T3))))
        Dim KVnew As Double = (Znew - 0.7) - Math.Exp(-0.7487 - 3.295 * (Znew - 0.7) + 0.6119 * Math.Pow((Znew - 0.7), 2) - 0.3193 * Math.Pow((Znew - 0.7), 3))
        Return KVnew
    End Function
    Public Function CaCalc(VGC As Decimal, Rint As Decimal)
        Dim A As Decimal = -7056447783.039 * Math.Pow(Rint, 4) + 29998425845.085 * Math.Pow(Rint, 3) - 47816290087.455 * Math.Pow(Rint, 2) + 33869264564.334 * Rint - 8995039235.23
        Dim B As Decimal = 12257713050.75 * Math.Pow(Rint, 4) - 52150924635.135 * Math.Pow(Rint, 3) + 83190281029.278 * Math.Pow(Rint, 2) - 58969692987.707 * Rint + 15672780130.06
        Dim C As Decimal = -5295114092.196 * Math.Pow(Rint, 4) + 22548683797.957 * Math.Pow(Rint, 3) - 36001147318.886 * Math.Pow(Rint, 2) + 25541641995.208 * Rint - 6794135037.906
        Dim Ca As Decimal = A * Math.Pow(VGC, 2) + B * VGC + C
        Return Ca
    End Function
    Public Function CnCalc(VGC As Decimal, Rint As Decimal)
        Dim A As Decimal = -5836850871.938 * Math.Pow(Rint, 4) + 25129932243.905 * Math.Pow(Rint, 3) - 40549060109.811 * Math.Pow(Rint, 2) + 29063266393.553 * Rint - 7807399922.174
        Dim B As Decimal = 8463984958.336 * Math.Pow(Rint, 4) - 36658365480.789 * Math.Pow(Rint, 3) + 59485482272.528 * Math.Pow(Rint, 2) - 42864364267.765 * Rint + 11573452051.452
        Dim C As Decimal = -2743401865.352 * Math.Pow(Rint, 4) + 12020312896.337 * Math.Pow(Rint, 3) - 19716727328.626 * Math.Pow(Rint, 2) + 14351222809.113 * Rint - 3911484775.02
        Dim Cn As Decimal = A * Math.Pow(VGC, 2) + B * VGC + C
        Return Cn
    End Function
    Public Sub Form1_KeyUp(sender As Object, e As EventArgs) Handles MyBase.KeyUp

        Try
            Dim D40 As Decimal?
            Dim D50 As Decimal?
            Dim D100 As Decimal?
            Dim KV40 As Decimal?
            Dim KV100 As Decimal?
            Dim KVX As Decimal?
            Dim KVXC As Decimal?
            Dim RI20 As Decimal?
            Dim DC As Decimal?
            Dim D15 As Decimal?
            Dim D20 As Decimal?
            Dim D90 As Decimal?
            Dim DatC As Decimal?
            Dim VGC As Decimal?
            Dim Rint As Decimal?
            Dim KVC As Decimal?
            TextBoxKVX.Enabled = True
            TextBoxKVXC.Enabled = True
            TextBoxD50.Enabled = True
            CheckBoxLR.Enabled = False
            If IsNumeric(TextBoxD40.Text) Then
                TextBoxD50.Enabled = False
                D40 = CDec(TextBoxD40.Text)
            End If
            If IsNumeric(TextBoxD50.Text) Then
                D50 = CDec(TextBoxD50.Text)
            End If
            If IsNumeric(TextBoxD100.Text) Then
                D100 = CDec(TextBoxD100.Text)
            End If
            If IsNumeric(TextBoxKV40.Text) Then
                TextBoxKVX.Enabled = False
                TextBoxKVXC.Enabled = False
                KV40 = CDec(TextBoxKV40.Text)
            ElseIf IsNumeric(TextBoxKV100.Text) And IsNumeric(TextBoxKVX.Text) And IsNumeric(TextBoxKVXC.Text) Then
                KV40 = CDec(D341(CDec(TextBoxKVX.Text), CDec(TextBoxKV100.Text), CDec(TextBoxKVXC.Text), 100, 40))
            End If
            If IsNumeric(TextBoxKVX.Text) Then
                KVX = CDec(TextBoxKVX.Text)
            End If
            If IsNumeric(TextBoxKVXC.Text) Then
                KVXC = CDec(TextBoxKVXC.Text)
            End If
            If IsNumeric(TextBoxKV100.Text) Then
                KV100 = CDec(TextBoxKV100.Text)
            End If
            If IsNumeric(TextBoxDC.Text) Then
                DC = CDec(TextBoxDC.Text)
            End If
            If IsNumeric(TextBoxRI20.Text) Then
                RI20 = CDec(TextBoxRI20.Text)
            End If
            If IsNumeric(TextBoxKVC.Text) Then
                KVC = CDec(TextBoxKVC.Text)
            End If
            If D40.HasValue Then
                D15 = 0.982262 * D40 + 32.819048
                D20 = 0.9864 * D40 + 25.74
            ElseIf D50.HasValue Then
                D15 = 0.9757 * D50 + 45.343
                D20 = 0.9786 * D50 + 39.41
            Else TextBoxD15.Text = ""
                TextBoxD20.Text = ""
            End If
            If D15.HasValue And KV40.HasValue Then
                VGC = (D15 / 1000 - 0.0664 - 0.1154 * Math.Log10(KV40 - 5.5)) / (0.94 - 0.109 * Math.Log10(KV40 - 5.5))
            End If
            If RI20.HasValue And D20.HasValue Then
                Rint = RI20 - D20 / 2000
            End If
            If D100.HasValue Then
                TextBoxD90.BackColor = Color.White
                Label10.Text = "All good"
                D90 = 0.995 * D100 + 10.65
            ElseIf D50.HasValue Then
                D90 = 1.021205 * D50 - 44.168493
                TextBoxD90.BackColor = Color.FromArgb(255, 180, 180)
                Label10.Text = "Not precise"
            Else TextBoxD90.Text = ""
                Label10.Text = "All good"
                TextBoxD90.BackColor = Color.White
            End If
            If DC.HasValue And Not CheckBoxLR.Enabled Then
                If DC > 50 Then
                    If D100.HasValue Then
                        DatC = 0.0006131 * D100 * DC - 1.18863 * DC + 0.9415 * D100 + 115.77
                    ElseIf D50.HasValue Then
                        DatC = (-0.0000016261 * DC * DC + 0.0007692386 * DC + 0.965822) * D50 + 0.0020669643 * DC * DC - 1.4046249943 * DC + 64.8564282
                    ElseIf D40.HasValue Then
                        DatC = (-0.0000016262 * DC * DC + 0.000771045 * DC + 0.9717856) * D40 + 0.0020771687 * DC * DC - 1.404813825 * DC + 52.8557156
                    Else TextBoxDatC.Text = ""
                    End If
                ElseIf DC > 40 Then
                    If D50.HasValue Then
                        DatC = (-0.0000016261 * DC * DC + 0.0007692386 * DC + 0.965822) * D50 + 0.0020669643 * DC * DC - 1.4046249943 * DC + 64.8564282
                    ElseIf D40.HasValue Then
                        DatC = (-0.0000016262 * DC * DC + 0.000771045 * DC + 0.9717856) * D40 + 0.0020771687 * DC * DC - 1.404813825 * DC + 52.8557156
                    ElseIf D100.HasValue Then
                        DatC = 0.0006131 * D100 * DC - 1.18863 * DC + 0.9415 * D100 + 115.77
                    Else TextBoxDatC.Text = ""
                    End If
                ElseIf DC <= 40 Then
                    If D40.HasValue Then
                        DatC = (-0.0000016262 * DC * DC + 0.000771045 * DC + 0.9717856) * D40 + 0.0020771687 * DC * DC - 1.404813825 * DC + 52.8557156
                    ElseIf D50.HasValue Then
                        DatC = (-0.0000016261 * DC * DC + 0.0007692386 * DC + 0.965822) * D50 + 0.0020669643 * DC * DC - 1.4046249943 * DC + 64.8564282
                    ElseIf D100.HasValue Then
                        DatC = 0.0006131 * D100 * DC - 1.18863 * DC + 0.9415 * D100 + 115.77
                    Else TextBoxDatC.Text = ""
                    End If
                End If
            End If
            If D15.HasValue Then
                    TextBoxD15.Text = CStr(Math.Round(CDec(D15), 1))

                End If
            If D20.HasValue Then
                TextBoxD20.Text = CStr(Math.Round(CDec(D20), 1))
            End If
            If D90.HasValue Then
                TextBoxD90.Text = CStr(Math.Round(CDec(D90), 1))
            End If
            If DatC.HasValue Then
                TextBoxDatC.Text = CStr(Math.Round(CDec(DatC), 1))
            End If
            If VGC.HasValue Then
                TextBoxVGC.Text = CStr(Math.Round(CDec(VGC), 4))
            End If
            If KV40.HasValue And KV100.HasValue And KVC.HasValue Then
                TextBoxKVatC.Text = CStr(Math.Round(D341(CDec(KV40), CDec(KV100), 40, 100, CDec(KVC)), 3))
            ElseIf KVX.HasValue And KV100.HasValue And KVC.HasValue And KVXC.HasValue Then
                TextBoxKVatC.Text = CStr(Math.Round(D341(KVX, KV100, KVXC, 100, KVC), 3))
            Else TextBoxKVatC.Text = ""
            End If
            If VGC.HasValue And Rint.HasValue Then
                TextBoxCa.Text = CStr(Math.Round(CaCalc(CDec(VGC), CDec(Rint)), 2))
                TextBoxCn.Text = CStr(Math.Round(CnCalc(CDec(VGC), CDec(Rint)), 2))
                TextBoxCp.Text = Math.Round(100 - CStr(CaCalc(CDec(VGC), CDec(Rint))) - CStr(CnCalc(CDec(VGC), CDec(Rint))), 2)
            End If
            If Rint.HasValue Then
                TextBoxRint.Text = CStr(Math.Round(CDec(Rint), 4))
            End If
            If D40.HasValue And D100.HasValue Or D50.HasValue And D100.HasValue Then
                CheckBoxLR.Enabled = True
            End If
            If CheckBoxLR.Checked And DC.HasValue Then
                If D40.HasValue And D100.HasValue Then
                    DatC = Math.Round(LRDensity(D40, D100, 273.15 + 40, 273.15 + 100, DC + 273.15), 1)
                ElseIf D50.HasValue And D100.HasValue Then
                    DatC = Math.Round(LRDensity(D50, D100, 273.15 + 50, 273.15 + 100, DC + 273.15), 1)
                ElseIf D40.HasValue And D50.HasValue Then
                    DatC = Math.Round(LRDensity(D40, D50, 273.15 + 40, 273.15 + 50, DC + 273.15), 1)
                Else TextBoxDatC.Text = ""
                End If
            End If
            If DatC.HasValue Then
                TextBoxDatC.Text = CStr(Math.Round(CDec(DatC), 1))
            End If
            Sts2.Text = "OK"
        Catch ex As Exception

            Sts2.Text = ex.Message

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim a As Control
        For Each a In Me.Controls
            If TypeOf a Is TextBox Then
                a.Text = Nothing
            End If
        Next
    End Sub

    Private Sub CheckBoxLR_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLR.CheckedChanged
        Form1_KeyUp(sender, e)
    End Sub
End Class
