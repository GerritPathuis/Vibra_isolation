'Imports System.Math
'Imports System.IO
'Imports System.Text
Imports System.Globalization
Imports System.Threading
'Imports Word = Microsoft.Office.Interop.Word
'Imports System.Windows.Forms.DataVisualization.Charting
'Imports System.Windows.Forms


Public Class Form1
    Public _weight As Double
    Public _speed_Hz As Double

    'Calculation Based on
    'http://www.astrotex.com/pdf/engineering-guide.pdf

    'Isolator data
    'http://www.gmt-benelux.nl/uploads/datasheet_240.pdf
    '"Model; dia; height; shore; N/mm; N; smax",
    Public iso_dia() As String = {
        "50x45 C 40°Sh; 50; 45; 40°Sh; 129; 753;5.9",
        "50x45 C 55°Sh; 50; 45; 55°Sh; 239; 1396;5.9",
        "50x45 C 70°Sh; 50; 45; 70°Sh; 387; 2265;5.9",
        "75x55 C 40°Sh; 75; 55; 40°Sh; 248; 1821;7.4",
        "75x55 C 55°Sh; 75; 55; 55°sh; 459; 3374;7.4",
        "75x55 C 70°Sh; 75; 55; 70°Sh; 745; 5476;7.4",
        "80x55 C 40°Sh; 80; 66; 40°Sh; 207; 1864;9.0",
        "80x55 C 55°Sh; 80; 66; 55°sh; 384; 3453;9.0",
        "80x55 C 70°Sh; 80; 66; 70°Sh; 623; 5605;9.0",
        "90x55 C 40°Sh; 90; 55; 40°Sh; 371; 2728;7.4",
        "90x55 C 55°Sh; 90; 55; 55°Sh; 688; 5053;7.4",
        "90x55 C 70°Sh; 90; 55; 70°Sh; 1116; 8202;7.4",
        "100x75 C 40°Sh; 100; 75; 40°Sh; 303; 3050;10.1",
        "100x75 C 55°Sh; 100; 75; 55°Sh; 562; 5650;10.1",
        "100x75 C 70°Sh; 100; 75; 70°Sh; 912; 9170;10.1",
        "100x100 C 40°Sh; 100; 100; 40°Sh; 197; 2714;13.8",
        "100x100 C 55°Sh; 100; 100; 55°Sh; 364; 5026;13.8",
        "100x100 C 70°Sh; 100; 100; 70°Sh; 591; 8159;13.8",
        "125x75 C 40°Sh; 125; 60; 40°Sh; 1016; 7622;7.5",
        "125x75 C 55°Sh; 125; 60; 55°Sh; 1882; 14118;7.5",
        "125x75 C 70°Sh; 125; 60; 70°Sh; 3055; 22915;7.5",
        "150x75 C 40°Sh; 150; 75; 40°Sh; 889; 8666;9.8",
        "150x75 C 55°Sh; 150; 75; 55°Sh; 1646; 16051;9.8",
        "150x75 C 70°Sh; 150; 75; 70°Sh; 2672; 26053;9.8",
        "150x100 C 40°Sh; 150; 100; 40°Sh; 535; 7216;9.8",
        "150x100 C 55°Sh; 150; 100; 55°Sh; 990; 13366;9.8",
        "150x100 C 70°Sh; 150; 100; 70°Sh; 1607; 21695;9.8"}

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged
        Calc_iso()
        Update_iso()
        Update_iso()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        ComboBox1.Items.Clear()
        '-------Fill combobox------------------
        For hh = 0 To (UBound(iso_dia) - 1)              'Fill combobox1
            words = iso_dia(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 13

    End Sub
    Private Sub Calc_iso()
        Dim speed As Double
        Dim iso_load As Double
        Dim Isolation As Double = 0.8   'Isolation
        Dim Tr As Double                 'transmissibility
        Dim fn_iso_system As Double
        Dim ds As Double                'required static deflection
        Dim k_spring As Double          'required spring constant
        Dim k_spring_each As Double          'required spring constant
        Dim no_isolators As Double


        no_isolators = NumericUpDown3.Value
        Isolation = NumericUpDown4.Value
        _weight = NumericUpDown1.Value
        speed = NumericUpDown2.Value
        _speed_Hz = speed / 60  '[Hz]
        iso_load = _weight / no_isolators
        Tr = 1 - Isolation              'transmissibility
        fn_iso_system = _speed_Hz / Math.Sqrt((1 / Tr) + 1)
        ds = (25.4 * 9.81) / fn_iso_system ^ 2      '[mm]
        k_spring = 9.81 * iso_load / ds                   '[]
        k_spring_each = k_spring / no_isolators

        TextBox1.Text = _speed_Hz.ToString("0")
        TextBox2.Text = iso_load.ToString("0")
        TextBox3.Text = Tr.ToString("0.0")
        TextBox4.Text = fn_iso_system.ToString("0.0")
        TextBox5.Text = ds.ToString("0.00")          'required static deflection
        TextBox6.Text = k_spring.ToString("0")
        TextBox18.Text = k_spring_each.ToString("0")
        TextBox9.Text = _weight * 9.81.ToString("0")   '[N]
    End Sub
    Private Sub Update_iso()
        Dim c1, c_total, no_spring As Double
        Dim f1, f_total As Double
        Dim words() As String
        Dim req_deflec, instal_deflec As Double
        Dim req_srate, instal_srate As Double

        If (ComboBox1.SelectedIndex > -1) Then
            Try
                words = iso_dia(ComboBox1.SelectedIndex).Split(CType(";", Char()))
                TextBox10.Text = words(1) 'Diameter [mm]
                TextBox11.Text = words(2) 'Height [mm]
                TextBox12.Text = words(3) 'Shore [degree]
                TextBox13.Text = words(4) 'Spring rate [N/mm]
                TextBox14.Text = words(5) 'Force max [N]
                TextBox15.Text = words(6) 'Deflection smax [mm]

                Double.TryParse(TextBox13.Text, c1)
                Double.TryParse(TextBox14.Text, f1)
                Double.TryParse(TextBox15.Text, instal_deflec)
                Double.TryParse(TextBox16.Text, instal_srate)

                Double.TryParse(TextBox5.Text, req_deflec)
                Double.TryParse(TextBox6.Text, req_srate)

                no_spring = NumericUpDown3.Value    'Spring rate one spring
                c_total = no_spring * c1            'Spring rate total
                f_total = no_spring * f1            'Allowed force total

                TextBox16.Text = c_total.ToString
                TextBox17.Text = f_total.ToString

                '-----------Checks----------
                TextBox17.BackColor = CType(IIf(NumericUpDown1.Value * 10 > f_total, Color.Red, Color.LightGreen), Color)
                TextBox15.BackColor = CType(IIf(req_deflec > instal_deflec, Color.Red, Color.LightGreen), Color)

                TextBox16.BackColor = CType(IIf(instal_srate * 0.8 < req_srate, Color.LightGreen, Color.Red), Color)
            Catch ex As Exception
                MessageBox.Show(ex.Message)  ' Show the exception's message.
            End Try
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Calc_iso()
        Update_iso()
        Update_iso()
    End Sub
End Class
