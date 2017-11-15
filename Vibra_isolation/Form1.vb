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
        "50x45C_40°Sh; 50; 45; 40Sh; 129; 753;5.9",
        "50x45C_55°Sh; 50; 45; 55Sh; 239; 1396;5.9",
        "50x45C_70°Sh; 50; 45; 70Sh; 387; 2265;5.9",
        "75x55C_40°Sh; 75; 55; 40Sh; 248; 1821;7.4",
        "75x55C_55°Sh; 75; 55; 55sh; 459; 3374;7.4",
        "75x55C_70°Sh; 75; 55; 70Sh; 745; 5476;7.4",
        "100x75C_40°Sh; 100; 75; 40Sh; 303; 3050;10.1",
        "100x75C_55°Sh; 100; 75; 55Sh; 562; 5650;10.1",
        "100x75C_70°Sh; 100; 75; 70Sh; 912; 9170;10.1",
        "150x75C_40°Sh; 150; 75; 40Sh; 889; 8666;9.8",
        "150x75C_55°Sh; 150; 75; 55Sh; 1646; 16051;9.8",
        "150x75C_70°Sh; 150; 75; 70Sh; 2672; 26053;9.8"}

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged
        Dim speed As Double
        Dim iso_load As Double
        Dim Isolation As Double = 0.8   'Isolation
        Dim Tr As Double                 'transmissibility
        Dim fn_iso_system As Double
        Dim ds As Double                'required static deflection
        Dim k_spring As Double          'required spring constant

        Isolation = NumericUpDown4.Value
        _weight = NumericUpDown1.Value
        speed = NumericUpDown2.Value
        _speed_Hz = speed / 60  '[Hz]
        iso_load = _weight / NumericUpDown3.Value
        Tr = 1 - Isolation              'transmissibility
        fn_iso_system = _speed_Hz / Math.Sqrt((1 / Tr) + 1)
        ds = (25.4 * 9.81) / fn_iso_system ^ 2      '[mm]
        k_spring = 9.81 * iso_load / ds                   '[]

        TextBox1.Text = _speed_Hz.ToString("0")
        TextBox2.Text = iso_load.ToString("0")
        TextBox3.Text = Tr.ToString("0.0")
        TextBox4.Text = fn_iso_system.ToString("0.0")
        TextBox5.Text = ds.ToString("0.00")          'required static deflection
        TextBox6.Text = k_spring.ToString("0")

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
        ComboBox1.SelectedIndex = 2
    End Sub

    Private Sub Update_iso()
        Dim c1, c_total, no_spring As Double
        Dim f1, f_total As Double
        Dim words() As String
        Dim req_deflec, instal_deflec As Double
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
            Double.TryParse(TextBox5.Text, req_deflec)

            no_spring = NumericUpDown3.Value    'Spring rate one spring
            c_total = no_spring * c1            'Spring rate total
            f_total = no_spring * f1            'Allowed force total

            TextBox16.Text = c_total.ToString
            TextBox17.Text = f_total.ToString

            '-----------Checks----------
            TextBox17.BackColor = CType(IIf(NumericUpDown1.Value * 10 > f_total, Color.Red, Color.LightGreen), Color)
            TextBox15.BackColor = CType(IIf(req_deflec > instal_deflec, Color.Red, Color.LightGreen), Color)
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Update_iso()
    End Sub
End Class
