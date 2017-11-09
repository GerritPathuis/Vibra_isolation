Public Class Form1
    Public _weight As Double
    Public _speed_Hz As Double
    'Based on
    'http://www.astrotex.com/pdf/engineering-guide.pdf
    '
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged
        Dim speed As Double
        Dim iso_load As Double
        Dim Isolation As Double = 0.8     'Isolation
        Dim T As Double             'transmissibility
        Dim fn_iso_system As Double
        Dim ds As Double                'required static deflection
        Dim k_spring As Double          'required spring constant

        Isolation = NumericUpDown4.Value
        _weight = NumericUpDown1.Value
        speed = NumericUpDown2.Value
        _speed_Hz = speed / 60  '[Hz]
        iso_load = _weight / NumericUpDown3.Value
        T = 1 - Isolation              'transmissibility
        fn_iso_system = _speed_Hz / Math.Sqrt((1 / T) + 1)
        ds = (25.4 * 9.81) / fn_iso_system ^ 2      '[mm]
        k_spring = 9.81 * iso_load / ds                   '[]

        TextBox1.Text = _speed_Hz.ToString("0")
        TextBox2.Text = iso_load.ToString("0")
        TextBox3.Text = T.ToString("0.0")
        TextBox4.Text = fn_iso_system.ToString("0.0")
        TextBox5.Text = ds.ToString("0.00")          'required static deflection
        TextBox6.Text = k_spring.ToString("0")
    End Sub
End Class
