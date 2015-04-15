Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim WithEvents Calculator1 As Calculator


    Public Delegate Sub FHandler(ByVal Value As Double, ByVal _
     Calculations As Double)
    Public Delegate Sub A2Handler(ByVal Value As Integer, ByVal _
       Calculations As Double)
    Public Delegate Sub LDhandler(ByVal Calculations As Double, ByVal _
       Count As Integer)

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        ' Creates a new instance of Calculator.
        Calculator1 = New Calculator()
        AddHandler Calculator1.FactorialComplete, AddressOf FactHandler
        AddHandler Calculator1.AddTwoComplete, AddressOf Add2Handler
        AddHandler Calculator1.FactorialMinusComplete, AddressOf _
           Fact1Handler
        AddHandler Calculator1.LoopComplete, AddressOf LDoneHandler

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblFactorial1 As System.Windows.Forms.Label
    Friend WithEvents lblFactorial2 As System.Windows.Forms.Label
    Friend WithEvents lblAddTwo As System.Windows.Forms.Label
    Friend WithEvents lblRunLoops As System.Windows.Forms.Label
    Friend WithEvents lblTotalCalculations As System.Windows.Forms.Label
    Friend WithEvents btnFactorial1 As System.Windows.Forms.Button
    Friend WithEvents btnFactorial2 As System.Windows.Forms.Button
    Friend WithEvents btnRunLoops As System.Windows.Forms.Button
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents btnAddTwo As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label4 As System.Windows.Forms.Label


    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.lblFactorial1 = New System.Windows.Forms.Label
        Me.lblFactorial2 = New System.Windows.Forms.Label
        Me.lblAddTwo = New System.Windows.Forms.Label
        Me.lblRunLoops = New System.Windows.Forms.Label
        Me.lblTotalCalculations = New System.Windows.Forms.Label
        Me.btnFactorial1 = New System.Windows.Forms.Button
        Me.btnFactorial2 = New System.Windows.Forms.Button
        Me.btnAddTwo = New System.Windows.Forms.Button
        Me.btnRunLoops = New System.Windows.Forms.Button
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblFactorial1
        '
        Me.lblFactorial1.Location = New System.Drawing.Point(104, 216)
        Me.lblFactorial1.Name = "lblFactorial1"
        Me.lblFactorial1.Size = New System.Drawing.Size(144, 23)
        Me.lblFactorial1.TabIndex = 0
        '
        'lblFactorial2
        '
        Me.lblFactorial2.Location = New System.Drawing.Point(104, 184)
        Me.lblFactorial2.Name = "lblFactorial2"
        Me.lblFactorial2.Size = New System.Drawing.Size(144, 23)
        Me.lblFactorial2.TabIndex = 1
        '
        'lblAddTwo
        '
        Me.lblAddTwo.Location = New System.Drawing.Point(104, 120)
        Me.lblAddTwo.Name = "lblAddTwo"
        Me.lblAddTwo.Size = New System.Drawing.Size(144, 23)
        Me.lblAddTwo.TabIndex = 2
        '
        'lblRunLoops
        '
        Me.lblRunLoops.Location = New System.Drawing.Point(104, 152)
        Me.lblRunLoops.Name = "lblRunLoops"
        Me.lblRunLoops.Size = New System.Drawing.Size(144, 23)
        Me.lblRunLoops.TabIndex = 3
        '
        'lblTotalCalculations
        '
        Me.lblTotalCalculations.Location = New System.Drawing.Point(24, 72)
        Me.lblTotalCalculations.Name = "lblTotalCalculations"
        Me.lblTotalCalculations.Size = New System.Drawing.Size(216, 23)
        Me.lblTotalCalculations.TabIndex = 4
        '
        'btnFactorial1
        '
        Me.btnFactorial1.BackColor = System.Drawing.Color.DarkGreen
        Me.btnFactorial1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFactorial1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFactorial1.ForeColor = System.Drawing.Color.White
        Me.btnFactorial1.Location = New System.Drawing.Point(8, 208)
        Me.btnFactorial1.Name = "btnFactorial1"
        Me.btnFactorial1.Size = New System.Drawing.Size(80, 23)
        Me.btnFactorial1.TabIndex = 5
        Me.btnFactorial1.Text = "Factorial"
        Me.ToolTip1.SetToolTip(Me.btnFactorial1, "Run factorial calculation")
        '
        'btnFactorial2
        '
        Me.btnFactorial2.BackColor = System.Drawing.Color.SeaGreen
        Me.btnFactorial2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFactorial2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFactorial2.ForeColor = System.Drawing.Color.White
        Me.btnFactorial2.Location = New System.Drawing.Point(8, 176)
        Me.btnFactorial2.Name = "btnFactorial2"
        Me.btnFactorial2.Size = New System.Drawing.Size(80, 23)
        Me.btnFactorial2.TabIndex = 6
        Me.btnFactorial2.Text = "Factorial - 1"
        Me.ToolTip1.SetToolTip(Me.btnFactorial2, "Run factorial minus one calculation")
        '
        'btnAddTwo
        '
        Me.btnAddTwo.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.btnAddTwo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddTwo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddTwo.ForeColor = System.Drawing.Color.White
        Me.btnAddTwo.Location = New System.Drawing.Point(8, 112)
        Me.btnAddTwo.Name = "btnAddTwo"
        Me.btnAddTwo.Size = New System.Drawing.Size(80, 23)
        Me.btnAddTwo.TabIndex = 7
        Me.btnAddTwo.Text = "Add Two"
        Me.ToolTip1.SetToolTip(Me.btnAddTwo, "Add two to the input value")
        '
        'btnRunLoops
        '
        Me.btnRunLoops.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.btnRunLoops.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRunLoops.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRunLoops.ForeColor = System.Drawing.Color.White
        Me.btnRunLoops.Location = New System.Drawing.Point(8, 144)
        Me.btnRunLoops.Name = "btnRunLoops"
        Me.btnRunLoops.Size = New System.Drawing.Size(80, 23)
        Me.btnRunLoops.TabIndex = 8
        Me.btnRunLoops.Text = "Run a Loop"
        Me.ToolTip1.SetToolTip(Me.btnRunLoops, "Execute a looping command. Used to show multi-threaded nature of program")
        '
        'txtValue
        '
        Me.txtValue.BackColor = System.Drawing.Color.White
        Me.txtValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtValue.Location = New System.Drawing.Point(24, 24)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(216, 20)
        Me.txtValue.TabIndex = 9
        Me.txtValue.Text = ""
        Me.txtValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(256, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(130, 144)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Multithread Factorial Calculator is a simple factorial calculator that serves as " & _
        "an example for thread programming in action. The calculator allows you to execut" & _
        "e multiple operations simultaneously. "
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(288, 200)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "About   |"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(336, 200)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Exit"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Firebrick
        Me.Label4.Location = New System.Drawing.Point(24, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(216, 16)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Input Value   (Positive Integer > 0)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(392, 273)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.btnRunLoops)
        Me.Controls.Add(Me.btnAddTwo)
        Me.Controls.Add(Me.btnFactorial2)
        Me.Controls.Add(Me.btnFactorial1)
        Me.Controls.Add(Me.lblTotalCalculations)
        Me.Controls.Add(Me.lblRunLoops)
        Me.Controls.Add(Me.lblAddTwo)
        Me.Controls.Add(Me.lblFactorial2)
        Me.Controls.Add(Me.lblFactorial1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(400, 307)
        Me.MinimumSize = New System.Drawing.Size(400, 307)
        Me.Name = "Form1"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Multithread Factorial Calculator"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected Sub btnFactorial1_Click(ByVal sender As Object, ByVal e As _
       System.EventArgs) Handles btnFactorial1.Click
        Try
            If IsNumeric(txtValue.Text) Then
                ' Passes the value typed in the txtValue to Calculator.varFact1.
                Calculator1.varFact1 = CInt(txtValue.Text)
                ' Disables the btnFactorial1 until this calculation is complete.
                btnFactorial1.Enabled = False
                '  Calculator1.Factorial()
                ' Passes the value 1 to Calculator1, thus directing it to start the ' correct thread.
                Calculator1.ChooseThreads(1)
            End If
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Protected Sub btnFactorial2_Click(ByVal sender As Object, ByVal e _
       As System.EventArgs) Handles btnFactorial2.Click
        Try
            If IsNumeric(txtValue.Text) Then
                Calculator1.varFact2 = CInt(txtValue.Text)
                btnFactorial2.Enabled = False
                'Calculator1.FactorialMinusOne()
                ' Calculator1.FactorialMinusOne()
                Calculator1.ChooseThreads(2)
            End If
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Protected Sub btnAddTwo_Click(ByVal sender As Object, ByVal e As _
       System.EventArgs) Handles btnAddTwo.Click
        Try
            If IsNumeric(txtValue.Text) Then
                Calculator1.varAddTwo = CInt(txtValue.Text)
                btnAddTwo.Enabled = False
                '   Calculator1.AddTwo()
                ' Calculator1.AddTwo()
                Calculator1.ChooseThreads(3)
            End If
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Protected Sub btnRunLoops_Click(ByVal sender As Object, ByVal e As _
       System.EventArgs) Handles btnRunLoops.Click
        Try
            If IsNumeric(txtValue.Text) Then
                Calculator1.varLoopValue = CInt(txtValue.Text)
                btnRunLoops.Enabled = False
                ' Lets the user know that a loop is running.
                lblRunLoops.Text = "Looping"
                '  Calculator1.RunALoop()
                Calculator1.ChooseThreads(4)
            End If
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub
    Protected Sub FactorialHandler(ByVal Value As Double, ByVal _
       Calculations As Double)

    End Sub

    Protected Sub Factorial1Handler(ByVal Value As Double, ByVal _
       Calculations As Double)

    End Sub

    Protected Sub AddTwoHandler(ByVal Value As Integer, ByVal _
       Calculations As Double)

    End Sub

    Protected Sub LoopDoneHandler(ByVal Calculations As Double, ByVal _
       Count As Integer)

    End Sub
    Public Sub FactHandler(ByVal Value As Double, ByVal Calculations As _
       Double)
        Try
            ' Displays the returned value in the appropriate label.
            lblFactorial1.Text = Value.ToString
            ' Re-enables the button so it can be used again.
            btnFactorial1.Enabled = True
            ' Updates the label that displays the total calculations performed
            lblTotalCalculations.Text = "Total Calculations are " & _
               Calculations.ToString
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub
    Public Sub Fact1Handler(ByVal Value As Double, ByVal Calculations As _
       Double)
        Try
            lblFactorial2.Text = Value.ToString
            btnFactorial2.Enabled = True
            lblTotalCalculations.Text = "Total Calculations are " & _
                Calculations.ToString
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub
    Public Sub Add2Handler(ByVal Value As Integer, ByVal Calculations As _
       Double)
        Try
            lblAddTwo.Text = Value.ToString
            btnAddTwo.Enabled = True
            lblTotalCalculations.Text = "Total Calculations are " & _
                Calculations.ToString
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub
    Public Sub LDoneHandler(ByVal Calculations As Double, ByVal Count As _
       Integer)
        Try
            btnRunLoops.Enabled = True
            lblRunLoops.Text = Count.ToString
            lblTotalCalculations.Text = "Total Calculations are " & _
               Calculations.ToString
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Private Sub txtValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtValue.TextChanged
        Try
            If (IsNumeric(txtValue.Text) = True And Val(txtValue.Text) > 0) Or txtValue.Text = "" Then
            Else
                MsgBox("Please enter a positive integer greater than zero as a value", MsgBoxStyle.OKOnly, "Error detected")
                txtValue.Text = ""
            End If
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            txtValue.Focus()
            txtValue.Select()
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Try
            Application.Exit()
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Try
            Dim result As DialogResult
            Dim about As New About_Splash_Screen
            result = about.ShowDialog()
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Private Sub error_handler(ByVal message As String)
        Try
            MsgBox("Multithread Factorial Calculator has trapped the following error: " & vbCrLf & message, MsgBoxStyle.Exclamation, "Error Encountered")
        Catch ex As Exception
            MsgBox("Multithread Factorial Calculator has trapped the following error: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error Encountered")
        End Try
    End Sub
End Class
