Public Class Calculator
    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.
    Public FactorialThread As System.Threading.Thread
    Public FactorialMinusOneThread As System.Threading.Thread
    Public AddTwoThread As System.Threading.Thread
    Public LoopThread As System.Threading.Thread


    Public varAddTwo As Integer
    Public varFact1 As Integer
    Public varFact2 As Integer
    Public varLoopValue As Integer
    Public varTotalCalculations As Double = 0

    Public Event FactorialComplete(ByVal Factorial As Double, ByVal _
       TotalCalculations As Double)
    Public Event FactorialMinusComplete(ByVal Factorial As Double, ByVal _
       TotalCalculations As Double)
    Public Event AddTwoComplete(ByVal Result As Integer, ByVal _
       TotalCalculations As Double)
    Public Event LoopComplete(ByVal TotalCalculations As Double, ByVal _
       Counter As Integer)
   
#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    Private Sub error_handler(ByVal message As String)
        Try
            MsgBox("Multithread Factorial Calculator has trapped the following error: " & vbCrLf & message, MsgBoxStyle.Exclamation, "Error Encountered")
        Catch ex As Exception
            MsgBox("Multithread Factorial Calculator has trapped the following error: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error Encountered")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    FactorialThread = New System.Threading.Thread(AddressOf _
                       Factorial)
                    ' Starts the thread.
                    FactorialThread.Start()
                Case 2
                    FactorialMinusOneThread = New _
                       System.Threading.Thread(AddressOf FactorialMinusOne)
                    FactorialMinusOneThread.Start()
                Case 3
                    AddTwoThread = New System.Threading.Thread(AddressOf AddTwo)
                    AddTwoThread.Start()
                Case 4
                    LoopThread = New System.Threading.Thread(AddressOf RunALoop)
                    LoopThread.Start()
            End Select
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub
    ' This sub will calculate the value of a number minus 1 factorial 
    ' (varFact2-1!).
    Public Sub FactorialMinusOne()
        Try
            Dim varX As Integer = 1
            Dim varTotalAsOfNow As Double
            Dim varResult As Double = 1
            ' Performs a factorial calculation on varFact2 - 1.
            For varX = 1 To varFact2 - 1
                varResult *= varX
                ' Increments varTotalCalculations and keeps track of the current
                ' total as of this instant.
                SyncLock Me
                    varTotalCalculations += 1
                    varTotalAsOfNow = varTotalCalculations
                End SyncLock

            Next varX
            ' Signals that the method has completed, and communicates the 
            ' result and a value of total calculations performed up to this 
            ' point
            RaiseEvent FactorialMinusComplete(varResult, varTotalAsOfNow)
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    ' This sub will calculate the value of a number factorial (varFact1!).
    Public Sub Factorial()
        Try
            Dim varX As Integer = 1
            Dim varResult As Double = 1
            Dim varTotalAsOfNow As Double = 0
            For varX = 1 To varFact1
                varResult *= varX
                SyncLock Me
                    varTotalCalculations += 1
                    varTotalAsOfNow = varTotalCalculations
                End SyncLock

            Next varX
            RaiseEvent FactorialComplete(varResult, varTotalAsOfNow)
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    ' This sub will add two to a number (varAddTwo + 2).
    Public Sub AddTwo()
        Try
            Dim varResult As Integer
            Dim varTotalAsOfNow As Double
            varResult = varAddTwo + 2
            SyncLock Me
                varTotalCalculations += 1
                varTotalAsOfNow = varTotalCalculations
            End SyncLock

            RaiseEvent AddTwoComplete(varResult, varTotalAsOfNow)
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    ' This method will run a loop with a nested loop varLoopValue times.
    Public Sub RunALoop()
        Try
            Dim varX As Integer
            Dim varY As Integer
            Dim varTotalAsOfNow As Double
            For varX = 1 To varLoopValue
                ' This nested loop is added solely for the purpose of slowing
                ' down the program and creating a processor-intensive
                ' application.
                For varY = 1 To 500
                    SyncLock Me
                        varTotalCalculations += 1
                        varTotalAsOfNow = varTotalCalculations
                    End SyncLock

                Next
            Next
            RaiseEvent LoopComplete(varTotalAsOfNow, varX - 1)
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

End Class
