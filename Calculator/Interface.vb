'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Windows.Forms

'This module contains this program's main interface.
Public Class InterfaceWindow
   'This procedure initializes this window when it is created.
   Public Sub New()
      Try
         InitializeComponent()

         AddHandler DecimalPointButton.Click, AddressOf ButtonHandler
         AddHandler Number0Button.Click, AddressOf ButtonHandler
         AddHandler Number1Button.Click, AddressOf ButtonHandler
         AddHandler Number2Button.Click, AddressOf ButtonHandler
         AddHandler Number3Button.Click, AddressOf ButtonHandler
         AddHandler Number4Button.Click, AddressOf ButtonHandler
         AddHandler Number5Button.Click, AddressOf ButtonHandler
         AddHandler Number6Button.Click, AddressOf ButtonHandler
         AddHandler Number7Button.Click, AddressOf ButtonHandler
         AddHandler Number8Button.Click, AddressOf ButtonHandler
         AddHandler Number9Button.Click, AddressOf ButtonHandler
         AddHandler AdditionButton.Click, AddressOf ButtonHandler
         AddHandler ClearButton.Click, AddressOf ButtonHandler
         AddHandler DivisionButton.Click, AddressOf ButtonHandler
         AddHandler EqualsButton.Click, AddressOf ButtonHandler
         AddHandler MultiplicationButton.Click, AddressOf ButtonHandler
         AddHandler SubtractionButton.Click, AddressOf ButtonHandler

         My.Application.ChangeCulture("en-US")
         Me.Text = My.Application.Info.Title
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure extends the number being displayed with the specified digit or a decimal point.
   Private Sub ButtonHandler(sender As Object, e As EventArgs)
      Try
         DisplayBox.Text = GetCurrentNumber(sender)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure prevents the display from having the focus.
   Private Sub Display_GotFocus(sender As Object, e As EventArgs) Handles DisplayBox.GotFocus
      Try
         EqualsButton.Focus()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure handles the user's keystrokes.
   Private Sub Form_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
      Try
         Select Case e.KeyCode
            Case Keys.F1
               MessageBox.Show(My.Application.Info.Description, ProgramInformation(), MessageBoxButtons.OK, MessageBoxIcon.Information)
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure further initializes this window when it is loaded.
   Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
      Try
         DisplayBox.Text = GetCurrentNumber(Number0Button)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Class
