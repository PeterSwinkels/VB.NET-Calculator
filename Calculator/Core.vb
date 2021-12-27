'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Text
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module CoreModule
   Private Const DEFAULT_MASK As String = "#########0"   'Contains the default mask used to format a number.
   Private Const MAXIMUM_LENGTH As Integer = 10          'Contains the maximum length for a number.

   'This enumeration lists the operations supported by this program.
   Private Enum OperationsE As Integer
      None            'None.
      Addition        'Addition.
      Division        'Division.
      Multiplication  'Multiplication.
      Subtraction     'Subtraction.
   End Enum

   'This structure defines the current number.
   Private Structure CurrentNumberStr
      Public HasError As Boolean          'Indicates whether an error occurred.
      Public Mask As Stringbuilder        'Defines the current number's display mask.
      Public Value As Double?             'Defines the current number.
   End Structure

   'This procedure formats the specified number using the specified mask.
   Private Function FormatNumber(Number As Double, Mask As String) As String
      Try
         Return $"{Number.ToString(Mask)}{If(Mask.EndsWith("."), ".", "")}"
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the current number.
   Public Function GetCurrentNumber(sender As Object) As String
      Try
         With PerformOperation(DirectCast(sender, Control).Text)
            Return If(.HasError, "E", FormatNumber(.Value.Value, .Mask.ToString()))
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure handles any errors that occur.
   Public Sub HandleError(ExceptionO As Exception)
      Try
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      Catch
         Application.Exit()
      End Try
   End Sub

   'This procedure performs the specified calculation on the specified numbers.
   Private Function PerformCalculation(LeftNumber As Double?, RightNumber As Double?, Operation As OperationsE) As Double?
      Try
         Select Case Operation
            Case OperationsE.Addition
               Return LeftNumber + RightNumber
            Case OperationsE.Division
               Return LeftNumber / RightNumber
            Case OperationsE.Multiplication
               Return LeftNumber * RightNumber
            Case OperationsE.Subtraction
               Return LeftNumber - RightNumber
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure performsthe specified operation.
   Private Function PerformOperation(ButtonText As String) As CurrentNumberStr
      Try
         Static CurrentNumber As New CurrentNumberStr With {.HasError = False, .Mask = New StringBuilder(DEFAULT_MASK), .Value = New Double}
         Static LeftNumber As Double? = Nothing
         Static Operation As OperationsE = OperationsE.None
         Static RightNumber As Double? = Nothing

         With CurrentNumber
            If .HasError Then ButtonText = "C"

            Select Case ButtonText.Trim().ToUpper()
               Case "C"
                  .HasError = False
                  .Mask = New StringBuilder(DEFAULT_MASK)
                  .Value = New Double
                  LeftNumber = Nothing
                  Operation = OperationsE.None
                  RightNumber = Nothing
               Case "+"
                  Operation = OperationsE.Addition
               Case "-"
                  Operation = OperationsE.Subtraction
               Case "X"
                  Operation = OperationsE.Multiplication
               Case "/"
                  Operation = OperationsE.Division
               Case "="
                  If (Not Operation = OperationsE.None) Then
                     If LeftNumber Is Nothing Then
                        LeftNumber = .Value
                        If RightNumber Is Nothing Then RightNumber = LeftNumber
                     Else
                        RightNumber = .Value
                     End If
                     .Value = PerformCalculation(LeftNumber, RightNumber, Operation)
                     .Mask = New StringBuilder(DEFAULT_MASK)
                     If Not Integer.TryParse(CStr(.Value), Nothing) Then
                        .Mask.Append(".")
                        For NewDigit As Integer = 1 To MAXIMUM_LENGTH
                           If CDbl(.Value).ToString(.Mask.ToString() & "0").EndsWith("0") Then Exit For
                           .Mask.Append("0")
                        Next NewDigit
                     End If

                     LeftNumber = Nothing
                  End If
               Case Else
                  If "0123456789.".Contains(ButtonText) Then
                     If Not Operation = OperationsE.None Then
                        If LeftNumber Is Nothing Then
                           LeftNumber = .Value
                           RightNumber = New Double
                           .Mask = New StringBuilder(DEFAULT_MASK)
                           .Value = RightNumber
                        End If
                     End If

                     If CDbl(.Value).ToString(.Mask.ToString()).Length < MAXIMUM_LENGTH Then
                        If ButtonText = "." Then
                           .Mask.Append(".")
                        Else
                           .Value = CDbl($"{FormatNumber(.Value.Value, .Mask.ToString())}{ButtonText}")
                           If .Mask.ToString().Contains(".") Then .Mask.Append("0")
                        End If
                     End If
                  End If
            End Select

            .HasError = (Double.IsInfinity(CDbl(.Value)) OrElse Double.IsNaN(CDbl(.Value)))
         End With

         Return CurrentNumber
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns information about this program.
   Public Function ProgramInformation() As String
      Try
         With My.Application.Info
            Return $"{ .Title} v{ .Version} - by: { .CompanyName}"
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
