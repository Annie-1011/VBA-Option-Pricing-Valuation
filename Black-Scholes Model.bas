Attribute VB_Name = "Module1"
Option Explicit

Function BSPRICE(Stock As Double, Exercise As Double, _
                    Interest As Double, Sigma As Double, _
                    Time As Double, Optional opttype As Variant) _
                    As Variant

Dim s As Double, x As Double
Dim r As Double, v As Double, t As Double

Dim d1 As Double, d2 As Double
Dim BSCall As Double, BSPut As Double

    s = Stock: x = Exercise
    r = Interest: v = Sigma
    t = Time
    
    If IsMissing(opttype) Then opttype = "Call"
    
    d1 = (Log(s / x) + r * t) / (v * Sqr(t)) + 0.5 * v * Sqr(t)
    d2 = d1 - v * Sqr(t)

With Application.WorksheetFunction
    BSCall = s * .NormSDist(d1) - x * Exp(-t * r) * .NormSDist(d2)
    BSPut = BSCall + x * Exp(-r * t) - s
                    
    If opttype = "Call" Then
        BSPRICE = BSCall
    ElseIf opttype = "Put" Then
        BSPRICE = BSPut
    Else
        BSPRICE = CVErr(xlErrValue)
    End If
End With
                    
End Function

Sub OptionPrice()

Dim s As Double, x As Double
Dim r As Double, v As Double, t As Double
Dim o As Variant
Dim Price As Variant
Dim App As Application
Set App = Application

    s = App.InputBox("Enter the Stock Price", "Step 1 of 6", , , , , , 1)
    x = App.InputBox("Enter the Exercise Price", "Step 2 of 6", , , , , , 1)
    r = App.InputBox("Enter the Interest Rate as a Decimal", "Step 3 of 6", , , , , , 1)
    v = App.InputBox("Enter the Stock Price Volatility as a Decimal", "Step 4 of 6", , , , , , 1)
    t = App.InputBox("Enter the Time to Expiration", "Step 5 of 6", , , , , , 1)
    o = App.InputBox("Enter the Option Type:" & vbNewLine & "Call, or Put", "Step 6 of 6", , , , , , 2)
    
    Price = BSPRICE(s, x, r, v, t, o)
    
    MsgBox "The option price is: " & Price & vbNewLine & vbNewLine & _
        "With inputs " & vbNewLine & vbNewLine & _
        "Stock price: " & Format(s, "$#,##0.00") & vbNewLine & _
        "Exerciee price: " & Format(x, "$#,##0.00") & vbNewLine & _
        "Interest rate: " & Format(r, "#0.00%") & vbNewLine & _
        "Volatility: " & Format(v, "#0.00%") & vbNewLine & _
        "Time: " & Format(t, "#,##0.00") & vbNewLine & _
        "Option Type " & o
End Sub

Sub OptionPricer()

    FrmOptionPricing.Show

End Sub
