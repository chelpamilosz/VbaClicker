Attribute VB_Name = "Module1"
Function Upgrade(clicks As Variant, perSecond As Variant, price As Variant, amount As Variant, CpS As Double, basePrice As LongLong) As Variant
    Dim varData(4) As Variant
    
    If clicks >= price Then
        clicks = clicks - price
        perSecond = perSecond + CpS
        price = Application.WorksheetFunction.RoundUp(basePrice * (1.15 ^ (amount + 1)), 0)
        amount = amount + 1
    Else
        MsgBox "You can't afford this upgrade!"
    End If
    
    varData(0) = clicks
    varData(1) = perSecond
    varData(2) = price
    varData(3) = amount
    
    Upgrade = varData
End Function

Sub BuyUpgrade(upgradeRow As Byte, CpS As Double, basePrice As LongLong)
    Dim arrUpgrade As Variant

    arrUpgrade = Upgrade(Cells(50, 5), Cells(15, 5), Cells(upgradeRow, 10), Cells(upgradeRow, 11), CpS, basePrice)

    Cells(50, 5) = arrUpgrade(0)
    Cells(15, 5) = arrUpgrade(1)
    Cells(upgradeRow, 10) = arrUpgrade(2)
    Cells(upgradeRow, 11) = arrUpgrade(3)
End Sub

Sub ButtonClicker_Click()
    clicks = Cells(50, 5)

    clicks = clicks + 1
    
    Cells(50, 5) = clicks
End Sub

Sub ButtonBuyUpgrade1_Click()
    Call BuyUpgrade(5, 0.1, 15)
End Sub

Sub ButtonBuyUpgrade2_Click()
    Call BuyUpgrade(6, 1, 100)
End Sub

Sub ButtonBuyUpgrade3_Click()
    Call BuyUpgrade(7, 8, 1100)
End Sub

Sub ButtonBuyUpgrade4_Click()
    Call BuyUpgrade(8, 47, 12000)
End Sub

Sub ButtonBuyUpgrade5_Click()
    Call BuyUpgrade(9, 260, 130000)
End Sub

Sub ButtonBuyUpgrade6_Click()
    Call BuyUpgrade(10, 1400, 1400000)
End Sub

Sub ButtonBuyUpgrade7_Click()
    Call BuyUpgrade(11, 7800, 20000000)
End Sub

Sub ButtonBuyUpgrade8_Click()
    Call BuyUpgrade(12, 44000, 330000000)
End Sub

Sub ButtonBuyUpgrade9_Click()
    Call BuyUpgrade(13, 260000, 5100000000#)
End Sub

Sub ButtonReset_Click()
    Cells(50, 5) = 0
    Cells(15, 5) = 0
    
    Cells(5, 10) = 15
    Cells(5, 11) = 0
    
    Cells(6, 10) = 100
    Cells(6, 11) = 0
    
    Cells(7, 10) = 1100
    Cells(7, 11) = 0
    
    Cells(8, 10) = 12000
    Cells(8, 11) = 0
    
    Cells(9, 10) = 130000
    Cells(9, 11) = 0
    
    Cells(10, 10) = 1400000
    Cells(10, 11) = 0
    
    Cells(11, 10) = 20000000
    Cells(11, 11) = 0
    
    Cells(12, 10) = 330000000
    Cells(12, 11) = 0
    
    Cells(13, 10) = 5100000000#
    Cells(13, 11) = 0
End Sub

Public Sub EventMacro()
    perSecond = Cells(15, 5)
    clicks = Cells(50, 5)
    
    clicks = clicks + perSecond
    
    Cells(50, 5) = clicks
    
    alertTime = Now + TimeValue("00:00:01")
    Application.OnTime alertTime, "EventMacro"
End Sub



