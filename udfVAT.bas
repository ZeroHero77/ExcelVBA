Attribute VB_Name = "udfVAT"
Function udf_VAT(Amount As Double, Rate As Double)

'Formula to calculate VAT (amount,rate)

udf_VAT = Amount * (Rate / 100)

End Function
