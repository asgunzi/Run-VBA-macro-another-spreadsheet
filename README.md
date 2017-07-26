# Run-VBA-macro-another-spreadsheet
A simple example of how to run a VBA macro in another spreadsheet

To run a VBA macro in another spreadsheet, we use the function Application.Run(NomePlanilha!NomeMacro)

Some examples: 
```VB

Application.Run ("PlotaVal.xlsb" & "!plotahora") 'Runs macro plotahora in the spreadsheet PlotaVal.xlsb
 
Application.Run ("PlotaVal.xlsb" & "!plotavalor(10)") 'Runs macro plotavalor with numeric parameter (10), in the spreadsheet PlotaVal.xlsb
 
Application.Run (strName & "!plotavalor(" & """" & "abc" & """" & ")") 'Runs macro plotavalor with string parameter "abc", in the spreadsheet PlotaVal.xlsb

```

A complete example: the macro in spreadsheet RodaMacroOutraPlanilha.xlsb runs a macro in the spreadsheet PlotaVal.xlsb.
```VB
Sub rodaMacroOutraPlan()
 
Dim strName As String
 
Application.DisplayAlerts = False
  
strName = "PlotaVal.xlsb" 'Name of the spreadsheet
   
Workbooks.Open "C:\Testes\" & strName 'Open spreadsheet

'Runs it
'Application.Run (strName & "!plotahora") 'Runs without parameters
 
'Application.Run (strName & "!plotavalor(10)") 'Runs macro with numeric parameter (10)
Application.Run (strName & "!plotavalor(" & """" & "abc" & """" & ")") 'Runs macro with texto parameter (“abc”)
 

Workbooks(strName).Save 'Save
  
Workbooks(strName).Close 'Closes
 
End Sub
```

Macro in spreadsheet PlotaVal.xlsb
```VB
Sub plotaValor(val As String)
 
Range("a2") = val
 
End Sub
 
Sub plotaHora()
 
Range("a1") = DateTime.Now
 
End Sub
```
