| Button Label | Format Code Example | Purpose                          |
| ------------ | ------------------- | -------------------------------- |
| **0dp**      | `#,##0`             | No decimal places                |
| **1dp/m**    | `#,##0.0,,"m"`      | One decimal place, in millions   |
| **\$0.0m**   | `$#,##0.0,,"m"`     | Currency in millions             |
| **1dp/k**    | `#,##0.0,"k"`       | One decimal place, in thousands  |
| **FX 4dp**   | `#,##0.0000`        | Four decimal places for FX rates |

Sub Format_0dp()
    Selection.NumberFormat = "#,##0"
End Sub

Sub Format_1dp_m()
    Selection.NumberFormat = "#,##0.0,,""m"""
End Sub

Sub Format_dollar_m()
    Selection.NumberFormat = "$#,##0.0,,""m"""
End Sub

Sub Format_1dp_k()
    Selection.NumberFormat = "#,##0.0,""k"""
End Sub

Sub Format_FX4dp()
    Selection.NumberFormat = "#,##0.0000"
End Sub
