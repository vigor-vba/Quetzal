# Quetzal

A simple ADO manipulator for lazy folks.

## Example

To fetch a recordset take this code as below:

```VBS
Dim qz As New Quetzal

If qz.Connect(dbOracle, "Src", "Usr", "PW").IsConnect Then

    Call qz.Query("select * from table1")
    
    If qz.HasRecordset Then
        Set rec = qz.Recordset
    End If

End If
```

You can use chainmethod to shorten this.

```VBS
Dim qz As New Quetzal

Set rec = qz.Connect(dbOracle, "Src", "Usr", "PW") _
            .Query("select * from table1") _
            .Recordset
```
