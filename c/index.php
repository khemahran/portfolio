<pre style="color: rgb(0, 0, 0); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; text-decoration-style: initial; text-decoration-color: initial; word-wrap: break-word; white-space: pre-wrap;">Sub Macro6()
&apos;
    Range(&quot;F5:P999&quot;).Select
    Selection.ClearContents
    Columns(&quot;B:B&quot;).Select
    Application.CutCopyMode = False
    Range(&quot;B1:B18&quot;).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range( _
        &quot;E4&quot;), Unique:=True
If Range(&quot;F1&quot;).Value &gt;= 1 Then

    For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;F3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;F&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;F&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
    End If
    If Range(&quot;F1&quot;).Value &gt;= 2 Then
    For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;G3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;G&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;G&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
    End If
    If Range(&quot;F1&quot;).Value &gt;= 3 Then
        For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;H3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;H&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;H&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
     End If
    If Range(&quot;F1&quot;).Value &gt;= 4 Then
        For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;I3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;I&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;I&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
     End If
    If Range(&quot;F1&quot;).Value &gt;= 5 Then
        For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;J3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;J&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;J&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
    
     End If
    If Range(&quot;F1&quot;).Value &gt;= 6 Then
            For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;K3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;K&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;K&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
    
     End If
    If Range(&quot;F1&quot;).Value &gt;= 7 Then
       
            For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;L3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;L&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;L&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
    
     End If
    If Range(&quot;F1&quot;).Value &gt;= 8 Then
       
            For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;M3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;M&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;M&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
    
        End If
    If Range(&quot;F1&quot;).Value &gt;= 9 Then
            For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;N3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;N&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;N&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
        End If
    If Range(&quot;F1&quot;).Value &gt;= 10 Then
            For k = 5 To 7
                    test = False
                    For i = 1 To 15
                            If Range(&quot;A&quot; &amp; i).Value = Range(&quot;O3&quot;).Value And Range(&quot;B&quot; &amp; i).Value = Range(&quot;E&quot; &amp; k).Value Then
                            test = True
                            Range(&quot;O&quot; &amp; k).Select
                            ActiveCell.FormulaR1C1 = &quot;x&quot;
                            ElseIf test = False Then
                                 Range(&quot;O&quot; &amp; k).Select
                                 ActiveCell.FormulaR1C1 = &quot;-&quot;
                         End If
                    Next i
    Next k
    
End If
 Range(&quot;P5&quot;).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = &quot;=COUNTIF(RC[-10]:RC[-1],&quot;&quot;x&quot;&quot;)&quot;
    Selection.AutoFill Destination:=Range(&quot;P5:P7&quot;), Type:=xlFillDefault
        Range(&quot;E4:P4&quot;).Select
    Selection.AutoFilter

End Sub</pre>