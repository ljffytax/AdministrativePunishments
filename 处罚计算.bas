Attribute VB_Name = "模块1"
'依照深圳税务局处罚规则生成纳税人逾期申报处罚单
'作者 lianjie
'2018-07-29,2018-08-01,2018-09-06,2018-09-29,
'2018-10-28,2018-11-05,2018-12-03,2018-12-28
'2019-01-11,2019-02-27,2019-03-29,2019-04-01
'2019-04-15,2019-04-18,2019-04-28,2019-07-01
'2019-08-03,2020-08-10,2020-12-17,2021-01-13
'2021-02-22,2021-06-10
'如您使用此程序或源代码，请保留此版权申明，谢谢
'水平有限，bug肯定有，如果您发现了bug请自行修复，勿反馈给作者，谢谢
'When I wrote this down only god and I knew, now only god knows...
Dim qfirst As Date
Dim qlast As Date
Dim qnfirst As Date
Dim qnlast As Date
Dim zfirst As Date
Dim zlast As Date
Dim grjfirst As Date
Dim grjlast As Date
Dim grnfirst As Date
Dim grnlast As Date
Dim grgzfirst As Date
Dim grgzlast As Date
Dim taxpayerType As String
Dim iRow As Integer
Dim d1 As Date
Dim d2 As Date
Dim total As Integer
Dim earliestTax As String

Sub DoIt()
    qfirst = ThisWorkbook.Sheets(1).Cells(2, 2) '企业所得税季度
    qlast = ThisWorkbook.Sheets(1).Cells(2, 3)
    qnfirst = ThisWorkbook.Sheets(1).Cells(6, 2) '企业所得税年度
    qnlast = ThisWorkbook.Sheets(1).Cells(6, 3)
    zfirst = ThisWorkbook.Sheets(1).Cells(3, 2) '增值税
    zlast = ThisWorkbook.Sheets(1).Cells(3, 3)
    grjfirst = ThisWorkbook.Sheets(1).Cells(4, 2) '个人所得税季度
    grjlast = ThisWorkbook.Sheets(1).Cells(4, 3)
    grnfirst = ThisWorkbook.Sheets(1).Cells(7, 2) '个人所得税年度
    grnlast = ThisWorkbook.Sheets(1).Cells(7, 3)
    grgzfirst = ThisWorkbook.Sheets(1).Cells(5, 2)     '工资个税
    grgzlast = ThisWorkbook.Sheets(1).Cells(5, 3)
    taxpayerType = ThisWorkbook.Sheets(1).Cells(8, 2)
    iRow = 11
    d1 = "2000-01"
    d2 = "2000-01"

    If qlast < qfirst Or qnlast < qnfirst Or zlast < zfirst Then
        res = MsgBox("输入错误！", vbOKOnly)
        Exit Sub
    End If
    If grjlast < grjfirst Or grnlast < grnfirst Or grgzlast < grgzfirst Then
        res = MsgBox("输入错误！", vbOKOnly)
        Exit Sub
    End If
    If taxpayerType = "" Then
        res = MsgBox("输入错误！", vbOKOnly)
        Exit Sub
    End If
    If qlast > Now() Or qnlast > Now() Or zlast > Now() Or _
        grjlast > Now() Or grgzlast > Now() Or grnlast > Now() Then
        res = MsgBox("大佬！处罚属期大于今天！是否继续？", vbYesNo)
        If 7 = res Then
            Exit Sub
        End If
    End If
    ThisWorkbook.Sheets(1).Range("A11:D50") = ""
    
    '===================================================
    '处罚规则，申报期截止于2018-01-01前的所有申报可以按最
    '早应申报时间合并处罚，最高处罚公司800，个人/个体45元
    '申报期2018-01-01及以后可以按申报期合并处罚即不同税种
    '如果申报相同，则可以合并一起处罚。
    '===================================================
    '===================================================
    '处罚规则，申报期截止于2018-01-01前的所有申报可以按最
    '早应申报时间合并处罚，最高处罚公司1000，个人/个体50元
    '申报期2018-01-01及以后可以按申报期合并处罚即不同税种
    '如果申报相同，则可以合并一起处罚。(2020-08-10)
    '一些边界：2016-07-03，2016-12-30，2017-04-04，2017-07-03
    '===================================================
    '===================================================
    '处罚规则，申报期截止于2018-01-01前的所有申报可以按最
    '早应申报时间合并处罚，处罚公司50，个人/个体20元
    '申报期2018-01-01及以后可以按申报期合并处罚即不同税种
    '如果申报相同，则可以合并一起处罚。(2020-10-15日以后)
    '===================================================
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '小规模公司处罚规则
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If taxpayerType = "小规模纳税人(公司)" Then
        '公司不应该有生产经营个人所得税
        If (grjfirst <> "0:0:0" Or grnfirst <> "0:0:0") Then
            res = MsgBox("企业不应该有生产经营个人所得税！", vbOKOnly)
            Exit Sub
        End If
        earliestTax = getEarliestTax("smallScale")
        oldRule taxpayerType, earliestTax
        If qnfirst <> "0:0:0" And qnlast >= "2017-01-01" Then
            If qnfirst < "2017-01-01" Then
                d1 = "2017-01-01"
            Else
                d1 = qnfirst
            End If
            Do
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税年报 属期：" & Year(d1)
                ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByYear(d1, "enterprise")
                iRow = iRow + 1
                d1 = DateAdd("yyyy", 1, d1)
            Loop Until (d1 > qnlast)
        End If
        If qlast > "2017-09-30" Or zlast > "2017-09-30" Or grgzlast > "2017-11-30" Then
        '获取所有税种里最早的申报时间
            If qfirst <= "2017-09-30" Or zfirst <= "2017-09-30" Then
                d1 = "2017-10-01"
            ElseIf grgzfirst <= "2017-11-30" Then
                d1 = "2017-12-01"
            Else
                If qfirst < zfirst Then
                    d1 = qfirst
                Else
                    d1 = zfirst
                End If
                If grgzfirst < d1 And grgzfirst <> "0:0:0" Then
                    d1 = grgzfirst
                End If
            End If
            If qlast > zlast Then
                d2 = qlast
            Else
                d2 = zlast
            End If
            If grgzlast > d2 Then
                d2 = grgzlast
            End If
            Do
                If d1 <= qlast And d1 >= qfirst And d1 <= zlast And d1 >= zfirst And d1 <= grgzlast _
                And d1 >= grgzfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税 增值税" & date2Season(d1) & _
                        "扣缴个税" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    ElseIf d1 >= "2017-12-01" Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If

                ElseIf d1 <= qlast And d1 >= qfirst And d1 <= zlast And d1 >= zfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税 增值税" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                ElseIf d1 <= qlast And d1 >= qfirst And d1 <= grgzlast And d1 >= grgzfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税 扣缴个税" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    ElseIf d1 >= "2017-12-01" Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                ElseIf d1 <= zlast And d1 >= zfirst And d1 <= grgzlast And d1 >= grgzfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 扣缴个税" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    ElseIf d1 >= "2017-12-01" Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                ElseIf d1 <= zlast And d1 >= zfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                ElseIf d1 <= grgzlast And d1 >= grgzfirst And d1 >= "2017-12-01" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                    iRow = iRow + 1
                ElseIf d1 <= qlast And d1 >= qfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税季度 属期：" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                End If
                d1 = DateAdd("m", 1, d1)
            Loop Until (d1 > d2)
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
    '小规模个体处罚规则
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf taxpayerType = "小规模纳税人(个体)" Then
        '个体可以同时存在扣缴个税与增值税 或者 单独的生产经营个税
        If grjfirst <> "0:0:0" And (qfirst <> "0:0:0" Or qnfirst <> "0:0:0" Or _
        grgzfirst <> "0:0:0" Or zfirst <> "0:0:0") Then
            res = MsgBox("生产经营个税与其它税种不能同时存在!", vbOKOnly)
            Exit Sub
        End If
        If qfirst <> "0:0:0" Or qnfirst <> "0:0:0" Or (grgzfirst <> "0:0:0" And _
        (grjfirst <> "0:0:0" Or grnfirst <> "0:0:0")) Then
            res = MsgBox("个体不应该存在企业所得税，同时不应该扣缴个税与生产经营同时存在!", vbOKOnly)
            Exit Sub
        End If

        '处理增值税与扣缴个人所得税
        earliestTax = getEarliestTax("smallScale")
        If earliestTax = "zfirst" Then
            If zfirst <> "0:0:0" And zfirst <= "2017-09-30" Then
                If zlast >= "2017-09-30" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & "~2017-9"
                Else
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & "~" & _
                    date2Month(zlast)
                End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedBySeason(zfirst, "individual")
            iRow = iRow + 1
            End If
        ElseIf earliestTax = "grgzfirst" Then
            If grgzfirst <> "0:0:0" And grgzfirst <= "2017-11-30" Then
                If grgzlast > "2017-11-30" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(grgzfirst) & "~2017-11"
                Else
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(grgzfirst) & "~" & _
                    date2Month(grgzlast)
                End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByMonth(grgzfirst, "individual")
            iRow = iRow + 1
            End If
        '处理生产经营个人所得税
        ElseIf earliestTax = "grjfirst" Then
            onlyGrscjy "grjfirst"
        ElseIf earliestTax = "grnfirst" Then
            onlyGrscjy "grnfirst"
        End If

        If grnlast >= "2017-01-01" Then
            If grnfirst < "2018-01-01" Then
                d1 = "2017-01-01"
            Else
                d1 = grnfirst
            End If
            Do
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税年报 属期：" & Year(d1)
                ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByYear(d1, "individual")
                iRow = iRow + 1
                d1 = DateAdd("yyyy", 1, d1)
            Loop Until (d1 > grnlast)
        End If
        If zlast > "2017-09-30" Or grgzlast > "2017-11-30" Then
            If zfirst <= "2017-09-30" Or grgzfirst <= "2017-11-30" Then
                d1 = "2017-12-01"
            Else
                If grgzfirst < zfirst Then
                    d1 = grgzfirst
                Else
                    d1 = zfirst
                End If
            End If
            If grgzlast > zlast Then
                d2 = grgzlast
            Else
                d2 = zlast
            End If
            Do
                If d1 <= grgzlast And d1 >= grgzfirst And d1 <= zlast And d1 >= zfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 " & date2Season(d1) & "扣缴个税 " & date2Month(d1)
                    Else
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                    End If
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                    iRow = iRow + 1
                ElseIf d1 <= zlast And d1 >= zfirst And (Month(d1) Mod 3) = 0 Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Season(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                    iRow = iRow + 1
                ElseIf d1 <= grgzlast And d1 >= grgzfirst Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                    iRow = iRow + 1
                End If
                d1 = DateAdd("m", 1, d1)
            Loop Until (d1 > d2)
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    '一般纳税人公司处罚规则
    ''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf taxpayerType = "一般纳税人(公司)" Then
        '公司不应该有生产经营个人所得税
        If (grjfirst <> "0:0:0" Or grnfirst <> "0:0:0") Then
            res = MsgBox("企业不应该有生产经营个人所得税！", vbOKOnly)
            Exit Sub
        End If
        earliestTax = getEarliestTax("generalTaxpayer")
        oldRule taxpayerType, earliestTax
        If qnfirst <> "0:0:0" And qnlast >= "2017-01-01" Then
            If qnfirst < "2017-01-01" Then
                d1 = "2017-01-01"
            Else
                d1 = qnfirst
            End If
            Do
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税年报 属期：" & Year(d1)
                ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByYear(d1, "enterprise")
                iRow = iRow + 1
                d1 = DateAdd("yyyy", 1, d1)
            Loop Until (d1 > qnlast)
        End If
        If qlast > "2017-09-30" Or zlast > "2017-11-30" Or grgzlast > "2017-11-30" Then
        '获取所有税种里最早的申报时间
            If qfirst <= "2017-09-30" Or zfirst <= "2017-11-30" Then
                d1 = "2017-10-01"
            ElseIf grgzfirst <= "2017-11-30" Then
                d1 = "2017-12-01"
            Else
                If qfirst < zfirst Then
                    d1 = qfirst
                Else
                    d1 = zfirst
                End If
                If grgzfirst < d1 And grgzfirst <> "0:0:0" Then
                    d1 = grgzfirst
                End If
            End If
            If qlast > zlast Then
                d2 = qlast
            Else
                d2 = zlast
            End If
            If grgzlast > d2 Then
                d2 = grgzlast
            End If
            Do
                If d1 <= qlast And d1 >= qfirst And d1 <= zlast And d1 >= zfirst And d1 <= grgzlast _
                And d1 >= grgzfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税" & date2Season(d1) & _
                        "增值税 扣缴个税" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    ElseIf d1 >= "2017-12-01" Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 增值税 属期：" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                ElseIf d1 <= qlast And d1 >= qfirst And d1 <= zlast And d1 >= zfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税 增值税" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    ElseIf d1 >= "2017-12-01" Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                ElseIf d1 <= qlast And d1 >= qfirst And d1 <= grgzlast And d1 >= grgzfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税 扣缴个税" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    ElseIf d1 >= "2017-12-01" Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                ElseIf d1 <= zlast And d1 >= zfirst And d1 <= grgzlast And d1 >= grgzfirst And d1 >= "2017-12-01" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 扣缴个税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                    iRow = iRow + 1
                ElseIf d1 <= zlast And d1 >= zfirst And d1 >= "2017-12-01" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                    iRow = iRow + 1
                ElseIf d1 <= grgzlast And d1 >= grgzfirst And d1 >= "2017-12-01" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                    iRow = iRow + 1
                ElseIf d1 <= qlast And d1 >= qfirst Then
                    If (Month(d1) Mod 3) = 0 Then
                        ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税季度 属期：" & date2Season(d1)
                        ThisWorkbook.Sheets(1).Cells(iRow, 4) = 50
                        iRow = iRow + 1
                    End If
                End If
                d1 = DateAdd("m", 1, d1)
            Loop Until (d1 > d2)
        End If
    '''''''''''''''''''''''''''''''''''''''''''''''
    '一般纳税人个体处罚规则
    '''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf taxpayerType = "一般纳税人(个体)" Then
        '个体可以同时存在扣缴个税与增值税 或者 单独的生产经营个税
        If grjfirst <> "0:0:0" And (qfirst <> "0:0:0" Or qnfirst <> "0:0:0" Or grgzfirst <> "0:0:0" Or zfirst <> "0:0:0") Then
            res = MsgBox("生产经营个税与其它税种不能同时存在!", vbOKOnly)
            Exit Sub
        End If
        If qfirst <> "0:0:0" Or qnfirst <> "0:0:0" Or (grgzfirst <> "0:0:0" And _
        (grjfirst <> "0:0:0" Or grnfirst <> "0:0:0")) Then
            res = MsgBox("个体不应该存在企业所得税，同时不应该扣缴个税与生产经营同时存在!", vbOKOnly)
            Exit Sub
        End If

        '处理增值税与扣缴个人所得税
        earliestTax = getEarliestTax("generalTaxpayer")
        If earliestTax = "zfirst" Then
            If zfirst <> "0:0:0" And zfirst <= "2017-11-30" Then
                If zlast >= "2017-11-30" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & "~2017-11"
                Else
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & "~" & _
                    date2Month(zlast)
                End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByMonth(zfirst, "individual")
            iRow = iRow + 1
            End If
        ElseIf earliestTax = "grgzfirst" Then
            If grgzfirst <> "0:0:0" And grgzfirst <= "2017-11-30" Then
                If grgzlast > "2017-11-30" Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(grgzfirst) & "~2017-11"
                Else
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(grgzfirst) & "~" & _
                    date2Month(grgzlast)
                End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByMonth(zfirst, "individual")
            iRow = iRow + 1
            End If
        '处理生产经营个人所得税
        ElseIf earliestTax = "grjfirst" Then
            onlyGrscjy "grjfirst"
        ElseIf earliestTax = "grnfirst" Then
            onlyGrscjy "grnfirst"
        End If
        
        If grnlast >= "2017-01-01" Then
            If grnfirst < "2018-01-01" Then
                d1 = "2017-01-01"
            Else
                d1 = grnfirst
            End If
            Do
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税年报 属期：" & Year(d1)
                ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByYear(d1, "individual")
                iRow = iRow + 1
                d1 = DateAdd("yyyy", 1, d1)
            Loop Until (d1 > grnlast)
        End If
        If zlast > "2017-11-30" Or grgzlast > "2017-11-30" Then
            If zfirst <= "2017-11-30" Or grgzfirst <= "2017-11-30" Then
                d1 = "2017-12-01"
            Else
                If grgzfirst < zfirst Then
                    d1 = grgzfirst
                Else
                    d1 = zfirst
                End If
            End If
            If grgzlast > zlast Then
                d2 = grgzlast
            Else
                d2 = zlast
            End If
            Do
                If d1 <= grgzlast And d1 >= grgzfirst And d1 <= zlast And d1 >= zfirst Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 扣缴个税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                    iRow = iRow + 1
                ElseIf d1 <= zlast And d1 >= zfirst Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                    iRow = iRow + 1
                ElseIf d1 <= grgzlast And d1 >= grgzfirst Then
                    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(d1)
                    ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                    iRow = iRow + 1
                End If
                d1 = DateAdd("m", 1, d1)
            Loop Until (d1 > d2)
        End If
    End If

    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "合计罚款金额"
    iRow = iRow + 2
    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "常用语"
    iRow = iRow + 1
    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "改正上述违法行为"
    iRow = iRow + 1
    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "经系统核实，该违法行为属首次违法，不予罚款。"
    iRow = iRow + 1
    ThisWorkbook.Sheets(1).Cells(iRow, 1) = "======================"
    res = delCOVID2019()
    iRow = 11
    Do While (ThisWorkbook.Sheets(1).Cells(iRow, 1) <> "合计罚款金额")
        iRow = iRow + 1
    Loop
    ThisWorkbook.Sheets(1).Cells(iRow, 4) = addMoney()
End Sub

Function oldRule(taxpayerType As String, taxName As String) As Integer
    If taxName = "qfirst" Then
        If qfirst <= "2017-09-30" Then
            If qlast >= "2017-09-30" Then
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税 属期：" & date2Month(qfirst) & _
                "~2017-9"
            Else
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税 属期：" & date2Month(qfirst) & _
                "~" & date2Month(qlast)
            End If
        ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedBySeason(qfirst, "enterprise")
        iRow = iRow + 1
        End If
    ElseIf taxName = "zfirst" And taxpayerType = "小规模纳税人(公司)" Then
        If zfirst <= "2017-09-30" Then
            If zlast >= "2017-09-30" Then
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & _
                "~2017-9"
            Else
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & "~" _
                & date2Month(zlast)
            End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedBySeason(zfirst, "enterprise")
            iRow = iRow + 1
        End If
    ElseIf taxName = "zfirst" And taxpayerType = "一般纳税人(公司)" Then
        If zfirst <= "2017-11-30" Then
            If zlast >= "2017-11-30" Then
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & _
                "~2017-11"
            Else
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "增值税 属期：" & date2Month(zfirst) & "~" _
                & date2Month(zlast)
            End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByMonth(zfirst, "enterprise")
            iRow = iRow + 1
        End If
    ElseIf taxName = "grgzfirst" Then
        If grgzfirst <= "2017-11-30" Then
            If grgzlast >= "2017-11-30" Then
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(grgzfirst) & _
                "~2017-11"
            Else
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "扣缴个税 属期：" & date2Month(grgzfirst) & "~" _
                & date2Month(grgzlast)
            End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByMonth(grgzfirst, "enterprise")
            iRow = iRow + 1
        End If
    ElseIf taxName = "qnfirst" Then
        If qnfirst < "2017-01-01" Then
            ThisWorkbook.Sheets(1).Cells(iRow, 1) = "企业所得税年报 属期：" & Year(qnfirst) & "-" & _
            Month(qnfirst) & "~2016-12"
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByYear(qnfirst, "enterprise")
            iRow = iRow + 1
        End If
    End If
End Function

'以下注释掉的代码是为适应2020-10-15开始的新处罚裁量基准
Function punishedBySeason(startTime As Date, taxpayerType As String) As Integer
    If taxpayerType = "enterprise" Then
        punishedBySeason = 50
        'If startTime >= "2017-07-01" And startTime <= "2017-09-30" Then
        '    punishedBySeason = 50
        'ElseIf startTime >= "2017-04-01" And startTime <= "2017-06-30" Then
        '    punishedBySeason = 100
        'ElseIf startTime >= "2017-01-01" And startTime <= "2017-03-31" Then
        '    punishedBySeason = 200
        'ElseIf startTime >= "2016-10-01" And startTime <= "2016-12-31" Then
        '    punishedBySeason = 400
        'ElseIf startTime >= "2016-07-01" And startTime <= "2016-09-30" Then
        '    punishedBySeason = 600
        'ElseIf startTime >= "2016-04-01" And startTime <= "2016-06-30" Then
        '    punishedBySeason = 800
        'ElseIf startTime <= "2016-03-31" Then
        '    punishedBySeason = 1000
        'End If
    Else
        punishedBySeason = 20
        'If startTime >= "2017-07-01" And startTime <= "2017-09-30" Then
        '    punishedBySeason = 20
        'ElseIf startTime >= "2017-04-01" And startTime <= "2017-06-30" Then
        '    punishedBySeason = 25
        'ElseIf startTime >= "2017-01-01" And startTime <= "2017-03-31" Then
        '    punishedBySeason = 30
        'ElseIf startTime >= "2016-10-01" And startTime <= "2016-12-31" Then
        '    punishedBySeason = 35
        'ElseIf startTime >= "2016-07-01" And startTime <= "2016-09-30" Then
        '    punishedBySeason = 40
        'ElseIf startTime >= "2016-04-01" And startTime <= "2016-06-30" Then
        '    punishedBySeason = 45
        'ElseIf startTime <= "2016-03-31" Then
        '    punishedBySeason = 50
        'End If
    End If
End Function

Function punishedByMonth(startTime As Date, taxpayerType As String) As Integer
    If taxpayerType = "enterprise" Then
        punishedByMonth = 50
        'If startTime >= "2017-09-01" And startTime <= "2017-11-30" Then
        '    punishedByMonth = 50
        'ElseIf startTime >= "2017-06-01" And startTime <= "2017-08-31" Then
        '    punishedByMonth = 100
        'ElseIf startTime >= "2017-03-01" And startTime <= "2017-05-31" Then
        '    punishedByMonth = 200
        'ElseIf startTime >= "2016-12-01" And startTime <= "2017-02-28" Then
        '    punishedByMonth = 400
        'ElseIf startTime >= "2016-09-01" And startTime <= "2016-11-30" Then
        '    punishedByMonth = 600
        'ElseIf startTime >= "2016-06-01" And startTime <= "2016-08-31" Then
        '    punishedByMonth = 800
        'ElseIf startTime <= "2016-05-31" Then
        '    punishedByMonth = 1000
        'End If
    Else
        punishedByMonth = 20
        'If startTime >= "2017-09-01" And startTime <= "2017-11-30" Then
        '    punishedByMonth = 20
        'ElseIf startTime >= "2017-06-01" And startTime <= "2017-08-31" Then
        '    punishedByMonth = 25
        'ElseIf startTime >= "2017-03-01" And startTime <= "2017-05-31" Then
        '    punishedByMonth = 30
        'ElseIf startTime >= "2016-12-01" And startTime <= "2017-02-28" Then
        '    punishedByMonth = 35
        'ElseIf startTime >= "2016-09-01" And startTime <= "2016-11-30" Then
        '    punishedByMonth = 40
        'ElseIf startTime >= "2016-06-01" And startTime <= "2016-08-31" Then
        '    punishedByMonth = 45
        'ElseIf startTime <= "2016-05-31" Then
        '    punishedByMonth = 50
        'End If
    End If
End Function

Function punishedByYear(startTime As Date, taxpayerType As String) As Integer
    If taxpayerType = "enterprise" Then
        y = Year(startTime)
        punishedByYear = 50
        'If y <= 2015 Then
        '    punishedByYear = 1000
        'ElseIf y = 2016 Then
        '    punishedByYear = 200
        'Else
        '    punishedByYear = 50
        'End If
    ElseIf taxpayerType = "individual" Then
        y = Year(startTime)
        punishedByYear = 20
        'If y <= 2015 Then
        '    punishedByYear = 50
        'ElseIf y = 2016 Then
        '    punishedByYear = 30
        'Else
        '    punishedByYear = 20
        'End If
    End If
End Function

Function date2Season(dt As Date) As String
    mMonth = Month(dt)
    If mMonth >= 1 And mMonth <= 3 Then
        date2Season = Year(dt) & "-1~" & Year(dt) & "-3"
    ElseIf mMonth >= 4 And mMonth <= 6 Then
        date2Season = Year(dt) & "-4~" & Year(dt) & "-6"
    ElseIf mMonth >= 7 And mMonth <= 9 Then
        date2Season = Year(dt) & "-7~" & Year(dt) & "-9"
    Else
        date2Season = Year(dt) & "-10~" & Year(dt) & "-12"
    End If
End Function

Function date2EndOfSeason(dt As Date) As Date
    mMonth = Month(dt)
    If mMonth >= 1 And mMonth <= 3 Then
        date2EndOfSeason = Year(dt) & "-3-31"
    ElseIf mMonth >= 4 And mMonth <= 6 Then
        date2EndOfSeason = Year(dt) & "-6-30"
    ElseIf mMonth >= 7 And mMonth <= 9 Then
        date2EndOfSeason = Year(dt) & "-9-30"
    Else
        date2EndOfSeason = Year(dt) & "-12-31"
    End If
End Function

Function date2Month(dt As Date) As String
    date2Month = Year(dt) & "-" & Month(dt)
End Function

Function getEarliestTax(taxpayerType As String) As String
    '将属期转换到申报期限，再返回最早申报期的那个税种
    Dim earliestDate As Date
    earliestDate = "2025-01-01"
    
    If qfirst <> "0:0:0" Then
        sbqfirst = DateAdd("d", 14, date2EndOfSeason(qfirst))
    Else
        sbqfirst = "0:0:0"
    End If
    If grjfirst <> "0:0:0" Then
        sbgrjfirst = DateAdd("d", 14, date2EndOfSeason(grjfirst))
    Else
        sbgrjfirst = "0:0:0"
    End If
    If qnfirst <> "0:0:0" Then
        sbqnfirst = DateAdd("yyyy", 1, Year(qnfirst) & "-01-01")
        sbqnfirst = DateAdd("m", 6, sbqnfirst)
        sbqnfirst = DateAdd("d", -1, sbqnfirst)
    Else
        sbqnfirst = "0:0:0"
    End If
    If grnfirst <> "0:0:0" Then
        sbgrnfirst = DateAdd("yyyy", 1, Year(grnfirst) & "-01-01")
        sbgrnfirst = DateAdd("m", 4, sbgrnfirst)
        sbgrnfirst = DateAdd("d", -1, sbgrnfirst)
    Else
        sbgrnfirst = "0:0:0"
    End If
    If grgzfirst <> "0:0:0" Then
        sbgrgzfirst = DateAdd("m", 1, grgzfirst)
        sbgrgzfirst = Year(sbgrgzfirst) & "-" & Month(sbgrgzfirst)
        sbgrgzfirst = DateAdd("d", 14, sbgrgzfirst)
    Else
        sbgrgzfirst = "0:0:0"
    End If
    '小规模纳税人
    If taxpayerType = "smallScale" Then
        If zfirst <> "0:0:0" Then
            sbzfirst = DateAdd("d", 14, date2EndOfSeason(zfirst))
        Else
            sbzfirst = "0:0:0"
        End If
    Else
        If zfirst <> "0:0:0" Then
            sbzfirst = DateAdd("m", 1, zfirst)
            sbzfirst = Year(sbzfirst) & "-" & Month(sbzfirst)
            sbzfirst = DateAdd("d", 14, sbzfirst)
        Else
            sbzfirst = "0:0:0"
        End If
    End If

    If sbqfirst <> "0:0:0" And sbqfirst < earliestDate Then
        earliestDate = sbqfirst
        getEarliestTax = "qfirst"
    End If
    If sbzfirst <> "0:0:0" And sbzfirst < earliestDate Then
        earliestDate = sbzfirst
        getEarliestTax = "zfirst"
    End If
    If sbgrjfirst <> "0:0:0" And sbgrjfirst < earliestDate Then
        earliestDate = sbgrjfirst
        getEarliestTax = "grjfirst"
    End If
    If sbgrgzfirst <> "0:0:0" And sbgrgzfirst < earliestDate Then
        earliestDate = sbgrgzfirst
        getEarliestTax = "grgzfirst"
    End If
    If sbqnfirst <> "0:0:0" And sbqnfirst < earliestDate Then
        earliestDate = sbqnfirst
        getEarliestTax = "qnfirst"
    End If
    If sbgrnfirst <> "0:0:0" And sbgrnfirst < earliestDate Then
        earliestDate = sbqnfirst
        getEarliestTax = "grnfirst"
    End If
End Function

Function onlyGrscjy(earliestTax As String) As Boolean '仅处理个人生产经营
    d1 = "2000-01"
    d2 = "2000-01"
    If earliestTax = "grjfirst" Then
        If grjfirst <> "0:0:0" And grjfirst <= "2017-09-30" Then
            If grjlast >= "2017-09-30" Then
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税 属期：" & date2Month(grjfirst) & "~2017-9"
            Else
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税 属期：" & date2Month(grjfirst) & _
                "~" & date2Month(grjlast)
            End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedBySeason(grjfirst, "individual")
            iRow = iRow + 1
        End If
        If grjlast > "2017-09-30" Then
            If grjfirst < "2017-10-01" Then
                d1 = "2017-10-01"
            Else
                d1 = grjfirst
            End If
            Do
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税 属期：" & date2Season(d1)
                ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                iRow = iRow + 1
                d1 = DateAdd("q", 1, d1)
            Loop Until (d1 > grjlast)
        End If
    ElseIf earliestTax = "grnfirst" Then
        If grnfirst <> "0:0:0" And grnfirst <= "2016-12-31" Then
            If grnlast >= "2016-12-31" Then
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税年报 属期：" & Year(grnfirst) & "~2016"
            Else
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税年报 属期：" & Year(grnfirst) & _
                "~" & Year(grnlast)
            End If
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = punishedByYear(grnfirst, "individual")
            iRow = iRow + 1
        End If
        If grjlast > "2017-09-30" Then
            If grjfirst < "2017-10-01" Then
                d1 = "2017-10-01"
            Else
                d1 = grjfirst
            End If
            Do
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = "生产经营个税 属期：" & date2Season(d1)
                ThisWorkbook.Sheets(1).Cells(iRow, 4) = 20
                iRow = iRow + 1
                d1 = DateAdd("q", 1, d1)
            Loop Until (d1 > grjlast)
        End If
    End If
    onlyGrscjy = True
End Function

Function delCOVID2019() As Boolean
    iRow = 11
    iFast = 11
    iSlow = 11
    Do
        st = ThisWorkbook.Sheets(1).Cells(iRow, 1)
        If (Left(Right(st, 2), 1) = "-" And Left(Right(st, 6), 4) = "2020") Then '2020-1~2020-9
            ThisWorkbook.Sheets(1).Cells(iRow, 1) = ""
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = ""
        Else
            If (Right(st, 7) = "2020-10" Or Right(st, 4) = "2019") Then
                ThisWorkbook.Sheets(1).Cells(iRow, 1) = ""
                ThisWorkbook.Sheets(1).Cells(iRow, 4) = ""
            End If
        End If
        iRow = iRow + 1
    Loop Until (ThisWorkbook.Sheets(1).Cells(iRow, 1) = "合计罚款金额" Or iRow > 50)
    
    Do
        st = ThisWorkbook.Sheets(1).Cells(iFast, 1)
        If (ThisWorkbook.Sheets(1).Cells(iSlow, 1) = "") Then
            Do
                iFast = iFast + 1
            Loop Until (ThisWorkbook.Sheets(1).Cells(iFast, 1) <> "")
            ThisWorkbook.Sheets(1).Cells(iSlow, 1) = ThisWorkbook.Sheets(1).Cells(iFast, 1)
            ThisWorkbook.Sheets(1).Cells(iSlow, 4) = ThisWorkbook.Sheets(1).Cells(iFast, 4)
            ThisWorkbook.Sheets(1).Cells(iFast, 1) = ""
            ThisWorkbook.Sheets(1).Cells(iFast, 4) = ""
        End If
        iSlow = iSlow + 1
        iFast = iSlow
    Loop Until (ThisWorkbook.Sheets(1).Cells(iSlow - 1, 1) = "======================" Or iSlow > 50)
    delCOVID2019 = True
End Function

Function addMoney() As Integer
    total = 0
    iRow = 11
    Do While (ThisWorkbook.Sheets(1).Cells(iRow, 1) <> "合计罚款金额")
        If ((ThisWorkbook.Sheets(1).Cells(iRow, 4) = 0 Or ThisWorkbook.Sheets(1).Cells(iRow, 4) = "首违") And ThisWorkbook.Sheets(1).Cells(iRow, 4) <> "") Then
            ThisWorkbook.Sheets(1).Cells(iRow, 4) = "首违"
        Else
            total = total + ThisWorkbook.Sheets(1).Cells(iRow, 4)
        End If
        iRow = iRow + 1
    Loop
    addMoney = total
End Function

'======打印处罚单======
Sub printPaper()
    If (iRow >= 11) Then
        PrintArea = "$A$10:$D$" + CStr(iRow - 5)
        ActiveSheet.PageSetup.PrintArea = PrintArea
        ActiveWindow.SelectedSheets.PrintPreview
    End If
End Sub

Sub totalMoney()
    iRow = 11
    Do While (ThisWorkbook.Sheets(1).Cells(iRow, 1) <> "合计罚款金额")
        iRow = iRow + 1
    Loop
    ThisWorkbook.Sheets(1).Cells(iRow, 4) = addMoney()
End Sub

Sub exitWithoutSave()
    'ActiveWorkbook.Close savechanges:=False
    Application.DisplayAlerts = False
    Application.Quit
End Sub

