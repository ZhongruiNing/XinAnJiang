Attribute VB_Name = "Module1"
'新安江日模型计算程序
Public Function xaj_day(para() As Single, data() As Single, result() As Single, intial() As Single)
    'para: 模型参数（16个变量）
    'data: p，e数据（2维）
    'result: 模拟结果（7列）：径流量、蒸发量、土壤含水量（上，中，下，总）、自由水蓄水量，（单位：mm）

'********************dim model parameter********************
    Dim kc As Single, um As Integer, lm As Integer, c As Single  '蒸发
    Dim wm As Integer, b As Single, im As Single '产流
    Dim sm As Single, ex As Single, kg As Single, ki As Single '分水源
    Dim cs As Single, ci As Single, cg As Single, cr As Single, lr As Integer '汇流
    
    Dim dm As Integer '下层蓄水容量
    Dim wmm As Single '流域最大点蓄水容量
    Dim smm As Single  '流域最大点自由水蓄水容量
    
'********************模型中间变量********************
    Dim p() As Single, pan_e() As Single '降水和蒸发资料
    Dim ep() As Single '流域蒸发能力
    Dim eu() As Single, el() As Single, ed() As Single, e() As Single '3层及总实际蒸发
    Dim wu() As Single, wl() As Single, wd() As Single, w() As Single '3层及总的土壤含水量
    Dim pe() As Single '净雨
    Dim a() As Single '蓄水容量曲线纵坐标
    Dim r() As Single ' 总径流量
    Dim fr() As Single '产流面积
    Dim au() As Single '自由水容量曲线纵坐标
    Dim s() As Single '自由水蓄水量
    Dim rs() As Single, ri() As Single, rg() As Single '3种径流成分
    Dim qs() As Single, qi() As Single, qg() As Single '单元出口3种径流汇流
    Dim qds() As Single '单元出口总径流
    Dim qsim() As Single '流域总出流
    
    Dim num As Long '总时间步长数
    Dim ns As Integer '自由水蓄水库水量平衡时段划分
    Dim r2 As Single '时段细化后入流量
    Dim pe2 As Single
    Dim fr2 As Single
    Dim s2() As Single '时段细化后自由水蓄水量
    Dim au2() As Single
    Dim ki2 As Single '时段细化后壤中流出流系数
    Dim kg2 As Single '时段细化后地下径流出流系数
    Dim rs2() As Single '时段细化后地表径流
    Dim rg2() As Single '时段细化后地下径流
    Dim ri2() As Single '时段细化后壤中流
    
    Dim i As Long, j As Integer, temp1 As Long
        
'********************read model parameter********************
    kc = para(1): um = para(2): lm = para(3): c = para(4)
    wm = para(5): b = para(6): im = para(7)
    sm = para(8): ex = para(9): kg = para(10): ki = para(11)
    cs = para(12): ci = para(13): cg = para(14): cr = para(15): lr = para(16)
    dm = wm - um - lm
    wmm = wm * (1 + b)
    smm = sm * (1 + ex)
    
    num = UBound(data, 2)
    ReDim Preserve p(num), pan_e(num)
    ReDim Preserve ep(num)
    ReDim Preserve eu(num), el(num), ed(num)
    ReDim Preserve e(num)
    ReDim Preserve wu(num + 1), wl(num + 1), wd(num + 1), w(num + 1)
    ReDim Preserve pe(num)
    ReDim Preserve a(num + 1)
    ReDim Preserve r(num)
    ReDim Preserve fr(num + 1)
    ReDim Preserve au(num + 1)
    ReDim Preserve s(num + 1)
    ReDim Preserve rs(num), ri(num), rg(num)
    ReDim Preserve qs(num), qi(num), qg(num)
    ReDim Preserve qds(num), qsim(num)
        
'********************初始土壤状态设定********************
    wu(1) = intial(1)
    wl(1) = intial(2)
    wd(1) = intial(3)
    w(1) = wu(1) + wl(1) + wd(1)
    a(1) = wmm * (1 - (1 - w(1) / wm) ^ (1 / (1 + b)))
    s(1) = 0
    au(1) = 0
    fr(1) = 1 - (1 - a(1) / wmm) ^ b
    
    For i = 1 To num
        p(i) = data(1, i) * (1 - im)
        pan_e(i) = data(2, i)
        
'********************蒸散发计算********************
        ep(i) = kc * pan_e(i)
        If wu(i) + p(i) >= ep(i) Then
            eu(i) = ep(i): el(i) = 0: ed(i) = 0
        Else
            eu(i) = wu(i) + p(i)
            If wl(i) >= c * lm Then
                el(i) = (ep(i) - eu(i)) * wl(i) / lm
                ed(i) = 0
            Else
                If wl(i) >= c * (ep(i) - eu(i)) Then
                    el(i) = c * (ep(i) - eu(i))
                    ed(i) = 0
                Else
                    el(i) = wl(i)
                    If wd(i) >= c * (ep(i) - eu(i)) - el(i) Then
                        ed(i) = c * (ep(i) - eu(i)) - el(i)
                    Else
                        ed(i) = wd(i)
                    End If
                    
                End If
            End If
        End If
        e(i) = eu(i) + el(i) + ed(i)
        pe(i) = p(i) - e(i)
                
'********************产流计算********************
        If pe(i) > 0 Then
            If a(i) + pe(i) < wmm Then
                r(i) = pe(i) + w(i) - wm + wm * (1 - (pe(i) + a(i)) / wmm) ^ (b + 1)
            Else
                r(i) = pe(i) + w(i) - wm
            End If
        Else
            r(i) = 0
        End If
          
'********************后一时刻土壤含水量计算********************
        If wu(i) + p(i) - eu(i) - r(i) <= um Then
            wu(i + 1) = wu(i) + p(i) - eu(i) - r(i)
            wl(i + 1) = wl(i) - el(i)
            wd(i + 1) = wd(i) - ed(i)
        Else
            wu(i + 1) = um
            If wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) <= lm Then
                wl(i + 1) = wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um)
                wd(i + 1) = wd(i) - ed(i)
            Else
                wl(i + 1) = lm
                If wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm <= dm Then
                    wd(i + 1) = wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm
                Else
                    wd(i + 1) = dm
                End If
            End If
        End If
        w(i + 1) = wu(i + 1) + wl(i + 1) + wd(i + 1)
        If w(i + 1) >= wm Then
            a(i + 1) = wmm
        Else
            a(i + 1) = wmm * (1 - (1 - w(i + 1) / wm) ^ (1 / (1 + b)))
        End If
        
        If r(i) > 0 Then
            fr(i + 1) = r(i) / pe(i)
        Else
            fr(i + 1) = 1 - (1 - a(i + 1) / wmm) ^ b
        End If
        
        If fr(i + 1) > 1 Then
            fr(i + 1) = 1
        ElseIf fr(i + 1) < 10 ^ (-4) Then
            fr(i + 1) = 0
        End If

'********************水源划分********************
'不考虑产流面积
'        If r(i) > 0 Then
'            If r(i) / 5 > 1 Then
'                ns = Int(r(i) / 5 + 1)
'                r2 = r(i) / ns
'                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
'                kg2 = ki2 * kg / ki
'                ReDim Preserve s2(ns + 1)
'                ReDim Preserve au2(ns)
'                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
'                s2(1) = s(i)
'                For j = 1 To ns
'                    au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
'                    If au2(j) + r2 >= smm Then
'                        rs2(j) = (r2 + s2(j) - sm)
'                        s2(j) = sm
'                    Else
'                        rs2(j) = (r2 + s2(j) - sm + sm * ((1 - (r2 + au2(j)) / smm) ^ (ex + 1)))
'                        s2(j) = s2(j) + (r2 - rs2(j))
'                    End If
'                    ri2(j) = s2(j) * ki2
'                    rg2(j) = s2(j) * kg2
'                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
'                Next j
'                s(i + 1) = s2(ns + 1)
'                For j = 1 To ns
'                    rs(i) = rs(i) + rs2(j)
'                    ri(i) = ri(i) + ri2(j)
'                    rg(i) = rg(i) + rg2(j)
'                Next j
'            Else
'                au(i) = smm * (1 - (1 - s(i) / sm) ^ (1 / (1 + ex)))
'                If au(i) + r(i) >= smm Then
'                    rs(i) = (r(i) + s(i) - sm)
'                    s(i) = sm
'                Else
'                    rs(i) = (r(i) + s(i) - sm + sm * ((1 - (r(i) + au(i)) / smm) ^ (ex + 1)))
'                    s(i) = s(i) + (r(i) - rs(i))
'                End If
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            End If
'        Else
'            rs(i) = 0
'            If s(i) > 0 Then
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            Else
'                ri(i) = 0
'                rg(i) = 0
'                s(i + 1) = 0
'            End If
'        End If

'变动产流面积
'        if s(i)=1.#INF then
'            s(i) = 0
'        End If
        
        If r(i) > 10 ^ (-4) Then
            If r(i) / 5 > 1 Then
                ns = Int(r(i) / 5 + 1)
                r2 = r(i) / ns
                pe2 = pe(i) / ns
                fr2 = fr(i + 1)
                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
                kg2 = ki2 * kg / ki
                ReDim Preserve s2(ns + 1)
                ReDim Preserve au2(ns)
                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
                s2(1) = s(i) * fr(i) / fr2
                For j = 1 To ns
                    If s2(j) >= sm Then
                        rs2(j) = (pe2 + s2(j) - sm) * fr2
                        s2(j) = sm
                    Else
                        au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
                        If au2(j) + pe2 >= smm Then
                            rs2(j) = (pe2 + s2(j) - sm) * fr2
                            s2(j) = sm
                        Else
                            rs2(j) = (pe2 + s2(j) - sm + sm * ((1 - (pe2 + au2(j)) / smm) ^ (ex + 1))) * fr2
                            s2(j) = s2(j) + (r2 - rs2(j))
                        End If
                    End If
                    ri2(j) = s2(j) * ki2 * fr2
                    rg2(j) = s2(j) * kg2 * fr2
                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
                Next j
                s(i + 1) = s2(ns + 1)
                For j = 1 To ns
                    rs(i) = rs(i) + rs2(j)
                    ri(i) = ri(i) + ri2(j)
                    rg(i) = rg(i) + rg2(j)
                Next j
                If s(i + 1) < 10 ^ (-6) Then
                    s(i + 1) = 0
                End If
            Else
                If s(i) * fr(i) / fr(i + 1) >= sm Then
                    rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                    s(i) = sm
                Else
                    au(i) = smm * (1 - (1 - s(i) * fr(i) / fr(i + 1) / sm) ^ (1 / (1 + ex)))
                    If au(i) + pe(i) >= smm Then
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                        s(i) = sm
                    Else
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm + sm * ((1 - (pe(i) + au(i)) / smm) ^ (ex + 1))) * fr(i + 1)
                        s(i) = s(i) * fr(i) / fr(i + 1) + (r(i) - rs(i)) / fr(i + 1)
                    End If
                    ri(i) = s(i) * ki * fr(i + 1)
                    rg(i) = s(i) * kg * fr(i + 1)
                    s(i + 1) = s(i) * (1 - ki - kg)
                    If s(i + 1) < 10 ^ (-6) Then
                        s(i + 1) = 0
                    End If

                End If
            End If
        Else
            rs(i) = 0
            If s(i) > 0 Then
                If fr(i + 1) > 0 Then
                    s(i) = s(i) * fr(i) / fr(i + 1)
                End If
                    
                ri(i) = s(i) * ki * fr(i + 1)
                rg(i) = s(i) * kg * fr(i + 1)
                s(i + 1) = s(i) * (1 - ki - kg)
                If s(i + 1) < 10 ^ (-6) Then
                    s(i + 1) = 0
                End If
            Else
                ri(i) = 0
                rg(i) = 0
                s(i + 1) = 0
            End If
        End If

        rs(i) = rs(i) + im * data(1, i)
        
'********************汇流计算********************
        qs(i) = qs(i - 1) * cs + (1 - cs) * rs(i)
        qi(i) = qi(i - 1) * ci + (1 - ci) * ri(i)
        qg(i) = qg(i - 1) * cg + (1 - cg) * rg(i)
        qds(i) = qs(i) + qi(i) + qg(i)
        If i < lr Then
            qsim(i) = cr * qsim(i - 1)
        Else
            qsim(i) = cr * qsim(i - 1) + (1 - cr) * qds(i - lr)
        End If
        
        result(1, i) = qsim(i)
        
        result(2, i) = e(i)
        result(3, i) = wu(i + 1)
        result(4, i) = wl(i + 1)
        result(5, i) = wd(i + 1)
        result(6, i) = w(i + 1)
        result(7, i) = s(i + 1) * fr(i + 1)
    Next i
    
End Function

'新安江洪水模型计算程序
Public Function xaj_flood(para() As Single, data() As Single, q_sim() As Single, intial() As Single, u As Single)
    'para: 模型参数（16个变量）
    'data: p，e数据（2维）
    'q_sim: 模拟结果（1维）：出口流量，（单位：m3/s）
    'intial:初始流量、土壤含水量（上、中、下、总）、自由水蓄水量
    'u:径流深与流量换算系数

'********************dim model parameter********************
    Dim kc As Single, um As Integer, lm As Integer, c As Single  '蒸发
    Dim wm As Integer, b As Single, im As Single '产流
    Dim sm As Single, ex As Single, kg As Double, ki As Double '分水源
    Dim cs As Double, ci As Double, cg As Double, cr As Double, lr As Integer     '汇流
    
    Dim dm As Integer '下层蓄水容量
    Dim wmm As Single '流域最大点蓄水容量
    Dim smm As Single  '流域最大点自由水蓄水容量
    
'********************模型中间变量********************
    Dim p() As Single, pan_e() As Single '降水和蒸发资料
    Dim ep() As Single '流域蒸发能力
    Dim eu() As Single, el() As Single, ed() As Single, e() As Single '3层及总实际蒸发
    Dim wu() As Single, wl() As Single, wd() As Single, w() As Single '3层及总的土壤含水量
    Dim pe() As Single '净雨
    Dim a() As Single '蓄水容量曲线纵坐标
    Dim r() As Single ' 总径流量
    Dim fr() As Single '产流面积
    Dim au() As Single '自由水容量曲线纵坐标
    Dim s() As Single '自由水蓄水量
    Dim rs()  As Double, ri() As Double, rg() As Double    '3种径流成分
    Dim qs() As Double, qi() As Double, qg() As Double    '单元出口3种径流汇流
    Dim qds() As Double  '单元出口总径流
    Dim qsim() As Double  '流域总出流
    
    Dim num As Long '总时间步长数
    Dim ns As Integer '自由水蓄水库水量平衡时段划分
    Dim r2 As Single '时段细化后入流量
    Dim pe2 As Single
    Dim fr2 As Single
    Dim s2() As Single '时段细化后自由水蓄水量
    Dim au2() As Single
    Dim ki2 As Double '时段细化后壤中流出流系数
    Dim kg2 As Double '时段细化后地下径流出流系数
    Dim rs2() As Single '时段细化后地表径流
    Dim rg2() As Single '时段细化后地下径流
    Dim ri2() As Single '时段细化后壤中流
    
    Dim i As Long, j As Integer, temp1 As Long
    Dim c1 As Double, c2 As Double
        
'********************read model parameter********************
    kc = para(1): um = para(2): lm = para(3): c = para(4)
    wm = para(5): b = para(6): im = para(7)
    sm = para(8): ex = para(9): kg = para(10): ki = para(11)
    cs = para(12): ci = para(13): cg = para(14): cr = para(15): lr = para(16)
    dm = wm - um - lm
    wmm = wm * (1 + b)
    smm = sm * (1 + ex)
    
    c1 = (1 - (1 - (ki + kg)) ^ (1 / 24)) / (1 + kg / ki)
    c2 = c1 * kg / ki
    ki = c1
    kg = c2
    
    cs = cs ^ (1 / 24)
    ci = ci ^ (1 / 24)
    cg = cg ^ (1 / 24)
    
    num = UBound(data, 2)
    ReDim Preserve p(num), pan_e(num)
    ReDim Preserve ep(num)
    ReDim Preserve eu(num), el(num), ed(num)
    ReDim Preserve e(num)
    ReDim Preserve wu(num + 1), wl(num + 1), wd(num + 1), w(num + 1)
    ReDim Preserve pe(num)
    ReDim Preserve a(num + 1)
    ReDim Preserve r(num)
    ReDim Preserve fr(num + 1)
    ReDim Preserve au(num + 1)
    ReDim Preserve s(num + 1)
    ReDim Preserve rs(num), ri(num), rg(num)
    ReDim Preserve qs(num), qi(num), qg(num)
    ReDim Preserve qds(num), qsim(num)
        
'********************初始土壤状态设定********************
    If intial(5) > wm Then
        intial(2) = um
        intial(3) = lm
        intial(4) = dm
        intial(5) = wm
    ElseIf intial(5) <= 0 Then
        intial(2) = um * 0.7
        intial(3) = lm * 0.7
        intial(4) = dm * 0.7
        intial(5) = wm * 0.7
    End If
    wu(1) = intial(2)
    wl(1) = intial(3)
    wd(1) = intial(4)
    w(1) = intial(5)
    a(1) = wmm * (1 - (1 - w(1) / wm) ^ (1 / (1 + b)))
    fr(1) = 1 - (1 - a(1) / wmm) ^ b
    s(1) = intial(6) / fr(1)
    If s(1) < sm Then
        au(1) = smm * (1 - (1 - s(1) / sm) ^ (1 / (1 + ex)))
    Else
        au(1) = smm
    End If
    
    qsim(0) = intial(1)
    qg(0) = qsim(0) / u

    
    For i = 1 To num
        p(i) = data(1, i) * (1 - im)
        pan_e(i) = data(2, i)
        
'********************蒸散发计算********************
        ep(i) = kc * pan_e(i)
        If wu(i) + p(i) >= ep(i) Then
            eu(i) = ep(i): el(i) = 0: ed(i) = 0
        Else
            eu(i) = wu(i) + p(i)
            If wl(i) >= c * lm Then
                el(i) = (ep(i) - eu(i)) * wl(i) / lm
                ed(i) = 0
            Else
                If wl(i) >= c * (ep(i) - eu(i)) Then
                    el(i) = c * (ep(i) - eu(i))
                    ed(i) = 0
                Else
                    el(i) = wl(i)
                    ed(i) = c * (ep(i) - eu(i)) - el(i)
                End If
            End If
        End If
        e(i) = eu(i) + el(i) + ed(i)
        pe(i) = p(i) - e(i)
                
'********************产流计算********************
        If pe(i) > 0 Then
            If a(i) + pe(i) < wmm Then
                r(i) = pe(i) + w(i) - wm + wm * (1 - (pe(i) + a(i)) / wmm) ^ (b + 1)
            Else
                r(i) = pe(i) + w(i) - wm
            End If
        Else
            r(i) = 0
        End If
          
'********************后一时刻土壤含水量计算********************
        If wu(i) + p(i) - eu(i) - r(i) <= um Then
            wu(i + 1) = wu(i) + p(i) - eu(i) - r(i)
            wl(i + 1) = wl(i) - el(i)
            wd(i + 1) = wd(i) - ed(i)
        Else
            wu(i + 1) = um
            If wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) <= lm Then
                wl(i + 1) = wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um)
                wd(i + 1) = wd(i) - ed(i)
            Else
                wl(i + 1) = lm
                If wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm <= dm Then
                    wd(i + 1) = wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm
                Else
                    wd(i + 1) = dm
                End If
            End If
        End If
        w(i + 1) = wu(i + 1) + wl(i + 1) + wd(i + 1)
        If w(i + 1) >= wm Then
            a(i + 1) = wmm
        Else
            a(i + 1) = wmm * (1 - (1 - w(i + 1) / wm) ^ (1 / (1 + b)))
        End If
        
        If r(i) > 0 Then
            fr(i + 1) = r(i) / pe(i)
        Else
            fr(i + 1) = 1 - (1 - a(i + 1) / wmm) ^ b
        End If
        
        If fr(i + 1) > 1 Then
            fr(i + 1) = 1
        End If

'********************水源划分********************
'不考虑产流面积
'        If r(i) > 0 Then
'            If r(i) / 5 > 1 Then
'                ns = Int(r(i) / 5 + 1)
'                r2 = r(i) / ns
'                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
'                kg2 = ki2 * kg / ki
'                ReDim Preserve s2(ns + 1)
'                ReDim Preserve au2(ns)
'                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
'                s2(1) = s(i)
'                For j = 1 To ns
'                    au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
'                    If au2(j) + r2 >= smm Then
'                        rs2(j) = (r2 + s2(j) - sm)
'                        s2(j) = sm
'                    Else
'                        rs2(j) = (r2 + s2(j) - sm + sm * ((1 - (r2 + au2(j)) / smm) ^ (ex + 1)))
'                        s2(j) = s2(j) + (r2 - rs2(j))
'                    End If
'                    ri2(j) = s2(j) * ki2
'                    rg2(j) = s2(j) * kg2
'                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
'                Next j
'                s(i + 1) = s2(ns + 1)
'                For j = 1 To ns
'                    rs(i) = rs(i) + rs2(j)
'                    ri(i) = ri(i) + ri2(j)
'                    rg(i) = rg(i) + rg2(j)
'                Next j
'            Else
'                au(i) = smm * (1 - (1 - s(i) / sm) ^ (1 / (1 + ex)))
'                If au(i) + r(i) >= smm Then
'                    rs(i) = (r(i) + s(i) - sm)
'                    s(i) = sm
'                Else
'                    rs(i) = (r(i) + s(i) - sm + sm * ((1 - (r(i) + au(i)) / smm) ^ (ex + 1)))
'                    s(i) = s(i) + (r(i) - rs(i))
'                End If
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            End If
'        Else
'            rs(i) = 0
'            If s(i) > 0 Then
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            Else
'                ri(i) = 0
'                rg(i) = 0
'                s(i + 1) = 0
'            End If
'        End If

'变动产流面积
        If r(i) > 0 Then
            If r(i) / 5 > 1 Then
                ns = Int(r(i) / 5 + 1)
                r2 = r(i) / ns
                pe2 = pe(i) / ns
                fr2 = fr(i + 1)
                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
                kg2 = ki2 * kg / ki
                ReDim Preserve s2(ns + 1)
                ReDim Preserve au2(ns)
                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
                s2(1) = s(i) * fr(i) / fr2
                For j = 1 To ns
                    If s2(j) >= sm Then
                        rs2(j) = (pe2 + s2(j) - sm) * fr2
                        s2(j) = sm
                    Else
                        au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
                        If au2(j) + pe2 >= smm Then
                            rs2(j) = (pe2 + s2(j) - sm) * fr2
                            s2(j) = sm
                        Else
                            rs2(j) = (pe2 + s2(j) - sm + sm * ((1 - (pe2 + au2(j)) / smm) ^ (ex + 1))) * fr2
                            s2(j) = s2(j) + (r2 - rs2(j))
                        End If
                    End If
                    ri2(j) = s2(j) * ki2 * fr2
                    rg2(j) = s2(j) * kg2 * fr2
                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
                Next j
                s(i + 1) = s2(ns + 1)
                For j = 1 To ns
                    rs(i) = rs(i) + rs2(j)
                    ri(i) = ri(i) + ri2(j)
                    rg(i) = rg(i) + rg2(j)
                Next j
            Else
                If s(i) * fr(i) / fr(i + 1) >= sm Then
                    rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                    s(i) = sm
                Else
                    au(i) = smm * (1 - (1 - s(i) * fr(i) / fr(i + 1) / sm) ^ (1 / (1 + ex)))
                    If au(i) + pe(i) >= smm Then
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                        s(i) = sm
                    Else
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm + sm * ((1 - (pe(i) + au(i)) / smm) ^ (ex + 1))) * fr(i + 1)
                        s(i) = s(i) * fr(i) / fr(i + 1) + (r(i) - rs(i)) / fr(i + 1)
                    End If
                    ri(i) = s(i) * ki * fr(i + 1)
                    rg(i) = s(i) * kg * fr(i + 1)
                    s(i + 1) = s(i) * (1 - ki - kg)
                End If
            End If
        Else
            rs(i) = 0
            If s(i) > 0 Then
                s(i) = s(i) * fr(i) / fr(i + 1)
                ri(i) = s(i) * ki * fr(i + 1)
                rg(i) = s(i) * kg * fr(i + 1)
                s(i + 1) = s(i) * (1 - ki - kg)
            Else
                ri(i) = 0
                rg(i) = 0
                s(i + 1) = 0
            End If
        End If

        rs(i) = rs(i) + im * data(1, i)
        
'********************汇流计算********************
        qs(i) = qs(i - 1) * cs + (1 - cs) * rs(i)
        qi(i) = qi(i - 1) * ci + (1 - ci) * ri(i)
        qg(i) = qg(i - 1) * cg + (1 - cg) * rg(i)
        qds(i) = (qs(i) + qi(i) + qg(i)) * u
        If i <= lr Then
            qsim(i) = cr * qsim(i - 1) + (1 - cr) * qg(0) / (cg ^ (lr - i)) * u
        Else
            qsim(i) = cr * qsim(i - 1) + (1 - cr) * qds(i - lr)
        End If
        q_sim(i) = qsim(i)
    Next i
    
End Function

'新安江洪水模型计算程序-nash瞬时单位线汇流
Public Function xaj_flood_uh(para() As Single, data() As Single, q_sim() As Single, intial() As Single, u As Single)
    'para: 模型参数（16个变量）
    'data: p，e数据（2维）
    'q_sim: 模拟结果（1维）：出口流量，（单位：m3/s）
    'intial:初始流量、土壤含水量（上、中、下、总）、自由水蓄水量
    'u:径流深与流量换算系数

'********************dim model parameter********************
    Dim kc As Single, um As Integer, lm As Integer, c As Single  '蒸发
    Dim wm As Integer, b As Single, im As Single '产流
    Dim sm As Single, ex As Single, kg As Double, ki As Double '分水源
    Dim cs As Single, ci As Single, cg As Single, cr As Single, lr As Integer '汇流
    
    Dim nn As Single, kk As Single 'nash瞬时单位线汇流两个参数
    
    
    Dim dm As Integer '下层蓄水容量
    Dim wmm As Single '流域最大点蓄水容量
    Dim smm As Single  '流域最大点自由水蓄水容量
    
'********************模型中间变量********************
    Dim p() As Single, pan_e() As Single '降水和蒸发资料
    Dim ep() As Single '流域蒸发能力
    Dim eu() As Single, el() As Single, ed() As Single, e() As Single '3层及总实际蒸发
    Dim wu() As Single, wl() As Single, wd() As Single, w() As Single '3层及总的土壤含水量
    Dim pe() As Single '净雨
    Dim a() As Single '蓄水容量曲线纵坐标
    Dim r() As Single ' 总径流量
    Dim fr() As Single '产流面积
    Dim au() As Single '自由水容量曲线纵坐标
    Dim s() As Single '自由水蓄水量
    Dim rs() As Single, ri() As Single, rg() As Single '3种径流成分
    Dim qs() As Single, qi() As Single, qg() As Single '单元出口3种径流汇流
    Dim qds() As Single '单元出口总径流
    Dim qsim() As Single '流域总出流
    
    Dim num As Long '总时间步长数
    Dim ns As Integer '自由水蓄水库水量平衡时段划分
    Dim r2 As Single '时段细化后入流量
    Dim pe2 As Single
    Dim fr2 As Single
    Dim s2() As Single '时段细化后自由水蓄水量
    Dim au2() As Single
    Dim ki2 As Double '时段细化后壤中流出流系数
    Dim kg2 As Double '时段细化后地下径流出流系数
    Dim rs2() As Single '时段细化后地表径流
    Dim rg2() As Single '时段细化后地下径流
    Dim ri2() As Single '时段细化后壤中流
    
    Dim i As Long, j As Integer, temp1 As Long
    Dim c1 As Double, c2 As Double
        
'********************read model parameter********************
    kc = para(1): um = para(2): lm = para(3): c = para(4)
    wm = para(5): b = para(6): im = para(7)
    sm = para(8): ex = para(9): kg = para(10): ki = para(11)
    cs = para(12): ci = para(13): cg = para(14): cr = para(15): lr = para(16)
    
    nn = para(15): kk = para(16)
    
    dm = wm - um - lm
    wmm = wm * (1 + b)
    smm = sm * (1 + ex)
    
    c1 = (1 - (1 - (ki + kg)) ^ (1 / 24)) / (1 + kg / ki)
    c2 = c1 * kg / ki
    ki = c1
    kg = c2
    
    cs = cs ^ (1 / 24)
    ci = ci ^ (1 / 24)
    cg = cg ^ (1 / 24)
    
    num = UBound(data, 2)
    ReDim Preserve p(num), pan_e(num)
    ReDim Preserve ep(num)
    ReDim Preserve eu(num), el(num), ed(num)
    ReDim Preserve e(num)
    ReDim Preserve wu(num + 1), wl(num + 1), wd(num + 1), w(num + 1)
    ReDim Preserve pe(num)
    ReDim Preserve a(num + 1)
    ReDim Preserve r(num)
    ReDim Preserve fr(num + 1)
    ReDim Preserve au(num + 1)
    ReDim Preserve s(num + 1)
    ReDim Preserve rs(num), ri(num), rg(num)
    ReDim Preserve qs(num), qi(num), qg(num)
    ReDim Preserve qds(num), qsim(num)
        
'********************初始土壤状态设定********************
    If intial(5) > wm Then
        intial(2) = um
        intial(3) = lm
        intial(4) = dm
        intial(5) = wm
    ElseIf intial(5) <= 0 Then
        intial(2) = um * 0.7
        intial(3) = lm * 0.7
        intial(4) = dm * 0.7
        intial(5) = wm * 0.7
    End If
    wu(1) = intial(2)
    wl(1) = intial(3)
    wd(1) = intial(4)
    w(1) = intial(5)
    a(1) = wmm * (1 - (1 - w(1) / wm) ^ (1 / (1 + b)))
    fr(1) = 1 - (1 - a(1) / wmm) ^ b
    s(1) = intial(6) / fr(1)
    If s(1) < sm Then
        au(1) = smm * (1 - (1 - s(1) / sm) ^ (1 / (1 + ex)))
    Else
        au(1) = smm
    End If
    
    qsim(0) = intial(1)
    qg(0) = qsim(0) / u

    
    For i = 1 To num
        p(i) = data(1, i) * (1 - im)
        pan_e(i) = data(2, i)
        
'********************蒸散发计算********************
        ep(i) = kc * pan_e(i)
        If wu(i) + p(i) >= ep(i) Then
            eu(i) = ep(i): el(i) = 0: ed(i) = 0
        Else
            eu(i) = wu(i) + p(i)
            If wl(i) >= c * lm Then
                el(i) = (ep(i) - eu(i)) * wl(i) / lm
                ed(i) = 0
            Else
                If wl(i) >= c * (ep(i) - eu(i)) Then
                    el(i) = c * (ep(i) - eu(i))
                    ed(i) = 0
                Else
                    el(i) = wl(i)
                    ed(i) = c * (ep(i) - eu(i)) - el(i)
                End If
            End If
        End If
        e(i) = eu(i) + el(i) + ed(i)
        pe(i) = p(i) - e(i)
                
'********************产流计算********************
        If pe(i) > 0 Then
            If a(i) + pe(i) < wmm Then
                r(i) = pe(i) + w(i) - wm + wm * (1 - (pe(i) + a(i)) / wmm) ^ (b + 1)
            Else
                r(i) = pe(i) + w(i) - wm
            End If
        Else
            r(i) = 0
        End If
          
'********************后一时刻土壤含水量计算********************
        If wu(i) + p(i) - eu(i) - r(i) <= um Then
            wu(i + 1) = wu(i) + p(i) - eu(i) - r(i)
            wl(i + 1) = wl(i) - el(i)
            wd(i + 1) = wd(i) - ed(i)
        Else
            wu(i + 1) = um
            If wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) <= lm Then
                wl(i + 1) = wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um)
                wd(i + 1) = wd(i) - ed(i)
            Else
                wl(i + 1) = lm
                If wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm <= dm Then
                    wd(i + 1) = wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm
                Else
                    wd(i + 1) = dm
                End If
            End If
        End If
        w(i + 1) = wu(i + 1) + wl(i + 1) + wd(i + 1)
        If w(i + 1) >= wm Then
            a(i + 1) = wmm
        Else
            a(i + 1) = wmm * (1 - (1 - w(i + 1) / wm) ^ (1 / (1 + b)))
        End If
        
        If r(i) > 0 Then
            fr(i + 1) = r(i) / pe(i)
        Else
            fr(i + 1) = 1 - (1 - a(i + 1) / wmm) ^ b
        End If
        
        If fr(i + 1) > 1 Then
            fr(i + 1) = 1
        End If

'********************水源划分********************
'不考虑产流面积
'        If r(i) > 0 Then
'            If r(i) / 5 > 1 Then
'                ns = Int(r(i) / 5 + 1)
'                r2 = r(i) / ns
'                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
'                kg2 = ki2 * kg / ki
'                ReDim Preserve s2(ns + 1)
'                ReDim Preserve au2(ns)
'                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
'                s2(1) = s(i)
'                For j = 1 To ns
'                    au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
'                    If au2(j) + r2 >= smm Then
'                        rs2(j) = (r2 + s2(j) - sm)
'                        s2(j) = sm
'                    Else
'                        rs2(j) = (r2 + s2(j) - sm + sm * ((1 - (r2 + au2(j)) / smm) ^ (ex + 1)))
'                        s2(j) = s2(j) + (r2 - rs2(j))
'                    End If
'                    ri2(j) = s2(j) * ki2
'                    rg2(j) = s2(j) * kg2
'                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
'                Next j
'                s(i + 1) = s2(ns + 1)
'                For j = 1 To ns
'                    rs(i) = rs(i) + rs2(j)
'                    ri(i) = ri(i) + ri2(j)
'                    rg(i) = rg(i) + rg2(j)
'                Next j
'            Else
'                au(i) = smm * (1 - (1 - s(i) / sm) ^ (1 / (1 + ex)))
'                If au(i) + r(i) >= smm Then
'                    rs(i) = (r(i) + s(i) - sm)
'                    s(i) = sm
'                Else
'                    rs(i) = (r(i) + s(i) - sm + sm * ((1 - (r(i) + au(i)) / smm) ^ (ex + 1)))
'                    s(i) = s(i) + (r(i) - rs(i))
'                End If
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            End If
'        Else
'            rs(i) = 0
'            If s(i) > 0 Then
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            Else
'                ri(i) = 0
'                rg(i) = 0
'                s(i + 1) = 0
'            End If
'        End If

'变动产流面积
        If r(i) > 0 Then
            If r(i) / 5 > 1 Then
                ns = Int(r(i) / 5 + 1)
                r2 = r(i) / ns
                pe2 = pe(i) / ns
                fr2 = fr(i + 1)
                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
                kg2 = ki2 * kg / ki
                ReDim Preserve s2(ns + 1)
                ReDim Preserve au2(ns)
                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
                s2(1) = s(i) * fr(i) / fr2
                For j = 1 To ns
                    If s2(j) >= sm Then
                        rs2(j) = (pe2 + s2(j) - sm) * fr2
                        s2(j) = sm
                    Else
                        au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
                        If au2(j) + pe2 >= smm Then
                            rs2(j) = (pe2 + s2(j) - sm) * fr2
                            s2(j) = sm
                        Else
                            rs2(j) = (pe2 + s2(j) - sm + sm * ((1 - (pe2 + au2(j)) / smm) ^ (ex + 1))) * fr2
                            s2(j) = s2(j) + (r2 - rs2(j))
                        End If
                    End If
                    ri2(j) = s2(j) * ki2 * fr2
                    rg2(j) = s2(j) * kg2 * fr2
                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
                Next j
                s(i + 1) = s2(ns + 1)
                For j = 1 To ns
                    rs(i) = rs(i) + rs2(j)
                    ri(i) = ri(i) + ri2(j)
                    rg(i) = rg(i) + rg2(j)
                Next j
            Else
                If s(i) * fr(i) / fr(i + 1) >= sm Then
                    rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                    s(i) = sm
                Else
                    au(i) = smm * (1 - (1 - s(i) * fr(i) / fr(i + 1) / sm) ^ (1 / (1 + ex)))
                    If au(i) + pe(i) >= smm Then
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                        s(i) = sm
                    Else
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm + sm * ((1 - (pe(i) + au(i)) / smm) ^ (ex + 1))) * fr(i + 1)
                        s(i) = s(i) * fr(i) / fr(i + 1) + (r(i) - rs(i)) / fr(i + 1)
                    End If
                    ri(i) = s(i) * ki * fr(i + 1)
                    rg(i) = s(i) * kg * fr(i + 1)
                    s(i + 1) = s(i) * (1 - ki - kg)
                End If
            End If
        Else
            rs(i) = 0
            If s(i) > 0 Then
                s(i) = s(i) * fr(i) / fr(i + 1)
                ri(i) = s(i) * ki * fr(i + 1)
                rg(i) = s(i) * kg * fr(i + 1)
                s(i + 1) = s(i) * (1 - ki - kg)
            Else
                ri(i) = 0
                rg(i) = 0
                s(i + 1) = 0
            End If
        End If

        rs(i) = rs(i) + im * data(1, i)
        
'********************汇流计算********************
        qs(i) = qs(i - 1) * cs + (1 - cs) * rs(i)
        qi(i) = qi(i - 1) * ci + (1 - ci) * ri(i)
        qg(i) = qg(i - 1) * cg + (1 - cg) * rg(i)
        qds(i) = (qs(i) + qi(i) + qg(i)) * u
        If i <= lr Then
            qsim(i) = cr * qsim(i - 1) + (1 - cr) * qg(0) / (cg ^ (lr - i)) * u
        Else
            qsim(i) = cr * qsim(i - 1) + (1 - cr) * qds(i - lr)
        End If
        q_sim(i) = qsim(i)
    Next i
    
End Function


'新安江洪水模型计算程序-30分钟
Public Function xaj_flood_min(para() As Single, data() As Single, q_sim() As Single, intial() As Single, u As Single)
    'para: 模型参数（16个变量）
    'data: p，e数据（2维）
    'q_sim: 模拟结果（1维）：出口流量，（单位：m3/s）
    'intial:初始流量、土壤含水量（上、中、下、总）、自由水蓄水量
    'u:径流深与流量换算系数

'********************dim model parameter********************
    Dim kc As Single, um As Integer, lm As Integer, c As Single  '蒸发
    Dim wm As Integer, b As Single, im As Single '产流
    Dim sm As Single, ex As Single, kg As Double, ki As Double '分水源
    Dim cs As Single, ci As Single, cg As Single, cr As Single, lr As Integer '汇流
    
    Dim dm As Integer '下层蓄水容量
    Dim wmm As Single '流域最大点蓄水容量
    Dim smm As Single  '流域最大点自由水蓄水容量
    
'********************模型中间变量********************
    Dim p() As Single, pan_e() As Single '降水和蒸发资料
    Dim ep() As Single '流域蒸发能力
    Dim eu() As Single, el() As Single, ed() As Single, e() As Single '3层及总实际蒸发
    Dim wu() As Single, wl() As Single, wd() As Single, w() As Single '3层及总的土壤含水量
    Dim pe() As Single '净雨
    Dim a() As Single '蓄水容量曲线纵坐标
    Dim r() As Single ' 总径流量
    Dim fr() As Single '产流面积
    Dim au() As Single '自由水容量曲线纵坐标
    Dim s() As Single '自由水蓄水量
    Dim rs() As Single, ri() As Single, rg() As Single '3种径流成分
    Dim qs() As Single, qi() As Single, qg() As Single '单元出口3种径流汇流
    Dim qds() As Single '单元出口总径流
    Dim qsim() As Single '流域总出流
    
    Dim num As Long '总时间步长数
    Dim ns As Integer '自由水蓄水库水量平衡时段划分
    Dim r2 As Single '时段细化后入流量
    Dim pe2 As Single
    Dim fr2 As Single
    Dim s2() As Single '时段细化后自由水蓄水量
    Dim au2() As Single
    Dim ki2 As Double '时段细化后壤中流出流系数
    Dim kg2 As Double '时段细化后地下径流出流系数
    Dim rs2() As Single '时段细化后地表径流
    Dim rg2() As Single '时段细化后地下径流
    Dim ri2() As Single '时段细化后壤中流
    
    Dim i As Long, j As Integer, temp1 As Long
    Dim c1 As Double, c2 As Double
        
'********************read model parameter********************
    kc = para(1): um = para(2): lm = para(3): c = para(4)
    wm = para(5): b = para(6): im = para(7)
    sm = para(8): ex = para(9): kg = para(10): ki = para(11)
    cs = para(12): ci = para(13): cg = para(14): cr = para(15): lr = para(16)
    dm = wm - um - lm
    wmm = wm * (1 + b)
    smm = sm * (1 + ex)
    
    c1 = (1 - (1 - (ki + kg)) ^ (1 / 48)) / (1 + kg / ki)
    c2 = c1 * kg / ki
    ki = c1
    kg = c2
    
    cs = cs ^ (1 / 48)
    ci = ci ^ (1 / 48)
    cg = cg ^ (1 / 48)
    
    num = UBound(data, 2)
    ReDim Preserve p(num), pan_e(num)
    ReDim Preserve ep(num)
    ReDim Preserve eu(num), el(num), ed(num)
    ReDim Preserve e(num)
    ReDim Preserve wu(num + 1), wl(num + 1), wd(num + 1), w(num + 1)
    ReDim Preserve pe(num)
    ReDim Preserve a(num + 1)
    ReDim Preserve r(num)
    ReDim Preserve fr(num + 1)
    ReDim Preserve au(num + 1)
    ReDim Preserve s(num + 1)
    ReDim Preserve rs(num), ri(num), rg(num)
    ReDim Preserve qs(num), qi(num), qg(num)
    ReDim Preserve qds(num), qsim(num)
        
'********************初始土壤状态设定********************
    If intial(5) > wm Then
        intial(2) = um
        intial(3) = lm
        intial(4) = dm
        intial(5) = wm
    ElseIf intial(5) <= 0 Then
        intial(2) = um * 0.7
        intial(3) = lm * 0.7
        intial(4) = dm * 0.7
        intial(5) = wm * 0.7
    End If
    wu(1) = intial(2)
    wl(1) = intial(3)
    wd(1) = intial(4)
    w(1) = intial(5)
    a(1) = wmm * (1 - (1 - w(1) / wm) ^ (1 / (1 + b)))
    fr(1) = 1 - (1 - a(1) / wmm) ^ b
    s(1) = intial(6) / fr(1)
    If s(1) < sm Then
        au(1) = smm * (1 - (1 - s(1) / sm) ^ (1 / (1 + ex)))
    Else
        au(1) = smm
    End If
    
    qsim(0) = intial(1)
    qg(0) = qsim(0) / u

    
    For i = 1 To num
        p(i) = data(1, i) * (1 - im)
        pan_e(i) = data(2, i)
        
'********************蒸散发计算********************
        ep(i) = kc * pan_e(i)
        If wu(i) + p(i) >= ep(i) Then
            eu(i) = ep(i): el(i) = 0: ed(i) = 0
        Else
            eu(i) = wu(i) + p(i)
            If wl(i) >= c * lm Then
                el(i) = (ep(i) - eu(i)) * wl(i) / lm
                ed(i) = 0
            Else
                If wl(i) >= c * (ep(i) - eu(i)) Then
                    el(i) = c * (ep(i) - eu(i))
                    ed(i) = 0
                Else
                    el(i) = wl(i)
                    ed(i) = c * (ep(i) - eu(i)) - el(i)
                End If
            End If
        End If
        e(i) = eu(i) + el(i) + ed(i)
        pe(i) = p(i) - e(i)
                
'********************产流计算********************
        If pe(i) > 0 Then
            If a(i) + pe(i) < wmm Then
                r(i) = pe(i) + w(i) - wm + wm * (1 - (pe(i) + a(i)) / wmm) ^ (b + 1)
            Else
                r(i) = pe(i) + w(i) - wm
            End If
        Else
            r(i) = 0
        End If
          
'********************后一时刻土壤含水量计算********************
        If wu(i) + p(i) - eu(i) - r(i) <= um Then
            wu(i + 1) = wu(i) + p(i) - eu(i) - r(i)
            wl(i + 1) = wl(i) - el(i)
            wd(i + 1) = wd(i) - ed(i)
        Else
            wu(i + 1) = um
            If wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) <= lm Then
                wl(i + 1) = wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um)
                wd(i + 1) = wd(i) - ed(i)
            Else
                wl(i + 1) = lm
                If wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm <= dm Then
                    wd(i + 1) = wd(i) - ed(i) + wl(i) - el(i) + (wu(i) + p(i) - eu(i) - r(i) - um) - lm
                Else
                    wd(i + 1) = dm
                End If
            End If
        End If
        w(i + 1) = wu(i + 1) + wl(i + 1) + wd(i + 1)
        If w(i + 1) >= wm Then
            a(i + 1) = wmm
        Else
            a(i + 1) = wmm * (1 - (1 - w(i + 1) / wm) ^ (1 / (1 + b)))
        End If
        
        If r(i) > 0 Then
            fr(i + 1) = r(i) / pe(i)
        Else
            fr(i + 1) = 1 - (1 - a(i + 1) / wmm) ^ b
        End If
        
        If fr(i + 1) > 1 Then
            fr(i + 1) = 1
        End If

'********************水源划分********************
'不考虑产流面积
'        If r(i) > 0 Then
'            If r(i) / 5 > 1 Then
'                ns = Int(r(i) / 5 + 1)
'                r2 = r(i) / ns
'                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
'                kg2 = ki2 * kg / ki
'                ReDim Preserve s2(ns + 1)
'                ReDim Preserve au2(ns)
'                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
'                s2(1) = s(i)
'                For j = 1 To ns
'                    au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
'                    If au2(j) + r2 >= smm Then
'                        rs2(j) = (r2 + s2(j) - sm)
'                        s2(j) = sm
'                    Else
'                        rs2(j) = (r2 + s2(j) - sm + sm * ((1 - (r2 + au2(j)) / smm) ^ (ex + 1)))
'                        s2(j) = s2(j) + (r2 - rs2(j))
'                    End If
'                    ri2(j) = s2(j) * ki2
'                    rg2(j) = s2(j) * kg2
'                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
'                Next j
'                s(i + 1) = s2(ns + 1)
'                For j = 1 To ns
'                    rs(i) = rs(i) + rs2(j)
'                    ri(i) = ri(i) + ri2(j)
'                    rg(i) = rg(i) + rg2(j)
'                Next j
'            Else
'                au(i) = smm * (1 - (1 - s(i) / sm) ^ (1 / (1 + ex)))
'                If au(i) + r(i) >= smm Then
'                    rs(i) = (r(i) + s(i) - sm)
'                    s(i) = sm
'                Else
'                    rs(i) = (r(i) + s(i) - sm + sm * ((1 - (r(i) + au(i)) / smm) ^ (ex + 1)))
'                    s(i) = s(i) + (r(i) - rs(i))
'                End If
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            End If
'        Else
'            rs(i) = 0
'            If s(i) > 0 Then
'                ri(i) = s(i) * ki
'                rg(i) = s(i) * kg
'                s(i + 1) = s(i) * (1 - ki - kg)
'            Else
'                ri(i) = 0
'                rg(i) = 0
'                s(i + 1) = 0
'            End If
'        End If

'变动产流面积
        If r(i) > 0 Then
            If r(i) / 5 > 1 Then
                ns = Int(r(i) / 5 + 1)
                r2 = r(i) / ns
                pe2 = pe(i) / ns
                fr2 = fr(i + 1)
                ki2 = (1 - (1 - (ki + kg)) ^ (1 / ns)) / (1 + kg / ki)
                kg2 = ki2 * kg / ki
                ReDim Preserve s2(ns + 1)
                ReDim Preserve au2(ns)
                ReDim Preserve rs2(ns), ri2(ns), rg2(ns)
                s2(1) = s(i) * fr(i) / fr2
                For j = 1 To ns
                    If s2(j) >= sm Then
                        rs2(j) = (pe2 + s2(j) - sm) * fr2
                        s2(j) = sm
                    Else
                        au2(j) = smm * (1 - (1 - s2(j) / sm) ^ (1 / (1 + ex)))
                        If au2(j) + pe2 >= smm Then
                            rs2(j) = (pe2 + s2(j) - sm) * fr2
                            s2(j) = sm
                        Else
                            rs2(j) = (pe2 + s2(j) - sm + sm * ((1 - (pe2 + au2(j)) / smm) ^ (ex + 1))) * fr2
                            s2(j) = s2(j) + (r2 - rs2(j))
                        End If
                    End If
                    ri2(j) = s2(j) * ki2 * fr2
                    rg2(j) = s2(j) * kg2 * fr2
                    s2(j + 1) = s2(j) * (1 - ki2 - kg2)
                Next j
                s(i + 1) = s2(ns + 1)
                For j = 1 To ns
                    rs(i) = rs(i) + rs2(j)
                    ri(i) = ri(i) + ri2(j)
                    rg(i) = rg(i) + rg2(j)
                Next j
            Else
                If s(i) * fr(i) / fr(i + 1) >= sm Then
                    rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                    s(i) = sm
                Else
                    au(i) = smm * (1 - (1 - s(i) * fr(i) / fr(i + 1) / sm) ^ (1 / (1 + ex)))
                    If au(i) + pe(i) >= smm Then
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm) * fr(i + 1)
                        s(i) = sm
                    Else
                        rs(i) = (pe(i) + s(i) * fr(i) / fr(i + 1) - sm + sm * ((1 - (pe(i) + au(i)) / smm) ^ (ex + 1))) * fr(i + 1)
                        s(i) = s(i) * fr(i) / fr(i + 1) + (r(i) - rs(i)) / fr(i + 1)
                    End If
                    ri(i) = s(i) * ki * fr(i + 1)
                    rg(i) = s(i) * kg * fr(i + 1)
                    s(i + 1) = s(i) * (1 - ki - kg)
                End If
            End If
        Else
            rs(i) = 0
            If s(i) > 0 Then
                s(i) = s(i) * fr(i) / fr(i + 1)
                ri(i) = s(i) * ki * fr(i + 1)
                rg(i) = s(i) * kg * fr(i + 1)
                s(i + 1) = s(i) * (1 - ki - kg)
            Else
                ri(i) = 0
                rg(i) = 0
                s(i + 1) = 0
            End If
        End If

        rs(i) = rs(i) + im * data(1, i)
        
'********************汇流计算********************
        qs(i) = qs(i - 1) * cs + (1 - cs) * rs(i)
        qi(i) = qi(i - 1) * ci + (1 - ci) * ri(i)
        qg(i) = qg(i - 1) * cg + (1 - cg) * rg(i)
        qds(i) = (qs(i) + qi(i) + qg(i)) * u
        If i <= lr Then
            qsim(i) = cr * qsim(i - 1) + (1 - cr) * qg(0) / (cg ^ (lr - i)) * u
        Else
            qsim(i) = cr * qsim(i - 1) + (1 - cr) * qds(i - lr)
        End If
        q_sim(i) = qsim(i)
    Next i
    
End Function
