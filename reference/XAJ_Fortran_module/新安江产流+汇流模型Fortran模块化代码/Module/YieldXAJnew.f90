!+++++++++++++++三水源新安江产流模型中的产流模块++++++++++++++++++++   
!-------------输入：净雨量PE
!------------输入： 时段初上层、下层、深层张力水蓄量WU,WL,WD,总蓄量W
!-------------输出：总产流量R， 时变上层、下层、深层张力水蓄量WU,WL,WD,总蓄量W
!------------参数：不透水面积IM,张力水蓄水容量WM,张力水蓄水容量曲线的方次B,张力水蓄水容量WM,上层蓄水容量WUM,下层蓄水容量WLM
!-----------中间变量：WMM,A 
    Subroutine YieldXAJnew(PE,W,WU,WL,WD,WM,WUM,WLM,R,B,IM) 
        Implicit none
        Real  PE
        Real W,WU,WL,WD
        Real WM,WUM,WLM,WDM
        Real R  
        Real B,IM
!=========中间变量=====        
        Real WMM,A 
        Real RRI 
 
            WDM=WM-WUM-WLM
            WMM=(1.0+B)*WM/(1.0-IM)         
    
            if (PE .le. 0.0) Then         !无径雨量，不产流
                R= 0.0
            else
                A = WMM * (1.0 - (1.0 - W / WM) ** (1.0 / (1.0 + B)))               !计算A值
                R = 0.0            
                A = A + PE                     
                R = PE - WM + W
                If (A .lt.  WMM) Then
                    R = R + WM * (1.0 - A / WMM) ** (1.0 + B)
                Endif
 
                if ((WU + PE - R) .le. WUM) Then 
                    WU=WU+PE-R
                    WL=WL
                    WD=WD
                else
                    WU=WUM
                    if ((WU+WL+PE-R-WUM) .ge. WLM) then             
                        WL=WLM
                        WD=W+PE-R-WU-WL
                        if (WD .gt. WDM) then
                            WD=WDM
                        endif
                    else
                        WL=WU+WL+PE-R-WUM
                        WD=WD
                    endif
                    
                endif
            endif
            W=WU+WL+WD
            
            Return
  
    End Subroutine YieldXAJnew
    
    

    
    
