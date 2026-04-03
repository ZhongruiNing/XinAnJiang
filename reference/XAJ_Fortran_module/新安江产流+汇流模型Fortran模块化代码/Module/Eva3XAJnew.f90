!++++++++++++++三水源新安江产流模型中的蒸散发模块+++++++++++++++
!--------输入：降雨P,水面蒸发能力EPW
!--------输入：前一时刻上层土壤含水量WU,下层土壤水WL和下层土壤水WD,总土壤含水量W
!--------参数：下层张力水容量WLM
!--------参数：蒸散发折算系数K，深层蒸散发系数 C
!--------输出：上、下、深三层时变的流域蒸散发量EU、EL、ED，总蒸发量ET  
!--------输出:  上、下、深三层时变的张力水蓄水量WU、WL、WD, 总土壤含水量W
    
    Subroutine Eva3XAJ(P,EP,EU,EL,ED,ET,W,WU,WL,WD,WLM,C)
                          
        Implicit none
        Real P
        Real EP,EU,EL,ED,ET
        Real W,WU,WL,WD
        Real WM,WUM,WLM,WDM
        Real K,C
 
        
 !===========三层蒸散发计算========================
        if ((P-EP) .ge. 0.0) then
            EU=EP
            EL=0.0
            ED=0.0
        else
            if  ((WU+P)  .ge.  EP) then                !上层含水量满足蒸发
                EU=EP
                EL=0.0
                ED=0.0
                WU=WU+P-EP
                WL=WL
                WD=WD
            else 
                EU=WU+P
                if  (WL .ge. C*WLM) then               !下层含水量满足补给  
                    EL=(EP-EU)*WL/WLM  
                    ED=0.0
                    WU=0.0
                    WL=WL-EL
                    WD=WD
                else
                    if  (WL .ge. C*(EP-EU)) then           !下层蒸发量与剩余蒸散发能力之比不小于深层蒸散发系数 
                        EL=C*(EP-EU)
                        ED=0.0
                        WU=0.0
                        WL=WL-EL
                        WD=WD                          
                    else                    
                        EL=WL
                        ED=C*(EP-EU)-WL
                        WU=0.0
                        WL=0.0
                        if (WD .lt.  ED) then
                            ED=WD
                            WD=0.
                        else
                            WD=WD-ED  
                        endif
                    endif
                endif
            endif
        endif
            
    
        ET=EU+EL+ED                 !总蒸发量
        W=WU+WL+WD              !总土壤含水量
 
  
    Return
    End  Subroutine Eva3XAJ