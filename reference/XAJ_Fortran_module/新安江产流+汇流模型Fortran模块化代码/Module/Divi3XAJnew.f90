!++++++++++++++++++++三水源新安江产流模型中的水源划分模块++++++++++++++++++++
!---------------------------输入：降雨扣除蒸发量PE,产流量R
!--------------------------输入：时段初自由水蓄量S,时段初产流面积FR
!--------------------------输出：地表径流深RS,壤中流径流深RI,地下径流深RG,更新的自由水蓄量S,更新的产流面积比FR
!------------------------参数：不透水面积IM,自由水蓄水容量SM,壤中流出流系数KI,地下径流出流系数KG,流域自由水容量分布曲线指数EX
 Subroutine  Divi3XAJnew(PE,R,S,FR,RS,RI,RG,KG,KI,IM,SM,EX)
                            
     
        Implicit none
        Real PE,R
        Real S,FR
        Real RS,RG,RI 
        Real SM,EX,IB,IM,KG,KI
!-------中间变量-------        
        Real RB,KGd,KId,xx,td,ff
        Real AU,SMM 
        Integer j
        Real SS
        Real div,Qdiv,Qall
        Integer nd
 
 
        
        !==========model parameters=============    
  
            div=5.0            !径流分段门槛
 
         
            SMM = (1.0 + EX) * SM          !SMM-流域单点最大自由水蓄水容量
            If (PE .le. 0.0) Then           !无地表径流
                RS = 0.0
                RG = S * KG * FR
                RI = S * KI * FR
                S = S * (1.0 - KG - KI)
            else
                RB = IM *PE                   !不透水面积产流
                RS = 0.0
                RI = 0.0
                RG = 0.0
                
                td=R-IM*PE
                xx=FR
                FR=td/PE
            
                S = xx * S / FR
                SS=S
                Qall=R/FR
                nd=int(Qall/div)+1
                Qdiv=Qall/float(nd)
                
                KId = (1.- (1. - (KG + KI)) ** (1.0 / nd)) / (KG + KI)           !基于径流深分段计算时的参数转换
                KGd = KId * KG
                KId = KId * KI
            
                
                RS = 0.0
                RI = 0.0
                RG = 0.0
                
                do j=1,nd                         !分段进行三种径流深计算，并对S进行更新
                    AU = SMM * (1.0 - (1.0 - S / SM) ** (1.0 / (1.0 + EX)))
                    if (AU + Qdiv .lt. SMM) then
                        RS=(Qdiv-SM+S+SM*(1-(Qdiv+AU)/SMM)** (1 + EX))*FR+RS
                    else
                        RS=(Qdiv+S-SM)*FR+RS
                    endif
                    S=j*Qdiv-RS/FR+S
                    RG = S * KGd * FR + RG             !时段地下径流
                    RI = S * KId * FR + RI                !时段地下径流
                    S=j*Qdiv+SS-(RS+RI+RG)/FR
                enddo
                
                RS=RS+RB
            endif
        
        Return
            
    End Subroutine DIVI3XAJnew
    
    
    