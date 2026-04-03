Subroutine XAJ3EvaYie(Pobs1,ETP1,W1,WU1,WL1,B1,C1,WM1,WUM1,WLM1,IM1,ntime,nsub,pep1,epp1,rp1,wp1,wup1,wlp1)
        Implicit none                           
        Integer ntime,nsub,itime,isub
        Real Pobs1(ntime,nsub),ETP1(ntime,nsub)
        Real W1(nsub),WU1(nsub),WL1(nsub)
        Real B1(nsub),C1(nsub),WM1(nsub),WUM1(nsub),WLM1(nsub),IM1(nsub)
 
        
        Real wp1(ntime,nsub),wup1(ntime,nsub),wlp1(ntime,nsub),epp1(ntime,nsub)

        Real pep1(ntime,nsub),rp1(ntime,nsub) 
                  
        Real W,WU,WL,WD
        Real EU,EL,ED,ET
        Real B,C,WM,WUM,WLM,IM  
        Real R,EP,P,PE
  
  
       
!--------------------逐个单元流域循环--------------        
        do isub=1,nsub   
            W=W1(isub)
            WU=WU1(isub)
            WL=WL1(isub)
            WD=W-WU-WL
            WM=WM1(isub)
            WUM=WUM1(isub)
            WLM=WLM1(isub)
            C=C1(isub)
            B=B1(isub)
            IM=IM1(isub)
            
            do itime=1,ntime
                P=Pobs1(itime,isub)
                EP=ETP1(itime,isub)
!====三层蒸散发计算模块：按照时段，根据降雨P，蒸发能力EP,上一个时刻土壤含水量状态W,WU,WL,WD，参数WLM,C，计算实际蒸发量EU,EL,ED,ET,并对土壤含水量状态W,WU,WL,WD进行更新==========================     
                call Eva3XAJ(P,EP,EU,EL,ED,ET,W,WU,WL,WD,WLM,C)
                
                PE=P-ET   !净雨量
                
!=============产流计算===============================     
             
!===============三水源产流计算模块：根据净雨量PE,土壤含水量状态W,WU,WL,WD,参数WM,WUM,WLM,B,IM ,计算产流量R           
                call YieldXAJnew(PE,W,WU,WL,WD,WM,WUM,WLM,R,B,IM) 
                    
                
! 将计算得到本时段单元流域实际蒸发量ET、净雨量PE、产流量R、土壤含水量状态W、WU、WL、WD赋值给数组epp1，pep1,rp1，wp1,wup1,wlp1,wdp1         
                epp1(itime,isub)=ET     
                pep1(itime,isub)=PE
                rp1(itime,isub)=R
                wp1(itime,isub)=W
                wup1(itime,isub)=WU
                wlp1(itime,isub)=WL
                 
            enddo
        enddo
        
    Return
    end subroutine XAJ3EvaYie