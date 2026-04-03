  !=============根据水面蒸发能力ETW和折算系数K，计算蒸发能力EP======================= 
    
    Subroutine ETPcal(Eobs1,K1,ETP1,ntime,nsub)
        Implicit none
        Real Eobs1(ntime,nsub)
        Real K1(nsub)
        Real ETP1(ntime,nsub)
        Integer ntime, nsub,itime,isub
        
        do isub=1,nsub
            do itime=1,ntime
                ETP1(itime,isub)=Eobs1(itime,isub)*K1(isub)
            enddo
        enddo
        
        Return
    
    end subroutine ETPcal
    