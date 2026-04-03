    Subroutine XAJ3Div(pep1,rp1,S1,FR1, KG1,KI1,IM1,SM1,EX1,DeltaT,ntime,nsub,rsp1,rip1,rgp1,sp1,frp1)
       
        Implicit none               
        Integer ntime,nsub,itime,isub                           
        Real pep1(ntime,nsub),rp1(ntime,nsub)
        Real S1(nsub),FR1(nsub)
        Real KG1(nsub),KI1(nsub),IM1(nsub),SM1(nsub),EX1(nsub)
        Real rsp1(ntime,nsub),rip1(ntime,nsub),rgp1(ntime,nsub),sp1(ntime,nsub),frp1(ntime,nsub)
        Real PE,RE
 
      
        Real S,FR
        Real RS,RG,RI,R 
        Real KG,KI,IM,SM,EX
        Real DeltaT
        

        Real bb1,bb2

!--------------------逐个单元流域循环--------------        
            do isub=1,nsub   
                KG=KG1(isub)
            
                KI=KI1(isub)
                
                IM=IM1(isub)
                SM=SM1(isub)
                EX=EX1(isub)
                
                S=S1(isub)
                FR=FR1(isub)
                
                !根据时段长，进行KG,KI参数的转换
                         
                bb1=KG+KI
                bb2=KG/KI
                KI=(1.0-(1.0-bb1)**(DeltaT/24.0))/(1+bb2)
                KG=KI*bb2 
                
 
                 do itime=1,ntime          !逐时刻进行演算
                   
                    PE=pep1(itime,isub)    !总雨量
                    R=rp1(itime,isub)     !总产流量
            

 !=============分水源计算==========================        
                    call Divi3XAJnew(PE,R,S,FR,RS,RI,RG,KG,KI,IM,SM,EX)     
                              
                   
                    rsp1(itime,isub)=RS
                    rip1(itime,isub)=RI
                    rgp1(itime,isub)=RG
                    sp1(itime,isub)=S
                    frp1(itime,isub)=FR
                
 
                enddo
            enddo
    
        Return
    end subroutine XAJ3Div