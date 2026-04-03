 !++++++++++++++三水源滞后演算模型中的坡地汇流模块++++++++++++++++++++++++++++
!-----------输入：参数CI1,CG1;  面积Area1; 计算时段长DeltaT;系列长度Ntime；单元流域个数nsub
!-----------输入：地表径流深rsp1,壤中流径流深rip1,地下径流深rgp1，初始时刻的壤中径流量QI1和地下径流量QG1
!-----------输出:  地表流量qsp1,壤中流流量qip1,地下径流流量gqp1 
!----------中间变量：U 
!--------- Program by Yingchun Huang on 2020-12-20-------------------        
    Subroutine LagHillslopenew(CI1,CG1,Area1,DeltaT,ntime,nsub,rsp1,rip1,rgp1,QI1,QG1,qsp1,qip1,qgp1)
                                     
         
        Implicit none
        Integer ntime,nsub,itime,isub
        Real CI1(nsub),CG1(nsub)         
        Real DeltaT,Area1(nsub),U
 
        Real rsp1(ntime,nsub),rgp1(ntime,nsub),rip1(ntime,nsub)
        Real QI1(nsub),QG1(nsub)
        Real qsp1(ntime,nsub),qip1(ntime,nsub),qgp1(ntime,nsub)
        Integer i,j
        Real CG,CI
        
        
        do isub=1,nsub
            
            
            CG = CG1(isub)**(DeltaT/24.0)
            CI = CI1(isub) **(DeltaT/24.0)
          
 !========单位转换系数============       
            U=Area1(isub)/(3.6*DeltaT)
        
 !=======初始时刻壤中流量QI,地下径流QG赋值======

           qip1(1,isub)=QI1(isub)
           qgp1(1,isub)=QG1(isub)

!===============坡地汇流计算==========
        Do i=2,ntime
            qsp1(i,isub)=rsp1(i,isub)*U                                 
            qip1(i,isub)=CI*qip1(i-1,isub)+(1.0-CI)*rip1(i,isub)*U
            qgp1(i,isub)=CG*qgp1(i-1,isub)+(1.0-CG)*rgp1(i,isub)*U                  
        enddo
       
        enddo
        
        
     
        Return
    End Subroutine LagHillslopenew
     
    
