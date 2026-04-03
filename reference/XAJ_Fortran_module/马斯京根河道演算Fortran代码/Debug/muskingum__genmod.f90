        !COMPILER-GENERATED INTERFACE MODULE: Sun Sep 05 21:03:21 2021
        ! This source file is for reference only and may not completely
        ! represent the generated interface used by the compiler.
        MODULE MUSKINGUM__genmod
          INTERFACE 
            SUBROUTINE MUSKINGUM(MK,MX,N,DELTAT,QIN,QOUT1,NTIME,QOUT)
              INTEGER(KIND=4) :: NTIME
              REAL(KIND=4) :: MK
              REAL(KIND=4) :: MX
              INTEGER(KIND=4) :: N
              REAL(KIND=4) :: DELTAT
              REAL(KIND=4) :: QIN(NTIME)
              REAL(KIND=4) :: QOUT1
              REAL(KIND=4) :: QOUT(NTIME)
            END SUBROUTINE MUSKINGUM
          END INTERFACE 
        END MODULE MUSKINGUM__genmod
