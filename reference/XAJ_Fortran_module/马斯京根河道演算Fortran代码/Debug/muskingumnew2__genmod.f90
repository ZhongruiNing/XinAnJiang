        !COMPILER-GENERATED INTERFACE MODULE: Sun Sep 05 21:03:21 2021
        ! This source file is for reference only and may not completely
        ! represent the generated interface used by the compiler.
        MODULE MUSKINGUMNEW2__genmod
          INTERFACE 
            SUBROUTINE MUSKINGUMNEW2(NSUB,NTIME,DT,REACD,SECTN,REASCD,  &
     &REAECD,RVMSKN,RVMSKXE,RVMSKKE,QINP1,QOUTINI,QTOT)
              INTEGER(KIND=4) :: NTIME
              INTEGER(KIND=4) :: NSUB
              REAL(KIND=4) :: DT
              CHARACTER(LEN=100) :: REACD(NSUB)
              INTEGER(KIND=4) :: SECTN(NSUB)
              INTEGER(KIND=4) :: REASCD(NSUB)
              INTEGER(KIND=4) :: REAECD(NSUB)
              INTEGER(KIND=4) :: RVMSKN(NSUB)
              REAL(KIND=4) :: RVMSKXE(NSUB)
              REAL(KIND=4) :: RVMSKKE(NSUB)
              REAL(KIND=4) :: QINP1(NTIME,NSUB)
              REAL(KIND=4) :: QOUTINI(NSUB)
              REAL(KIND=4) :: QTOT(NTIME,NSUB+1)
            END SUBROUTINE MUSKINGUMNEW2
          END INTERFACE 
        END MODULE MUSKINGUMNEW2__genmod
