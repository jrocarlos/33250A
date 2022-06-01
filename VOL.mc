my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33250A-5
DATE:                  2019-03-08 12:08:36
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       126
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
#-------------------CLEAR-----------------------
  1.001  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#-------------------VARIÁVEIS-----------------------
  1.002  MATH         P = 0
  1.003  MATH         LP = 2
  1.004  MATH         CP = 1
  1.005  MATH         T  = 0
  1.006  MATH         L = 0
  1.007  MATH         LEV =0
  1.008  MATH         EXP = 0
  1.009  MATH         LINHA = 2
  1.010  MATH         COLUNA = 5
  1.011  MATH         TEMPO = 20
  1.012  MATH         EX = 0
  1.013  MATH         Y = 0
#-------------------CONFIG EXCEL-----------------------
  1.014  LIB          COM xlWS = xlApp.Worksheets["VOL"];
  1.015  LIB          xlWS.Select();
#-------------------CAL MULTIMETER-----------------------
  1.016  OPBR         DESEJA REALIZAR O AUTOCAL?
  1.017  JMPT         1.020
  1.018  JMPF         1.025
  1.019  JMP          1.093
  1.020  IEEE         [@23]RESET
  1.021  IEEE         ACAL ALL
  1.022  WAIT         -t 15:00 AUTOCAL RUN
  1.023  MATH         Y = Y + 1
  1.024  JMP          1.090
#-------------------SETUP-----------------------
  1.025  IF           Y <= 1
  1.026  IEEE         [@23]RESET
  1.027  JMP          1.090
  1.028  ENDIF
  1.029  IEEE         [@23][TERM CR]
  1.030  IEEE         FUNC ACV
  1.031  IEEE         SETACV SYNC
  1.032  IEEE         LFILTER ON
  1.033  IEEE         NDIG 8
  1.034  IEEE         RANGE AUTO
  1.035  IEEE         TARM AUTO
  1.036  IEEE         RES 0.00001
#-------------------CONFIG  Nº MEAS----------------
  1.037  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.038  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.039  DO
  1.040  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.041  LIB          PONTO = P1.Value2;
  1.042  IF           PONTO == 0
  1.043  JMP          1.093
  1.044  ENDIF
  1.045  MATH         CP = CP + 1
  1.046  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.047  LIB          TEX = T1.Value2;
  1.048  MATH         P = PONTO&TEX
 #----------------LEVEL-------------------
  1.049  MATH         CP = CP + 1
  1.050  LIB          COM L1 = xlApp.Cells[LP,CP];
  1.051  LIB          LEV = L1.Value2;
  1.052  MATH         CP = CP + 1
  1.053  LIB          COM T2 = xlApp.Cells[LP,CP];
  1.054  LIB          EXP = T2.Value2;
  1.055  MATH         L = LEV&EXP
  1.056  MATH         Z1 = CMP  (EXP,"mV")
  1.057  MATH         Z2 = CMP  (EXP,"V")
#----------------------END-------------------------------
  1.058  IF           P == 00
  1.059  JMP          1.093
  1.060  ENDIF
#----------------------------CONFIG OUT GENERATOR--------------
  1.061  DO
  1.062  IEEE         [@13]:FREQ [V P]
  1.063  IEEE         :VOLT:UNIT VRMS
  1.064  IEEE         :VOLT:LEV [V L]
  1.065  IEEE         OUTP ON
  1.066  WAIT         [D2000]
#------------CONFIG MULT----------------
  1.067  WAIT         -t [V TEMPO] Please Standby
  1.068  IEEE         [@23]ACV[I]
  1.069  IF           Z1 == 1
  1.070  MATH         EX = 1E-3
  1.071  ENDIF
  1.072  IF           Z2 == 1
  1.073  MATH         EX = 1E0
  1.074  ENDIF
  1.075  MATH         MEM = MEM / EX
#------------------SAVE DATE----------------
  1.076  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.077  LIB          selectedCell.Select();
  1.078  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.079  MATH         T = T + 1
  1.080  MATH         COLUNA = COLUNA + 1
  1.081  MATH         CP = CP + 1
  1.082  UNTIL        T == A
  1.083  MATH         T  = 0
  1.084  MATH         COLUNA = 5
  1.085  MATH         LINHA = LINHA + 1
  1.086  MATH         CP = 1
  1.087  MATH         LP = LP + 1
  1.088  UNTIL        PONTO == 0
  1.089  JMP          1.093
#-------------------SETUP-----------------------
  1.090  DISP         Connect the generator to the UUT as follows:
  1.090  DISP
  1.090  DISP         [32]   Generator         to         Counter
  1.090  DISP         [32]
  1.090  DISP         [32]     OUTPUT -------------------> CHANNEL 1
  1.090  DISP         [32]
  1.090  DISP         [32]     GPIB CONTADOR 3458A = 23
  1.090  DISP         [32]     GPIB GERADOR 33250A = 13
  1.091  PIC          SETUP4
  1.092  JMP          1.029
#------------------RESET------------------
  1.093  IEEE         [@23]*RST
  1.094  IEEE         [@13]*RST
