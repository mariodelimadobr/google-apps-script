=IF(ISBLANK(K2);    "";
IF(D2 <> TRUE;  "CPF Nulo";
IF(COUNTIFS(T3:T;T2)>= 1;    "Cancelada";    "Validada")))

=ARRAYFORMULA(IF(ISBLANK($K$2:$K);""; 
ARRAYFORMULA("Lista 1")))

=IF(ISBLANK(K2);""; 
IF(BI2  = "2 anos";    BJ2;
IF(BI2  = "2,5 anos";  BK2;
IF(BI2  = "3 anos";    BL2;
IF(BI2  = "3,5 anos";  BM2;
IF(BI2  = "4 anos";    BN2;
IF(BI2  = "4,5 anos";  BO2;
IF(BI2  = "5 anos";    BP2;
"Verificar"))))))

=FILTER(BI2:BI;K2:K <> "")
