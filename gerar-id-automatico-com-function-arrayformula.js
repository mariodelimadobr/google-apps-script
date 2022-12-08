#function gerar id da inscrição
=ARRAYFORMULA(IF(ISBLANK($K$2:$K);""; 
ARRAYFORMULA(ROW($A$2:$A)-1 + 12022010)))
