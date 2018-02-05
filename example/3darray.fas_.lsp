(defun main 
FasStringtables 0
FasStringtables 1
(defun main 
nil
(setq AI_ABORT <Func> AI_ABORT)
AI_ABORT
(cond T (
(cond (NOT (FINDFILE "ai_utils.lsp")) (
(cond (EQ "failed" (LOAD "ai_utils" "failed")) (
normal cond
nil
normal cond
(AI_ABORT "3DREIHE" "Kann Datei AI_UTILS.LSP nicht finden")
normal cond
(AI_ABORT "3DREIHE" (STRCAT "Kann Datei AI_UTILS.LSP nicht finden." "\n Support-Verzeichnis überprüfen."))
Then OR Else
(setq MODES <Func> MODES)
MODES
(setq MODER <Func> MODER)
MODER
(setq 3DAERR <Func> 3DAERR)
3DAERR
(setq P-ARRAY <Func> P-ARRAY)
P-ARRAY
(setq R-ARRAY <Func> R-ARRAY)
R-ARRAY
(setq C:3DARRAY <Func> C:3DARRAY)
(vl-ACAD-defun C:3DARRAY)
C:3DARRAY
(PRINC "  3DREIHE geladen.")
(defun AI_ABORT
(APP  MSG)
(defun *ERROR*
(S)
(setq *ERROR* OLD_ERROR)
(setq *ERROR* <Func> *ERROR*)
(ALERT (STRCAT " Anwendungsfehler: " APP " \n\n  " MSG "  \n"))
(defun MODES
(A)
(setq MLST nil)
nil
(setq MLST (APPEND MLST (LIST (LIST (CAR A) (GETVAR (CAR A))))))
(setq A (CDR A))
(defun MODER
nil
(SETVAR (CAAR MLST) (CADAR MLST))
(setq MLST (CDR MLST))
(defun 3DAERR
(ST)
(PRINC (STRCAT "\nFehler: " S))
(ads-cmd Then OR Else)
(ads-cmd "_E")
(AI_UNDO_OFF )
Retval of jmp1 MODER at 233()
(setq *ERROR* OLDERR)
(defun P-ARRAY
(_al-bind-alist '(N AF YN CEN C RA))
(setq N 0)
(INITGET (+ 1 2 4))
(setq N (GETINT "\nAnzahl der Elemente in der Anordnung angeben: "))
(PROMPT "\nElementanzahl muß höher sein als 1")
(INITGET 2)
(setq AF (GETREAL "\nAuszufüllenden Winkel angeben (+=ccw, -=cw) <360>: "))
(setq AF 360)
(INITGET Then OR Else)
(setq YN (GETKWORD "\nAngeordnete Objekte drehen? [Ja/Nein] <J>: "))
(setq YN "Ja")
(setq YN Then OR Else)
(INITGET 17)
(setq CEN (GETPOINT "\nMittelpunkt der Anordnung angeben: "))
(setq C (TRANS CEN 1 0))
(INITGET 17)
(setq RA (GETPOINT CEN "\nZweiten Punkt auf Drehachse angeben: "))
(PRINC "\nUngültiger Punkt: Zweiter Punkt darf nicht mit Mittelpunkt übereinstimmen.")
(INITGET 17)
(setq RA (GETPOINT CEN "\nBitte versuchen Sie es erneut: "))
(SETVAR "UCSFOLLOW" 0)
(SETVAR "GRIDMODE" 0)
(ads-cmd "_.UCS")
(ads-cmd "_ZAXIS")
(ads-cmd CEN)
(ads-cmd RA)
(setq CEN (TRANS C 0 1))
(ads-cmd "_.ARRAY")
(ads-cmd SS)
(ads-cmd "")
(ads-cmd "_P")
(ads-cmd CEN)
(ads-cmd N)
(ads-cmd AF)
(ads-cmd YN)
(ads-cmd "_.UCS")
(defun R-ARRAY
(_al-bind-alist '(NR NC NL FLAG X Y Z C EL EN SS2 E))
(cond (= NR NC NL nil) (
(cond (= NR NC NL 1) (
it's OR skip next 6 bytes -> 977
it's OR skip next 6 bytes -> 977
(setq NR 1)
(INITGET (+ 2 4))
(setq NR (GETINT "\nZeilenanzahl eingeben (---) <1>: "))
(setq NR 1)
(INITGET (+ Then OR Else 4))
(setq NC (GETINT "\nSpaltenanzahl eingeben (|||) <1>: "))
(setq NC 1)
(INITGET (+ Then OR Else 4))
(setq NL (GETINT "\nEbenenanzahl eingeben (...) <1>: "))
(setq NL 1)
(PRINC "\nAnordnung mit einem Element, nichts zu verarbeiten.\nBitte versuchen Sie es erneut")
(setq MAXLIMIT 100000)
(setq MAXLIMIT Then OR Else)
(setq NE (SSLENGTH SS))
(PRINC "\nDies würde ")
(PRINC (- (* NC NR NL NE) 1))
(PRINC " Objekte erzeugen, also die erlaubte Anzahl von ")
(PRINC MAXLIMIT)
(SETVAR "ORTHOMODE" 1)
(SETVAR "HIGHLIGHT" 0)
(setq FLAG 0)
(INITGET (+ 1 2))
(setq Y (GETDIST "\nZeilenabstand eingeben (---): "))
(setq FLAG 1)
(INITGET (+ 1 2))
(setq X (GETDIST "\nSpaltenabstand eingeben (|||): "))
(setq FLAG (+ FLAG 2))
(INITGET (+ 1 2))
(setq Z (GETDIST "\nEbenenabstand eingeben (...): "))
(SETVAR Then OR Else 0)
(setq C 1)
(setq EL (ENTLAST ))
(setq EN (ENTNEXT EL))
(setq EL EN)
(setq EN (ENTNEXT EL))
(ads-cmd "_.COPY")
(ads-cmd SS)
(ads-cmd "")
(ads-cmd "0,0,0")
(ads-cmd (APPEND (LIST 0 0) (LIST (* C Z))))
(setq C (1+ C))
(setq SS2 (SSADD ))
(setq E (ENTNEXT EL))
(setq ED (ENTGET E))
(cond (= (CDR (NTH 1 ED)) "VERTEX") (
(cond (= (CDR (NTH 1 ED)) "ATTRIB") (
(cond (= (CDR (NTH 1 ED)) "SEQEND") (
it's OR skip next 6 bytes -> 1829
it's OR skip next 6 bytes -> 1829
it's OR skip next 6 bytes -> 1829
(SSADD E SS2)
(setq E (ENTNEXT Then OR Else))
(cond (= FLAG 1) (
(cond (= FLAG 2) (
(cond (= FLAG 3) (
normal cond
nil
(ads-cmd "_.ARRAY")
(ads-cmd SS)
(ads-cmd SS2)
(ads-cmd "")
(ads-cmd "_R")
(ads-cmd NR)
(ads-cmd NC)
(ads-cmd Y)
normal cond
(ads-cmd X)
(ads-cmd "_.ARRAY")
(ads-cmd SS)
(ads-cmd SS2)
(ads-cmd "")
(ads-cmd "_R")
(ads-cmd "1")
(ads-cmd NC)
normal cond
(ads-cmd X)
(ads-cmd "_.ARRAY")
(ads-cmd SS)
(ads-cmd SS2)
(ads-cmd "")
(ads-cmd "_R")
(ads-cmd NR)
(ads-cmd "1")
(defun C:3DARRAY
(_al-bind-alist '(OLDERR SS XX UNDO_SETTING))
(PROMPT "\n *** Unzulässiger Befehl im Papierbereich ***\n")
(setq OLDERR *ERROR*)
(setq *ERROR* 3DAERR)
Retval of jmp1 MODES at 121('("cmdecho" "blipmode" "highlight" "orthomode" "ucsfollow" "gridmode"))
(SETVAR "CMDECHO" 0)
(AI_UNDO_ON )
(ads-cmd "_.UNDO")
(ads-cmd "_GROUP")
(GRAPHSCR )
(setq SS nil)
(setq SS (AI_SSGET (SSGET )))
(INITGET 0 "Rechteckige Polare Kreisförmige")
(setq XX (GETKWORD "\nAnordnungstyp eingeben [Rechteckig/Polar] <R>:"))
(cond (EQ XX "Rechteckige") (
(cond (EQ XX nil) (
it's OR skip next 6 bytes -> 2426
it's OR skip next 6 bytes -> 2426
(cond T (
(cond T (
normal cond
T
Retval of jmp1 P-ARRAY at 409()
normal cond
Retval of jmp1 P-ARRAY at 409()
Retval of jmp1 R-ARRAY at 914()
(ads-cmd "_.UNDO")
(ads-cmd "_E")
(AI_UNDO_OFF )
Retval of jmp1 MODER at 233()
(setq *ERROR* OLDERR)
