	(princ "Start")

	(defun TESTExt ( Blah1 Blah2)
	  (princ "Start")
	)

	(defun Local ( / Blah1 Blah2)
	  (princ "Start")
	)


	(defun TEST ( Blah1 Blah2 / ANG1 ANG2)
 
	     (setq ANG1 "Monday")
	     (setq ANG2 "Tuesday")
 
	     (princ (strcat "\nANG1 has the value " ANG1))
	     (princ (strcat "\nANG2 has the value " ANG2))
	   (princ)
	   
	   
		(defun DDSTEEL ( / p1 p2 p3)
				(princ "func Inside func")
		)
	);defun
	
	(TEST 1 2)
	(princ "End")
