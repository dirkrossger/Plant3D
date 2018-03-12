(progn
  (setq txh (cd:USR_Abfrage "\nEnter Textheight: " (if (null txh) 1.0 txh)))

; =========================================================================================== ;
; Abfrage eines Wertes 'INT oder 'STR
; (cd:USR_Abfrage "\nEingabe DateiPrefix" (if (null DEFAULT) " " DEFAULT))
; =========================================================================================== ;

(defun cd:USR_Abfrage (MSG DEFAULT / VALUE)
  (if(or(=(type MSG)'STR)(setq MSG ""))
    (cond
      ((=(type DEFAULT) 'STR)
        (if(=(setq VALUE(getstring 't (strcat "\n" MSG "<" DEFAULT ">: ")))"")
          (setq VALUE DEFAULT)
          VALUE
        )
      ) 
      ((=(type DEFAULT) 'INT)
        (if(not(setq VALUE(getint (strcat "\n" MSG "<" (vl-princ-to-string DEFAULT) ">: "))))
          (setq VALUE DEFAULT)
          VALUE
        )  
      )
      ((=(type DEFAULT) 'REAL)
          (if(not(setq VALUE(getreal (strcat "\n" MSG "<" (vl-princ-to-string DEFAULT)">: "))))
            (setq VALUE DEFAULT)
            VALUE
          )  
        )	  
    )
  )
)
  (defun createMLeader ()
  (setq currMl (getvar "CMLEADERSTYLE"))
  (if (null (setq allMl '()))
    (foreach n (DICTSEARCH (NAMEDOBJDICT) "ACAD_MLEADERSTYLE")
      (cond
	((= (car n) 3)
	 (setq allMl (append allMl (list (cdr n))))
	 )
	)
      )
    )

   (setq odir (getvar "UCSXDIR"))
   (while
    (not
      (and
	(setq sel (entsel "\nSelect a Height:"))

	(setq e (car sel))
	(= (cdr (assoc 0 (entget e))) "INSERT")
;;;	(null result)
	(cond
	    ((= (strcase (cdr (assoc 2 (entget e)))) "LEVEL")
             (command-s "_undo" "_be")

	     ; Get block scale for select mleader style
	     (if (setq scale (list (cdr (assoc 41 (entget e))) (cdr (assoc 42 (entget e)))))
	       (cond
		 ((equal '(1 1) scale) (setvar "CMLEADERSTYLE" "M-400-2.5"))
	         ((equal '(0.5 0.5) scale) (setvar "CMLEADERSTYLE" "M-200-2.5"))
		 ((equal '(0.25 0.25) scale) (setvar "CMLEADERSTYLE" "M-100-2.5"))
		 ((equal '(2.5 2.5) scale) (setvar "CMLEADERSTYLE" "M-1000-2.5"))
		 (T (setvar "CMLEADERSTYLE" "Standard"))
		 )
	       )
	     
	     (setq zdir (trans (list 0 0 0) 1 0 T)
		   xang (cdr (assoc 50 (entget e)))) ;(angle (list 0 0 0) (trans (getvar "UCSXDIR") 0 zdir))
	     (command-s "_ucs" "_ob" e); zdir xang "")

	     (setq p1 (cdr (assoc 10 (entget e))))
	     (setq p2 (getpoint (trans p1 0 1) "\nSpecify leader arrowhead location: "))
	     (setq o (vlax-ename->vla-object e))
	     ; build mleader with block "plushöjd"
;;;	     (command-s "._mleader" "_o" "_c" "_b" "plushöjd" "_x" (trans p1 0 1) p2 (strcat (_getatt o "NIVÅ")))

	     ; use mleader with mtext
	     (cond
	       ((= (substr (getvar "acadver") 1 2) "21") ; Ausnahme in 2017 version
		(command-s "._mleader" "_l" p2 (trans p1 0 1) (rtos (findReal (_getatt o "Elev")) 2 2))
		)
	       (T (command-s "._mleader" (trans p1 0 1) p2 (rtos (findReal (_getatt o "Elev")) 2 2)))
	       )

	     
	     (command-s "_ucs" "_p")
	     (if (= "MULTILEADER" (cdr (assoc 0 (entget (entlast)))))
	       (progn
		 (entdel e)
		 )
	       )
	     (command-s "_ucs" zdir xang "")
;;;	     (princ "\nMultileader successfully")
	     (setq result t)
	     (command-s "_undo" "_e")
	     )
	    (T
	     (setq test 1)
	     (princ "\nSelected block must named 'Level' with Attribut!")
	     (setq result nil)
	     )
	    )
	)
      )
;;;    (princ "\nWrong Choice ...")
    )
   )