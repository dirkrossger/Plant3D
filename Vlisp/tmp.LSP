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

(defun cd:CreateMLeader (string /)
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
  (command-s "_undo" "_be")
  (if (findfile (strcat (getvar "DWGPREFIX") "mleaderstyles.dwg"))
    (command "_insert" (strcat (getvar "DWGPREFIX") "mleaderstyles.dwg") '(0 0 0) 1 1 0)
    (princ "\nMultileaderstyles not found!")
    )

  ; Get block scale for select mleader style
  (if (setq scale txh)
    (cond
      ((equal 1 scale) (setvar "CMLEADERSTYLE" "M-400-2.5"))
      ((equal 0.5 scale) (setvar "CMLEADERSTYLE" "M-200-2.5"))
      ((equal 0.25 scale) (setvar "CMLEADERSTYLE" "M-100-2.5"))
      ((equal 2.5 scale) (setvar "CMLEADERSTYLE" "M-1000-2.5"))
      (T (setvar "CMLEADERSTYLE" "Standard"))
       )
    )

;;;  (setq zdir (trans (list 0 0 0) 1 0 T)
;;;        xang (cdr (assoc 50 (entget e)))) ;(angle (list 0 0 0) (trans (getvar "UCSXDIR") 0 zdir))
;;;  (command-s "_ucs" "_ob" e); zdir xang "")
  (setq p1 (getpoint "\nSpecify leader arrowhead location: "))
  (setq p2 (getpoint "\nSpecify leader text location: "))
;;;  (setq o (vlax-ename->vla-object e))

  ; use mleader with mtext
  (cond
    ((= (substr (getvar "acadver") 1 2) "21") ; Ausnahme in 2017 version
     (command-s "._mleader" "_l" p2 (trans p1 0 1) string))
    (T (command-s "._mleader" (trans p1 0 1) p2 string)))

  (command-s "_ucs" "_p")
  (command-s "_undo" "_e")
  (princ "\nFinished.")

  (princ)
  )

(defun c:xx()
  (setq txh (cd:USR_Abfrage "\nEnter Textheight: " (if (null txh) 1.0 txh)))
  (cd:CreateMLeader "test")

  (princ)
  )