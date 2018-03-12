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

(defun cd:CreateMLeader (string p1 p2 /)
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

  (if (and (findfile (strcat (getvar "DWGPREFIX") "mleaderstyles.dwg"))
	   (null (tblobjname "block" "mleaderstyles")))
    (command "_insert" (strcat (getvar "DWGPREFIX") "mleaderstyles.dwg") '(0 0 0) 1 1 0)
    (if (null (findfile (strcat (getvar "DWGPREFIX") "mleaderstyles.dwg")))
      (princ "\nMultileaderstyles not found!")
      )
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

;;;  (setq p1 (getpoint "\nSpecify leader arrowhead location: "))
;;;  (setq p2 (getpoint "\nSpecify leader text location: "))

  ; use mleader with mtext
  (cond
    ((= (substr (getvar "acadver") 1 2) "21") ; Ausnahme in 2017 version
     (command-s "._mleader" "_l" p2 (trans p1 0 1) string))
    (T (command-s "._mleader" (trans p1 0 1) p2 string)))

;;;  (princ "\nFinished.")

  (princ)
  )

(defun c:xx()
  (command-s "_undo" "_be")

  (setq osm (getvar "OSMODE"))
  (setvar "OSMODE" 0)
  (setvar "CMDDIA" 0)
  (setvar "CMDECHO" 0)
  (setq txh (cd:USR_Abfrage "\nEnter Textheight: " (if (null txh) 1.0 txh)))
  (if (setq ss (ssget "X" '((0 . "LINE"))))
    (repeat (setq i (sslength ss))
      (setq e (ssname ss (setq i (1- i))))
      (setq layname (cdr (assoc 8 (entget e))))
      (setq inspt (cdr (assoc 10 (entget e))))
      (cd:CreateMLeader layname inspt (list (+ txh (car inspt)) (+ txh (cadr inspt)) 0))
      )
    )
  (command-s "_undo" "_e")

  (setvar "OSMODE" osm)
  
  )
  (princ)
  )