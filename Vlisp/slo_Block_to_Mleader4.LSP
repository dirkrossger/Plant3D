(defun c:fAttr (/
		test
		sel
		e
		o
	        p1
	        p2
	        result
		findReal
		_getatt
		*error*
	     )
  
  (defun _getatt (block tag / att result)
    (foreach att (vlax-invoke block 'getattributes)
      (if (eq (strcase tag) (strcase (vla-get-tagstring att)))
	(progn (setq result (vla-get-textstring att)))
      )
    )
    result
    )
  
  (defun findReal (x)
    (distof (apply 'strcat (mapcar 'chr(vl-remove-if-not '(lambda(X)(member X (vl-string->list "0123456789.")))(vl-string->list x)))))
  )

  
  (defun *error*(msg)
    (cond
      ((= test 1)
       (alert
	 "Block definition could be wrong!
Blockname = Level
Block Attribute tag = Elev")
       )
      ((= test 2)
       (alert "Import Multileaderstyle missed!")
       )
      )
    (princ msg)
    (if (/= "(T T T)" (vl-princ-to-string (mapcar '= odir (getvar "UCSXDIR"))))
      (command-s "_ucs" "_p" "")
      )
    (command-s "_undo" "_e")
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

  ; ---------------------------------------------------------------------------------------
  ;                                   MAIN / START ROUTINE                                ;
  ; ---------------------------------------------------------------------------------------

 
  (princ "\nSet flag to Height Attribute: Pick Attribute with new Position!")
  (setvar "CMDDIA" 0)
  (setvar "CMDECHO" 0)
  (setq test 0)

  (command-s "_undo" "_be")
  (setq result nil)

  ; insert standard multileaders from Drawing
  (if (null (tblsearch "block" "Standards"))
    (if (setq file (findfile (strcat (getenv "iso\\library") "Standards.dwg"))) ; load definition mleaders in Standards.dwg
       (command-s "_insert" file '(0 0) 1 1 0)
       (alert
	 "Import Multileaderstyle missed!
Repair Infratools installation"))
    )

  (if (tblsearch "block" "Standards")
    (createMleader))

  (princ)
  )




