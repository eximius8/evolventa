;Written by Trunov Mikhail
; Excel link taken from https://www.ozon.ru/context/detail/id/2644304/?partner=22andxy&utm_content=link
(defun ex_set_connect (file_path /) 
  (setq g_oex (vlax-get-or-create-object "Excel.Application"))
  (vlax-put-property g_oex "Visible" :vlax-false)
  (setq g_wkbs (vlax-get-property g_oex "Workbooks")) 
  (setq g_awb (vlax-invoke-method g_wkbs "Open" file_path))
  (setq g_shs (vlax-get-property g_awb "Worksheets"))
  (setq g_mainsh (vlax-get-property g_shs "Item" 1))
);defun ex_set_connect

(defun ex_break_connect (file_path /)
 (vlax-invoke-method g_awb "SaveCopyAs"  file_path) 
 (vlax-invoke-method g_awb "Close" :vlax-false :vlax-false)
 (vlax-invoke-method g_oex "Quit")
  (mapcar
    (function
      (lambda (x)
        (if
          (and x (not (vlax-object-released-p x)))
          (vlax-release-object x)
        );_end of if
      );_end of lambda
    );_end of function
    (list g_mainsh g_shs g_awb g_wkbs g_oex celk)
  );_end of mapcar
  (setq g_oex nil
	g_mainsh nil
	g_shs nil
	g_awb nil
	g_wkbs nil
	celk nil
  )
  (gc)
)

(defun ex_put (znachen cel / celk)  
  (setq celk (vlax-get-property g_mainsh "Range" cel))  
  (vlax-put-property celk "Formula" znachen)
  (vlax-release-object celk)  
  (gc)
)

(defun tg (alfa)
  (/ (sin alfa) (cos alfa))
)

(defun arccos (cosalf)    
   (cond
     ((> (abs cosalf) 1)
      (alert "\nacos not calculated")
     )
     ((= cosalf 0.)
      (* pi 0.5)
     )
     ((>= cosalf 0.)
      (atan (/ (sqrt (- 1 (* cosalf cosalf))) cosalf))
     )
     ((< cosalf 0.)
      (+ pi (atan (/ (sqrt (- 1 (* cosalf cosalf))) cosalf)))
     )
   )
)

(defun inv (ugol)
  (- (tg ugol) ugol)
)

(defun add_pl (list_1 / list_2)
  (setq list_2 (mapcar '(lambda (a) (trans a 1 0)) list_1))
  (entmakex (append	      
	      (list '(0 . "LWPOLYLINE")
		    '(100 . "AcDbEntity")
		    '(100 . "AcDbPolyline")
		    (cons 90 (length list_1))
		    '(70 . 0)
	      ); list
	      (mapcar '(lambda (a) (cons 10 a))
		      list_2
	      ); mapcar
	    ); append
  ); entmakex
); defun

(defun add_tex (t1 text_ )  
  (entmakex (list '(0 . "MTEXT")
		    '(100 . "AcDbEntity")
		    '(100 . "AcDbMText")
		    '(41 . 7)
		    (cons 7 (getvar "TEXTSTYLE"))
		    (cons 8 (getvar "CLAYER"))
		    (cons 10 (trans t1 1 0))
		    (cons 1 text_)
	      ); list
    ); entmakex
); defun

(defun li (list_of cvet / i list_new)
  (if (> (length list_of) 1)
    (progn
      (entmakex (list '(0 . "line")
			'(48 . 0.59)
		      (cons 62 cvet)
		(cons 8 (getvar "CLAYER"))
		(cons 10 (trans (nth 0 list_of) 1 0))
		(cons 11 (trans (nth 1 list_of) 1 0))
	    ); list
      ); entmakex
      (setq i 1)
      (repeat (1- (length list_of))
	(setq list_new (cons (nth i list_of) list_new)
	      i (1+ i)
	); setq
      ); repeat
      (li list_new cvet)
    ); progn
  ); if
); defun


(defun add_arc (center a1 a2 / b1 b2 cent)
  (setq cent (trans center 1 0)
	b1 (trans a1 1 0)
	b2 (trans a2 1 0)
  ); setq
  (entmakex (list '(0 . "ARC")
		  '(100 . "AcDbEntity")
		  '(67 . 0)
		  '(100 . "AcDbCircle")
		  (cons 10 cent)
		  (cons 40 (distance cent b1))
		  '(210 0.0 0.0 1.0)
		  '(100 . "AcDbArc")		  
		  (cons 50 (angle cent b1))
		  (cons 51 (angle cent b2))
	    ); list
  ); entmakex
); defun

(defun add_arc2 (center z d param cvet / cent ugol1 ugol2)
  (setq cent (trans center 1 0)
	ugol1 (- (if (eq param 1)
		     (* pi 1.5)
		     (* pi 0.5)
		 ); if
		 (+ (* 0.1 pi) (/ p d 0.5))
	      ); -
	ugol2 (+ (if (eq param 1)
		     (* pi 1.5)
		     (* pi 0.5)
		 ); if
		 (+ (* 0.1 pi) (/ p d 0.5))
	      ); +
  ); setq
  (entmakex (list '(0 . "ARC")
		  '(100 . "AcDbEntity")
		  '(67 . 0)
		  '(100 . "AcDbCircle")
		  (cons 10 cent)
		  (cons 40 (* 0.5 d))
		  '(210 0.0 0.0 1.0)
		  '(100 . "AcDbArc")
		  (cons 62 cvet)
		  (cons 50 ugol1)
		  (cons 51 ugol2)
	    ); list
  ); entmakex
); defun

(defun add_dimarc (center a1 a2 tex)
  (vla-adddimarc (vla-get-ModelSpace (vla-get-ActiveDocument (vlax-get-acad-object)))
    		 (vlax-3d-point (trans center 1 0))
   		 (vlax-3d-point (trans a1 1 0))
    		 (vlax-3d-point (trans a2 1 0))
   		 (vlax-3d-point (trans a2 1 0))
  ); vla-adddimarc
  (mapcar '(lambda (x)
	     (vlax-put-property
	       (vlax-ename->vla-object (entlast))
	       (car x)
	       (cadr x)
	     )
	  )
	  (list (list "SymbolPosition" acSymNone)
		(list "TextPrefix" tex)
		(list "ExtensionLineExtend" 0)
		(list "ExtensionLineOffset" 0)
		(list "Textgap" 1.2)
	  )
  )
); defun

(defun add_diametr (centr d prefix ugol da)
  (vla-AddDimDiametric (vla-get-ModelSpace (vla-get-ActiveDocument (vlax-get-acad-object)))
    (vlax-3d-point (trans (polar centr ugol (* 0.5 d)) 1 0))
    (vlax-3d-point (trans (polar centr (+ pi ugol) (* 0.5 d)) 1 0))
    (* -0.15 da)
  )
  (mapcar '(lambda (x)
	     (vlax-put-property
	       (vlax-ename->vla-object (entlast))
	       (car x)
	       (cadr x)
	     )
	  )
	  (list (list "TextPrefix" (strcat "\\A1;d\\H0.7x;\\S^" prefix ";\\H1.42857x; "))		
		(list "Textgap" 1.3)
	  )
  )
)

(defun add_dimal (a1 a2 a3 tex)
   (vla-AddDimAligned (vla-get-ModelSpace (vla-get-ActiveDocument (vlax-get-acad-object)))
    		     (vlax-3d-point (trans a1 1 0))
    		     (vlax-3d-point (trans a2 1 0))
    		     (vlax-3d-point (trans a3 1 0))
  ); vla-AddDimAligned
  (vla-put-textprefix (vlax-ename->vla-object (entlast))
    tex
  ); vlax-put-property
  (vla-put-Textgap (vlax-ename->vla-object (entlast))
    1.2
  ); vla-put-Textgap
); defun

(defun add_ang (center a1 a2 tex)
  (vla-adddimangular (vla-get-ModelSpace (vla-get-ActiveDocument (vlax-get-acad-object)))
     (vlax-3d-point (trans center 1 0))
     (vlax-3d-point (trans a1 1 0))
     (vlax-3d-point (trans a2 1 0))
     (vlax-3d-point (trans (polar a1 (angle a1 a2) (* 0.5 (distance a1 a2))) 1 0))
  ); vla-adddimangular
  (vla-put-textprefix (vlax-ename->vla-object (entlast))
    tex
  ); vla-put-textprefix
  (vla-put-Textgap (vlax-ename->vla-object (entlast))
    1.2
  ); vla-put-Textgap
); defun

(defun add_dimdi (a1 a2 dlina tex)
  (vla-AddDimDiametric (vla-get-ModelSpace (vla-get-ActiveDocument (vlax-get-acad-object)))
    (vlax-3d-point (trans a1 1 0))
    (vlax-3d-point (trans a2 1 0))
    dlina
  ); vla-adddimordinate
  (vla-put-textprefix (vlax-ename->vla-object (entlast))
    tex
  ); vla-put-textprefix
); defun

(defun mak_mirr (obj p1 p2 /)
  (vla-mirror (vlax-ename->vla-object obj)
	      (vlax-3d-point (trans p1 1 0))
    	      (vlax-3d-point (trans p2 1 0))
  ); vla-mirror
); defun

(defun ugpoinv2 (inva / ugmax ugmin alfw)
  (setq ugmax (/ pi 6.)
	ugmin (/ pi 9.)
  )
  (repeat 100
    (setq alfw (* 0.5 (+ ugmin ugmax)))
    (if (> (inv alfw) inva)
      (setq ugmax alfw)
      (setq ugmin alfw)
    )
  )
  (setq alfw alfw)
);_end of defun

(defun tetdel (dt dbt)  
  (inv (arccos (/ dbt dt)))
)

(defun dl (db x d / rol)
  (setq rol (- (* 0.5 d (sin (/ pi 9.))) (/ (* (- 1 x) m) (sin (/ pi 9.)))))
  (sqrt (+ (* db db) (* 4 rol rol)))
)

(defun vvod (/ dety y)
  (initget 7)
  (setq m (getreal "\nVvedite modul zaceplenija v [mm]: "))  
  (initget 7)
  (setq z1 (getreal "\n[Z1]Vvedite chislo zubjev kolesa 1: "))  
  (initget 7)
  (setq z2 (getreal "\n[Z2]Vvedite chislo zubjev kolesa 2: "))  
 ;(initget 7)
  (setq x1 (getreal "\n[x1]Vvedite koefficient smeschenija kolesa 1: "))
  ;(initget 7)
  
  (setq x2 (getreal "\n[x2]Vvedite koefficient smeschenija kolesa 2: "))  
  (setq invalfw (+ (inv (/ pi 9.)) (/ (* 2. (tg (/ pi 9.)) (+ x1 x2)) (+ z1 z2)))
  ;(ugpoinv invalfw)
  	d1 (* m z1)
	d2 (* m z2)
	alfw (ugpoinv2 invalfw)
	aw (/ (* 0.5 m (+ z1 z2) (cos (/ pi 9.))) (cos alfw))
	dw1 (/ (* m z1 (cos (/ pi 9.))) (cos alfw))
	dw2 (/ (* m z2 (cos (/ pi 9.))) (cos alfw))
	y (/ (- aw (* 0.5 (+ d1 d2))) m)
	dety (- (+ x1 x2) y)
	da1 (+ d1 (* 2. m (+ x1 (- 1. dety))))
	da2 (+ d2 (* 2. m (- (1+ x2) dety)))
	df1 (- d1 (* 2. m (- 1.25 x1)))
	df2 (- d2 (* 2. m (- 1.25 x2)))
	db1 (* d1 (cos (/ pi 9.)))
	db2 (* d2 (cos (/ pi 9.)))
	rof (* 0.38 m)  
	dl1 (dl db1 x1 d1)
	dl2 (dl db2 x2 d2)
	p (* m pi)
	s1 (+ (* 0.5 pi m) (* 2 x1 m (tg (/ pi 9.))))
	s2 (+ (* 0.5 pi m) (* 2 x2 m (tg (/ pi 9.))))
	eps_a (/ (+ (sqrt (- (* da1 da1) (* db1 db1)))
		    (sqrt (- (* da2 da2) (* db2 db2)))
		    (* -1 (sin alfw) (+ dw1 dw2))
		 )
		 (* 2. p (cos (/ pi 9.)))
	      )
  ); setq
  (if (< eps_a 1.11)
    (alert "Koefficient torcevogo perekritija men'she 1,11!")
  ); if
)

(defun zerkal (obj)
  (vla-put-LineWeight (vlax-ename->vla-object obj)
    		aclnwt100
  ); vla-put-LineWeight
  (mak_mirr obj centr (polar centr (+ (- povor invalfw) (tetdel d db) (/ s d)) 5))
  (mak_mirr (entlast) centr (polar centr (+ (- povor invalfw (/ pi z)) (tetdel d db) (/ s d)) 5))
  (mak_mirr (entlast) centr (polar centr (+ (- povor invalfw (/ (* 2. pi) z)) (tetdel d db) (/ s d)) 5))
  (mak_mirr (entlast) centr (polar centr (+ (- povor invalfw) (tetdel d db) (/ s d)) 5))
  (mak_mirr (entlast) centr (polar centr (+ (- povor invalfw) (tetdel d db) (/ (* 2. pi) z) (/ s d)) 5))
)

(defun vseokr (/)
  (initget 65)
  (setq o1 (getpoint "Ukajite gde chertit' zaceplenie: ")
	o2 (polar o1 (* pi (/ 3. 2.)) aw))
  ;Koleso 1
  (add_arc2 o1 z1 d1 1 1)
  (add_arc2 o1 z1 dw1 1 3)
  (add_arc2 o1 z1 da1 1 5)
  (add_arc2 o1 z1 db1 1 6)
  (add_arc2 o1 z1 df1 1 7)
  (add_diametr o1 d1 "1" (- (* pi 1.4) (/ p d1 0.5)) da1)
  (add_diametr o1 dw1 "w1" (- (* pi 1.45) (/ p d1 0.5)) da1)  
  (add_diametr o1 da1 "a1" (- (* pi 1.5) (/ p d1 0.5)) da1) 
  (add_diametr o1 db1 "b1" (- (* pi 1.35) (/ p d1 0.5)) da1)  
  (add_diametr o1 df1 "f1" (- (* pi 1.3) (/ p d1 0.5)) da1)
  ; Koleso 2
  (add_arc2 o2 z2 d2 2 1)
  (add_arc2 o2 z2 dw2 2 3)
  (add_arc2 o2 z2 da2 2 5)
  (add_arc2 o2 z2 db2 2 6)
  (add_arc2 o2 z2 df2 2 7)
  (add_diametr o2 d2 "2" (+ (* pi 0.7) (/ p d2 0.5)) da2)  
  (add_diametr o2 dw2 "w2" (+ (* pi 0.65) (/ p d2 0.5)) da2)  
  (add_diametr o2 da2 "a2" (+ (* pi 0.6) (/ p d2 0.5)) da2) 
  (add_diametr o2 db2 "b2" (+ (* pi 0.75) (/ p d2 0.5)) da2)  
  (add_diametr o2 df2 "f2" (+ (* pi 0.8) (/ p d2 0.5)) da2)
  (li (list o1 o2) 7)
); defun

(defun profil (povor da dl db dw z tocho / sha tet rad)
  (initget 7)
  (setq sha (/ (* 0.5 (- da dl)) (getreal "\nTochek na evolvente ne menee: "))
	rad (* dl 0.5)
	mas nil
  ); setq
  (while (<= rad (* da 0.5))
    (setq tet (+ povor (- (inv (arccos (/ db rad 2.))) invalfw))
	  mas (cons (polar tocho tet rad) mas)
	  rad (+ rad sha)
    )
  ); while
  (setq mas (cons (polar tocho (+ povor (- (inv (arccos (/ db da))) invalfw)) (* da 0.5)) mas)
	mas (cons (polar tocho povor (* dw 0.5)) mas)
	mas (cons (polar tocho (+ povor (- (inv (/ pi 9.)) invalfw)) (* 0.5 m z)) mas)
	mas (vl-sort mas '(lambda (e1 e2) (> (car e1) (car e2))))
  )
  mas
); defun

(defun cherchu (massiv centr df db d dl z povor s da / gam ksi alfa_n ugdug rr)
  ;Переходная кривая
  (add_pl massiv) 
  (zerkal (entlast))
  (if (or (and (< (abs (* 0.5 (- df dl)))
	          rof
      	       ); <
	       (> db df)
      	  ); and
	  (< db df)
      ); or
         (if (equal d d1)
           (progn
     	     (setq ugdug (arccos (/ (+ (* df rof) (/ (* dl dl) 4) (/ (* df df) 4)) (* dl (+ rof (/ df 2))))))
     	     (add_arc (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug)  (+ rof (/ df 2)))
                      (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ dl 2))
	              (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug) (/ df 2))
	     );add_arc
	     (setq gam (- povor invalfw (* (tetdel dl db) -1) ugdug))
	     (zerkal (entlast))
    	   );progn
    	   (progn
      	     (setq ugdug (arccos (/ (+ (* df rof) (/ (* dl dl) 4) (/ (* df df) 4)) (* dl (+ rof (/ df 2))))))
      	     (add_arc (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug)  (+ rof (/ df 2)))
	              (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ dl 2))
	       	      (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug) (/ df 2))
      	     ); add_arc
	     (setq gam (- povor invalfw (* (tetdel dl db) -1) ugdug))
	     (zerkal (entlast))
   	   );progn
  	 );if
   (progn
     (setq rr (* (sqrt (- (expt (+ rof (/ df 2)) 2) (* rof rof))) 2))
         (if (equal d d1)
           (progn
     	     (setq ugdug (arccos (/ (+ (* df rof) (/ (* rr rr) 4) (/ (* df df) 4)) (* rr (+ rof (/ df 2))))))
      	     (add_pl (list (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ dl 2))
	     	     	   (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ rr 2))
		     ); list
	     )
	     (zerkal (entlast))
	     (add_arc (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug) (+ rof (/ df 2)))
		      (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ rr 2))
	              (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug) (/ df 2))
      	     ); add_arc
	     (setq gam (- povor invalfw (* (tetdel dl db) -1) ugdug))
	     (zerkal (entlast))
    	   );progn
    	   (progn
   	     (setq ugdug (arccos (/ (+ (* df rof) (/ (* rr rr) 4) (/ (* df df) 4)) (* rr (+ rof (/ df 2))))))
      	     (add_pl (list (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ dl 2))
                       	   (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ rr 2))
		     ); list
	     ); add_pl
	     (zerkal (entlast))
	     (add_arc  (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug)  (+ rof (/ df 2)))
		       (polar centr (- povor invalfw (* (tetdel dl db) -1)) (/ rr 2))
	       	       (polar centr (- povor invalfw (* (tetdel dl db) -1) ugdug) (/ df 2))
             ); add_arc
	     (setq gam (- povor invalfw (* (tetdel dl db) -1) ugdug))
	     (zerkal (entlast))
           );progn
         );if
   );progn
  ); if
  (add_arc centr
	   (polar centr (+ povor (- (tetdel da db) invalfw)) (* da 0.5))
	   (polar centr (+ povor (- (tetdel da db) invalfw) (/ (* da (+ (/ s d) (- (inv (/ pi 9)) (tetdel da db)))) (* 0.5 da))) (* da 0.5))
  ); add_arc
  (vla-put-LineWeight (vlax-ename->vla-object (entlast))
    aclnwt100
  ); vla-put-LineWeight
  (mak_mirr (entlast) centr (polar centr (+ (- povor invalfw (/ pi z)) (tetdel d db) (/ s d)) 5))
  (mak_mirr (entlast) centr (polar centr (+ (- povor invalfw) (tetdel d db) (/ s d)) 5))  
  (add_dimarc centr
      	       (polar centr (+ povor (/ p (* 0.5 d)) (- (inv (/ pi 9)) invalfw)) (* 0.5 d))
    	       (polar centr (+ povor (/ p (* 0.5 d)) (/ s (* 0.5 d)) (- (inv (/ pi 9)) invalfw)) (* 0.5 d))
    	       (strcat "\\A1;S\\H0.7x;\\S^"
		       (if (= d d1)
			 "1"
			 "2"
		       ); if
		       ";\\H1.42857x;="
	       ); strcat
  ); add_dimarc
  (add_dimarc centr
      	       (polar centr (+ povor (/ s (* 0.5 d)) (- (inv (/ pi 9)) invalfw)) (* 0.5 d))
    	       (polar centr (+ povor (/ p (* 0.5 d)) (- (inv (/ pi 9)) invalfw)) (* 0.5 d))
    	       (strcat "\\A1;e\\H0.7x;\\S^"
		       (if (= d d1)
			 "1"
			 "2"
		       ); if
		       ";\\H1.42857x;="
	       ); strcat
  ); add_dimarc
  (li (list centr (polar centr (+ povor (/ p (* -0.5 d)) (- (inv (/ pi 9)) invalfw)) (* 0.5 d))) 7)
  (li (list centr (polar centr (+ povor (- (inv (/ pi 9)) invalfw)) (* 0.5 d))) 7)
  (add_ang centr
	   (polar centr (+ povor (/ p (* -0.5 d)) (- (inv (/ pi 9)) invalfw)) (* 0.25 d))
	   (polar centr (+ povor (- (inv (/ pi 9)) invalfw)) (* 0.25 d))
	   (strcat "\\A1;\U+03C4\\H0.7x;\\S^"
		   (if (= d d1)
			 "1"
			 "2"
		   ); if
		   ";\\H1.42857x;="
   	   ); strcat
  ); add_ang
  (setq ksi (- (+ (- povor invalfw (/ pi z)) (tetdel d db) (/ s d))
	      gam
	   )
	alfa_n (- (/ pi z 0.5)
		  (* 2. ksi)
		  (/ p d 0.5)
	       )
  )
  (add_arc centr
	   (polar centr (- gam alfa_n) (/ df 2.))
	   (polar centr gam (/ df 2.))
  )
  (vla-put-LineWeight (vlax-ename->vla-object (entlast))
    		aclnwt100
  ); vla-put-LineWeight
  (mak_mirr (entlast) centr (polar centr (+ (- povor invalfw) (tetdel d db) (/ s d)) 5))
)

(defun liam (/ tochka shagl x labpr shag2 x2 l1 l2 liam1 liam2 liam1_ex liam2_ex
	     liam1ob liam2ob)
  (initget 7)
  (setq liam1 nil
	liam2 nil
	liam1_ex nil
	liam2_ex nil
	labpr (* (+ db1 db2) (tg alfw) 0.5)
	shagl (/ labpr (getreal "\nTochek na krivikh skoljenija ne menee: "))
  ); setq
  (initget 64)
  (setq x (* 4 shagl)
	shag2 (/ labpr 9.0)
	x2 shag2
  ); setq
  (initget 7)
  (setq	muliam (getreal "\nVvedite vo skol'ko raz uvelichivat' epuri skoljenija: ")
  );setq
  (while (< x (- labpr (* 4 shagl)))
    (setq l1 (+ 1 (/ z2 z1) (* -1 (/ z2 z1) (/ labpr x)))
	  l2 (+ 1 (/ z1 z2) (* -1 (/ z1 z2) (/ labpr (- labpr x))))
	  liam1 (cons (list x (* muliam l1)) liam1)
	  liam2 (cons (list x (* muliam l2)) liam2)
	  x (+ x shagl)
    )
  );while
  (setq liam1 (cons (list (/ (* z2 labpr) (+ z2 z1)) 0.) liam1)
	liam1 (vl-sort liam1 '(lambda (e1 e2) (> (car e1) (car e2))))
	liam2 (cons (list (/ (* z2 labpr) (+ z2 z1)) 0.) liam2)
	liam2 (vl-sort liam2 '(lambda (e1 e2) (> (car e1) (car e2))))
  )
  (while (< x2 labpr)
    (setq l1 (+ 1 (/ z2 z1) (* -1 (/ z2 z1) (/ labpr x2)))
	  l2 (+ 1 (/ z1 z2) (* -1 (/ z1 z2) (/ labpr (- labpr x2))))
	  liam1_ex (cons (list x2 l1) liam1_ex)
	  liam2_ex (cons (list x2 l2) liam2_ex)
	  x2 (+ x2 shag2)
    )
  );while
  (setq liam1_ex (reverse liam1_ex))
  (li (list o2 (polar o2 alfw labpr)) 7)
  (add_pl liam1)
  (setq liam1ob (vlax-ename->vla-object (entlast)))
  (add_pl liam2)
  (setq liam2ob (vlax-ename->vla-object (entlast)))
  (mapcar '(lambda (x)
	     (vla-rotate x
	       (vlax-3d-point (trans '(0 0) 1 0))
	       alfw
	     ); vla-rotate
	   ); lambda
	  (list liam1ob liam2ob)
  )
  (mapcar '(lambda (x)
	     (vla-move x
		       (vlax-3d-point (trans '(0 0) 1 0))
		       (vlax-3d-point (trans o2 1 0))
	     )
	   ); lambda
	  (list liam1ob liam2ob)
  )
); defun

(defun c:zubex (/ activedoc o old oldt i j mas0 mas1 spiska
		da1 d1 df1 dl1 db1 dw1 o1 s1
		da2 d2 df2 dl2 db2 dw2 o2 s2
		invalfw ;z1 z2 m x1 x2 alfw
		p eps_a rof aw zacep_dial)
  (vl-load-com)
  (setq old (getvar "OSMODE")
	activedoc (vla-get-ActiveDocument (vlax-get-acad-object))
  )
  (setvar "OSMODE" 20903)
  (setq oldt (entget (tblobjname "style" "standard"))
	oldt (vl-remove (assoc 3 oldt) oldt)
	oldt (append oldt (list (cons 3 "TIMESI.TTF" )))
  );_end of setq
  (entmod oldt)  
  (vvod)
  (vseokr)  
  (cherchu (profil (* pi -0.5) da1 dl1 db1 dw1 z1 o1)
	   o1 df1 db1 d1 dl1 z1 (* pi 1.5) s1 da1
  )
  (cherchu (profil (* pi 0.5) da2 dl2 db2 dw2 z2 o2)
	   o2 df2 db2 d2 dl2 z2 (* pi 0.5) s2 da2
  )
  (setq o (polar o2 (* 0.5 pi) (* 0.5 dw2)))  
  (li (list o (polar o (- alfw (/ pi 2)) (/ (* 0.5 dw2) (cos alfw)))) 7)
  (li (list (polar o 0. 50) (polar o 0. -50)) 3)
  (li (list (polar o alfw (+ (* 0.5 dw1 (sin alfw)) 10)) (polar o (+ pi alfw) (+ (* 0.5 dw2 (sin alfw)) 10))) 6)  
  (add_tex (polar o alfw (* 0.5 dw1 (sin alfw))) "B")
  (add_tex (polar o (+ pi alfw) (* 0.5 dw2 (sin alfw))) "A")
  (li (list o1 (polar o alfw (* 0.5 dw1 (sin alfw)))) 7)
  (li (list o2 (polar o (+ pi alfw) (* 0.5 dw2 (sin alfw)))) 7)
  (add_tex o1 "\\A1;O\\H0.7x;\\S^1;\\H1.4286x;")
  (add_tex o2 "\\A1;O\\H0.7x;\\S^2;\\H1.4286x;")
  (add_tex o "p")
  (add_dimarc o2
      	      (polar o2 (+ (* 0.5 pi) (/ p (* -0.5 d2)) (- (inv (/ pi 9)) invalfw)) (* 0.5 d2))
    	      (polar o2 (+ (* 0.5 pi) (- (inv (/ pi 9)) invalfw)) (* 0.5 d2))
    	      "p="
  ); add_dimarc
  (initget 1)
  (add_dimal o1 o2    
    (getpoint "\nMejosevoe rasstojanie: ")
    "\\A1;a\\H0.7x;\\S^W;\\H1.42857x;="
  ); add_dimal
  (initget 1)
  (add_dimal (polar o1 (* -0.5 pi) (* 0.5 da1))
    (polar o2 (* 0.5 pi) (* 0.5 df2))
    (getpoint "\nRadial'ni zazor: ")
    "C="
  ); add_dimal
  (liam)
  (setvar "OSMODE" old)
  (initget 64)
  (vla-addtable (vla-get-ModelSpace activedoc)
    (vlax-3d-point (trans (getpoint "\nTablica osnovnikh parametrov: ") 1 0))
    2 9 10. 11.5
  ); vla-addtable
  (vla-UnmergeCells (vlax-ename->vla-object (entlast))
    0 0 1 9
  ); vla-UnmergeCells
  (vla-settextheight (vlax-ename->vla-object (entlast))
    actitlerow 2.5
  ); vla-settextheight
  (vla-settextheight (vlax-ename->vla-object (entlast))
    acheaderrow 2.5
  ); vla-settextheight
  (setq i 0
	j 0
	mas0 '("m, [mm]"
	       "\\A1;Z\\H0.7x;\\S^1;"
	       "\\A1;Z\\H0.7x;\\S^2;"
	       "\U+03B1"
	       "\\A1;h\\H0.7x;\\S^a;"
	       "C*"
	       "\\A1;x\\H0.7x;\\S^1;"
	       "\\A1;x\\H0.7x;\\S^2;"
	       "\\A1;\U+03B5\\H0.7x;\\S^a;"
	       )
	mas1 (list
	       (rtos m 2 1)
	       (rtos z1 2 0)
	       (rtos z2 2 0)
	       "20%%D"
	       "1"
	       "0.25"
	       (rtos x1 2 3)
	       (rtos x2 2 3)
	       (rtos eps_a 2 2)
	      )
  )
  (repeat 2
    (repeat 9
      (vla-settext (vlax-ename->vla-object (entlast))
        j i (nth i (eval (read (strcat "mas" (itoa j)))))
      ); vla-settext
      (setq i (1+ i))
    ); repeat
    (setq j (1+ j)
	  i 0
    )
  ); repeat
  (initget 1 "Yes No")
  (if (equal "Yes" (getkword "\nExport parametrov v excel? [Yes/No]: "))
    (progn
      (defun exeee ( / ref)
  	(setq ref (getfiled "Ukajite file s shablonom" "c:\\Zacep.xlsx" "xlsx" 128))
  	(ex_set_connect ref)
  	(ex_put m "K7")
  	(ex_put z1 "K2")
  	(ex_put z2 "M2")
  	(ex_put x1 "K3")
 	(ex_put x2 "M3")
  	(ex_put alfw "K5")
  	(ex_break_connect (getfiled "Gde sokhranit' file s tablicei" ref "xlsx" 129))
      ); defun
      (exeee)
    )
  )
)
(c:zubex)
