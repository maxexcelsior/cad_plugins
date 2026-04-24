;;; --- JZDYZ: Boundary Point Annotation Command ---
;;; --- Author: CMX | Date: 2026 ---
;;; --- Commands: JZDYZ, JZD_TABLE, JZD_PTS, JZD_CL, JZD_INFO ---
;;; --- Note: ANSI encoding required for Chinese text in CAD ---

(defun c:JZDYZ (/ acad_obj doc mspace lay_id ss i ent pts raw_pts clean_pts txt_h pt pt_text mleader_obj pt_list x_coord y_coord choice_mleader choice_pt pt_obj fuzz exists)
  (setvar "cmdecho" 0)
  (vl-load-com)

  ;; --- 1. Environment Initialization ---
  (setq acad_obj (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acad_obj))
  (setq mspace (vla-get-ModelSpace doc))
  (setq lay_id "BoundaryPointNo")
  (setq fuzz 0.001)
  
  (if (not (regapp "JZD_COORD_DATA")) (regapp "JZD_COORD_DATA"))
  (if (not (regapp "JZD_POINT_DATA")) (regapp "JZD_POINT_DATA"))

  (if (not (tblsearch "LAYER" lay_id)) (command "-layer" "n" lay_id "c" "7" lay_id ""))
  (if (not (tblsearch "STYLE" "HZ_GB")) (command "-style" "HZ_GB" "txt.shx,gbcbig.shx" "0" "0.8" "0" "n" "n" "n"))

  ;; --- 2. Select and Extract Unique Coordinates ---
  (princ "\n请选择多段线 (LWPOLYLINE): ")
  (setq ss (ssget '((0 . "LWPOLYLINE"))))
  
  (if ss
    (progn
      (setq raw_pts '() i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq pts (mapcar '(lambda (x) (list (cadr x) (caddr x) 0.0)) 
                    (vl-remove-if-not '(lambda (x) (= (car x) 10)) (entget ent))))
        (setq raw_pts (append raw_pts pts))
        (setq i (1+ i))
      )

      ;; Remove duplicate points
      (setq clean_pts '())
      (foreach p raw_pts
        (setq exists nil)
        (foreach cp clean_pts (if (equal p cp fuzz) (setq exists t)))
        (if (not exists) (setq clean_pts (cons p clean_pts)))
      )
      (setq clean_pts (reverse clean_pts))

      (setq txt_h (getdist "\n请输入标注文字高度 <2.5>: "))
      (if (null txt_h) (setq txt_h 2.5))

      ;; Ask user if generate MLeader
      (initget "Yes No")
      (setq choice_mleader (getkword "\n是否生成引线标注? [是(Y)/否(N)] <Y>: "))
      (if (null choice_mleader) (setq choice_mleader "Yes"))

      ;; Ask user if generate POINT entities
      (initget "Yes No")
      (setq choice_pt (getkword "\n是否在拐点生成POINT实体? [是(Y)/否(N)] <Y>: "))
      (if (null choice_pt) (setq choice_pt "Yes"))

      (setq i 1)
      (foreach pt clean_pts
        ;; CAD X = Horizontal, CAD Y = Vertical
        ;; Surveying: Y(East) = CAD X, X(North) = CAD Y
        (setq x_coord (rtos (car pt) 2 3))   ;; CAD X -> Horizontal (East)
        (setq y_coord (rtos (cadr pt) 2 3))  ;; CAD Y -> Vertical (North)
        
        ;; Generate MLeader if user selected Yes
        (if (= choice_mleader "Yes")
          (progn
            (setq pt_text (list (+ (car pt) (* txt_h 2.0)) (+ (cadr pt) (* txt_h 2.0)) 0.0))
            (setq pt_list (vlax-make-safearray vlax-vbDouble '(0 . 5)))
            (vlax-safearray-fill pt_list (append pt pt_text))
            (setq mleader_obj (vla-AddMLeader mspace (vlax-make-variant pt_list) 0))
            
            (vla-put-ContentType mleader_obj 2)
            (vla-put-TextString mleader_obj (strcat "J" (itoa i)))
            (vla-put-TextHeight mleader_obj txt_h)
            (vla-put-Layer mleader_obj lay_id)
            (vla-put-ArrowheadType mleader_obj 19)
            (vla-put-ArrowheadSize mleader_obj (/ txt_h 2.5))
            (vla-put-DoglegLength mleader_obj 0.0)

            ;; Store extended data: PointNo, Y(Vertical), X(Horizontal)
            (setq ent (entget (vlax-vla-object->ename mleader_obj)))
            (setq ent (append ent (list (list -3 (list "JZD_COORD_DATA" 
              (cons 1000 (strcat "J" (itoa i)))
              (cons 1000 y_coord)
              (cons 1000 x_coord))))))
            (entmod ent)
          )
        )

        ;; Generate POINT entity if user selected Yes
        (if (= choice_pt "Yes")
          (progn
            (setq pt_obj (vla-AddPoint mspace (vlax-3d-point pt)))
            (vla-put-Layer pt_obj lay_id)
            (setq ent (entget (vlax-vla-object->ename pt_obj)))
            (setq ent (append ent (list (list -3 (list "JZD_POINT_DATA"
              (cons 1000 (strcat "J" (itoa i)))
              (cons 1000 y_coord)
              (cons 1000 x_coord))))))
            (entmod ent)
          )
        )
        (setq i (1+ i))
      )
      (princ (strcat "\n[完成] 原始点数: " (itoa (length raw_pts)) " | 去重后点数: " (itoa (length clean_pts))))
    )
  )
  (princ)
)

;;; --- JZD_TABLE: Generate Boundary Point Coordinate Table ---
(defun c:JZD_TABLE (/ acad_obj doc mspace ss ss_ml i ent data xdata_list pt_list sort_list pt_ins obj_tbl row txt_h pt_no pt_x pt_y ml_obj col_count pts_per_col total_pts col row_offset col_width row_height)
  (vl-load-com)
  (setq acad_obj (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acad_obj))
  (setq mspace (vla-get-ModelSpace doc))
  
  (princ "\n正在搜索界址点数据...")
  
  ;; 1. Get text height from MLeader annotation
  (setq ss_ml (ssget "X" '((0 . "MULTILEADER") (-3 ("JZD_COORD_DATA")))))
  (if ss_ml
    (progn
      (setq ml_obj (vlax-ename->vla-object (ssname ss_ml 0)))
      (setq txt_h (vla-get-TextHeight ml_obj))
    )
    (setq txt_h 2.5)
  )
  
  (princ (strcat "\n文字高度: " (rtos txt_h 2 2)))
  
  ;; 2. Collect all POINT entities with extended data
  (setq ss (ssget "X" '((0 . "POINT") (-3 ("JZD_POINT_DATA")))))
  
  (if ss
    (progn
      (setq pt_list '() i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq data (entget ent '("JZD_POINT_DATA")))
        (setq xdata_list (cdr (cadr (assoc -3 data))))
        (setq pt_no (cdr (nth 0 xdata_list)))
        (setq pt_y (cdr (nth 1 xdata_list)))  ;; Y = Vertical (North)
        (setq pt_x (cdr (nth 2 xdata_list)))  ;; X = Horizontal (East)
        (if (and pt_no pt_x pt_y)
          (setq pt_list (cons (list pt_no pt_y pt_x) pt_list))
        )
        (setq i (1+ i))
      )

      (princ (strcat "\n找到 " (itoa (length pt_list)) " 个界址点。"))

      ;; 3. Sort by point number (J1, J2, J3...)
      (if (> (length pt_list) 0)
        (setq sort_list (vl-sort pt_list 
          '(lambda (a b) 
            (< (atoi (substr (car a) 2)) (atoi (substr (car b) 2)))))
        )
      )

      ;; 4. Ask user for points per column
      (setq pts_per_col (getint "\n请输入每列拐点数量 <20>: "))
      (if (null pts_per_col) (setq pts_per_col 20))
      
      (setq total_pts (length sort_list))
      (setq col_count (1+ (fix (/ (- total_pts 1) pts_per_col))))
      
      (princ (strcat "\n共 " (itoa total_pts) " 个点，分 " (itoa col_count) " 列显示。"))

      ;; 5. Create table
      (if (> (length sort_list) 0)
        (progn
          (setq pt_ins (getpoint "\n请指定表格插入点: "))
          (if pt_ins
            (progn
              (setq row_height (* txt_h 2.5))
              (setq col_width (* txt_h 12.0))
              
              ;; Create table: rows = pts_per_col + 2 (title + header), cols = col_count * 3
              (setq obj_tbl (vla-addtable mspace 
                (vlax-3d-point pt_ins) 
                (+ pts_per_col 2) 
                (* col_count 3) 
                row_height 
                col_width))
              
              ;; Set column widths for each column group
              (setq i 0)
              (repeat col_count
                (vla-SetColumnWidth obj_tbl (* i 3) (* txt_h 6.0))     ;; Point No.
                (vla-SetColumnWidth obj_tbl (+ (* i 3) 1) (* txt_h 12.0)) ;; Vertical Y
                (vla-SetColumnWidth obj_tbl (+ (* i 3) 2) (* txt_h 12.0)) ;; Horizontal X
                (setq i (1+ i))
              )

              ;; Set text height for all rows
              (vla-SetTextHeight obj_tbl acTitleRow txt_h)
              (vla-SetTextHeight obj_tbl acHeaderRow txt_h)
              (vla-SetTextHeight obj_tbl acDataRow txt_h)

              ;; Fill title row
              (vla-SetText obj_tbl 0 0 "界址点坐标表")
              
              ;; Fill header rows for each column
              (setq i 0)
              (repeat col_count
                (vla-SetText obj_tbl 1 (* i 3) "拐点序号")
                (vla-SetText obj_tbl 1 (+ (* i 3) 1) "纵坐标(Y)")
                (vla-SetText obj_tbl 1 (+ (* i 3) 2) "横坐标(X)")
                (setq i (1+ i))
              )

              ;; Fill data rows - distribute points across columns
              (setq row 2)
              (setq col 0)
              (setq row_offset 0)
              (foreach item sort_list
                ;; Calculate current column and row
                (setq col (fix (/ row_offset pts_per_col)))
                (setq row (+ 2 (- row_offset (* col pts_per_col))))
                
                ;; Fill data in correct column position
                (vla-SetText obj_tbl row (* col 3) (car item))      ;; Point No.
                (vla-SetText obj_tbl row (+ (* col 3) 1) (cadr item)) ;; Y (Vertical)
                (vla-SetText obj_tbl row (+ (* col 3) 2) (caddr item)) ;; X (Horizontal)
                
                (setq row_offset (1+ row_offset))
              )
              
              (princ (strcat "\n[成功] 已生成界址点坐标表，共 " (itoa total_pts) " 个点。"))
            )
          )
        )
        (princ "\n[错误] 未提取到有效的界址点数据。")
      )
    )
    (princ "\n[错误] 未找到界址点数据。请先运行 JZDYZ 命令并选择生成POINT实体。")
  )
  (princ)
)

;;; --- JZD_PTS: Check Polyline Vertex Count ---
(defun c:JZD_PTS (/ ss i ent pts)
  (princ "\n请选择要检查节点数的多段线: ")
  (if (setq ss (ssget '((0 . "LWPOLYLINE"))))
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq pts (vl-remove-if-not '(lambda (x) (= (car x) 10)) (entget ent)))
        (princ (strcat "\n对象 " (itoa (1+ i)) " (句柄 " (cdr (assoc 5 (entget ent))) "): 节点总数 = " (itoa (length pts))))
        (setq i (1+ i))
      )
    )
    (princ "\n未选中多段线。")
  )
  (princ)
)

;;; --- JZD_CL: Clean Duplicate Vertices ---
(defun c:JZD_CL (/ ss i ent elist pts p clean_pts fuzz)
  (setq fuzz 0.0001)
  (princ "\n请选择要清理重叠节点的多段线: ")
  (if (setq ss (ssget '((0 . "LWPOLYLINE"))))
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq elist (entget ent))
        (setq pts '())
        (foreach x elist (if (= (car x) 10) (setq pts (cons (cdr x) pts))))
        (setq pts (reverse pts))
        (setq clean_pts (list (car pts)))
        (foreach p (cdr pts)
          (if (not (equal p (car clean_pts) fuzz)) (setq clean_pts (cons p clean_pts)))
        )
        (setq clean_pts (reverse clean_pts))
        (if (/= (length pts) (length clean_pts))
          (progn
            (setq elist (vl-remove-if '(lambda (x) (= (car x) 10)) elist))
            (foreach p clean_pts (setq elist (append elist (list (cons 10 p)))))
            (entmod elist)
            (princ (strcat "\n对象 " (itoa (1+ i)) ": 清理了 " (itoa (- (length pts) (length clean_pts))) " 个重复节点。"))
          )
          (princ (strcat "\n对象 " (itoa (1+ i)) ": 无重复节点。"))
        )
        (setq i (1+ i))
      )
    )
  )
  (princ)
)

;;; --- JZD_INFO: Query Point Data ---
(defun c:JZD_INFO (/ ent data xdata)
  (if (setq ent (car (entsel "\n请点击标注或点实体查看数据: ")))
    (progn
      (setq data (entget ent '("JZD_COORD_DATA" "JZD_POINT_DATA")))
      (setq xdata (assoc -3 data))
      (if xdata
        (progn
          (princ "\n========== 界址点数据 ==========")
          (foreach app_list (cdr xdata)
            (foreach itm (cdr app_list)
              (if (= (car itm) 1000) (princ (strcat "\n>> " (cdr itm))))
            )
          )
          (princ "\n=================================")
        )
        (princ "\n[提示] 该实体不包含界址点扩展数据。")
      )
    )
  )
  (princ)
)
