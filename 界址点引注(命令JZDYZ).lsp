;;; --- JZDYZ: 界址点标注与属性工具 (表头优化版) ---
;;; --- 指令列表: JZDYZ, JZD_TABLE, JZD_PTS, JZD_CL, JZD_INFO ---

(defun c:JZDYZ (/ acad_obj doc mspace lay_id ss i ent pts raw_pts clean_pts txt_h pt pt_text mleader_obj pt_list x_coord y_coord choice_mleader choice_pt pt_obj fuzz exists mtxt_obj)
  (setvar "cmdecho" 0)
  (vl-load-com)

  ;; --- 1. 环境初始化 ---
  (setq acad_obj (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acad_obj))
  (setq mspace (vla-get-ModelSpace doc))
  (setq lay_id "界址点序号")
  (setq fuzz 0.001)
  
  (if (not (regapp "JZD_COORD_DATA")) (regapp "JZD_COORD_DATA"))
  (if (not (regapp "JZD_POINT_DATA")) (regapp "JZD_POINT_DATA"))

  (if (not (tblsearch "LAYER" lay_id)) (command "-layer" "n" lay_id "c" "7" lay_id ""))
  (if (not (tblsearch "STYLE" "HZ_GB")) (command "-style" "HZ_GB" "txt.shx,gbcbig.shx" "0" "0.8" "0" "n" "n" "n"))

  ;; --- 2. 选择并提取去重坐标 ---
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

      (setq clean_pts '())
      (foreach p raw_pts
        (setq exists nil)
        (foreach cp clean_pts (if (equal p cp fuzz) (setq exists t)))
        (if (not exists) (setq clean_pts (cons p clean_pts)))
      )
      (setq clean_pts (reverse clean_pts))

      (setq txt_h (getdist "\n请输入标注文字高度 <2.5>: "))
      (if (null txt_h) (setq txt_h 2.5))

      (initget "Yes No")
      (setq choice_mleader (getkword "\n是否生成引线? [是(Y)/否(N)] <Y>: "))
      (if (null choice_mleader) (setq choice_mleader "Yes"))

      (initget "Yes No")
      (setq choice_pt (getkword "\n是否在对应拐点生成 POINT 实体? [是(Y)/否(N)] <Y>: "))
      (if (null choice_pt) (setq choice_pt "Yes"))

      (setq i 1)
      (foreach pt clean_pts
        (setq x_coord (rtos (car pt) 2 3))
        (setq y_coord (rtos (cadr pt) 2 3))
        
        (if (= choice_mleader "Yes")
          (progn
            (setq pt_text (list (+ (car pt) (* txt_h 1.2)) (+ (cadr pt) (* txt_h 1.2)) 0.0))
            (setq pt_list (vlax-make-safearray vlax-vbDouble '(0 . 5)))
            (vlax-safearray-fill pt_list (append pt pt_text))
            (setq mleader_obj (vla-AddMLeader mspace (vlax-make-variant pt_list) 0))
            (vla-put-TextString mleader_obj (strcat "J" (itoa i)))
            (vla-put-TextHeight mleader_obj txt_h)
            (vla-put-Layer mleader_obj lay_id)
            (vla-put-ArrowheadType mleader_obj 19)
            (vla-put-ArrowheadSize mleader_obj (/ txt_h 2.5))
            (vla-put-DoglegLength mleader_obj 0.0)
            
            (setq ent (entget (vlax-vla-object->ename mleader_obj)))
            (setq ent (append ent (list (list -3 (list "JZD_COORD_DATA" (cons 1000 (strcat "J" (itoa i))) (cons 1000 y_coord) (cons 1000 x_coord))))))
            (entmod ent)
          )
          (progn
            (setq mtxt_obj (vla-AddMText mspace (vlax-3d-point pt) 0.0 (strcat "J" (itoa i))))
            (vla-put-Height mtxt_obj txt_h)
            (vla-put-Layer mtxt_obj lay_id)
            (vla-put-AttachmentPoint mtxt_obj 5) 
            (vla-put-InsertionPoint mtxt_obj (vlax-3d-point pt))
            
            (setq ent (entget (vlax-vla-object->ename mtxt_obj)))
            (setq ent (append ent (list (list -3 (list "JZD_COORD_DATA" (cons 1000 (strcat "J" (itoa i))) (cons 1000 y_coord) (cons 1000 x_coord))))))
            (entmod ent)
          )
        )

        (if (= choice_pt "Yes")
          (progn
            (setq pt_obj (vla-AddPoint mspace (vlax-3d-point pt)))
            (vla-put-Layer pt_obj lay_id)
            (setq ent (entget (vlax-vla-object->ename pt_obj)))
            (setq ent (append ent (list (list -3 (list "JZD_POINT_DATA" (cons 1000 (strcat "J" (itoa i))) (cons 1000 y_coord) (cons 1000 x_coord))))))
            (entmod ent)
          )
        )
        (setq i (1+ i))
      )
      (princ (strcat "\n[完成] 处理点数: " (itoa (length clean_pts))))
    )
  )
  (princ)
)

;;; --- JZD_TABLE: 生成界址点坐标成果表 (样式深度定制版) ---
(defun c:JZD_TABLE (/ acad_obj doc mspace ss i ent data xdata_list pt_list sort_list pt_ins obj_tbl row txt_h pt_no pt_x pt_y col_count pts_per_col total_pts col row_offset col_width row_height r c)
  (vl-load-com)
  (setq acad_obj (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acad_obj))
  (setq mspace (vla-get-ModelSpace doc))
  
  (princ "\n正在搜索界址点数据...")
  (setq ss (ssget "X" '((0 . "POINT") (-3 ("JZD_POINT_DATA")))))
  
  (if ss
    (progn
      (setq pt_list '() i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq data (entget ent '("JZD_POINT_DATA")))
        (setq xdata_list (cdr (cadr (assoc -3 data))))
        (setq pt_list (cons (list (cdr (nth 0 xdata_list)) (cdr (nth 1 xdata_list)) (cdr (nth 2 xdata_list))) pt_list))
        (setq i (1+ i))
      )

      ;; 按点号数字排序
      (setq sort_list (vl-sort pt_list '(lambda (a b) (< (atoi (substr (car a) 2)) (atoi (substr (car b) 2))))))
      
      (setq pts_per_col (getint "\n请输入每栏行数 <20>: "))
      (if (null pts_per_col) (setq pts_per_col 20))
      
      (setq total_pts (length sort_list))
      (setq col_count (1+ (fix (/ (- total_pts 1) pts_per_col))))
      (setq txt_h 2.5) 

      (setq pt_ins (getpoint "\n指定插入点: "))
      (if pt_ins
        (progn
          (setq row_height (* txt_h 2.5) col_width (* txt_h 12.0))
          (setq obj_tbl (vla-addtable mspace (vlax-3d-point pt_ins) (+ pts_per_col 2) (* col_count 3) row_height col_width))
          
          ;; 全表单元格初始化（对齐、字高、字体）
          (setq r 0)
          (repeat (+ pts_per_col 2)
            (setq c 0)
            (repeat (* col_count 3)
              (vla-SetCellAlignment obj_tbl r c 5) ; 正中对齐
              (vla-SetCellTextHeight obj_tbl r c txt_h)
              (vla-SetCellTextStyle obj_tbl r c "HZ_GB")
              (setq c (1+ c))
            )
            (setq r (1+ r))
          )

          ;; --- 修改点 1: 更新大标题 ---
          (vla-SetText obj_tbl 0 0 "界址点坐标表（广州2000坐标系）")
          
          ;; --- 修改点 2: 简化表头文字 ---
          (setq i 0)
          (repeat col_count
            (vla-SetText obj_tbl 1 (* i 3) "点号")
            (vla-SetText obj_tbl 1 (+ (* i 3) 1) "纵坐标")
            (vla-SetText obj_tbl 1 (+ (* i 3) 2) "横坐标")
            (setq i (1+ i))
          )

          ;; 数据填充
          (setq row_offset 0)
          (foreach item sort_list
            (setq col (fix (/ row_offset pts_per_col)))
            (setq row (+ 2 (- row_offset (* col pts_per_col))))
            (vla-SetText obj_tbl row (* col 3) (car item))
            (vla-SetText obj_tbl row (+ (* col 3) 1) (cadr item))
            (vla-SetText obj_tbl row (+ (* col 3) 2) (caddr item))
            (setq row_offset (1+ row_offset))
          )
          (princ "\n[成功] 成果表已生成。")
        )
      )
    )
    (princ "\n错误: 未在图中找到界址点属性数据。")
  )
  (princ)
)

;;; --- 辅助工具 (JZD_PTS, JZD_CL, JZD_INFO) 保持不变 ---
(defun c:JZD_PTS (/ ss i ent pts)
  (if (setq ss (ssget '((0 . "LWPOLYLINE"))))
    (progn (setq i 0) (repeat (sslength ss) (setq ent (ssname ss i)) (setq pts (vl-remove-if-not '(lambda (x) (= (car x) 10)) (entget ent))) (princ (strcat "\n多段线 " (itoa (1+ i)) ": 节点数 = " (itoa (length pts)))) (setq i (1+ i))))
  ) (princ)
)
(defun c:JZD_CL (/ ss i ent elist pts p clean_pts fuzz)
  (setq fuzz 0.0001)
  (if (setq ss (ssget '((0 . "LWPOLYLINE"))))
    (progn (setq i 0) (repeat (sslength ss) (setq ent (ssname ss i) elist (entget ent) pts '()) (foreach x elist (if (= (car x) 10) (setq pts (cons (cdr x) pts)))) (setq pts (reverse pts) clean_pts (list (car pts))) (foreach p (cdr pts) (if (not (equal p (car clean_pts) fuzz)) (setq clean_pts (cons p clean_pts)))) (setq clean_pts (reverse clean_pts)) (if (/= (length pts) (length clean_pts)) (progn (setq elist (vl-remove-if '(lambda (x) (= (car x) 10)) elist)) (foreach p clean_pts (setq elist (append elist (list (cons 10 p))))) (entmod elist) (princ (strcat "\n多段线 " (itoa (1+ i)) ": 已清理冗余点。")))) (setq i (1+ i))))
  ) (princ)
)
(defun c:JZD_INFO (/ ent data xdata)
  (if (setq ent (car (entsel "\n查询界址点属性: ")))
    (progn (setq data (entget ent '("JZD_COORD_DATA" "JZD_POINT_DATA")) xdata (assoc -3 data)) (if xdata (progn (princ "\n--- 属性 ---") (foreach app_list (cdr xdata) (foreach itm (cdr app_list) (if (= (car itm) 1000) (princ (strcat "\n>> " (cdr itm))))))) (princ "\n无数据。")))
  ) (princ)
)