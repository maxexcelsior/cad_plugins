(defun c:JZDYZ (/ acad_obj doc mspace lay_id ss i ent pts raw_pts clean_pts txt_h pt pt_text mleader_obj pt_list x_coord y_coord choice pt_obj fuzz exists)
  (setvar "cmdecho" 0)
  (vl-load-com)

  ;; --- 1. 环境初始化 ---
  (setq acad_obj (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acad_obj))
  (setq mspace (vla-get-ModelSpace doc))
  (setq lay_id "界址点序号")
  (setq fuzz 0.001) ; 坐标去重精度
  
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

      ;; --- 修复逻辑：使用 foreach 替代 member-if 进行全局去重 ---
      (setq clean_pts '())
      (foreach p raw_pts
        (setq exists nil)
        (foreach cp clean_pts
          (if (equal p cp fuzz) (setq exists t))
        )
        (if (not exists) (setq clean_pts (cons p clean_pts)))
      )
      (setq clean_pts (reverse clean_pts))

      (setq txt_h (getdist "\n请输入标注文字高度 <2.5>: "))
      (if (null txt_h) (setq txt_h 2.5))

      (initget "Yes No")
      (setq choice (getkword "\n是否在对应拐点生成 POINT 实体? [是(Y)/否(N)] <Y>: "))
      (if (null choice) (setq choice "Yes"))

      (setq i 1)
      (foreach pt clean_pts
        (setq x_coord (rtos (cadr pt) 2 3))
        (setq y_coord (rtos (car pt) 2 3))
        (setq pt_text (list (+ (car pt) (* txt_h 2.0)) (+ (cadr pt) (* txt_h 2.0)) 0.0))
        
        (setq pt_list (vlax-make-safearray vlax-vbDouble '(0 . 5)))
        (vlax-safearray-fill pt_list (append pt pt_text))
        (setq mleader_obj (vla-AddMLeader mspace (vlax-make-variant pt_list) 0))
        
        (vla-put-ContentType mleader_obj 2) 
        (vla-put-TextString mleader_obj (strcat "J" (itoa i)))
        (vla-put-TextHeight mleader_obj txt_h)
        (vla-put-Layer mleader_obj lay_id)
        (vla-put-ArrowheadType mleader_obj 19) ; 实心圆点
        (vla-put-ArrowheadSize mleader_obj (/ txt_h 2.5))
        (vla-put-DoglegLength mleader_obj 0.0) 

        (setq ent (entget (vlax-vla-object->ename mleader_obj)))
        (setq ent (append ent (list (list -3 (list "JZD_COORD_DATA" (cons 1000 (strcat "点号: J" (itoa i))) (cons 1000 (strcat "X坐标: " x_coord)) (cons 1000 (strcat "Y坐标: " y_coord)))))))
        (entmod ent)

        (if (= choice "Yes")
          (progn
            (setq pt_obj (vla-AddPoint mspace (vlax-3d-point pt)))
            (vla-put-Layer pt_obj lay_id)
            (setq ent (entget (vlax-vla-object->ename pt_obj)))
            (setq ent (append ent (list (list -3 (list "JZD_POINT_DATA" (cons 1000 (strcat "点号: J" (itoa i))) (cons 1000 (strcat "X坐标: " x_coord)) (cons 1000 (strcat "Y坐标: " y_coord)))))))
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

;;; --- 命令 1: 检查节点数 (保持现状) ---
(defun c:JZD_PTS (/ ss i ent pts)
  (princ "\n选择要检查节点数的多段线: ")
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

;;; --- 命令 2: 清理重叠节点 (保持现状) ---
(defun c:JZD_CL (/ ss i ent elist pts p clean_pts fuzz)
  (setq fuzz 0.0001)
  (princ "\n选择要清理重叠节点的多段线: ")
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
          (if (not (equal p (car clean_pts) fuzz))
            (setq clean_pts (cons p clean_pts))
          )
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

;;; --- 命令 3: 查询信息 ---
(defun c:JZD_INFO (/ ent data xdata)
  (if (setq ent (car (entsel "\n点击标注或物理点以查看内部数据: ")))
    (progn
      (setq data (entget ent '("JZD_COORD_DATA" "JZD_POINT_DATA")))
      (setq xdata (assoc -3 data))
      (if xdata
        (progn
          (princ "\n========== 界址点数据对象属性 ==========")
          (foreach app_list (cdr xdata)
            (foreach itm (cdr app_list)
              (if (= (car itm) 1000) (princ (strcat "\n>> " (cdr itm))))
            )
          )
          (princ "\n=========================================")
        )
        (princ "\n[提示] 该实体不包含任何界址点扩展属性。")
      )
    )
  )
  (princ)
)