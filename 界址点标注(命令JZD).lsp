(defun c:JZD (/ ss ent pts i pt txt_h tbl pt_ins row obj_tbl lay1 lay2 mspace)
  (setvar "cmdecho" 0)
  (vl-load-com)

  ;; --- 1. 环境初始化 ---
  (setq lay1 "界址点序号")
  (setq lay2 "界址点坐标表")
  (setq mspace (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))))

  ;; 创建图层 (颜色7为黑白自动转换)
  (if (not (tblsearch "LAYER" lay1)) (command "-layer" "n" lay1 "c" "7" lay1 ""))
  (if (not (tblsearch "LAYER" lay2)) (command "-layer" "n" lay2 "c" "7" lay2 ""))

  ;; 创建最兼容的 SHX 大字体样式
  ;; txt.shx 处理数字和字母，gbcbig.shx 处理中文 (CAD自带，兼容性最强)
  (if (not (tblsearch "STYLE" "HZ_GB"))
    (command "-style" "HZ_GB" "txt.shx,gbcbig.shx" "0" "0.8" "0" "n" "n" "n")
  )

  ;; --- 2. 选择对象 ---
  (setq ss (entsel "\n请选择多段线 (LWPOLYLINE): "))
  (if (and ss (setq ent (car ss)))
    (progn
      (if (= (cdr (assoc 0 (entget ent))) "LWPOLYLINE")
        (progn
          ;; 提取顶点坐标
          (setq pts (mapcar 'cdr 
                      (vl-remove-if-not '(lambda (x) (= (car x) 10)) (entget ent))))
          
          (setq txt_h (getdist "\n请输入标注文字高度 <2.5>: "))
          (if (null txt_h) (setq txt_h 2.5))

          ;; --- 3. 标注点号 (存入“界址点序号”图层) ---
          (setq i 1)
          (foreach pt pts
            (entmake (list
                       '(0 . "TEXT")
                       (cons 8 lay1)       ; 图层
                       (cons 7 "HZ_GB")    ; 使用兼容大字体样式
                       (cons 10 pt)
                       (cons 40 txt_h)
                       (cons 1 (strcat "J" (itoa i)))
                       '(72 . 1) (cons 11 pt) '(73 . 2)
                     ))
            (setq i (1+ i))
          )

          ;; --- 4. 生成坐标表 (存入“界址点坐标表”图层) ---
          (initget 1) 
          (setq pt_ins (getpoint "\n请指定坐标表插入点: "))
          
          ;; 创建表格对象
          (setq obj_tbl (vla-addtable mspace (vlax-3d-point pt_ins) (+ (length pts) 2) 3 (* txt_h 2.0) (* txt_h 10.0)))
          
          (vla-put-layer obj_tbl lay2)

          ;; 强制表格样式链接到 HZ_GB 字体样式
          (vla-SetTextStyle obj_tbl acTitleRow "HZ_GB")
          (vla-SetTextStyle obj_tbl acHeaderRow "HZ_GB")
          (vla-SetTextStyle obj_tbl acDataRow "HZ_GB")

          ;; 填充表格表头内容
          (vla-settext obj_tbl 0 0 "界址点坐标成果表")
          (vla-settext obj_tbl 1 0 "点号")
          (vla-settext obj_tbl 1 1 "X坐标(北)")
          (vla-settext obj_tbl 1 2 "Y坐标(东)")

          ;; 循环填充坐标
          (setq row 2 i 1)
          (foreach pt pts
            (vla-settext obj_tbl row 0 (strcat "J" (itoa i)))
            (vla-settext obj_tbl row 1 (rtos (cadr pt) 2 3)) ; Y坐标 -> 测量X
            (vla-settext obj_tbl row 2 (rtos (car pt) 2 3))  ; X坐标 -> 测量Y
            (setq row (1+ row) i (1+ i))
          )
          (princ (strcat "\n[成功] 已标注 " (itoa (length pts)) " 个点，表格字体已设为国标大字体。"))
        )
        (princ "\n[错误] 所选对象不是多段线。")
      )
    )
  )
  (princ)
)