(defun c:JZD (/ ss i ent pts all_pts txt_h choice table_type pt_ins obj_tbl row lay1 lay2 mspace file_path csv_file)
  (setvar "cmdecho" 0)
  (vl-load-com)

  ;; --- 1. 环境初始化 ---
  (setq lay1 "界址点序号" lay2 "界址点坐标表")
  (setq mspace (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))))

  (if (not (tblsearch "LAYER" lay1)) (command "-layer" "n" lay1 "c" "7" lay1 ""))
  (if (not (tblsearch "LAYER" lay2)) (command "-layer" "n" lay2 "c" "7" lay2 ""))
  (if (not (tblsearch "STYLE" "HZ_GB"))
    (command "-style" "HZ_GB" "txt.shx,gbcbig.shx" "0" "0.8" "0" "n" "n" "n")
  )

  ;; --- 2. 选择多段线 ---
  (princ "\n请选择多段线 (LWPOLYLINE): ")
  (setq ss (ssget '((0 . "LWPOLYLINE"))))
  
  (if ss
    (progn
      (setq all_pts '() i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq pts (mapcar 'cdr (vl-remove-if-not '(lambda (x) (= (car x) 10)) (entget ent))))
        (setq all_pts (append all_pts pts))
        (setq i (1+ i))
      )

      (setq txt_h (getdist "\n请输入标注文字高度 <2.5>: "))
      (if (null txt_h) (setq txt_h 2.5))

      ;; --- 3. 标注点号 ---
      (setq i 1)
      (foreach pt all_pts
        (entmake (list '(0 . "TEXT") (cons 8 lay1) (cons 7 "HZ_GB") (cons 10 pt) (cons 40 txt_h)
                       (cons 1 (strcat "J" (itoa i))) '(72 . 1) (cons 11 pt) '(73 . 2)))
        (setq i (1+ i))
      )
      (princ (strcat "\n[提示] 已完成 " (itoa (length all_pts)) " 个点的标注。"))

      ;; --- 4. 询问导出方式 ---
      (initget "Yes No")
      (setq choice (getkword "\n是否生成坐标成果表? [是(Y)/否(N)] <Y>: "))
      (if (or (null choice) (= choice "Yes"))
        (progn
          (initget "CAD Excel")
          (setq table_type (getkword "\n请选择生成方式 [在CAD中生成(C)/导出Excel(E)] <C>: "))
          
          (if (or (null table_type) (= table_type "CAD"))
            ;; --- 方案 A: CAD内生成表格 ---
            (progn
              (setq pt_ins (getpoint "\n请指定坐标表插入点: "))
              (setq obj_tbl (vla-addtable mspace (vlax-3d-point pt_ins) (+ (length all_pts) 2) 3 (* txt_h 2.0) (* txt_h 10.0)))
              (vla-put-layer obj_tbl lay2)
              (vla-SetTextStyle obj_tbl acTitleRow "HZ_GB")
              (vla-SetTextStyle obj_tbl acHeaderRow "HZ_GB")
              (vla-SetTextStyle obj_tbl acDataRow "HZ_GB")
              (vla-settext obj_tbl 0 0 "界址点坐标汇总表")
              (vla-settext obj_tbl 1 0 "点号") 
              (vla-settext obj_tbl 1 1 "X坐标(北)") 
              (vla-settext obj_tbl 1 2 "Y坐标(东)")
              (setq row 2 i 1)
              (foreach pt all_pts
                (vla-settext obj_tbl row 0 (strcat "J" (itoa i)))
                (vla-settext obj_tbl row 1 (rtos (cadr pt) 2 3))
                (vla-settext obj_tbl row 2 (rtos (car pt) 2 3))
                (setq row (1+ row) i (1+ i))
              )
              (princ "\n[成功] CAD表格已生成。")
            )
            ;; --- 方案 B: 导出 Excel (采用稳定的 CSV 文件写入) ---
            (progn
              (setq file_path (getfiled "保存坐标成果表" "Result.csv" "csv" 1))
              (if file_path
                (progn
                  (setq csv_file (open file_path "w"))
                  (if csv_file
                    (progn
                      ;; 写入 CSV 表头
                      (write-line "点号,X坐标(北),Y坐标(东)" csv_file)
                      (setq i 1)
                      (foreach pt all_pts
                        (write-line (strcat "J" (itoa i) "," (rtos (cadr pt) 2 3) "," (rtos (car pt) 2 3)) csv_file)
                        (setq i (1+ i))
                      )
                      (close csv_file)
                      (princ (strcat "\n[完成] 数据已保存至: " file_path))
                    )
                    (princ "\n[错误] 无法创建文件，请检查目录权限或文件是否被占用。")
                  )
                )
              )
            )
          )
        )
      )
    )
    (princ "\n[取消] 未选中多段线。")
  )
  (princ)
)