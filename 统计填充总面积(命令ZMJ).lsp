(defun c:ZMJ (/ ss i ent m2 total_area)
  (setvar "cmdecho" 0)
  (vl-load-com)

  (princ "\n请选择要统计面积的填充 (HATCH): ")
  
  ;; 1. 建立选择集，仅过滤出 HATCH (填充) 对象
  (setq ss (ssget '((0 . "HATCH"))))
  
  (if ss
    (progn
      (setq total_area 0.0)
      (setq i 0)
      
      ;; 2. 遍历选择集并累加面积
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        ;; 获取填充对象的面积属性
        (setq m2 (vla-get-area (vlax-ename->vla-object ent)))
        (setq total_area (+ total_area m2))
        (setq i (1+ i))
      )
      
      ;; 3. 在命令行输出结果
      (princ "\n================ 面积统计结果 ================")
      (princ (strcat "\n>>> 已选择填充数量: " (itoa (sslength ss))))
      (princ (strcat "\n>>> CAD原始总面积 : " (rtos total_area 2 3)))
      
      ;; 4. 增加不同单位制的参考说明
      (princ "\n----------------------------------------------")
      (princ (strcat "\n>>> 参考1 (若单位为mm): " (rtos (/ total_area 1000000.0) 2 3) " 平方米"))
      (princ (strcat "\n>>> 参考2 (若单位为 m): " (rtos total_area 2 3) " 平方米"))
      (princ "\n==============================================")
    )
    (princ "\n未选择任何填充对象。")
  )
  
  (princ)
)

(princ "\n加载成功！输入 ZMJ 运行统计面积命令。")
(princ)