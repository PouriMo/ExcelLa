(defun c:ExcelLa (/ ss_blocks ss_fill pt_left pt_right center_pt output_file f
                    depth gamma has_fill fill_depth fill_gamma
                    force_input force_indices force_list f_method vec_str vec_vals fx fy p1 p2 len
                    block_data fill_data sorted_blocks final_blocks
                    final_blocks_with_forces first_blk last_blk left_abut_edge right_abut_edge
                    all_interfaces line_lengths line_angles midpoints ds_values
                    node_rows interface_rows max_rows row_idx pt_row int_row force_row
                    old_delobj last_ent open_csv)

  (vl-load-com)

  ;; ==========================================
  ;; 1. GEOMETRY SELECTION
  ;; ==========================================
  (princ "\n\n--- STEP 1: CENTER & ABUTMENTS ---")
  (while (not (setq center_pt (getpoint "\nClick the center point of the arch geometry: "))))
  (while (not (setq pt_left (getpoint "\nClick anywhere ON the edge of the LEFT ground abutment: "))))
  (while (not (setq pt_right (getpoint "\nClick anywhere ON the edge of the RIGHT ground abutment: "))))
  
  (princ "\n\n--- STEP 2: BLOCKS ---")
  (princ "\nSelect all closed LWPOLYLINE blocks:")
  (while (not (setq ss_blocks (ssget '((0 . "LWPOLYLINE,POLYLINE")))))
    (princ "\nInvalid selection. Please select POLYLINE entities for your blocks:")
  )
  
  (princ "\n\n--- STEP 3: BACKFILL ---")
  (initget "Yes No")
  (setq has_fill (getkword "\nDo you have backfill geometry? [Yes/No] <No>: "))
  (if (null has_fill) (setq has_fill "No"))
  
  (if (= has_fill "Yes")
    (progn
      (princ "\nSelect all backfill closed LWPOLYLINES:")
      (while (not (setq ss_fill (ssget '((0 . "LWPOLYLINE,POLYLINE")))))
        (princ "\nInvalid selection. Please select POLYLINE entities for your backfill:")
      )
    )
  )

  ;; ==========================================
  ;; 2. PROPERTIES & FORCES INPUT
  ;; ==========================================
  (princ "\n\n--- STEP 4: NUMERICAL PROPERTIES ---")
  (while (not (setq depth (getreal "\nEnter blocks' depth (d) in mm: "))))
  (while (not (setq gamma (getreal "\nEnter blocks' specific weight (gamma) in kg/m^3: "))))
  
  (if (= has_fill "Yes")
    (progn
      (while (not (setq fill_depth (getreal "\nEnter backfill depth in mm: "))))
      (while (not (setq fill_gamma (getreal "\nEnter backfill specific weight in kg/m^3: "))))
    )
  )

  (princ "\n\n--- STEP 5: EXTERNAL FORCES ---")
  (setq force_input (getstring t "\nEnter block indices for external forces (comma-separated, e.g., 1,3,5) or press Enter for none: "))
  (setq force_indices (ParseIndices force_input))
  (setq force_list '())
  
  (if force_indices
    (progn
      (initget "Type Draw")
      (setq f_method (getkword "\nHow to define force vectors? [Type/Draw] <Type>: "))
      (if (null f_method) (setq f_method "Type"))
      
      (foreach idx force_indices
        (if (= f_method "Type")
          (progn
            (setq vec_str (getstring (strcat "\nEnter vector (Fx,Fy,Fz) for Block " (itoa idx) " (e.g., 0,-1,0): ")))
            (setq vec_vals (ParseVector vec_str))
            (setq fx (car vec_vals) fy (cadr vec_vals))
          )
          (progn
            (setq p1 (getpoint (strcat "\nClick START point of vector for Block " (itoa idx) ": ")))
            (setq p2 (getpoint p1 "\nClick END point: "))
            (setq len (distance p1 p2))
            (if (> len 0)
              (setq fx (/ (- (car p2) (car p1)) len)
                    fy (/ (- (cadr p2) (cadr p1)) len))
              (setq fx 0.0 fy 0.0)
            )
          )
        )
        (setq force_list (append force_list (list (list idx fx fy))))
      )
    )
  )

  ;; ==========================================
  ;; 3. PROCESSING
  ;; ==========================================
  (if ss_blocks
    (progn
      (setq output_file (getfiled "Save CSV" "" "csv" 1))
      (if (and output_file (setq f (open output_file "w")))
        (progn
          (setq old_delobj (getvar "DELOBJ"))
          (setvar "DELOBJ" 0)

          ;; Extract geometry properties and sort blocks left-to-right
          (setq block_data (ExtractGeomProps ss_blocks depth gamma))
          (setq sorted_blocks (vl-sort block_data '(lambda (a b) (< (car (nth 2 a)) (car (nth 2 b))))))

          (if (= has_fill "Yes")
            (setq fill_data (ExtractGeomProps ss_fill fill_depth fill_gamma))
          )

          ;; --- NEW ROBUST BACKFILL ASSIGNMENT LOGIC ---
          ;; Initialize blocks with format: (ent weight centroid pts fill_status list_of_fills)
          (setq final_blocks '())
          (foreach blk sorted_blocks
            (setq final_blocks (append final_blocks (list (append blk (list "No" '())))))
          )

          ;; Assign each piece of backfill via Shared Vertices
          (if fill_data
            (foreach fill fill_data
              (setq f_weight (nth 1 fill) f_cent (nth 2 fill) f_pts (nth 3 fill))
              
              (setq min_dist 1e99 closest_blk_idx -1 idx 0)
              (foreach blk final_blocks
                (setq b_cent (nth 2 blk) b_pts (nth 3 blk))
                
                ;; Crucial test: Do they share at least 2 points?
                (if (ShareVertices f_pts b_pts)
                  (progn
                    (setq dist (distance f_cent b_cent))
                    (if (< dist min_dist)
                      (progn (setq min_dist dist) (setq closest_blk_idx idx))
                    )
                  )
                )
                (setq idx (1+ idx))
              )
              
              ;; Fallback if no vertices perfectly align (X-axis proximity)
              (if (< closest_blk_idx 0)
                (progn
                  (setq min_dx 1e99 idx 0)
                  (foreach blk final_blocks
                    (setq b_cent (nth 2 blk))
                    (setq dx (abs (- (car f_cent) (car b_cent))))
                    (if (< dx min_dx)
                      (progn (setq min_dx dx) (setq closest_blk_idx idx))
                    )
                    (setq idx (1+ idx))
                  )
                  (princ "\nNote: A backfill did not perfectly share vertices with any block. Assigned to closest block by X-axis.")
                )
              )

              ;; Update the matched structural block
              (if (>= closest_blk_idx 0)
                (progn
                  (setq target_blk (nth closest_blk_idx final_blocks))
                  (setq b_ent (nth 0 target_blk) 
                        b_weight (nth 1 target_blk) 
                        b_cent (nth 2 target_blk) 
                        b_pts (nth 3 target_blk)
                        b_fills (nth 5 target_blk))
                  
                  (setq total_weight (+ b_weight f_weight))
                  (setq comb_x (/ (+ (* b_weight (car b_cent)) (* f_weight (car f_cent))) total_weight))
                  (setq comb_y (/ (+ (* b_weight (cadr b_cent)) (* f_weight (cadr f_cent))) total_weight))
                  
                  (setq b_fills (append b_fills (list f_pts))) ; Store the backfill points for plotting
                  
                  (setq updated_blk (list b_ent total_weight (list comb_x comb_y) b_pts "Yes" b_fills))
                  (setq final_blocks (SubstNth closest_blk_idx updated_blk final_blocks))
                )
              )
            )
          )
          ;; -------------------------------------------

          ;; Process Force Vectors
          (setq final_blocks_with_forces '())
          (setq block_num 1)
          (foreach blk final_blocks
             (setq f_entry (assoc block_num force_list))
             (if f_entry
               (setq fx (nth 1 f_entry) fy (nth 2 f_entry))
               (setq fx 0.0 fy 0.0)
             )
             ;; Append Fx and Fy
             (setq final_blocks_with_forces (append final_blocks_with_forces (list (append blk (list fx fy)))))
             (setq block_num (1+ block_num))
          )

          ;; Extract structural interfaces
          (setq all_interfaces '())
          
          ;; 1. Left Abutment
          (setq first_blk (car final_blocks_with_forces))
          (setq left_abut_edge (ClosestEdgeToPoint (nth 3 first_blk) pt_left))
          (setq all_interfaces (append all_interfaces (list left_abut_edge)))
          
          ;; 2. Internal Interfaces
          (setq i 0)
          (while (< i (- (length final_blocks_with_forces) 1))
            (setq entA (nth 0 (nth i final_blocks_with_forces)))
            (setq entB (nth 0 (nth (+ i 1) final_blocks_with_forces)))
            (setq shared (FindSharedEdge entA entB))
            (if shared
              (setq all_interfaces (append all_interfaces (list shared)))
              (princ (strcat "\nWarning: No shared edge found between blocks " (itoa (1+ i)) " and " (itoa (+ i 2))))
            )
            (setq i (1+ i))
          )
          
          ;; 3. Right Abutment
          (setq last_blk (last final_blocks_with_forces))
          (setq right_abut_edge (ClosestEdgeToPoint (nth 3 last_blk) pt_right))
          (setq all_interfaces (append all_interfaces (list right_abut_edge)))

          ;; Order Interface points
          (setq ordered_interfaces '() midpoints '() line_lengths '() line_angles '())
          (foreach intf all_interfaces
            (setq p1 (car intf) p2 (cadr intf))
            (setq ordered (if (< (distance p1 center_pt) (distance p2 center_pt)) (list p1 p2) (list p2 p1)))
            (setq ordered_interfaces (append ordered_interfaces (list ordered)))
            
            (setq len_val (distance p1 p2))
            (setq fi_val (* -1.0 (- 90.0 (* 180.0 (/ (angle (car ordered) (cadr ordered)) pi)))))
            (setq mid (list (/ (+ (car p1) (car p2)) 2.0) (/ (+ (cadr p1) (cadr p2)) 2.0)))
            
            (setq line_lengths (append line_lengths (list len_val)))
            (setq line_angles (append line_angles (list fi_val)))
            (setq midpoints (append midpoints (list mid)))
          )

          ;; Compute Ds values
          (setq ds_values '() i 0)
          (while (< i (- (length midpoints) 1))
            (setq ds_values (append ds_values (list (distance (nth i midpoints) (nth (1+ i) midpoints)))))
            (setq i (1+ i))
          )

          ;; Build CSV Output strings (Including Backfill Geometries)
          (setq node_rows '() block_num 1)
          (foreach blk final_blocks_with_forces
            (setq pts (nth 3 blk) fill_status (nth 4 blk) fills (nth 5 blk) fx (nth 6 blk) fy (nth 7 blk))
            
            (if (not (equal (car pts) (last pts) 1e-4))
              (setq pts (append pts (list (car pts))))
            )
            
            ;; Output structural block vertices
            (setq is_first_row T)
            (foreach pt pts
              (setq pt_row (strcat "Block " (itoa block_num) "," (rtos (car pt) 2 6) "," (rtos (cadr pt) 2 6)))
              (if is_first_row
                (setq force_str (strcat (rtos fx 2 6) "," (rtos fy 2 6) "," fill_status))
                (setq force_str ",,") 
              )
              (setq node_rows (append node_rows (list (list pt_row force_str))))
              (setq is_first_row nil)
            )
            
            ;; Output backfill vertices belonging to this block
            (foreach f_pts fills
              (if (not (equal (car f_pts) (last f_pts) 1e-4))
                (setq f_pts (append f_pts (list (car f_pts))))
              )
              (foreach pt f_pts
                (setq pt_row (strcat "Backfill " (itoa block_num) "," (rtos (car pt) 2 6) "," (rtos (cadr pt) 2 6)))
                (setq node_rows (append node_rows (list (list pt_row ",,"))))
              )
            )
            
            (setq block_num (1+ block_num))
          )

          (setq max_rows (max (length node_rows) (length line_lengths)))

          ;; Write Output (Perfectly aligned columns)
          (write-line "Block,X [mm],Y [mm],,Interface,Len [mm],Ds [mm],Fi [deg],Depth [mm],Weight [N],,Fx_lam [N],Fy_lam [N],Backfill_Added" f)
          (setq row_idx 0)
          (while (< row_idx max_rows)
            
            (if (< row_idx (length node_rows))
              (setq blk_data (nth row_idx node_rows) pt_row (car blk_data) force_row (cadr blk_data))
              (setq pt_row ",," force_row ",,")
            )
            
            (if (< row_idx (length line_lengths))
              (progn
                (setq len_val (nth row_idx line_lengths) fi_val (nth row_idx line_angles))
                (setq ds_val (if (< row_idx (length ds_values)) (rtos (nth row_idx ds_values) 2 6) ""))
                (setq weight_val (if (< row_idx (length final_blocks_with_forces)) (rtos (nth 1 (nth row_idx final_blocks_with_forces)) 2 6) ""))
                (setq int_row (strcat "Interface " (itoa (1+ row_idx)) "," (rtos len_val 2 6) "," ds_val "," (rtos fi_val 2 6) "," (rtos depth 2 6) "," weight_val))
              )
              (setq int_row ",,,,,")
            )
            
            (write-line (strcat pt_row ",," int_row ",," force_row) f)
            (setq row_idx (1+ row_idx))
          )

          (close f)
          (setvar "DELOBJ" old_delobj)
          (princ "\nExport completed successfully.")
          
          (initget "Yes No")
          (setq open_csv (getkword "\nDo you want to open the exported CSV file now? [Yes/No] <Yes>: "))
          (if (null open_csv) (setq open_csv "Yes"))
          
          (if (= open_csv "Yes")
            (startapp "explorer.exe" output_file)
          )
        )
      )
    )
    (princ "\nIncomplete selection. Command aborted.")
  )
  (princ)
)

;; === Helper Functions ===

(defun GetLinePts (ent / obj p1 p2)
  (setq obj (vlax-ename->vla-object ent))
  (setq p1 (vlax-safearray->list (vlax-variant-value (vla-get-StartPoint obj))))
  (setq p2 (vlax-safearray->list (vlax-variant-value (vla-get-EndPoint obj))))
  (list (list (car p1) (cadr p1)) (list (car p2) (cadr p2)))
)

(defun ExtractGeomProps (ss d g / i ent reg obj cent area weight pts data_list last_ent)
  (setq i 0 data_list '())
  (while (< i (sslength ss))
    (setq ent (ssname ss i))
    (setq pts (GetPolyPts ent))
    (setq last_ent (entlast))
    (command "_.REGION" ent "")
    
    (if (not (eq last_ent (entlast)))
      (progn
        (setq reg (entlast))
        (setq obj (vlax-ename->vla-object reg))
        (setq cent (vlax-safearray->list (vlax-variant-value (vla-get-Centroid obj))))
        (setq area (vla-get-Area obj))
        (vla-delete obj)
        (setq weight (* g area d 9.81 0.000000001))
        (setq data_list (append data_list (list (list ent weight cent pts))))
      )
      (princ "\nWarning: A polyline could not be converted to a Region. Ensure it is closed and planar.")
    )
    (setq i (1+ i))
  )
  data_list
)

(defun GetPolyPts (ent / obj pts pt_list i ubound)
  (setq obj (vlax-ename->vla-object ent))
  (setq pts (vlax-variant-value (vla-get-Coordinates obj)))
  (setq i 0 pt_list '())
  (if (vl-string-search "Polyline" (vla-get-ObjectName obj))
    (progn
      (setq ubound (vlax-safearray-get-u-bound pts 1))
      (while (< i ubound)
        (setq pt_list (append pt_list (list (list (vlax-safearray-get-element pts i)
                                                  (vlax-safearray-get-element pts (+ i 1))))))
        (setq i (+ i 2))
      )
    )
  )
  pt_list
)

(defun MakeSegments (pts / segs i)
  (setq segs '() i 0)
  (while (< i (length pts))
    (if (= i (- (length pts) 1))
      (setq segs (append segs (list (list (nth i pts) (nth 0 pts)))))
      (setq segs (append segs (list (list (nth i pts) (nth (+ i 1) pts)))))
    )
    (setq i (1+ i))
  )
  segs
)

(defun FindSharedEdge (entA entB / segsA segsB shared)
  (setq segsA (MakeSegments (GetPolyPts entA)))
  (setq segsB (MakeSegments (GetPolyPts entB)))
  (setq shared nil)
  (foreach a segsA
    (foreach b segsB
      (if (or (and (equal (car a) (car b) 1e-4) (equal (cadr a) (cadr b) 1e-4))
              (and (equal (car a) (cadr b) 1e-4) (equal (cadr a) (car b) 1e-4)))
        (setq shared a)
      )
    )
  )
  shared
)

;; NEW: Checks if two point lists share at least 2 identical coordinates
(defun ShareVertices (ptsA ptsB / count)
  (setq count 0)
  (foreach pa ptsA
    (foreach pb ptsB
      (if (< (distance pa pb) 1e-3)
        (setq count (1+ count))
      )
    )
  )
  (>= count 2)
)

(defun ClosestEdgeToPoint (pts pt / minDist bestSeg segs d)
  (setq minDist 1e99 bestSeg nil)
  (setq segs (MakeSegments pts))
  (foreach seg segs
    (setq d (DistPointToSegment pt (car seg) (cadr seg)))
    (if (< d minDist)
      (progn (setq minDist d) (setq bestSeg seg))
    )
  )
  bestSeg
)

(defun DistPointToSegment (p a b / l2 t_param proj)
  (setq l2 (+ (expt (- (car a) (car b)) 2) (expt (- (cadr a) (cadr b)) 2)))
  (if (equal l2 0.0 1e-6)
    (distance p a)
    (progn
      (setq t_param (/ (+ (* (- (car p) (car a)) (- (car b) (car a)))
                          (* (- (cadr p) (cadr a)) (- (cadr b) (cadr a))))
                       l2))
      (setq t_param (max 0.0 (min 1.0 t_param)))
      (setq proj (list (+ (car a) (* t_param (- (car b) (car a))))
                       (+ (cadr a) (* t_param (- (cadr b) (cadr a))))))
      (distance p proj)
    )
  )
)

(defun ParseIndices (str / str_list temp_str i char num result)
  (setq result '() str_list (vl-string->list str) temp_str "" i 0)
  (while (< i (length str_list))
    (setq char (chr (nth i str_list)))
    (if (or (equal char ",") (equal char " ") (equal char ";"))
      (progn
        (if (> (strlen temp_str) 0)
          (if (> (setq num (atoi temp_str)) 0) (setq result (append result (list num)))))
        (setq temp_str "")
      )
      (setq temp_str (strcat temp_str char))
    )
    (setq i (1+ i))
  )
  (if (> (strlen temp_str) 0)
    (if (> (setq num (atoi temp_str)) 0) (setq result (append result (list num)))))
  (vl-sort result '<)
)

(defun ParseVector (str / lst temp char i res fx fy)
  (setq lst (vl-string->list str) temp "" i 0 res '())
  (while (< i (length lst))
    (setq char (chr (nth i lst)))
    (if (or (= char ",") (= char " ") (= char ";"))
      (progn
        (if (> (strlen temp) 0) (setq res (append res (list (atof temp)))))
        (setq temp "")
      )
      (setq temp (strcat temp char))
    )
    (setq i (1+ i))
  )
  (if (> (strlen temp) 0) (setq res (append res (list (atof temp)))))
  (setq fx (if (> (length res) 0) (nth 0 res) 0.0))
  (setq fy (if (> (length res) 1) (nth 1 res) 0.0))
  (list fx fy)
)

(defun SubstNth (n new_val lst / i res)
  (setq i 0 res '())
  (foreach item lst
    (if (= i n)
      (setq res (append res (list new_val)))
      (setq res (append res (list item)))
    )
    (setq i (1+ i))
  )
  res
)

(princ "\nType ExcelLa to run the static block exporter.")
(princ)