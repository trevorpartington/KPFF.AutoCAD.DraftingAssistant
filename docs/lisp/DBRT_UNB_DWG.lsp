(defun c:UNBD (/ doc drawingName drawingIdentifier 
                  lList initialLayout layoutName layoutNum 
                  outputFile csvLine csvLines expfilePath
                  ptObj ptobjsAll ptobjsforLayout ptCoords ptcoordsforLayout 
                  mlObj mlobjsAll mlobjsforLayout mlX mlY mlattrsforLayout mlStyle
                  blObj blobjsAll blobjsforLayout blX blY blattrsforLayout
                  allattrsforLayout
                  arNotes arNotesExp)
  (vl-load-com)

  ;; Get the current document and its name
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq drawingName (getvar "DWGNAME"))
  
  ;; Set file path for csv exports
  (setq expfilepath "Z:/2300001-2309999/2300416 STA Division Street BRT WO2/CADD/Resources/Exports")
  
  ;; Extract the drawing identifier from the drawing name (e.g. "DBRT-PV1.dwg" to "PV")
  (setq drawingIdentifier (substr drawingName 6 2))
  ;; Get layout names
  (setq lList (layoutlist))

  ;; Get initial layout (to restore at end)
  ;; (setq initialLayout (vla-get-Name (vla-get-ActiveLayout doc)))
  
  ;; Initialize list variables
  (setq csvLines '())
  (setq ptobjsAll '())
  (setq mlobjsAll '())
  (setq blobjsAll '())
  
  ;; Get all of the COGO points
  (setq ptobjsAll (LM:ss->vla (ssget "_X" '((0 . "AECC_COGO_POINT")))))
  
  ;; Get all of the multileaders
  (setq mlobjsAll (LM:ss->vla (ssget "_X" '((0 . "MULTILEADER")))))
    
  ;; Get all of the FlagC blocks
  (setq blobjsAll (LM:ss->vla (ssget "_X" '((0 . "INSERT")(2 . "_TagHexagon")))))

  ;; Loop through each layout
  (foreach layoutName lList
    
    ;; Reset variables
    (setq csvLine nil)
    (setq layoutNum (substr layoutName 4)) ;; 101, 102, etc.
    (setq ptobjsforLayout '())
    (setq ptcoordsforLayout '())
    (setq mlobjsforLayout '())
    (setq mlattrsforLayout '())
    (setq blobjsforLayout '())
    (setq blattrsforLayout '())
    (princ (strcat "\nProcessing Layout: " layoutName))
        
;; For each point
(foreach ptObj ptobjsAll
  
  ;; Check each for layout number in name - exact match with dash (case-insensitive)
  (if (vl-string-search (strcase (strcat layoutNum "-")) (strcase (vla-get-name ptObj)))
    
    ;; Create list of point objects specific to the layout
    (setq ptobjsforLayout (cons ptObj ptobjsforLayout))
  )
)
    
    ;; Create list of point coordinates
    (setq ptcoordsforLayout (cogo-to-coordinates ptobjsforLayout))
        
    ;; For each multileader
    (foreach mlObj mlobjsAll

      ;(highlight-mleader mlObj) ;; For debugging
      
      ;; Set style and location variables
        (setq mlStyle (vla-get-stylename mlObj))
        (setq mlX (car(mleaderarrowpoint mlObj)))
        (setq mlY (cadr(mleaderarrowpoint mlObj)))
      
      ;(princ "\nProcessing completed successfully.") ;; For debugging
      ;(clear-highlight-mleader mlObj)
      
      ;; Check each for style and location
      (if (and (eq mlStyle "Arw-Straight-Hex_Anno_WSDOT") 
         (is-point-in-polygon mlX mlY ptcoordsforLayout))

        ;; Create list of multileader objects specific to the layout
        (setq mlobjsforLayout (cons mlObj mlobjsforLayout))
      )
    )
    
    ;; For each block
    (foreach blObj blobjsAll

      ;; Set location variables      
        (setq blX (car(getblockinsertionPoint blObj)))
        (setq blY (cadr(getblockinsertionPoint blObj)))
      
      ;; Check each for location
      (if (is-point-in-polygon blX blY ptcoordsforLayout) 

        ;; Create list of block objects specific to the layout
        (setq blobjsforLayout (cons blObj blobjsforLayout))
      )
    )
    
    ;; Convert list of multileaders to sorted list of attributes with duplicates removed
    (setq mlattrsforLayout (sort-values (ExtractTagNumbersFromMultileaders mlobjsforLayout)))
    
    ;; Convert list of blocks to sorted list of attributes with duplicates removed
    (setq blattrsforLayout (sort-values (ExtractTagNumbersFromBlocks blobjsforLayout)))
    
    ;; Combine lists of blocks and multileaders
    (setq allattrsforLayout (sort-values (remove-duplicates (append mlattrsforLayout blattrsforLayout))))
    
    ;; Construct CSV output
    (setq csvLine (cons layoutNum allattrsforLayout))
    (setq csvLine (ListToCSV csvLine))
    (setq csvLines (cons csvLine csvLines))
    
  )
  
  (setq csvlines (reverse-list csvlines))
  (write-csv csvLines drawingIdentifier)
    
  ;(extract-identifiers) ; Not used
  (load-csv-files drawingIdentifier expfilePath)
  (process-layouts arNotes arNotesExp drawingIdentifier)
  (princ "Notes updated successfully.")
  (princ)
  
  ;; Restore the initial active layout  
  ;(vla-put-ActiveLayout doc (vla-item layouts initialLayout))
)



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;LIST FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



(defun sort-values (lst)
  ;; Convert all elements to numbers, sort them, then convert back to strings
  (mapcar 'itoa (vl-sort (mapcar 'atoi lst) '<))
)

;; Example usage:
;; (sort-values '(4 2 10 3 1))
;; This will return ("1" "2" "3" "4" "10")


(defun calc-centroid (pts / x-sum y-sum centroid pt)
  ;; pts should be a list of 4 points, each point is a list of two values (x, y)
  (setq x-sum 0
        y-sum 0)
  
  ;; Iterate through each point, summing the x and y coordinates
  (foreach pt pts
    (setq x-sum (+ x-sum (car pt)))  ; sum x-coordinates
    (setq y-sum (+ y-sum (cadr pt))) ; sum y-coordinates
  )
  
  ;; Calculate the centroid by averaging the x and y sums
  (setq centroid (list (/ x-sum 4.0) (/ y-sum 4.0)))
  
  ;; Return the centroid
  centroid
)

;; Example usage:
;; (calc-centroid '((0.0 0.0) (10.0 0.0) (10.0 10.0) (0.0 10.0)))


(defun cogo-to-coordinates (cogo-points / point coords easting northing)
  ;; Initialize the list to hold the eastings and northings
  (setq coords '())
  
  ;; Iterate through each COGO point in the list
  (foreach point cogo-points
    (setq easting  (vlax-get point 'Easting)
          northing (vlax-get point 'Northing))
    ;; Append the eastings and northings as a pair to the coords list
    (setq coords (append coords (list (list easting northing))))
  )

  ;; Return the list of coordinates
  coords
)
;; Example usage:
;; (setq my-points (list cogo-point1 cogo-point2 cogo-point3))
;; (cogo-to-coordinates my-points)


(defun convert-numbers-to-strings (numList / result)
  (setq result '()) ; Initialize an empty list
  (while numList
    (setq result (cons (rtos (car numList)) result)) ; Convert first number and add it to result
    (setq numList (cdr numList)) ; Move to the next number in the list
  )
  (reverse result) ; Reverse the result list to maintain original order
)
   

(defun remove-duplicates (lst)
  (defun helper (input output)
    (if (null input)
      output
      (if (member (car input) output)
        (helper (cdr input) output)
        (helper (cdr input) (append output (list (car input)))))))
  (helper lst '())
)
;; Example Usage
;;(setq myList '("A" "B" "C" "B" "D" "A"))
;;(remove-duplicates myList)


(defun reverse-list (lst)
  (if (null lst)
    '()
    (append (reverse-list (cdr lst)) (list (car lst))))
)
;; Example Usage
;;(setq myList '("A" "B" "C" "D"))
;;(reverse-list myList)



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;CSV FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



(defun ListToCSV (lst / csv item)
  (setq csv "")
  (foreach item lst
    (setq csv (strcat csv item ","))
  )
  ;; Remove trailing comma
  (if (> (strlen csv) 0)
    (setq csv (substr csv 1 (- (strlen csv) 1)))
  )
  csv
)
;; Example usage:
;;(setq mylist '("apple" "banana" "cherry"))
;;(ListToCSV mylist) ;; Result: "apple,banana,cherry"
;;(princ)


(defun load-csv-files (di expFP)
  (setq arNotes (read-csv-to-array (strcat expFP "/DBRT_" di "_NOTES.csv") T))
  (setq arNotesExp (read-csv-to-array (strcat expFP "/DBRT_" di "_NOTES_FROM_DWG.csv") nil))
)


(defun parsecsv (line / result pos)
  (setq result '()) ; Initialize an empty list to hold the columns
  (while (setq pos (vl-string-search "," line))
    (setq result (append result (list (substr line 1 pos)))) ; Extract the substring before the comma
    (setq line (substr line (+ pos 2)))) ; Update the line to be the part after the comma
  (setq result (append result (list line))) ; Add the final column
  result
)


(defun read-csv-to-array (filepath has-header)
  (setq arData '()) ; Initialize empty list
  (if (findfile filepath)
    (progn
      (setq file (open filepath "r"))
      (if has-header
        (read-line file)) ; Skip header if exists
      (while (setq line (read-line file))
        (setq arData (append arData (list (parsecsv line)))))
      (close file)
      arData
    )
    (prompt (strcat "File not found: " filepath))
  )
)


(defun write-csv (strList di / filePath file)
  ;; Define the file path for the CSV file
  (setq filePath (strcat "Z:/2300001-2309999/2300416 STA Division Street BRT WO2/CADD/Resources/Exports/DBRT_" di "_NOTES_FROM_DWG.csv"))
  
  ;; Open the file for writing
  (setq file (open filePath "w"))
  
  ;; Loop through the list of strings and write each to the file
  (foreach line strList
    (write-line line file)
  )
  
  ;; Close the file after writing
  (close file)
  
  ;; Return a success message
  (princ "\nCSV file written successfully.")
)
;; Example usage:
;; (write-csv '("value1,value2,value3" "value4,value5,value6"))
  


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;OBJECT FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
  


(defun clear-block-attributes (layoutNum / ss blockName i j entObj attrs att)
  ;; Initialize i to 0 before entering the loop
  (setq i 0)
  
  ;; Loop through blocks (layoutNum)-HEX-NT01 thru (layoutNum)-HEX-NT24 in the given layout
  (repeat 24
    (setq blockName (strcat layoutNum "-HEX-NT" (if (< i 9) (strcat "0" (itoa (1+ i))) (itoa (1+ i)))))
    ;; Select all blocks (including anonymous blocks)
    (setq ss (ssget "_X" (list (cons 0 "INSERT"))))
    
    ;; Check if any blocks were found
    (if ss
      (progn
        (repeat (setq j (sslength ss))
          (setq entObj (vlax-ename->vla-object (ssname ss (setq j (1- j)))))
          ;; See if the effective name matches the current block name
          (if (and (vlax-property-available-p entObj 'EffectiveName)
                  (eq (strcase (vla-get-EffectiveName entObj)) (strcase blockName)))
            (progn
              ;(princ (strcat "\nFound block: " (vla-get-Name entObj)))
              (LM:SetVisibilityState entObj "No Hex")
              (setq attrs (vlax-invoke entObj 'GetAttributes))
              ;; Clear the attributes in the block
              (foreach att attrs
                (vla-put-TextString att "")
              )
            )
          )
        )
      )
      (princ (strcat "\nNo blocks found with name" blockName))
    )
    
    ;; Increment i for the next block
    (setq i (1+ i))
  )
  ;; Return nil to end function
  nil
)


(defun ExtractTagNumbersFromBlocks (bl-list / result bl attrList attrValue)
  ;; Initialize an empty list to store the tag numbers
  (setq result nil)
  
  ;; Loop through each block reference in the list
  (foreach bl bl-list

    (setq attrValue (LM:vl-getattributevalue bl "TAGNUMBER"))
    ;; Add the attribute value to the result list if it exists
    (if attrValue
      (setq result (cons attrValue result))
    )
  )
  
  ;; Remove duplicate entries from the result list
  (setq result (remove-duplicates result))
  
  ;; Return the list of unique tag numbers
  result
)



(defun ExtractTagNumbersFromMultileaders (ml-list / result ml BlockColl ent nm blockObj attrValue)
  ; Initialize an empty list to store the tag numbers
  (setq result nil)  

  ;; Get the blocks collection
  (setq BlockColl (vla-get-blocks (vla-get-ActiveDocument (vlax-get-acad-object))))
  
  (foreach ml ml-list
    
    (setq nm (vla-get-ContentBlockName ml))
    (setq blockObj (vla-item BlockColl nm))

    ;; Iterate through the block's entities to find the attribute
    (vlax-for ent blockObj
      (if (eq (vla-get-ObjectName ent) "AcDbAttributeDefinition")
        (setq attrValue (vla-getBlockAttributeValue ml (vla-get-ObjectID ent)))
      )
    )    
    (setq result (cons attrValue result))
  )
  (setq result (remove-duplicates result))
  result
)


;; Get Attribute Value  -  Lee Mac - TMP mod
;; Returns the value held by the specified tag within the supplied block, if present.
;; blk - [vla] VLA Block Reference Object
;; tag - [str] Attribute TagString
;; Returns: [str] Attribute value, else nil if tag is not found.

(defun LM:vl-getattributevalue ( blk tag / att)
    (setq tag (strcase tag))
    (vl-some '(lambda ( att ) (if (= tag (strcase (vla-get-tagstring att))) (vla-get-textstring att))) (vlax-invoke blk 'getattributes))
)


(defun GetValidLeaderPoint (ml / index max-index leader-vertices)
  (setq index 0)
  (setq max-index 100)
  (setq leader-vertices nil)
  
  ;; Loop through indices until we find a valid leader or reach max-index
  (while (and (not leader-vertices) (< index max-index))
    (setq leader-vertices (vl-catch-all-apply 'vlax-invoke (list ml 'getleaderlinevertices index)))
    ;; If the result is not an error, break out of the loop
    (if (not (vl-catch-all-error-p leader-vertices))
      (setq leader-vertices leader-vertices)
      ;; If it is an error, reset leader-vertices and try the next index
      (setq leader-vertices nil)
    )
    (setq index (1+ index))
  )
  
  ;; Return the found leader vertices or nil if none were found
  leader-vertices
)


(defun is-vla-object (obj)
  (eq (type obj) 'VLA-OBJECT)
)


;; Bounding Box  -  Lee Mac
;; Returns the point list describing the rectangular frame bounding the supplied object.
;; obj - [vla] VLA-Object

(defun LM:boundingbox ( obj / a b lst )
    (if
        (and
            (vlax-method-applicable-p obj 'getboundingbox)
            (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-getboundingbox (list obj 'a 'b))))
            (setq lst (mapcar 'vlax-safearray->list (list a b)))
        )
        (mapcar '(lambda ( a ) (mapcar '(lambda ( b ) ((eval b) lst)) a))
           '(
                (caar   cadar)
                (caadr  cadar)
                (caadr cadadr)
                (caar  cadadr)
            )
        )
    )
)


;;------------=={ SelectionSet -> VLA Objects }==-------------;;
;;                                                            ;;
;;  Converts a SelectionSet to a list of VLA Objects          ;;
;;------------------------------------------------------------;;
;;  Author: Lee Mac, Copyright © 2011 - www.lee-mac.com       ;;
;;------------------------------------------------------------;;
;;  Arguments:                                                ;;
;;  ss - Valid SelectionSet (Pickset)                         ;;
;;------------------------------------------------------------;;
;;  Returns:  List of VLA Objects, else nil                   ;;
;;------------------------------------------------------------;;

(defun LM:ss->vla ( ss / i l )
    (if ss
        (repeat (setq i (sslength ss))
            (setq l (cons (vlax-ename->vla-object (ssname ss (setq i (1- i)))) l))
        )
    )
)


(defun mleaderarrowpoint (ml / numLdrs)
  (setq numLdrs (vla-get-leadercount ml))
  (if (= numLdrs 0)
    ;; Case: No leaders, get the bounding box and centroid of multileader
    (calc-centroid(LM:boundingbox ml))
    
    ;; Case: One or more leaders, return the x,y coordinates of any leader
    (GetValidLeaderPoint ml)
  )
)

(defun getblockinsertionPoint (blockObj / insPt)
  ;; Check if the block object is a valid VLA-OBJECT and is a block reference
  (if (and blockObj (eq (vla-get-ObjectName blockObj) "AcDbBlockReference"))
    (progn
      ;; Get the insertion point of the block directly as a list
      (setq insPt (vlax-get blockObj 'InsertionPoint))
      
      ;; Return only the X and Y coordinates as a list
      (list (nth 0 insPt) (nth 1 insPt))
    )
    ;; Error handling: invalid object type
    nil  ;; Return nil if it's not a valid block reference
  )
)





;;(vlax-invoke ml 'getleaderlinevertices 0)


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;MISC FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


;; Not used
;(defun extract-identifiers ()
;  (setq filename (getvar 'dwgname)) ; Get the filename
;  (setq ShortDrawingIdentifier (substr filename 6 2)) ; Letters 6-7
;  (setq FullDrawingIdentifier (substr filename 6 3))  ; Letters 6-8
;)


(defun is-point-in-polygon (test-x test-y polygon / num-vertices polygon is-inside i current-point next-point x1 x2 y1 y2 test-y test-x)
  ;; Arguments:
  ;; test-x: X coordinate of the test point
  ;; test-y: Y coordinate of the test point
  ;; polygon: List of (x y) points representing the vertices of the polygon in order

  (setq num-vertices (length polygon))  ;; Number of vertices in the polygon
  (setq is-inside nil)                  ;; Result flag, initially nil (outside)
  (setq i 0)                            ;; Index for looping over vertices

  ;; Loop through all edges of the polygon
  (while (< i num-vertices)
    ;; Get current and next vertices
    (setq current-point (nth i polygon))
    (setq next-point (nth (my-mod (1+ i) num-vertices) polygon))  ;; Loop back to the start

    (setq x1 (car current-point)) ;; X of current vertex
    (setq y1 (cadr current-point)) ;; Y of current vertex
    (setq x2 (car next-point))     ;; X of next vertex
    (setq y2 (cadr next-point))    ;; Y of next vertex

    ;; Check if test point is between the y bounds of the edge and handle intersection
    (if (and (or (and (< y1 test-y) (<= test-y y2))  ;; y1 < test-y <= y2 or y2 < test-y <= y1
                 (and (< y2 test-y) (<= test-y y1)))
             (< test-x
                (+ x1
                   (* (/ (- test-y y1) (- y2 y1))
                      (- x2 x1)))))
      (setq is-inside (not is-inside))  ;; Toggle the inside flag
    )

    (setq i (1+ i))  ;; Increment index
  )

  is-inside  ;; Return result
)


(defun my-mod (a b)
  ;; Custom modulus function
  (- a (* b (fix (/ a b))))
)


(defun process-layouts (arNotes arNotesExp di)
  (setq acadObj (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acadObj))
  (setq layouts (vla-get-Layouts doc))

  ;; Loop through each layout based on arNotesExp data
  (foreach layoutRow arNotesExp
    (setq layoutNum (car layoutRow))  ;; First column of arNotesExp corresponds to layout number. E.g. "101"
    ;;(setq layoutName (strcat di "-" layoutNum))  ;; No need to convert layoutNum to string. E.g. "PV-101"
    
    ;; Switch to layout
    ;(vla-put-ActiveLayout doc (vla-Item layouts layoutName))
    
    ;; Clear attributes
    (clear-block-attributes layoutNum)
    
    ;; Get KeyNotes list for this layout
    (setq listKeyNotes (cdr layoutRow)) 
    
    ;; Loop through each KeyNoteNum and update blocks accordingly
    (setq i 0)  ;; Initialize index
    (while (< i (length listKeyNotes))
      (setq KeyNoteNum (nth i listKeyNotes))  ;; Get the current KeyNoteNum
      (setq constructionNote (cadr (nth (1- (atoi KeyNoteNum)) arNotes)))  ;; Extract the second element from the list (the note itself)
      
      ;; Create block name
      (setq blockName (strcat layoutNum "-HEX-NT" (if (< i 9) (strcat "0" (itoa (1+ i))) (itoa (1+ i)))))
      
      ;; Select all blocks (including anonymous blocks)
      (setq ss (ssget "_X" (list (cons 0 "INSERT"))))
      
      ;; If block exists, update attributes
      (if ss
        (progn
          (repeat (setq j (sslength ss))
            (setq entObj (vlax-ename->vla-object (ssname ss (setq j (1- j)))))
            
            ;; See if the effective name matches the current block name
            (if (and (vlax-property-available-p entObj 'EffectiveName)
                    (eq (strcase (vla-get-EffectiveName entObj)) (strcase blockName)))
              (progn
                ;(princ (strcat "\nFound block: " (vla-get-Name entObj)))

                ;; Turn on hexagon
                (LM:SetVisibilityState entObj "Hex")

                (setq attrs (vlax-invoke entObj 'GetAttributes))
                
                ;; Loop through block attributes and update based on tag
                (foreach att attrs
                  (setq tag (vla-get-TagString att))
                  (cond
                    ;; Update KeyNoteNum
                    ((= tag "NUM")
                    (vla-put-TextString att KeyNoteNum))
                    
                    ;; Update construction note
                    ((= tag "NOTE")
                    ;; Ensure constructionNote is properly a string
                    (setq constructionNoteStr (if (listp constructionNote)
                                                  (vl-princ-to-string constructionNote)
                                                  constructionNote))
                    (vla-put-TextString att constructionNoteStr))
                  )
                )
              )
            )
          )
        )
      )
      (setq i (1+ i))  ;; Increment KeyNoteNum counter
    )
  )
)



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;VISIBILITY STATE FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



;; Set Dynamic Block Visibility State  -  Lee Mac
;; Sets the Visibility Parameter of a Dynamic Block (if present) to a specific value (if allowed)
;; blk - [vla] VLA Dynamic Block Reference object
;; val - [str] Visibility State Parameter value
;; Returns: [str] New value of Visibility Parameter, else nil

(defun LM:SetVisibilityState ( blk val / vis )
    (if
        (and
            (setq vis (LM:getvisibilityparametername blk))
            (member (strcase val) (mapcar 'strcase (LM:getdynpropallowedvalues blk vis)))
        )
        (LM:setdynpropvalue blk vis val)
    )
)


;; Get Visibility Parameter Name  -  Lee Mac
;; Returns the name of the Visibility Parameter of a Dynamic Block (if present)
;; blk - [vla] VLA Dynamic Block Reference object
;; Returns: [str] Name of Visibility Parameter, else nil

(defun LM:getvisibilityparametername ( blk / vis )  
    (if
        (and
            (vlax-property-available-p blk 'effectivename)
            (setq blk
                (vla-item
                    (vla-get-blocks (vla-get-document blk))
                    (vla-get-effectivename blk)
                )
            )
            (= :vlax-true (vla-get-isdynamicblock blk))
            (= :vlax-true (vla-get-hasextensiondictionary blk))
            (setq vis
                (vl-some
                   '(lambda ( pair )
                        (if
                            (and
                                (= 360 (car pair))
                                (= "BLOCKVISIBILITYPARAMETER" (cdr (assoc 0 (entget (cdr pair)))))
                            )
                            (cdr pair)
                        )
                    )
                    (dictsearch
                        (vlax-vla-object->ename (vla-getextensiondictionary blk))
                        "ACAD_ENHANCEDBLOCK"
                    )
                )
            )
        )
        (cdr (assoc 301 (entget vis)))
    )
)


;; Get Dynamic Block Property Allowed Values  -  Lee Mac
;; Returns the allowed values for a specific Dynamic Block property.
;; blk - [vla] VLA Dynamic Block Reference object
;; prp - [str] Dynamic Block property name (case-insensitive)
;; Returns: [lst] List of allowed values for property, else nil if no restrictions

(defun LM:getdynpropallowedvalues ( blk prp )
    (setq prp (strcase prp))
    (vl-some '(lambda ( x ) (if (= prp (strcase (vla-get-propertyname x))) (vlax-get x 'allowedvalues)))
        (vlax-invoke blk 'getdynamicblockproperties)
    )
)


;; Set Dynamic Block Property Value  -  Lee Mac
;; Modifies the value of a Dynamic Block property (if present)
;; blk - [vla] VLA Dynamic Block Reference object
;; prp - [str] Dynamic Block property name (case-insensitive)
;; val - [any] New value for property
;; Returns: [any] New value if successful, else nil

(defun LM:setdynpropvalue ( blk prp val )
    (setq prp (strcase prp))
    (vl-some
       '(lambda ( x )
            (if (= prp (strcase (vla-get-propertyname x)))
                (progn
                    (vla-put-value x (vlax-make-variant val (vlax-variant-type (vla-get-value x))))
                    (cond (val) (t))
                )
            )
        )
        (vlax-invoke blk 'getdynamicblockproperties)
    )
)



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;DEBUGGING FUNCTIONS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



(defun clear-highlight-mleader (mleader-obj)
  (if (and mleader-obj (is-vla-object mleader-obj))
    (vla-highlight mleader-obj :vlax-false)
  )
)


(defun highlight-mleader (mleader-obj)
  (if (and mleader-obj (is-vla-object mleader-obj))
    (vla-highlight mleader-obj :vlax-true)
  )
)


(defun log-mleader-handle (mleader-obj)
  (if (and mleader-obj (is-vla-object mleader-obj))
    (progn
      (setq handle (vla-get-Handle mleader-obj))
      (princ (strcat "\nMultileader Handle: " handle))
      handle
    )
  )
)


(defun safe-process-mleader (mleader-obj)
  (if (and mleader-obj (is-vla-object mleader-obj))
    (progn
      (log-mleader-handle mleader-obj)
      (mapcar '(lambda ( x ) (vlax-invoke mleader-obj 'getleaderlinevertices x)) (vlax-invoke mleader-obj 'getleaderlineindexes 0))
    )
    (princ "\nInvalid multileader object.")
  )
)
