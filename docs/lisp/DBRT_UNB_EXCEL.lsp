(defun c:UNBE ( / csvFile row line blockID blockName dwgSeries values ss i entObj attrs att tag)
  ;; Open the CSV file for reading
  (setq csvFile (open "Z:/2300001-2309999/2300416 STA Division Street BRT WO2/CADD/Resources/Exports/DBRT_NOTES.csv" "r"))
  
  ;; Ensure the file was opened successfully
  (if csvFile
    (progn
      ;; Get the DWG series for filtering (e.g., PV-1 from DBRT-PV1.dwg)
      (setq dwgSeries (strcat (substr (getvar "DWGNAME") 6 2) "-" (substr (getvar "DWGNAME") 8 1))) ;; E.g., PV-1

      ;; Initialize row counter
      (setq row 0)

      ;; Loop through each line in the CSV file
      (while (setq line (read-line csvFile))
        ;; Skip the header row
        (if (> row 0)
          (progn
            ;; Split the line by commas into a list of values
            (setq values (split-string line ","))

            ;; Extract block info
            (setq blockID (nth 0 values)) ;; E.g., PV-101-HEX-NT01

            ;; Check if the row matches the DWG series
            (if (wcmatch blockID (strcat dwgSeries "*"))
              (progn
                ;; Extract blockName and attribute details
                (setq blockName (substr blockID 4)) ;; E.g., 101-HEX-NT01
                (setq attrName1 (nth 1 values))
                (setq attrValue1 (nth 2 values))
                (setq attrName2 (nth 3 values))
                (setq attrValue2 (nth 4 values))

                ;; Debugging: Print out the values to see what is being read
                (princ (strcat "\nRow " (itoa row) ": "
                               "blockID = " blockID ", "
                               "blockName = " blockName ", "
                               "AttrName1 = " attrName1 ", "
                               "AttrValue1 = " attrValue1 ", "
                               "AttrName2 = " attrName2 ", "
                               "AttrValue2 = " attrValue2))

                ;; Select all blocks (including anonymous blocks)
                (setq ss (ssget "_X" (list (cons 0 "INSERT"))))

                ;; Check if any blocks were found
                (if ss
                  (repeat (setq i (sslength ss))
                    (setq entObj (vlax-ename->vla-object (ssname ss (setq i (1- i)))))
                    ;; See if the effective name matches the current block name
                    (if (and (vlax-property-available-p entObj 'EffectiveName)
                             (eq (strcase (vla-get-EffectiveName entObj)) (strcase blockName)))
                      (progn
                        (princ (strcat "\nFound block: " (vla-get-Name entObj)))
                        (if (and attrValue2 (/= attrValue2 "")) ; Check if attrValue2 is empty
                            (progn
                              ;; Action if attrValue2 is not empty
                              (LM:SetVisibilityState entObj "Hex")
                            )
                            (progn
                              ;; Action if attrValue2 is empty
                              (LM:SetVisibilityState entObj "No Hex")
                            )
                        )
                        (princ) ; Clean exit
                        (setq attrs (vlax-invoke entObj 'GetAttributes))
                        (foreach att attrs
                          (setq tag (vla-get-TagString att))
                          (cond
                            ((= tag attrName1) (vla-put-TextString att attrValue1))
                            ((= tag attrName2) (vla-put-TextString att attrValue2))
                          )
                        )
                      )
                    )
                  )
                )
              )
            )
          )
        )
        (setq row (1+ row))
      )
      ;; Close the file after reading
      (close csvFile)
      (princ "\nAttribute values updated from CSV.\n")
    )
    (princ "\nFailed to open the CSV file.\n")
  )
  (princ)
)



(defun split-string (str delim / pos res)
  (setq pos 0 res nil)
  (while (setq pos (vl-string-search delim str))
    (setq res (cons (substr str 1 pos) res))
    (setq str (substr str (+ pos 2))))
  (reverse (cons str res))
)



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