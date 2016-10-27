(require 'cl-lib)
(require 'cl-macs)

(defmacro define-insert-functions (accent-name mappings)
  (let* ((cases (cl-mapcar (lambda (pair) `((,(car pair)) ,(cdr pair)))
                           mappings))
         (printable-mappings
          (cl-remove-if (lambda (c) (= ?\s (char-syntax c)))
                        mappings :key #'car))
         (insert-fn-name (intern (format "insert-%s" accent-name)))
         (other-fn-names
          (cl-mapcar (lambda (pair)
                       (intern (format "insert-%c%s"
                                       (car pair) accent-name)))
                     printable-mappings)))
    `(progn
       (defun ,insert-fn-name (char)
         (interactive "cCharacter:")
         (insert-char (cl-case char ,@cases (t char)) 1 t))
       ,@(cl-mapcar
           (lambda (name pair)
             `(defun ,name ()
                ,(format "Insert %c." (cdr pair))
                (interactive)
                (,insert-fn-name ,(car pair))))
           other-fn-names printable-mappings))))

(define-insert-functions grave
  ((?A . #xC0) (?a . #xE0)
   (?E . #xC8) (?e . #xE8)
   (?I . #xCC) (?i . #xEC)
   (?O . #xD2) (?o . #xF2)
   (?U . #xD9) (?u . #xF9)))

(define-insert-functions acute
  ((?A . #xC1) (?a . #xE1)
   (?E . #xC9) (?e . #xE9)
   (?I . #xCD) (?i . #xED)
   (?O . #xD3) (?o . #xF3)
   (?U . #xDA) (?u . #xFA)
   (?\s . #xB4)))

(define-insert-functions hat
  ((?A . #xC2) (?a . #xE2)
   (?E . #xCA) (?e . #xEA)
   (?I . #xCE) (?i . #xEE)
   (?O . #xD4) (?o . #xF4)
   (?U . #xDB) (?u . #xFB)
   (?\s . #x2C6)))

(define-insert-functions tilde
  ((?A . #xC3) (?a . #xE3)
   (?N . #xD1) (?n . #xF1)
   (?O . #xD5) (?o . #xF5)
   (?\s . #x2DC)))

(define-insert-functions umlaut
  ((?A . #xC4) (?a . #xE4)
   (?E . #xCB) (?e . #xEB)
   (?I . #xCF) (?i . #xEF)
   (?O . #xD6) (?o . #xF6)
   (?U . #xDC) (?u . #xFC)
   (?Y . #x178) (?y . #xFF)
   (?\s . #xA8)))

(defun insert-hellip ()
  (interactive)
  (insert-char #x2026 1 t))

(defun insert-bullet ()
  (interactive)
  (insert-char #x2022 1 t))

(defun insert-degree-sign ()
  (interactive)
  (insert-char #xB0 1 t))

(defun insert-inverted-bang ()
  (interactive)
  (insert-char #xA1 1 t))

(defun insert-inverted-question-mark ()
  (interactive)
  (insert-char #xBF 1 t))

(provide 'accented-characters)
