;;; This started off as a function I wrote for Judy that rotates
;;; columns of text.
(require 'rect)

(defun jrh/destructively-rotate-list (list count)
  "Rotate LIST COUNT times destructively.  Taken from
http://homepage1.nifty.com/bmonkey/emacs/elisp/mcomplete.el,
where it is called `mcomplete-rotate-list'.
\(destructively-rotate-list (list 1 2 3) 1) => (2 3 1).
\(destructively-rotate-list (list 1 2 3) -1) => (3 1 2)."
  (when list
    (let* ((len (length list))
           (count (mod count len))
           new-top new-last)
      (if (zerop count)
          list
        (setq new-last (nthcdr (1- count) list)
              new-top (cdr new-last))
        (setcdr (last new-top) list)
        (setcdr new-last nil)
        new-top))))

;;;###autoload
(defun rectangle-shift (start end &optional arg)
  "Shift the region-rectangle down ARG lines, wrapping around to the
top."
  (interactive "r\nP")
  (let ((repeat-key last-input-event)
        (deactivate-mark nil)
        (p (point))
        startcol endcol lines count)
    (save-excursion
      (goto-char start)
      (setq startcol (current-column))
      (goto-char end)
      (setq endcol (current-column))
      ;; Ensure the start column is the left one.
      (if (< endcol startcol)
          (let ((col startcol))
            (setq startcol endcol endcol col)))
      (setq lines (delete-extract-rectangle start end)
            count (if arg (- arg) -1))
      (setq lines (jrh/destructively-rotate-list lines count))
      (goto-char start)
      (insert-rectangle lines))
    (goto-char p)
    ;; Set up for repeating.
    (message "(Type %s to repeat)"
             (format-kbd-macro (vector repeat-key) nil))
    (set-transient-map
     (let ((map (make-sparse-keymap)))
       (define-key map (vector repeat-key)
         `(lambda () (interactive)
            (rectangle-shift ,start ,end)))
       map))))

;;;###autoload
(defun rectangle-yank-dup (start end &optional arg)
  "Calls `yank' at the start of each line of the region-rectangle."
  (interactive "r\nP")
  (apply-on-rectangle
   (lambda (startcol endcol &rest args)
     (save-excursion
       (forward-char startcol)
       (yank)))
   start end))

;;;###autoload
(defun downcase-rectangle (start end &optional arg)
  "Convert the region-rectangle to lower case."
  (interactive "r\nP")
  (let ((deactivate-mark nil))
    (save-excursion
      (apply-on-rectangle
       (lambda (startcol endcol &rest args)
         (downcase-region (progn (move-to-column startcol) (point))
                          (progn (move-to-column endcol) (point))))
       start end arg))))

;;;###autoload
(defun upcase-rectangle (start end &optional arg)
  "Convert the region-rectangle to upper case."
  (interactive "r\nP")
  (let ((deactivate-mark nil))
    (save-excursion
      (apply-on-rectangle
       (lambda (startcol endcol &rest args)
         (upcase-region (progn (move-to-column startcol) (point))
                        (progn (move-to-column endcol) (point))))
       start end arg))))

(provide 'rectangle-shift)
