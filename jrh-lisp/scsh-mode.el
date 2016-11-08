;; @file jrh-lisp/scsh-mode.el
;; Jamison Hope <jrh@theptrgroup.com>

(require 'scheme)

;;;###autoload
(add-to-list 'auto-mode-alist '("\\.scsh\\'" . scsh-mode))

;;;###autoload
(define-derived-mode scsh-mode scheme-mode "Scsh"
  "Major mode for Scheme Shell.
\\{scsh-mode-map}"
  ;; vertical bar is an ordinary symbol character, not a delimiter,
  ;; but #| |# comments should still work.
  (modify-syntax-entry ?\| "_ 23bn"))

(provide 'scsh-mode)
