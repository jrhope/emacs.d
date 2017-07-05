;; @file jrh-lisp/git-stat-all.el
;; @author Jamison Hope <jrh@theptrgroup.com>

(if (locate-library "subr-x")
    (require 'subr-x)
  (defsubst string-trim-right (string)
    "Remove trailing whitespace from STRING."
    (if (string-match "[ \t\n\r]+\\'" string)
        (replace-match "" t t string)
      string)))

(defun jrh/git-stat-all (parent-dir)
  (interactive "Dparent directory: ")

  (mapc
   (if (fboundp 'magit-status-internal)
       #'magit-status-internal
     #'git-status)
   (split-string
    (string-trim-right
     (shell-command-to-string
      (concat "find " parent-dir
              " -type d -name .git |"
              "sed -e 's/.git//'")))
    "\n")))

(provide 'git-stat-all)
