;; osd-mode.el --- major mode for editing Orocos state machine files
;; Jamison Hope 2010-12-06 15:37:14EST jrh osd-mode.el
;; Time-stamp: <2013-12-05 16:54:24EST jrh osd-mode.el>
;; $Id: $

;; Inspired by js.el and other things.

(eval-and-compile
  (require 'cc-mode)
  (require 'font-lock)
  (require 'paren))

(defconst osd--name-start-re "[a-zA-Z]"
  "Regexp matching the start of an OSD identifier, without grouping.")

;; (defconst osd--stmt-delim-chars "^;{}?:")

(defconst osd--name-re (concat osd--name-start-re
                               "\\(?:\\s_\\|\\sw\\)*")
  "Regexp matching an OSD identifier, without grouping.")

;; var TYPE NAME
(defconst osd--var-decl-re
  (concat "^\\s-*var\\s-+"
          "\\(" osd--name-re "\\)\\s-+"
          "\\(" osd--name-re "\\)"))

;; StateMachine NAME
(defconst osd--sm-heading-re
  (concat
   "^\\s-*StateMachine\\s-+\\(" osd--name-re "\\)")
  "Regexp matching the start of a StateMachine.
Match group 1 is the name of the SM.")

;; SubMachine TYPE NAME
(defconst osd--smvar-decl-re
  (concat "^\\s-*SubMachine\\s-+"
          "\\(" osd--name-re "\\)\\s-+"
          "\\(" osd--name-re "\\)"))

;; initial state NAME
(defconst osd--init-state-heading-re
  (concat
   "^\\s-*initial\\s-+state\\s-+\\(" osd--name-re "\\)")
  "Regexp matching the start of a state.
Match group 1 is the name of the state.")

;; state NAME
(defconst osd--state-heading-re
  (concat
   "^\\s-*state\\s-+\\(" osd--name-re "\\)")
  "Regexp matching the start of a state.
Match group 1 is the name of the state.")

;; final state NAME
(defconst osd--final-state-heading-re
  (concat
   "^\\s-*final\\s-+state\\s-+\\(" osd--name-re "\\)")
  "regexp matching the start of a state.
Match group 1 is the name of the state.")

;; select NAME
(defconst osd--select-state-re
  (concat "select\\s-+\\(" osd--name-re "\\)")
  "regexp matching the \"select STATE\" part of a transition
statement.
Match group 1 is the name of the state.")

;; ROOTMACHINE CLASS INSTANCE
(defconst osd--state-instantiation-re
  (concat
   "^\\s-*\\(RootMachine\\)\\s-+\\("
   osd--name-re "\\)\\s-+\\("
   osd--name-re "\\)"))

;; keywords
(defconst osd--keyword-re
  (regexp-opt
   '("new" "delete" "this" "do" "set" "select"
     "if" "then" "else" "switch"
     "throw" "try" "catch"
     "var" "const" "SubMachine"
     "StateMachine" "state" "entry" "handle" "run" "exit"
     "transition" "transitions" "preconditions"
     "initial" "final") 'symbols)
  "Regexp matching any OSD keyword.")

;; types
(defconst osd--basic-type-re
  (regexp-opt '("bool" "double" "array" "frame") 'symbols)
  "Regular expression matching any predefined type in OSD.")

;; constants
(defconst osd--constant-re
  (regexp-opt '("false" "true") 'symbols)
  "Regexp matching constants in OSD.")

(defconst osd--font-lock-keywords-1
  `((,osd--sm-heading-re          (1 font-lock-function-name-face))
    (,osd--init-state-heading-re  (1 font-lock-variable-name-face))
    (,osd--state-heading-re       (1 font-lock-variable-name-face))
    (,osd--final-state-heading-re (1 font-lock-variable-name-face))
    (,osd--state-instantiation-re (1 font-lock-keyword-face)
                                  (2 font-lock-function-name-face)
                                  (3 font-lock-variable-name-face)))
  "Minimal highlighting expressions for `osd-mode'.")

(defconst osd--font-lock-keywords-2
  (append osd--font-lock-keywords-1
          `((,osd--var-decl-re     (1 font-lock-type-face)
                                   (2 font-lock-variable-name-face))
            (,osd--smvar-decl-re   (1 font-lock-type-face)
                                   (2 font-lock-variable-name-face))
            (,osd--select-state-re (1 font-lock-variable-name-face))
            (,osd--keyword-re      (1 font-lock-keyword-face))
            (,osd--basic-type-re   (1 font-lock-type-face))
            (,osd--constant-re     (1 font-lock-constant-face))))
  "Level two font lock keywords for `osd-mode'.")

(defconst osd--font-lock-keywords-3
  (append osd--font-lock-keywords-2
          (list))
  "All the highlighting in `osd-mode'.")

(defvar osd-font-lock-keywords osd--font-lock-keywords-3
  "Default highlighting expressions for `osd-mode'.")

;;; KeyMap & Indent

(defun osd-electric-brace (arg)
  "Insert a brace and indent."
  (interactive "*P")
  (self-insert-command (prefix-numeric-value arg))
  (osd-indent-line))

(defun osd-electric-paren (arg)
  "Insert a paren and indent."
  (interactive "*P")
  (self-insert-command (prefix-numeric-value arg))
  (osd-indent-line))

(defun osd-electric-newline (&optional arg)
  "Indent and then insert newline."
  (interactive "*P")
  (osd-indent-line)
  (delete-trailing-whitespace)
  (newline arg))


(defvar osd-mode-map
  (let ((map (make-keymap)))
    (define-key map (kbd "{") 'osd-electric-brace)
    (define-key map (kbd "}") 'osd-electric-brace)
    (define-key map (kbd "(") 'osd-electric-paren)
    (define-key map (kbd ")") 'osd-electric-paren)
    (define-key map (kbd "RET") 'osd-electric-newline)
    map)
  "Keymap for OSD major mode")


(defun osd--column-of (char-as-string &optional from-end)
  (save-excursion
    (if from-end
        (progn ;; todo bound the search
          (end-of-line)
          (search-backward char-as-string)
          (current-column))
      (progn
        (search-forward char-as-string)
        (current-column)))))

(defun osd-indent-line ()
  "Indent current line as OSD code"
  (interactive)
  (let ((orig-indent (current-indentation))
        (orig-column (current-column)))
    (beginning-of-line)
    (if (bobp)                          ; beginning-of-buffer
        (progn
          (indent-line-to 0)
          (move-to-column (- orig-column orig-indent)))
      (let ((not-indented t) cur-indent)
        (cond ((looking-at "^[ \t]*}")
               ;; logic borrowed from show-paren-function
               (save-excursion
                 (save-restriction
                   (let (pos)
                     (search-forward "}" nil nil 1)
                     (condition-case ()
                         (setq pos (scan-sexps (point) -1))
                       (error (setq pos t cur-indent 0 not-indented nil)))
                     (when (integerp pos)
                       (unless (condition-case ()
                                   (eq (point) (scan-sexps pos 1))
                                 (error nil))
                         (setq pos nil)))
                     (when (integerp pos)
                       (goto-char pos)
                       (setq cur-indent (current-indentation)
                             not-indented nil)))))
               (when (< cur-indent 0) (setq cur-indent 0)))
              ((looking-at "^[ \t]*select\\b.*")
               (save-excursion
                 (while not-indented
                   (forward-line -1)
                   (cond ((looking-at "^[ \t]*transition.*$")
                          (setq cur-indent (+ (current-indentation)
                                              standard-indent)
                                not-indented nil))))))
              (t
               (save-excursion
                 (while not-indented
                   (forward-line -1)
                   (cond ((looking-at "^[ \t]*}")
                          (setq cur-indent (current-indentation))
                          (setq not-indented nil))
                         ((looking-at "^.*{[ \t]*$")
                          (setq cur-indent (+ (current-indentation)
                                              standard-indent))
                          (setq not-indented nil))
                         ((bobp)
                          (setq not-indented nil))
                         ((looking-at "^.*([^)]*$")
                          (setq cur-indent (osd--column-of "(")
                                not-indented nil)))))))
        (unless cur-indent (setq cur-indent 0))
        (indent-line-to cur-indent)
        (when (> orig-column orig-indent)
          (move-to-column (+ cur-indent (- orig-column orig-indent))))
        ;; (when (> orig-column cur-indent)
        ;;   (move-to-column (+ cur-indent (- orig-column orig-indent))))
))))

(defvar osd-mode-hook nil)

;;;###autoload
(add-to-list 'auto-mode-alist '("\\.osd\\'" . osd-mode))

(defvar osd-mode-syntax-table
  (let ((st (make-syntax-table)))
    (modify-syntax-entry ?/ ". 124b" st)
    (modify-syntax-entry ?* ". 23" st)
    (modify-syntax-entry ?\n "> b" st)
    st)
  "Syntax table for osd-mode")


(defun osd-mode ()
  "Major mode for editing OSD files"
  (interactive)
  (kill-all-local-variables)
  (set-syntax-table osd-mode-syntax-table)
  (use-local-map osd-mode-map)
  (set (make-local-variable 'font-lock-defaults)
       '(osd-font-lock-keywords))
  (set (make-local-variable 'indent-line-function)
       'osd-indent-line)
  (set (make-local-variable 'comment-start) "// ")
  (setq major-mode 'osd-mode)
  (setq mode-name "osd")
  (run-hooks 'osd-mode-hook))

(provide 'osd-mode)
