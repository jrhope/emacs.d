;;; bytecode-mode.el --- View Java class files disassembled by
;;; gnu.bytecode.dump.  -*- lexical-binding: t; -*-

;; Copyright (C) 2016  Jamison Hope <jhope@alum.mit.edu>

;; Permission is hereby granted, free of charge, to any person obtaining a copy
;; of this software and associated documentation files (the "Software"), to deal
;; in the Software without restriction, including without limitation the rights
;; to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
;; copies of the Software, and to permit persons to whom the Software is
;; furnished to do so, subject to the following conditions:
;; The above copyright notice and this permission notice shall be included in
;; all copies or substantial portions of the Software.

;; THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
;; IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
;; FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
;; AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
;; LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
;; OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
;; THE SOFTWARE.

;;; Inspired primarily by javap-mode.el and
;;; <http://nullprogram.com/blog/2012/08/01/>.


;; Bytecode Mode consists of two parts: `bytecode-mode', and
;; `bytecode-handler'.  `bytecode-mode' is a major mode which provides
;; syntax highlighting for text in the format of Kawa's disassembler,
;; gnu.bytecode.dump.  `bytecode-handler' is a file name handler for
;; Java class files (i.e. "\\.class\\'") which intercepts
;; `insert-file-contents' and instead of inserting the raw class file
;; bytes, inserts the results of invoking gnu.bytecode.dump.

;; There are three customization settings:
;; - `bytecode-mode-classpath', which should be a list of paths for
;;   constructing the classpath to invoke the disassembler.  It should
;;   typically include a Kawa jar in there somewhere.  This defaults
;;   to nil.
;; - `bytecode-mode-disassembler', which is the name of the Java
;;   program to execute.  This defaults to "gnu.bytecode.dump".
;; - `bytecode-mode-initially-verbose', which is either t or nil
;;   (defaults to t).
;; - `bytecode-mode-disassembler-options', which is a list of other
;;   flags to pass to the disassembler (they will appear before the
;;   class file in the argument list).  This defaults to
;;   '("--verbose").

(defgroup bytecode-mode nil
  "In-buffer viewer for Java class files disassembled by Kawa."
  :group 'tools
  :prefix "bytecode-mode-")

(defcustom bytecode-mode-classpath nil
  "Classpath containing the disassembler."
  :type '(repeat (file :tag "Path"))
  :group 'bytecode-mode)

(defcustom bytecode-mode-disassembler "gnu.bytecode.dump"
  "Java disassembler to use, as fully qualified class name."
  :type 'string
  :group 'bytecode-mode)

(defcustom bytecode-mode-initially-verbose t
  "Whether output should default to --verbose."
  :type 'boolean
  :group 'bytecode-mode)

;; (defcustom bytecode-mode-disassembler-options (list "--verbose")
;;   "Extra arguments to the disassembler."
;;   :type '(repeat string)
;;   :group 'bytecode-mode)


(defconst bytecode-font-lock-keywords
  (eval-when-compile
    `(
      ("^Reading \\.class from .*$" . font-lock-comment-face)
      ("^ *line: [0-9]+ at pc: [0-9]+" . font-lock-comment-face)
      ("\\(#[0-9]+\\)="  (1 font-lock-constant-face))
      ("^\\(#[0-9]+\\):" (1 font-lock-constant-face))
      ("^ *-?[0-9]+:" . font-lock-string-face)
      ("\\(\"\\w+\"\\)" . font-lock-string-face)
      ("\\<@?[a-zA-Z]+\\.[a-zA-Z0-9._$]+\\>" . font-lock-type-face)
      (,(concat "\\<"
                (regexp-opt '("public" "protected" "private"
                              "static" "super" "native" "abstract"))
                "\\>") . font-lock-keyword-face)
      (,(concat "\\<"
                (regexp-opt '("invokevirtual" "invokespecial"
                              "invokestatic" "invokeinterface"
                              "invokedynamic"))
                "\\>") . font-lock-function-name-face)
      (,(concat "\\<"
                (regexp-opt '("void" "boolean" "char" "short" "int" "long"
                              "float" "double"))
                "\\>")
       . font-lock-type-face)
      ("\\<[ilfd]const_m?[0-5]\\>" . font-lock-constant-face)
      ("\\<[bs]ipush\\>" . font-lock-keyword-face)
      (,(concat "\\<" (regexp-opt '("ldc" "ldc_w" "ldc2_w")) "\\>")
       . font-lock-keyword-face)
      ("\\<[ilfda]\\(load\\|store\\)\\(_[0-3]\\)?\\>"
       . font-lock-variable-name-face)
      ("\\<[ilfdabcs]a\\(load\\|store\\)\\>" . font-lock-variable-name-face)
      ("\\<pop2?\\>" . font-lock-keyword-face)
      ("\\<dup2?\\(_x[12]\\)?\\>" . font-lock-keyword-face)
      ("\\<[ilfd]\\(add\\|sub\\|mul\\|div\\|rem\\|neg\\)\\>"
       . font-lock-keyword-face)
      ("\\<[il]\\(shl\\|u?shr\\|and\\|or\\|xor\\)\\>"
       . font-lock-keyword-face)
      ("\\<i2[lfdbcs]\\>" . font-lock-keyword-face)
      ("\\<l2[ifd]\\>" . font-lock-keyword-face)
      ("\\<f2[ild]\\>" . font-lock-keyword-face)
      ("\\<d2[ilf]\\>" . font-lock-keyword-face)
      ("\\<[fd]cmp[lg]\\>" . font-lock-keyword-face)
      ("\\<if\\(_icmp\\)?\\(eq\\|ne\\|lt\\|ge\\|gt\\|le\\)\\>"
       . font-lock-keyword-face)
      ("\\<if_acmp\\(eq\\|ne\\)\\>" . font-lock-keyword-face)
      ("\\<if\\(non\\)?null\\>" . font-lock-keyword-face)
      ("\\<\\(goto\\|jsr\\)\\(_w\\)?\\>" . font-lock-keyword-face)
      ("\\<[ilfda]?return\\>" . font-lock-keyword-face)
      ("\\<\\(get\\|put\\)\\(static\\|field\\)\\>"
       . font-lock-variable-name-face)
      ("\\<\\(\\(multi\\)?a\\)?newarray\\>" . font-lock-keyword-face)
      ("\\<monitor\\(enter\\|exit\\)\\>" . font-lock-keyword-face)
      (,(concat "\\<" (regexp-opt
                       '("nop" "aconst_null" "swap" "iinc" "lcmp"
                         "ret" "tableswitch" "lookupswitch" "new"
                         "arraylength" "instanceof" "wide"))
                "\\>") . font-lock-keyword-face)
      (,(concat "\\<" (regexp-opt '("athrow" "checkcast")) "\\>")
       . font-lock-warning-face)))
  "Default expressions to highlight in bytecode mode.")

(defvar bytecode-mode-syntax-table
  (let ((table (make-syntax-table)))
    ;; suppress backslash-escaping
    (modify-syntax-entry ?\\ "w" table)
    table)
  "Syntax table for use in bytecode-mode.")

(defvar bytecode-mode-map
  (let ((map (make-sparse-keymap)))
    ;; We start with the special-mode-map set.
    (set-keymap-parent map special-mode-map)
    (define-key map (kbd "t") 'bytecode-toggle-verbosity)
    (define-key map (kbd "v") 'bytecode-toggle-verbosity)
    (define-key map (kbd "r") 'revert-buffer)
    (define-key map (kbd "n") 'forward-paragraph)
    (define-key map (kbd "p") 'backward-paragraph)
    map)
  "Keymap used by `bytecode-mode'.")

(put 'bytecode-verbose 'permanent-local t)
(defvar bytecode-verbose bytecode-mode-initially-verbose
  "If true, then output will be --verbose.

See `bytecode-toggle-verbosity'.")

(defun bytecode--get-classpath ()
  (if bytecode-mode-classpath
      (string-join bytecode-mode-classpath ":")
    (getenv "CLASSPATH")))

(defun bytecode-insert-file (buffer file)
  "Inserts the disassembly of the given class file into the specified
buffer."
  (interactive "bBuffer: \nfFile: ")
  (when (string-match "\\.jar:" file)
    (setq file (concat "jar:file:"
                       (replace-regexp-in-string "\\.jar:" ".jar!/"
                                                 file))))
  (let* ((cp (bytecode--get-classpath))
         (cp-args (if cp (list "-cp" cp) nil))
         (verbosity (if bytecode-verbose (list "--verbose") nil))
         (proc `("java" nil ,buffer nil
                 ,@cp-args
                 ,bytecode-mode-disassembler
                 ,@verbosity
                 ,file)))
    (apply #'call-process proc)))

(defun bytecode-mode--where-am-i ()
  "Tries to determine the current location within a class dump, in a
format which can be used to relocate the point after toggling
verbosity."
  ;; The sections of a dump are:
  ;; 1. "Reading .class from ..."
  ;; 2. "Classfile format major version ..."
  ;; 3. Constant Pool (only if verbose)
  ;; 4. Class info ("Access flags: ...")
  ;; 5. Fields
  ;; 6. Methods
  ;; 7. Attributes
  (let ((my-line (line-number-at-pos))
        (my-col (current-column)))
    (beginning-of-line)
    (cond
     ((bobp)
      `(top ,my-col))
     ((looking-at-p "^Classfile format major version:")
      `(version ,my-col))
     ((looking-at-p "^#[0-9]+:")
      `(cpool ,(- my-line 1) ,my-col))
     (t
      (beginning-of-buffer)
      (let* ((accs-line (progn (re-search-forward "^Access flags:")
                               (line-number-at-pos)))
             (flds-line (progn (re-search-forward "^Fields (count:")
                               (line-number-at-pos)))
             (mtds-line (progn (re-search-forward "^Methods (count:")
                               (line-number-at-pos)))
             (atts-line (progn (re-search-forward "^Attributes (count:")
                               (line-number-at-pos))))
        (cond ((< my-line flds-line)
               `(access ,(- my-line accs-line) ,my-col))
              ((< my-line mtds-line)
               `(fields ,(- my-line flds-line) ,my-col))
              ((< my-line atts-line)
               `(methods ,(- my-line mtds-line) ,my-col))
              (t `(attributes ,(- my-line atts-line) ,my-col))))))))

(defun bytecode-mode--go-back (loc)
  ;(message "going back: %s" loc)
  (beginning-of-buffer)
  (cond
   ((eq 'top (car loc))
    (move-to-column (cadr loc)))
   ((eq 'version (car loc))
    (when bytecode-verbose
      (forward-line 1)
      (move-to-column (cadr loc))))
   ((eq 'cpool (car loc))
    (when bytecode-verbose
      (re-search-forward "^#1:")
      (forward-line (cadr loc))
      (move-to-column (caddr loc))))
   ((eq 'access (car loc))
    (re-search-forward "^Access flags:")
    (forward-line (cadr loc))
    (move-to-column (caddr loc)))
   ((eq 'fields (car loc))
    (re-search-forward "^Fields (count:")
    (forward-line (cadr loc))
    (move-to-column (caddr loc)))
   ((eq 'methods (car loc))
    (re-search-forward "^Methods (count:")
    (forward-line (cadr loc))
    (move-to-column (caddr loc)))
   ((eq 'attributes (car loc))
    (re-search-forward "^Attributes (count:")
    (forward-line (cadr loc))
    (move-to-column (caddr loc))))
  ;(recenter)
  )

(defun bytecode-revert-buffer (file-name auto-save-p)
  (let ((loc (bytecode-mode--where-am-i)))
    ;(message "where-am-i: %s" loc)
    (setq buffer-read-only nil)
    (erase-buffer)
    (bytecode-insert-file (current-buffer) file-name)
    (setq buffer-read-only t)
    (set-buffer-modified-p nil)
    ;; This works when called from bytecode-toggle-verbosity, but not
    ;; when called from revert-buffer.
    (bytecode-mode--go-back loc)))

(defun bytecode-toggle-verbosity ()
  (interactive)
  (setq-local bytecode-verbose (not bytecode-verbose))
  (bytecode-revert-buffer (buffer-file-name) nil))

;;;###autoload
(define-derived-mode bytecode-mode special-mode "Bytecode"
  "A major mode for viewing bytecode files."
  :syntax-table bytecode-mode-syntax-table
  (setq-local font-lock-defaults '(bytecode-font-lock-keywords))
  (use-local-map bytecode-mode-map)
  ;; If we're looking at a raw class file, try to reload as
  ;; disassembly.  This happens if the class is inside a jar.
  (when (string-match "\312\376\272\276" ; CA FE BA BE
                      (buffer-substring-no-properties 1 5))
    (bytecode-revert-buffer (buffer-file-name) nil))
  (setq-local revert-buffer-insert-file-contents-function 'bytecode-revert-buffer)
  (goto-char (point-min)))

(defun bytecode-handler-real (op args)
  "Run the real handler without the bytecode handler installed."
  (let ((inhibit-file-name-handlers
         (cons 'bytecode-handler
               (and (eq inhibit-file-name-operation op)
                    inhibit-file-name-handlers)))
        (inhibit-file-name-operation op))
    ;(message "bytecode-handler-real op: %s, args: %s" op args)
    (apply op args)))

;;;###autoload
(defun bytecode-handler (op &rest args)
  "Handle .class files by putting the output of gnu.bytecode.dump in
the buffer."
  ;(message "bytecode-handler (%s, %s)" op args)
  (cond
   ((eq op 'insert-file-contents)
    (let ((file (car args)))
      (bytecode-revert-buffer file nil)
      (setq buffer-file-name file)
      (goto-char (point-min))
      (list (expand-file-name file) (- (point-max) (point-min)))))
   ((eq op 'file-writable-p) nil)
   ((eq op 'vc-registered)
    ;; Don't try to call vc-registered on a class file within a jar;
    ;; instead just call vc-registered on the containing jar.
    ;; (Without this, some step in the chain tries to cd into the
    ;; jar and then complains about no such directory.)
    (let ((file (car args)))
      (when (string-match "\\.jar:.*\\.class\\'" file)
        (setq file (replace-regexp-in-string "\\.jar:.*\\'" ".jar"
                                             file)))
      (bytecode-handler-real op (list file))))
   (t (bytecode-handler-real op args))))

(add-to-list 'auto-mode-alist '("\\.class\\'" . bytecode-mode))
(add-to-list 'file-name-handler-alist '("\\.class\\'" . bytecode-handler))

(provide 'bytecode-mode)
;;; bytecode-mode.el ends here
