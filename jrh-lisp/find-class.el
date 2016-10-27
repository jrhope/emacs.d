;;; Use ctags to locate the declaration of the given C++ or Java
;;; class.

(require 'cl)

(defvar fc--temp-file-alist nil
  "The mapping from directories to ctags temp files.")

(defun fc--lookup-temp-file (dir)
  "Returns a temp file for DIR or NIL."
  (cdr (cl-assoc (file-truename dir) fc--temp-file-alist
                 :test #'equal)))

(defun fc--store-temp-file (dir tmpfile)
  "Stores a (DIR . TMPFILE) pair in `fc--temp-file-alist'."
  (push (cons (file-truename dir) tmpfile) fc--temp-file-alist))

(defun fc--generate-tags-file (dir tmpfile)
  "Run ctags on headers rooted at DIR and store in TMPFILE."
  (shell-command
   (concat "ctags -R --exclude='*.html'"
           " --c-kinds=c --c++-kinds=c --java-kinds=c"
           (format " -f %s %s" tmpfile dir))))

(defun fc-find-class (classname dir)
  "Search for the given C++ or Java class."
  (interactive
   (let ((dflt (or (caar fc--temp-file-alist)
                   default-directory)))
     (list (read-string "Class name: ")
           (read-file-name "Directory: " dflt dflt))))
  (if (and classname dir (file-directory-p dir))
      (let ((tmpfile (fc--lookup-temp-file dir)))
        (unless tmpfile
          (setf tmpfile (make-temp-file "tags"))
          (fc--generate-tags-file dir tmpfile)
          (fc--store-temp-file dir tmpfile))
        (shell-command-to-string
         (format "grep -E %s %s"
                 (shell-quote-argument
                  (concat "^" classname "[[:space:]].*[[:space:]]\\bc\\b"))
                 tmpfile)))
    (error "fc-find-class: Bad dir: %s or class name: %s" dir
           classname)))

;;;###autoload
(defun find-file-for-class (classname dir)
  "Visit the file(s) declaring the given class."
  (interactive
   (let ((dflt (or (caar fc--temp-file-alist)
                   default-directory)))
     (list (read-string "Class name: ")
           (read-file-name "Directory: " dflt dflt))))
  (let* ((files (split-string (fc-find-class classname dir) "\n" t)))
    (if (not files)
        (minibuffer-message "No file found for class %s" classname)
      (let ((size (/ (window-height) (length files))))
        (let ((f-parts (split-string (car files) "\t")))
          (find-file (cadr f-parts))
          (beginning-of-buffer)
          (search-forward-regexp
           (replace-regexp-in-string "/\\(.*\\)$.*" "\\1" (caddr f-parts)))
          (beginning-of-line))
        (dolist (file (cdr files))
          (let ((f-parts (split-string file "\t")))
            (select-window (split-window nil size))
            (find-file (cadr f-parts))
            (beginning-of-buffer)
            (search-forward-regexp
             (replace-regexp-in-string "/\\(.*\\)$.*" "\\1"
                                       (caddr f-parts)))
            (beginning-of-line))))
      (if (= 1 (length files))
          (minibuffer-message "Opened file %s for class %s"
                              (file-name-nondirectory
                               (buffer-file-name)) classname)
        (minibuffer-message "Class %s found in %s files"
                            classname (length files))))))

(provide 'find-class)
