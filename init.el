;;; Emacs init file.

;;; Put MacPorts bin dir in PATH.
(when (eq system-type 'darwin)
  (push "/opt/local/bin" exec-path)
  (setenv "PATH" (concat "/opt/local/bin:" (getenv "PATH"))))

;;; Turn off Press and Hold
(when (and (fboundp 'ns-set-resource) (display-graphic-p))
  (ns-set-resource nil "ApplePressAndHoldEnabled" "NO"))

;;; Set frame title format to distinguish Emacs and EmacsMac windows
(when (and (display-graphic-p)
           (eq system-type 'darwin)
           (string-prefix-p "/Applications/MacPorts" invocation-directory))
  (setq frame-title-format
        `("" "%b" " (" ,(replace-regexp-in-string
                         "/Applications/MacPorts/\\(.*\\)\\.app/.*"
                         "\\1" invocation-directory) ")")))

;;; Put trash in ~/.Trash (Emacs 23.1+)
(when (>= emacs-major-version 23)
  (setq delete-by-moving-to-trash t)
  (when (eq system-type 'darwin)
    (setq trash-directory "~/.Trash")))

;;; Set email address for ChangeLog entries
(cond ((executable-find "git")
       (setq user-mail-address
             (shell-command-to-string
              "git config user.email | tr -d \" \\n\"")))
      ((executable-find "contacts")
       (setq user-mail-address
             (shell-command-to-string
              "contacts -Hmlf %we | tr -d \" \\n\""))))

;;; Always send mail via mailclient on Mac
(when (eq system-type 'darwin)
  (setq send-mail-function 'mailclient-send-it))

;;; Inhibit startup message
(setq inhibit-startup-screen t)

;;; Turn on Show Paren mode
(show-paren-mode t)

;;; Turn on syntax highlighting
(global-font-lock-mode t)

;;; Show trailing whitespace, except in eww and Buffer-menu
(setq-default show-trailing-whitespace t)
(let ((hide-whitespace
       (lambda () (setq show-trailing-whitespace nil))))
  (dolist (hook '(eww-mode-hook
                  Buffer-menu-mode-hook))
    (add-hook hook hide-whitespace)))

;;; Turn on Transient Mark mode
(setq transient-mark-mode t)

;;; Hide the toolbar
(when (fboundp 'tool-bar-mode)
  (tool-bar-mode -1))

;;; Display date/time in modeline
(setq display-time-day-and-date t)
(display-time)

;;; Show line/column numbers in modeline
(line-number-mode 1)
(column-number-mode 1)

;;; Show function name in modeline
(which-function-mode 1)

;;; Ensure files end with newline
(setq require-final-newline t)

;;; Don't use tab characters
(setq-default indent-tabs-mode nil)

;;; Indent by 2 spaces
(setq standard-indent 2)
(setq c-basic-offset 2)

;;; Autofill
(add-hook 'text-mode-hook 'turn-on-auto-fill)
(add-hook 'emacs-lisp-mode-hook 'turn-on-auto-fill)
(setq-default auto-fill-function 'do-auto-fill)

;;; SubWord mode
(when (locate-library "subword")
  (add-hook 'c-mode-common-hook 'subword-mode)
  (add-hook 'osd-mode-hook 'subword-mode))

;;; Don't reindent on newlines
(when (locate-library "electric")
  (setq electric-indent-chars '()))

;;; Locate with Spotlight -- http://emacswiki.org/emacs/MacOSTweaks
(when (eq system-type 'darwin)
  (setq locate-command "mdfind"))

;;; Make customize use a separate file
(if (>= emacs-major-version 24)
    (setq custom-file (expand-file-name "custom.el" user-emacs-directory))
  (setq custom-file (expand-file-name "~/.emacs.d/custom-old.el")))
(when (file-exists-p custom-file)
  (load custom-file))

;;; Now do package-y stuff on Emacs 24+.
(cond
 ((>= emacs-major-version 24)
  (require 'package)
  (setq package-enable-at-startup nil)
  (add-to-list 'package-archives '("melpa" . "https://melpa.org/packages/") t)
  (add-to-list 'package-archives '("org"   . "http://orgmode.org/elpa/")    t)
  (package-initialize)

  ;; Bootstrap org and org-plus-contrib
  (unless (package-installed-p 'org)
    (package-refresh-contents)
    (package-install 'org))
  (unless (package-installed-p 'org-plus-contrib)
    (package-refresh-contents)
    (package-install 'org-plus-contrib))
  ;; Also install other third party export backends.
  (unless (and (package-installed-p 'ox-mediawiki))
    (package-refresh-contents)
    (package-install 'ox-mediawiki))

  ;; These things must be set before loading org-mode.

  ;; Enable these export backends.
  (setq org-export-backends
        '(ascii beamer html icalendar latex man md odt org texinfo
                ;; mediawiki steals ?m from md in the export menu
                ;; so leave it off of the list
                ))

  ;; Allow alphabetical bullets in lists.
  (setq org-list-allow-alphabetical t)

  ;; Load these modules with org-mode.
  (setq org-modules
        '(org-bbdb org-bibtex org-docview org-gnus org-id org-info
          org-irc org-mhe org-protocol org-rmail org-w3m org-bookmark
          org-elisp-symbol org-mac-link org-man org-vm))

  ;; soffice won't be in PATH on Mac OS
  (when (eq system-type 'darwin)
    (setq org-odt-convert-processes
          '(("LibreOffice" "/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to %f%x --outdir %d %i")
            ("unoconv" "unoconv -f %f -o %d %i"))))

  ;; Open org files fully expanded
  (setq org-startup-folded 'showeverything)

  ;; allow open paren to be a postmatch, and close paren to be a
  ;; prematch
  (setq org-emphasis-regexp-components
        '("- \t()'\"{" "- \t.,:!?;'\"()}\\[" " \t\r\n" "." 1))

  ;; Enable these languages to be evaluated in code blocks.
  (org-babel-do-load-languages
   'org-babel-load-languages
   '((emacs-lisp . t)
     (ditaa . t)
     (dot . t)
     (latex . t)
     (plantuml . t)
     (shell . t)))

  ;; Load config.org, where the bulk of settings and package loads
  ;; are.
  (org-babel-load-file
   (expand-file-name "config.org" user-emacs-directory)))

 (t ;; old Emacs
  (load-file "~/.emacs.d/init-old.el")))
