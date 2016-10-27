;;; Emacs 22 init (for the old Mac /usr/bin/emacs)

;;; Use IswitchB
(when (locate-library "iswitchb")
  (require 'iswitchb)
  (iswitchb-mode 1)
  (setq iswitchb-case t))

;;; Uniquify buffer names
(when (locate-library "uniquify")
  (require 'uniquify)
  (setq uniquify-buffer-name-style 'post-forward))

;;; Git
(add-to-list 'load-path "/opt/local/share/git/contrib/emacs")
(when (locate-library "git")
  (autoload 'git-status "git"
    "Major mode for doing Git stuff." t))
(when (locate-library "git-blame")
  (autoload 'git-blame-mode "git-blame"
    "Minor mode for incremental blame for Git." t)
  (setq git-blame-prefix-format "%h %25A:"))
