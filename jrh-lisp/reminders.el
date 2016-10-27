;;; Interface with Reminders.app

(require 'org-capture)

(defvar jrh/org-capture-reminder-file
  (concat temporary-file-directory "reminder.org")
  "Temporary file for storing reminder data.")

(defconst jrh/reminder-file-template
  "Date:%^{Reminder Date}U%?\nTitle:%i\nNote:"
  "Template for Reminder org-capture file.")

;; see #'org-capture-set-target-location
(defun jrh/reminder-file ()
  "Prepares the Reminder data as an org-capture target buffer."
  (org-capture-put :jrh-create-reminder t)
  (set-buffer (org-capture-target-buffer jrh/org-capture-reminder-file))
  (org-capture-put-target-region-and-position)
  (widen)
  (goto-char (point-max)))

(defun jrh/create-reminder (name remind-date &optional note)
  "Uses parsed reminder file data to run an AppleScript which makes
the new reminder."
  (let ((maybe-note (if (and note
                             (> (length note) 0))
                        (concat ",body:\"" note "\"") "")))
    (do-applescript
     (concat
      "tell application \"Reminders\"\n"
      "tell default list\n"
      "make reminder with properties {"
      "name:\"" name "\","
      "remind me date:date \"" remind-date "\""
      maybe-note "}\nend tell\nend tell"))))

(defun jrh/org-date->applescript-date (date)
  "Convert an org-mode style date to AppleScript style date."
  (let ((year (substring date 1 5))
        (month (substring date 6 8))
        (day (substring date 9 11))
        (time (substring date 16 21)))
    (concat month "-" day "-" year " " time " AM")))

(defun jrh/maybe-create-reminder ()
  "An org-capture-prepare-finalize hook which will invoke
jrh/create-reminder if the capture was one of ours (i.e. if the
:jrh-create-reminder breadcrumb was set and the capture wasn't
aborted.)"
  (let ((create (org-capture-get :jrh-create-reminder))
        date name note m1 m2)
    (when (and create (not org-note-abort))
      (setq date (buffer-substring (+ (point-min) 5)
                                   (+ (point-min) 27)))
      (goto-char (point-min)) (forward-line 1)
      (forward-char 6) ; "Title:"
      (setq m1 (point))
      (end-of-line)
      (setq m2 (point))
      (setq name (buffer-substring m1 m2))
      (forward-line 1)
      (forward-char 5) ; "Note:"
      (setq note (buffer-substring (point) (point-max)))
      (jrh/create-reminder name (jrh/org-date->applescript-date date) note))))

(provide 'reminders)
