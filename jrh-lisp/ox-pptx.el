;;; ox-pptx.el --- PowerPoint Back-End for Org Export Engine
;;; -*- lexical-binding: t; -*-

;; Copyright (C) 2017  Jamison Hope <jhope at alum dot mit dot edu>

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

(require 'cl-lib)
(require 'ox)
(require 'ox-beamer)
(require 'ox-html)


;;; Define Back-End

(org-export-define-backend 'pptx
  '((bold . org-pptx-bold)
    (center-block . org-pptx-center-block)
    (clock . org-pptx-clock)
    (code . org-pptx-code)
    (drawer . org-pptx-drawer)
    (dynamic-block . org-pptx-dynamic-block)
    (entity . org-pptx-entity)
    (example-block . org-pptx-example-block)
    (export-block . org-pptx-export-block)
    (export-snippet . org-pptx-export-snippet)
    (fixed-width . org-pptx-fixed-width)
    (footnote-definition . org-pptx-footnote-definition)
    (footnote-reference . org-pptx-footnote-reference)
    (headline . org-pptx-headline)
    (horizontal-rule . org-pptx-horizontal-rule)
    (inline-src-block . org-pptx-inline-src-block)
    (inlinetask . org-pptx-inlinetask)
    (italic . org-pptx-italic)
    (item . org-pptx-item)
    (keyword . org-pptx-keyword)
    (latex-environment . org-pptx-latex-environment)
    (latex-fragment . org-pptx-latex-fragment)
    (line-break . org-pptx-line-break)
    (link . org-pptx-link)
    (node-property . org-pptx-node-property)
    (paragraph . org-pptx-paragraph)
    (plain-list . org-pptx-plain-list)
    (plain-text . org-pptx-plain-text)
    (planning . org-pptx-planning)
    (property-drawer . org-pptx-property-drawer)
    (quote-block . org-pptx-quote-block)
    (radio-target . org-pptx-radio-target)
    (section . org-pptx-section)
    (special-block . org-pptx-special-block)
    (src-block . org-pptx-src-block)
    (statistics-cookie . org-pptx-statistics-cookie)
    (strike-through . org-pptx-strike-through)
    (subscript . org-pptx-subscript)
    (superscript . org-pptx-superscript)
    (table . org-pptx-table)
    (table-cell . org-pptx-table-cell)
    (table-row . org-pptx-table-row)
    (target . org-pptx-target)
    (template . org-pptx-template)
    (timestamp . org-pptx-timestamp)
    (underline . org-pptx-underline)
    (verbatim . org-pptx-verbatim)
    (verse-block . org-pptx-verse-block))
  :menu-entry
  '(?p "Export to PowerPoint"
       ((?p "To file" org-pptx-export-to-pptx)
        (?o "To file and open"
            (lambda (a s v b)
              (if a (org-pptx-export-to-pptx t s v b)
                (let ((org-file-apps
                       (cons '("\\.pptx\\'" . system) org-file-apps)))
                  (org-open-file
                   (org-pptx-export-to-pptx nil s v b) nil)))))
        (?s "To Scheme file" org-pptx-export-to-scm))))


;;; User Configurable Variables

(defgroup org-export-pptx nil
  "Options for exporting Org mode files to PowerPoint."
  :tag "Org Export PowerPoint"
  :group 'org-export)

;;;; Templates

(defcustom org-pptx-prototype-pptx-file
  (expand-file-name "~/Documents/Org Presentation Template.pptx")
  "Location of PowerPoint file to use as presentation prototype."
  :group 'org-export-pptx
  :type 'file)

(defcustom org-pptx-template-kawa-file
  (expand-file-name "ox-pptx-template.scm"
                    (file-name-directory (or load-file-name
                                             buffer-file-name)))
  "Location of Kawa Scheme template for pptx export."
  :group 'org-export-pptx
  :type 'file)

(defcustom org-pptx-classpath
  (list
   (expand-file-name "~/java/Kawa/lib/kawa.jar")
   (expand-file-name "~/java/poi/poi-3.16/poi-3.16.jar")
   (expand-file-name "~/java/poi/poi-3.16/poi-ooxml-3.16.jar")
   (expand-file-name "~/java/poi/poi-3.16/poi-ooxml-schemas-3.16.jar")
   (expand-file-name "~/java/poi/poi-3.16/ooxml-lib/xmlbeans-2.6.0.jar")
   (expand-file-name "~/java/poi/ooxml-schemas-1.3.jar"))
  "List of jar files to use as classpath for Kawa + POI."
  :group 'org-export-pptx
  :type '(repeat (file :must-match t)))


;;; Internal Functions
(defun org-pptx--escape-text (text)
  (dolist (pair '(("\\\\" . "\\\\")
                  ("\""   . "\\\"")
                  )
                text)
    (setq text (replace-regexp-in-string (car pair) (cdr pair) text t t))))

(defun org-pptx--footnote-section (info)
  (let* ((fn-alist (org-export-collect-footnote-definitions info))
         (fn-alist
          (cl-loop for (n _type raw) in fn-alist collect
                   (cons n (org-trim (org-export-data raw info))))))
    (when fn-alist
      (concat
       "(create-section-header-slide '((plain-text \"Endnotes\")))\n\n"
       "(create-slide '((plain-text \"Endnotes\"))\n"
       " '((plain-list ordered 0\n"
       (format "%s"
               (mapconcat
                (lambda (fn)
                  (let ((n (car fn)) (def (cdr fn)))
                    (format "  (item %s. nil %s)\n" n def)))
                fn-alist
                "\n"))
       ")))"))))

(defun org-pptx--list-depth (plain-list info)
  (let* ((parent (org-export-get-parent plain-list))
         (ptype (org-element-type parent)))
    (case ptype
      ((item)
       (1+ (org-pptx--list-depth (org-export-get-parent parent) info)))
      ((section)
       (let* ((headline (org-export-get-parent parent))
              (d (org-export-low-level-p headline info)))
         (if (wholenump d) d 0)))
      (t 0))))


;;; Transcode Functions

;;;; Bold

(defun org-pptx-bold (bold contents _info)
  "Transcode BOLD from Org to PPTX.
CONTENTS is the text with bold markup.  INFO is a plist holding
contextual information."
  (let ((blank (org-element-property :post-blank bold)))
    (format "(bold %s %s)" blank contents)))

;;;; Center Block

(defun org-pptx-center-block (_center-block contents _info)
  "Transcode a CENTER-BLOCK element from Org to PPTX.
CONTENTS holds the contents of the center block.  INFO is a plist
holding contextual information."
  (format "(center %s)" contents))

;;;; Clock

(defun org-pptx-clock (clock _contents _info)
  "Transcode a CLOCK element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (format "(clock \"%s\" \"%s\" \"%s\")"
          org-clock-string
          (org-timestamp-translate (org-element-property :value clock))
          (org-element-property :duration clock)))

;;;; Code

(defun org-pptx-code (code _contents _info)
  "Transcode CODE from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let ((blank (org-element-property :post-blank code))
        (output (org-element-property :value code)))
    (setq output (org-pptx--escape-text output))
    (format "(code %s (plain-text \"%s\"))" blank output)))

;;;; Drawer

(defun org-pptx-drawer (drawer contents _info)
  "Transcode a DRAWER element from Org to PPTX.
CONTENTS holds the contents of the drawer.  INFO is a plist holding
contextual information."
  (format "(drawer %s %s)"
          (org-element-property :drawer-name drawer)
          contents))

;;;; Dynamic Block

(defun org-pptx-dynamic-block (_dynamic-block contents _info)
  "Transcode a DYNAMIC-BLOCK element from Org to PPTX.
CONTENTS holds the contents of the block.  INFO is a plist
holding contextual information.  See `org-export-data'."
  contents)

;;;; Entity

(defun org-pptx-entity (entity _contents _info)
  "Transcode an ENTITY element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let ((blank (org-element-property :post-blank entity)))
    (format "(entity %s (plain-text \"%s\"))" blank
            (org-element-property :utf-8 entity))))

;;;; Example Block

(defun org-pptx-example-block (example-block _contents info)
  "Transcode an EXAMPLE-BLOCK element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual"
  (let ((output (org-export-format-code-default example-block info)))
    (setq output (org-pptx--escape-text output))
    (format "(example-block \"%s\")" output)))

;;;; Export Block

(defun org-pptx-export-block (export-block _contents _info)
  "Transcode an EXPORT-BLOCK element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (when (string= (org-element-property :type export-block) "PPTX")
    (org-remove-indentation (org-element-property :value export-block))))

;;;; Export Snippet

(defun org-pptx-export-snippet (export-snippet _contents info)
  "Transcode an EXPORT-SNIPPET object from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (when (eq (org-export-snippet-backend export-snippet) 'pptx)
    (org-export-data (org-element-property :value export-snippet)
                     info)))

;;;; Fixed Width

(defun org-pptx-fixed-width (fixed-width _contents _info)
  "Transcode a FIXED-WIDTH element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let ((output (org-element-property :value fixed-width)))
    (setq output (org-pptx--escape-text output))
    (format "(fixed-width \"%s\")" output)))

;;;; Footnote Definition

; this isn't used
(defun org-pptx-footnote-definition (_footnote-definition _contents _info)
  nil)

;;;; Footnote Reference

(defun org-pptx-footnote-reference (footnote-reference _contents info)
  "Transcode a FOOTNOTE-REFERENCE element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let ((n (org-export-get-footnote-number footnote-reference info)))
    (format "(footnote-reference (plain-text \"[%s]\"))" n)))

;;;; Headline

(defun org-pptx-headline (headline contents info)
  "Transcode a HEADLINE element from Org to PPTX.
CONTENTS holds the contents of the headline.  INFO is a plist holding
contextual information."
  (unless (org-element-property :footnote-section-p headline)
    (let ((level (org-export-get-relative-level headline info))
          (frame-level (org-beamer--frame-level headline info))
          (text (org-export-data
                 (org-element-property :title headline) info)))
      (cond
       ((= level frame-level)
        (format "(create-slide '(%s)\n '(%s))\n\n" text (org-trim contents)))
       ((< level frame-level)
        (let ((hidden (org-element-property :PPTX-NO-SLIDE headline)))
          (if hidden
              (format "%s" contents)
            (format "(create-section-header-slide '(%s)%s)\n\n%s" text
                    (let ((sub (org-element-property :SUBTITLE headline)))
                      (if sub (format " (plain-text \"%s\")" sub) ""))
                    contents))))
       (t
        (let ((numberedp (org-export-numbered-headline-p headline info))
              (numbers (org-export-get-headline-number headline info)))
          (concat
           (when (org-export-first-sibling-p headline info)
             (format "(plain-list %sordered %s\n"
                     (if numberedp "" "un")
                     (1- (org-export-low-level-p headline info))))
           (format "(item %s nil (paragraph %s) %s)"
                   (if numberedp (car (last numbers)) "-")
                   text
                   (if contents (org-trim contents) ""))
           (when (org-export-last-sibling-p headline info)
             ")"))))))))

;;;; Horizontal Rule

(defun org-pptx-horizontal-rule (_horizontal-rule _contents _info)
  "Transcode a HORIZONTAL-RULE element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  nil)

;;;; Inline Src Block

(defun org-pptx-inline-src-block (inline-src-block _contents _info)
  "Transcode an INLINE-SRC-BLOCK element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let ((blank (org-element-property :post-blank inline-src-block))
        (output (org-element-property :value inline-src-block)))
    (setq output (org-pptx--escape-text output))
    (format "(code %s (plain-text \"%s\"))" blank output)))

;;;; Inlinetask

(defun org-pptx-inlinetask (_inlinetask _contents _info)
  "Transcode an INLINETASK element from Org to PPTX.
CONTENTS holds the contents of the block.  INFO is a plist holding
contextual information."
  nil)

;;;; Italic

(defun org-pptx-italic (italic contents _info)
  "Transcode ITALIC from Org to PPTX.
CONTENTS is the text with italic markup.  INFO is a plist holding
contextual information."
  (let ((blank (org-element-property :post-blank italic)))
    (format "(italic %s %s)" blank contents)))

;;;; Item

(defun org-pptx-item (item contents _info)
  "Transcode an ITEM element from Org to PPTX.
CONTENTS holds the contents of the item.  INFO is a plist holding
contextual information."
  (let ((bullet (org-element-property :bullet item))
        (checkbox (org-element-property :checkbox item)))
    (format "(item %s %s %s)" bullet checkbox (org-trim contents))))

;;;; Keyword

(defun org-pptx-keyword (_keyword _contents _info)
  "Transcode a KEYWORD element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  nil)

;;;; Latex Environment

(defun org-pptx-latex-environment (latex-environment _contents info)
  "Transcode a LATEX-ENVIRONMENT element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let* ((latex-frag (org-remove-indentation
                      (org-element-property :value latex-environment)))
         (formula-link (org-trim
                        (org-html-format-latex latex-frag 'dvipng info))))
    (when (and formula-link
               (string-match "file:\\([^]]*\\.png\\)" formula-link))
      (format "(latex-environment \"%s\" (code 0 (plain-text \"%s\")))"
              (match-string 1 formula-link)
              latex-frag))))

;;;; Latex Fragment

(defun org-pptx-latex-fragment (latex-fragment _contents info)
  "Transcode a LATEX-FRAGMENT element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let* ((latex-frag (org-element-property :value latex-fragment))
         (formula-link (org-trim
                        (org-html-format-latex latex-frag 'dvipng info))))
    (when (and formula-link
               (string-match "file:\\([^]]*\\.png\\)" formula-link))
      (format "(latex-fragment \"%s\" (code 1 (plain-text \"%s\")))"
              (match-string 1 formula-link)
              latex-frag))))

;;;; Line Break

(defun org-pptx-line-break (_line-break _contents _info)
  "Transcode a LINE-BREAK element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  "(line-break)")

;;;; Link

(defun org-pptx-link (link desc _info)
  "Transcode a LINK element from Org to PPTX.
DESC is the description part of the link, or the empty string.
INFO is a plist holding contextual information.  See
`org-export-data'."
  (let ((type (org-element-property :type link))
        (path (org-element-property :path link)))
    (if (and (string= type "file")
             (or (string-match "\\.png\\'" path)
                 (string-match "\\.svg\\'" path)
                 (string-match "\\.eps\\'" path)))
        (format "(img \"%s\")" path)
      (format "(link \"%s\" \"%s\" %s)" type path
              (or desc
                  (format "(plain-text \"%s:%s\")" type path))))))

;;;; Node Property

(defun org-pptx-node-property (node-property contents info)
  "Transcode a NODE-PROPERTY element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (format "(node-property \"%s:%s\")"
          (org-element-property :key node-property)
          (let ((value (org-element-property :value node-property)))
            (if value (concat " " value) ""))))

;;;; Paragraph
(defun org-pptx-paragraph (_paragraph contents _info)
  "Transcode a PARAGRAPH element from Org to PPTX.
CONTENTS is the contents of the paragraph, as a string.  INFO is a
plist holding contextual information."
  (format "(paragraph %s)" contents))

;;;; Plain List

(defun org-pptx-plain-list (plain-list contents info)
  "Transcode a PLAIN-LIST element from Org to PPTX.
CONTENTS is the contents of the list.  INFO is a plist holding
contextual information."
  (let ((type (org-element-property :type plain-list))
        (depth (org-pptx--list-depth plain-list info)))
    (format "(plain-list %s %s\n%s)" type depth (org-trim contents))))

;;;; Plain Text

(defun org-pptx-plain-text (text info)
  "Transcode a TEXT string from Org to PPTX.
TEXT is the string to transcode.  INFO is a plist holding
contextual information."
  (let ((output text))
    ;; activate smart quotes
    (when (plist-get info :with-smart-quotes)
      (setq output (org-export-activate-smart-quotes output :utf-8 info text)))
    ;; strip trailing newline
    (unless (org-export-get-next-element text info)
      (setq output (replace-regexp-in-string "\n\\'" "" output)))
    ;; convert other newlines to space
    (setq output (replace-regexp-in-string "\n" " " output))
    ;; escape quotation marks
    (setq output (replace-regexp-in-string "\"" "\\\\\"" output))
    ;; elide empty strings
    (unless (string= output "")
      (let ((print-level nil))
        (format "(plain-text \"%s\")" output)))))

;;;; Planning

(defun org-pptx-planning (planning _contents _info)
  "Transcode a PLANNING element from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (format
   "(planning %s)"
   (org-trim
    (mapconcat
     (lambda (pair)
       (let ((timestamp (cdr pair)))
         (when timestamp
           (let ((string (car pair)))
             (format "%s \"%s\"" string
                     (org-pptx--escape-text
                      (org-timestamp-translate timestamp)))))))
     `((,org-closed-string . ,(org-element-property :closed planning))
       (,org-deadline-string . ,(org-element-property :deadline planning))
       (,org-scheduled-string . ,(org-element-property :scheduled planning)))
     " "))))

;;;; Property Drawer

(defun org-pptx-property-drawer (_property-drawer contents _info)
  "Transcode a PROPERTY-DRAWER element from Org to PPTX.
CONTENTS holds the contents of the drawer.  INFO is a plist holding
contextual information."
  (and (org-string-nw-p contents)
       (format "(property-drawer \"%s\")"
               (org-pptx--escape-text contents))))

;;;; Quote Block

(defun org-pptx-quote-block (_quote-block contents _info)
  "Transcode a QUOTE-BLOCK element from Org to PPTX.
CONTENTS holds the contents of the block.  INFO is a plist holding
contextual information."
  (format "(quote-block \"%s\")" (org-pptx--escape-text contents)))

;;;; Radio Target

(defun org-pptx-radio-target (radio-target text info)
  "Transcode a RADIO-TARGET element from Org to PPTX.
TEXT is the text of the target.  INFO is a plist holding contextual
information."
  (format "(radio-target \"%s\" \"%s\")"
          (org-export-get-reference radio-target info)
          text))

;;;; Section

(defun org-pptx-section (_section contents _info)
  "Transcode a SECTION element from Org to PPTX.
CONTENTS holds the contents of the section.  INFO is a plist holding
contextual information."
  contents)

;;;; Special Block

(defun org-pptx-special-block (special-block contents _info)
  "Transcode a SPECIAL-BLOCK element from Org to PPTX.
CONTENTS holds the contents of the block.  INFO is a plist holding
contextual information."
  (format "(special-block \"%s\" \"%s\")"
          (org-element-property :type special-block)
          contents))

;;;; Src Block

(defun org-pptx-src-block (src-block _contents info)
  "Transcode a SRC-BLOCK element from Org to PPTX.
CONTENTS holds the contents of the block.  INFO is a plist holding
contextual information."
  (let ((output (org-export-format-code-default src-block info)))
    (setq output (org-pptx--escape-text output))
    (format "(src-block \"%s\")" output)))

;;;; Statistics Cookie

(defun org-pptx-statistics-cookie (statistics-cookie _contents _info)
  "Transcode a STATISTICS-COOKIE object from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual information."
  (let ((blank (org-element-property :post-blank statistics-cookie))
        (cookie-value (org-element-property :value statistics-cookie)))
    (format "(code %s (plain-text \"%s\"))" blank cookie-value)))

;;;; Strike-Through

(defun org-pptx-strike-through (_strike-through contents _info)
  "Transcode STRIKE-THROUGH from Org to PPTX.
CONTENTS is the text with strike-through markup.  INFO is a plist
holding contextual information."
  (format "(strike-through (plain-text \"%s\"))" contents))

;;;; Subscript

(defun org-pptx-subscript (_subscript contents _info)
  "Transcode a SUBSCRIPT object from Org to PPTX.
CONTENTS is the contents of the object.  INFO is a plist holding
contextual information."
  (format "(subscript (plain-text \"%s\"))" contents))

;;;; Superscript

(defun org-pptx-superscript (_superscript contents _info)
  "Transcode a SUPERSCRIPT object from Org to PPTX.
CONTENTS is the contents of the object.  INFO is a plist holding
contextual information."
  (format "(superscript (plain-text \"%s\"))" contents))

;;;; Table
(defun org-pptx-table (table contents info)
  "Transcode a TABLE element from Org to PPTX.
CONTENTS is the contents of the table.  INFO is a plist holding
contextual information."
  (if (eq (org-element-property :type table) 'table.el)
      ;; "table.el" table.
      "(plain-text \"table.el table not supported\")"
    ;; standard table
    (let ((caption (org-export-get-caption table)))
      (format "(table (%s) %s)"
              (if caption
                  (org-export-data caption info)
                "(plain-text \"\")")
              contents))))

;;;; Table Cell

(defun org-pptx-table-cell (_table-cell contents _info)
  "Transcode a TABLE-CELL element from Org to PPTX.
CONTENTS is the contents of the cell.  INFO is a plist used as a
communication channel."
  (format "(cell %s)" contents))

;;;; Table Row

(defun org-pptx-table-row (table-row contents _info)
  "Transcode a TABLE-ROW element from Org to PPTX.
CONTENTS is the contents of the row.  INFO is a plist used as a
communication channel."
  (if (eq (org-element-property :type table-row) 'standard)
      (format "(row %s)" contents)
    ""))

;;;; Target

(defun org-pptx-target (_target _contents _info)
  "Transcode a TARGET object from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual
information."
  "")

;;;; Template
(defun org-pptx-template (contents info)
  "Return complete document string after PPTX conversion.
CONTENTS is the transcoded contents string.  INFO is a plist holding
export options."
  (let* ((fn-sec (org-pptx--footnote-section info))
         (contents
          (if fn-sec (format "%s\n%s" contents fn-sec) contents)))
    (let ((template (org-file-contents org-pptx-template-kawa-file))
          (spec `((?T . ,org-pptx-prototype-pptx-file)
                  (?t . ,(org-export-data (plist-get info :title) info))
                  (?d . ,(let ((d (plist-get info :date)))
                           (if d (format " %s000" (org-export-data d info))
                             "")))
                  (?a . ,(org-export-data (plist-get info :author) info))
                  (?c . ,contents))))
      (format-spec template spec))))

;;;; Timestamp

(defun org-pptx-timestamp (timestamp _contents _info)
  "Transcode a TIMESTAMP object from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual
information."
  (org-timestamp-format timestamp "(plain-text \"%Y-%m-%d\")"))

;;;; Underline

(defun org-pptx-underline (_underline contents _info)
  "Transcode UNDERLINE from Org to PPTX.
CONTENTS is the text with underline markup.  INFO is a plist
holding contextual information."
  (format "(underline (plain-text \"%s\"))" contents))

;;;; Verbatim

(defun org-pptx-verbatim (verbatim _contents _info)
  "Transcode VERBATIM from Org to PPTX.
CONTENTS is nil.  INFO is a plist holding contextual
information."
  (let ((blank (org-element-property :post-blank verbatim))
        (output (org-element-property :value verbatim)))
    (setq output (org-pptx--escape-text output))
    (format "(verbatim %s (plain-text \"%s\"))" blank output)))

;;;; Verse Block

(defun org-pptx-verse-block (_verse-block contents _info)
  "Transcode a VERSE-BLOCK element from Org to PPTX.
CONTENTS is verse block contents.  INFO is a plist holding
contextual information."
  (format "(verse-block \"%s\")" (org-pptx--escape-text contents)))


;;; End-user functions

;;;###autoload
(defun org-pptx-export-to-scm
    (&optional async subtreep visible-only body-only ext-plist)
  "Export current buffer to Kawa."
  (interactive)
  (let ((outfile (org-export-output-file-name ".scm" subtreep)))
    (org-export-to-file 'pptx outfile
      async subtreep visible-only body-only ext-plist)))

;;;###autoload
(defun org-pptx-export-to-pptx
    (&optional async subtreep visible-only body-only ext-plist)
  "Export current buffer to Kawa then process through to PPTX."
  (interactive)
  (let ((outfile (org-export-output-file-name ".scm" subtreep)))
    (org-export-to-file 'pptx outfile
      async subtreep visible-only body-only ext-plist
      (lambda (file) (org-pptx-compile file)))))

(defun org-pptx-compile (scmfile)
  "Run the given Kawa Scheme file to generate a PPTX file."
  (message "Processing Scheme file %s..." (expand-file-name scmfile))
  (let* ((pptxfile (replace-regexp-in-string "\\.scm\\'" ".pptx" scmfile))
         (logbuf (format "*ox-pptx: %s" scmfile))
         (process
          `(,(concat "java -cp " (mapconcat 'identity org-pptx-classpath ":")
                     " -Dpptxfile=\"" pptxfile "\""
                     " kawa.repl " scmfile)))
         (msg (message "%s" (car process)))
         (outfile (org-compile-file scmfile process "pptx" nil logbuf)))
    outfile))

(provide 'ox-pptx)
