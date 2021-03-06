;; This file was auto-generated by ox-pptx.

(import (only (srfi 1) list-tabulate))

(define pptx ::org.apache.poi.xslf.usermodel.XMLSlideShow
  (org.apache.poi.xslf.usermodel.XMLSlideShow
   (java.io.FileInputStream "%T")))

(define-syntax plain-text (syntax-rules () ((_ e) e)))
(define (paragraph . rest) ::void #!void)

;;; Define constants.

(define-constant *title* ::String %t)
(define-constant *subtitle* ::String %s)
(define-constant *date*  ::java.util.Date (java.util.Date%d))
(define-constant *now*   ::java.util.Date (java.util.Date))
(define-constant *disp-date* ::String
  ((java.text.SimpleDateFormat "dd MMMM yyyy"):format *date*))

(define-constant *author* ::String %a)

(define-constant *table-style*
  (((->org.apache.poi.xslf.usermodel.XSLFTable
     ((->org.apache.poi.xslf.usermodel.XSLFSheet
       (pptx:slides 2)):shapes 2)):getCTTable):getTblPr))

(define-constant *text-shape-bottom-inset*
  (->org.apache.poi.xslf.usermodel.XSLFTextShape
   ((->org.apache.poi.xslf.usermodel.XSLFSheet
     (pptx:slides 2)):shapes 1)):bottom-inset)

(define-constant *text-shape-left-inset*
  (->org.apache.poi.xslf.usermodel.XSLFTextShape
   ((->org.apache.poi.xslf.usermodel.XSLFSheet
     (pptx:slides 2)):shapes 1)):left-inset)

(define-constant *text-shape-right-inset*
  (->org.apache.poi.xslf.usermodel.XSLFTextShape
   ((->org.apache.poi.xslf.usermodel.XSLFSheet
     (pptx:slides 2)):shapes 1)):right-inset)

(define-constant *text-shape-top-inset*
  (->org.apache.poi.xslf.usermodel.XSLFTextShape
   ((->org.apache.poi.xslf.usermodel.XSLFSheet
     (pptx:slides 2)):shapes 1)):top-inset)

(define-constant *text-paragraph-indents*
  (map *:.indent
       (->org.apache.poi.xslf.usermodel.XSLFTextShape
        ((->org.apache.poi.xslf.usermodel.XSLFSheet
          (pptx:slides 2)):shapes 1)):text-paragraphs))

(define-constant *text-paragraph-left-margins*
  (map *:.left-margin
       (->org.apache.poi.xslf.usermodel.XSLFTextShape
        ((->org.apache.poi.xslf.usermodel.XSLFSheet
          (pptx:slides 2)):shapes 1)):text-paragraphs))

(define-constant *text-paragraph-bullet-characters*
  (map *:.bullet-character
       (->org.apache.poi.xslf.usermodel.XSLFTextShape
        ((->org.apache.poi.xslf.usermodel.XSLFSheet
          (pptx:slides 2)):shapes 1)):text-paragraphs))

(define-constant *text-color*
  (->org.apache.poi.xslf.usermodel.XSLFTextRun
   ((->org.apache.poi.xslf.usermodel.XSLFTextParagraph
     ((->org.apache.poi.xslf.usermodel.XSLFTextShape
       ((->org.apache.poi.xslf.usermodel.XSLFSheet
         (pptx:slides 2)):shapes 1)):text-paragraphs 0)):text-runs
         0)):font-color)

(define-constant *page-size*
  (((->org.apache.poi.xslf.usermodel.XSLFSlideMaster
     (pptx:slide-masters 0)):get-layout
     org.apache.poi.xslf.usermodel.SlideLayout:TITLE_AND_CONTENT):get-placeholder 1):anchor)

;;; update-contents performs a search and replace for the title, name,
;;; and date strings within an array of shapes.

(define (update-contents shapes ::java.util.List[org.apache.poi.xslf.usermodel.XSLFShape]
                         #!optional (withsubtitle ::boolean #f))
  (for-each
   (lambda (shape ::org.apache.poi.xslf.usermodel.XSLFShape)
     (cond
      ((? s::org.apache.poi.xslf.usermodel.XSLFTextShape shape)
       (for-each
        (lambda (paragraph ::org.apache.poi.xslf.usermodel.XSLFTextParagraph)
          (for-each
           (lambda (run ::org.apache.poi.xslf.usermodel.XSLFTextRun)
             (when (run:raw-text:contains "Presentation Title")
               (set! run:text
                     (run:raw-text:replace-all "Presentation Title"
                                               (if withsubtitle
                                                   (string-append *title*
                                                                  "\n\n"
                                                                  *subtitle*)
                                                   *title*))))
             (when (run:raw-text:contains "13 June 2016")
               (set! run:text
                     (run:raw-text:replace-all "13 June 2016" *disp-date*)))
             (when (run:raw-text:contains "Presenter Name, Presenter Title")
               (set! run:text
                     (run:raw-text:replace-all "Presenter Name, Presenter Title"
                                               *author*)))
             (when (run:raw-text:contains "Date of presentation")
               (set! run:text
                     (run:raw-text:replace-all "Date of presentation"
                                               *disp-date*))))
           paragraph:text-runs))
        s:text-paragraphs))))
   shapes))

;;; Update Master

(update-contents (->org.apache.poi.xslf.usermodel.XSLFSlideMaster
                  (pptx:slide-masters 0)):shapes)

;;; Update Layouts

(for-each
 (lambda (layout ::org.apache.poi.xslf.usermodel.XSLFSlideLayout)
   (update-contents layout:shapes))
 (->org.apache.poi.xslf.usermodel.XSLFSlideMaster
  (pptx:slide-masters 0)):slide-layouts)

;;; Update Custom Title Page

(update-contents (->org.apache.poi.xslf.usermodel.XSLFSheet
                  (pptx:slides 0)):shapes
                  (not (eq? #!null *subtitle*)))

;;; Remove other pages

(let loop ((n ::int (- (pptx:slides:size) 1)))
     (when (> n 0)
       (pptx:removeSlide n)
       (loop (- n 1))))

;;; Remove Image 7 and Image 8, which aren't needed (only referenced
;;; in the (just-removed) last slide.

(let ((pkg pptx:package-part:package))
  (for-each
   (lambda (part ::org.apache.poi.openxml4j.opc.PackagePart)
     (pkg:delete-part-recursive part:part-name))
   (pkg:get-parts-by-name
    (java.util.regex.Pattern:compile "/ppt/media/image[78]\.png"))))

;;; Update core properties.

(let ((props ::org.apache.poi.ooxml.POIXMLProperties$CoreProperties
             pptx:properties:core-properties))
  (set! props:created (java.util.Optional:of *now*))
  (set! props:creator *author*)
  (set! props:modified (->String #!null))
  (set! props:revision "0")
  (set! props:title *title*)
  (if *subtitle* (set! props:subject-property *subtitle*))
  (let ((under ::org.apache.poi.openxml4j.opc.internal.PackagePropertiesPart
               props:underlying-properties))
    (set! under:last-modified-by-property *author*)))

;;; Save file, then re-open it -- this flushes some obsolete
;;; relationships for the slides we just deleted.

(let* ((pptxfile ::String (or (java.lang.System:get-property "pptxfile")
                              "file.pptx"))
       (out ::java.io.FileOutputStream (java.io.FileOutputStream
                                        pptxfile)))
  (pptx:write out)
  (out:close)
  (set! pptx (org.apache.poi.xslf.usermodel.XMLSlideShow
              (java.io.FileInputStream pptxfile))))

;;; Functions to create slides.

(define-constant SECTION_HEADER
  org.apache.poi.xslf.usermodel.SlideLayout:SECTION_HEADER)

(define-constant TITLE_AND_CONTENT
  org.apache.poi.xslf.usermodel.SlideLayout:TITLE_AND_CONTENT)

(define-constant TITLE_ONLY
  org.apache.poi.xslf.usermodel.SlideLayout:TITLE_ONLY)

(define (set-slide-title! (slide ::org.apache.poi.xslf.usermodel.XSLFSlide)
                          title)
  (let ((s (slide:get-placeholder 0)))
    (s:clear-text)
    (let ((p (s:add-new-text-paragraph)))
      (for-each (lambda (t) (add-text p t #t)) title))))

(define (create-section-header-slide title #!optional subtitle)
  (let* ((layout ((->org.apache.poi.xslf.usermodel.XSLFSlideMaster
                   (pptx:slide-masters 0)):get-layout SECTION_HEADER))
         (slide (pptx:create-slide layout)))
    (set-slide-title! slide title)
    (if subtitle
        (set! (slide:get-placeholder 1):text (->String subtitle))
      (slide:remove-shape (slide:get-placeholder 1)))))

(define (add-text (p ::org.apache.poi.xslf.usermodel.XSLFTextParagraph)
                  text #!optional bold italic)
  ::org.apache.poi.xslf.usermodel.XSLFTextRun
  (if (eq? text 'nil)
      (add-text p '(plain-text "") bold italic)
      (let ((tag (car text)))
        (case tag
          ((plain-text)
           (let* ((t (p:add-new-text-run)))
             (set! t:bold bold)
             (set! t:italic italic)
             (set! t:text (cadr text))
             t))
          ((entity)
           (let ((t (add-text p (caddr text) bold italic)))
             (when (> (cadr text) 0)
               (add-text p '(plain-text " ") bold italic))
             t))
          ((code verbatim)
           (let ((t (add-text p (caddr text) bold italic)))
             (set! t:font-family "Courier")
             (when (> (cadr text) 0)
               (add-text p '(plain-text " ") bold italic))
             t))
          ((bold)
           (let ((t ::org.apache.poi.xslf.usermodel.XSLFTextRun #!null))
             (do ((texts (cddr text) (cdr texts))) ((null? texts))
               (set! t (add-text p (car texts) #t italic)))
             (when (> (cadr text) 0)
               (add-text p '(plain-text " ") bold italic))
             t))
          ((italic)
           (let ((t ::org.apache.poi.xslf.usermodel.XSLFTextRun #!null))
             (do ((texts (cddr text) (cdr texts))) ((null? texts))
               (set! t (add-text p (car texts) bold #t)))
             (when (> (cadr text) 0)
               (add-text p '(plain-text " ") bold italic))
             t))
          ((link)
           (let* ((t (add-text p (cadddr text) bold italic))
                  (hl (t:create-hyperlink)))
             (hl:set-address
              (format #f "~a:~a" (cadr text) (caddr text)))
             t))
          ((subscript)
           (let ((t ::org.apache.poi.xslf.usermodel.XSLFTextRun #!null))
             (do ((texts (cddr text) (cdr texts))) ((null? texts))
               (set! t (add-text p (car texts) bold italic))
               (set! t:subscript #t))
             (when (> (cadr text) 0)
               (add-text p '(plain-text " ") bold italic))
             t))
          ((superscript)
           (let ((t ::org.apache.poi.xslf.usermodel.XSLFTextRun #!null))
             (do ((texts (cddr text) (cdr texts))) ((null? texts))
               (set! t (add-text p (car texts) bold italic))
               (set! t:superscript #t))
             (when (> (cadr text) 0)
               (add-text p '(plain-text " ") bold italic))
             t))
          ((footnote-reference)
           (add-text p (cadr text) bold italic))
          ((latex-environment latex-fragment)
           (add-text p (caddr text) bold italic))
          ((line-break)
           (p:add-line-break))))))

(define (get-string text) ::String
  (let ((tag (car text)))
    (case tag
      ((plain-text) (cadr text))
      ((entity) (get-string (caddr text)))
      ((code verbatim) (get-string (caddr text)))
      ((bold) (get-string (caddr text)))
      ((italic) (get-string (caddr text)))
      ((link) (get-string (cadddr text)))
      ((footnote-reference) (get-string (cadr text)))
      (else ""))))

(define (insert-list (s ::org.apache.poi.xslf.usermodel.XSLFTextShape)
                     lst #!optional (set-props? ::boolean #f))
  (do ((lst lst (cdr lst))) ((null? lst))
    (let ((tag (caar lst)))
      (case tag
        ((plain-list)
         (let ((ord (cadar lst))
               (depth (caddar lst))
               (data (cdddar lst)))
           (do ((data data (cdr data))) ((null? data))
             (let ((checkbox (caddar data))
                   (paragraph (car (cdddar data)))
                   (kids (cdr (cdddar data))))
               (let ((p (s:add-new-text-paragraph)))
                 (set! p:indent-level depth)
                 (if (eq? ord 'ordered)
                     (let ((num (cadar data)))
                       (p:set-bullet-auto-number
                        org.apache.poi.sl.usermodel.AutoNumberingScheme:arabicPeriod
                        (->int num)))
                     (set! p:bullet #t))
                 (when (and set-props?
                            (< depth (length *text-paragraph-indents*)))
                   (set! p:indent
                         (*text-paragraph-indents* depth))
                   (set! p:left-margin
                         (*text-paragraph-left-margins* depth))
                   (unless (eq? ord 'ordered)
                     (set! p:bullet-character
                           (*text-paragraph-bullet-characters* depth))))
                 (unless (eq? checkbox 'nil)
                   (let* ((x (cdr (assq checkbox
                                       '((on . #\x2611)
                                         (off . #\x2610)
                                         (trans . #\x2610)))))
                          (txt `(plain-text ,(format "~a " x)))
                          (t (add-text p txt)))
                     (when set-props?
                       (set! t:font-color *text-color*))))
                 (do ((texts (cdr paragraph) (cdr texts))) ((null? texts))
                   (let ((t (add-text p (car texts))))
                     (when set-props?
                       (set! t:font-color *text-color*))))
                 (when (not (null? kids))
                   (insert-list s kids set-props?)))))))
        (else
         (insert-list
          s
          `((plain-list
             unordered
             (item
              -
              nil
              (paragraph
               (plain-text
                ,(format #f "Cannot insert ~a here." tag))))))))))))

(define (create-list-slide title lst)
  (let* ((layout ((->org.apache.poi.xslf.usermodel.XSLFSlideMaster
                   (pptx:slide-masters 0)):get-layout TITLE_AND_CONTENT))
         (slide (pptx:create-slide layout)))
    (set-slide-title! slide title)
    (let ((s ::org.apache.poi.xslf.usermodel.XSLFTextShape
             (slide:get-placeholder 1)))
      (s:clear-text)
      (do ((lst lst (cdr lst))) ((null? lst))
        (insert-list s lst))
      (set! s:text-autofit
            org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL))))

(define (ext->imgtype ext::String)
  (cond ((string=? ext "png")
         org.apache.poi.sl.usermodel.PictureData$PictureType:PNG)
        ((string=? ext "svg")
         org.apache.poi.sl.usermodel.PictureData$PictureType:SVG)
        ((or (string=? ext "jpg") (string=? ext "jpeg"))
         org.apache.poi.sl.usermodel.PictureData$PictureType:JPEG)
        ((string=? ext "gif")
         org.apache.poi.sl.usermodel.PictureData$PictureType:GIF)
        ((string=? ext "tiff")
         org.apache.poi.sl.usermodel.PictureData$PictureType:TIFF)
        ((string=? ext "eps")
         org.apache.poi.sl.usermodel.PictureData$PictureType:EPS)
        ((string=? ext "bmp")
         org.apache.poi.sl.usermodel.PictureData$PictureType:BMP)
        ((string=? ext "emf")
         org.apache.poi.sl.usermodel.PictureData$PictureType:EMF)
        ((string=? ext "wmf")
         org.apache.poi.sl.usermodel.PictureData$PictureType:WMF)
        (else 0)))

(define (insert-image (slide ::org.apache.poi.xslf.usermodel.XSLFSlide)
                      (filename ::String))
  (let* ((data ::bytevector (path-bytes filename))
         (bytes ::byte[] data:buffer)
         (ext (path-extension filename))
         (type (ext->imgtype ext))
         (index (pptx:add-picture bytes type))
         (p (slide:create-picture index))
         (image-size index:image-dimension)
         (anchor p:anchor))
    (when (string=? ext "svg")
      ;; every svg shows 200x200 as its size, because it's dumb.  So
      ;; pull out the actual size from the XML.  Yes, DOM is
      ;; inefficient.
      (let* ((xmlfile (java.io.File filename))
             (factory (javax.xml.parsers.DocumentBuilderFactory:newInstance))
             (builder (factory:newDocumentBuilder))
             (doc (builder:parse xmlfile))
             (elem doc:document-element))
        (elem:normalize)
        (let ((widthstr  ::String (elem:get-attribute "width"))
              (heightstr ::String (elem:get-attribute "height")))
          (when (and (not (string=? widthstr  ""))
                     (not (string=? heightstr "")))
            (set! widthstr  (widthstr:replace-all  "px" ""))
            (set! heightstr (heightstr:replace-all "px" ""))
            (let ((w (java.lang.Integer:parseInt widthstr))
                  (h (java.lang.Integer:parseInt heightstr)))
              (set! image-size (java.awt.Dimension w h)))))))
    (when (> image-size:width *page-size*:width)
      (image-size:set-size
       *page-size*:width
       (/ (* image-size:height *page-size*:width) image-size:width)))
    (when (> image-size:height *page-size*:height)
      (image-size:set-size
       (/ (* image-size:width *page-size*:height) image-size:height)
       *page-size*:height))
    (anchor:set-rect
     (+ *page-size*:x (/ (- *page-size*:width image-size:width) 2))
     (+ *page-size*:y (/ (- *page-size*:height image-size:height) 2))
     image-size:width image-size:height)
    (set! p:anchor anchor)))

(define (insert-paragraph (slide ::org.apache.poi.xslf.usermodel.XSLFSlide)
                          lst s)
  (unless (null? lst)
    (if (equal? 'img (caar lst))
        (insert-image slide (cadar lst))
        (let* ((s ::org.apache.poi.xslf.usermodel.XSLFTextShape
                  (or s
                      (let ((s (slide:create-text-box)))
                        (s:clear-text)
                        (set! s:text-autofit
                              org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL)
                        s)))
               (p (s:add-new-text-paragraph)))
          (set! s:anchor *page-size*)
          (do ((texts lst (cdr texts))) ((null? texts))
            (let ((t (add-text p (car texts))))
              (set! t:font-color *text-color*)))))))

(define (insert-center (slide ::org.apache.poi.xslf.usermodel.XSLFSlide)
                       content)
  (let* ((s ::org.apache.poi.xslf.usermodel.XSLFTextShape
            (let ((s (slide:create-text-box)))
              (s:clear-text)
              (set! s:text-autofit
                    org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL)
              s))
         (p (s:add-new-text-paragraph)))
    (set! s:anchor *page-size*)
    (set! p:text-align
          org.apache.poi.sl.usermodel.TextParagraph$TextAlign:CENTER)
    (do ((texts (cdar content) (cdr texts))) ((null? texts))
      (let ((t (add-text p (car texts))))
        (set! t:font-color *text-color*)))))

(define (optimize-table-columns (tbl ::org.apache.poi.xslf.usermodel.XSLFTable))
  (let ((nr tbl:number-of-rows) (nc tbl:number-of-columns))
    (do ((c ::int 0 (+ c 1))) ((= c (- nc 1)))
      (let ((w-orig (tbl:get-column-width c))
            (w-orig+1 (tbl:get-column-width (+ c 1)))
            (h-orig tbl:anchor:height)
            (th-orig
             (apply min (list-tabulate
                         nr
                         (lambda (r) (tbl:get-cell r c):text-height)))))
        (let ((w w-orig) (w+1 w-orig+1) (th-best th-orig) (h-best h-orig))
          (do ((try 8 (- try 1)))
              ((= try 0)
               (tbl:set-column-width c w)
               (tbl:set-column-width (+ c 1) w+1)
               (tbl:update-cell-anchor))
            (let* ((this-w   (* try 0.25 w-orig))
                   (this-w+1 (+ w-orig+1 w-orig (- this-w))))
              (tbl:set-column-width c this-w)
              (tbl:set-column-width (+ c 1) this-w+1)
              (tbl:update-cell-anchor)
              (let ((th (apply min (list-tabulate
                                    nr
                                    (lambda (r) (tbl:get-cell r c):text-height)))))
                (when (and (<= th th-best)
                           (<= tbl:anchor:height h-best))
                  (set! h-best tbl:anchor:height)
                  (set! th-best th)
                  (set! w this-w)
                  (set! w+1 this-w+1))))))))))

(define (insert-table (slide ::org.apache.poi.xslf.usermodel.XSLFSlide)
                      caption
                      rows)
  (let ((s (slide:create-table)))
    (set! s:anchor *page-size*)
    ((s:getCTTable):setTblPr *table-style*)
    (do ((rows rows (cdr rows)) (i 0 (+ i 1))) ((null? rows))
      (unless (equal? (car rows) '(row nil))
        (let ((r (s:add-row)))
          (set! r:height 25)
          (do ((cells (cdar rows) (cdr cells))) ((null? cells))
            (let* ((c (r:add-cell))
                   (p (c:add-new-text-paragraph)))
              (set! c:text-autofit
                    org.apache.poi.sl.usermodel.TextShape$TextAutofit:SHAPE)
              (do ((texts (cdar cells) (cdr texts))) ((null? texts))
                (add-text p (car texts) (= i 0)))
              (c:resize-to-fit-text)
              (c:set-border-color
               org.apache.poi.sl.usermodel.TableCell$BorderEdge:top
               java.awt.Color:white)
              (c:set-border-color
               org.apache.poi.sl.usermodel.TableCell$BorderEdge:bottom
               java.awt.Color:white)
              (c:set-border-color
               org.apache.poi.sl.usermodel.TableCell$BorderEdge:left
               java.awt.Color:white)
              (c:set-border-color
               org.apache.poi.sl.usermodel.TableCell$BorderEdge:right
               java.awt.Color:white))))))
    (let ((nc ::int s:number-of-columns)
          (nr ::int s:number-of-rows)
          (a s:anchor))
      (let loop ((w ::double 0) (i ::int 0))
        (if (< i nc)
            ;; calculate the full width
            (loop (+ w (s:get-column-width i)) (+ i 1))
            (let loop2 ((h ::double 0) (i ::int 0))
              (if (< i nr)
                  ;; calculate the full height
                  (loop2 (+ h (s:get-row-height i)) (+ i 1))
                  (begin
                    ;; scale all columns to full page width
                    (do ((i ::int 0 (+ i 1))) ((= i nc))
                      (s:set-column-width
                       i (* (s:get-column-width i)
                            (/ *page-size*:width w))))
                    (set! w *page-size*:width)
                    (a:set-rect
                     (+ *page-size*:x (/ (- *page-size*:width w) 2)) *page-size*:y w h)
                    (set! s:anchor a)
                    (s:update-cell-anchor)
                    ;; adjust relative column widths to reduce height
                    ;; needed given a fixed font size
                    (optimize-table-columns s)
                    (when (not (equal? '((plain-text "")) caption))
                      (let* ((cap
                              (let ((s (slide:create-text-box)))
                                (s:clear-text)
                                (set! s:text-autofit
                                      org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL)
                                s))
                             (p (cap:add-new-text-paragraph))
                             (a2 cap:anchor))
                        (for-each
                         (lambda (t)
                           (let ((r (add-text p t)))
                             (set! r:font-size (* r:font-size 0.8d0))))
                         caption)
                        (set! p:text-align
                              org.apache.poi.sl.usermodel.TextParagraph$TextAlign:CENTER)
                        (a2:set-rect
                         (+ *page-size*:x (/ (- *page-size*:width w) 2)) h w 12)
                        (set! cap:anchor a2)))))))))))

(define (insert-block (slide ::org.apache.poi.xslf.usermodel.XSLFSlide)
                      block)
  (let* ((s
          (let ((s (slide:create-text-box)))
            (s:clear-text)
            (set! s:text-autofit
                  org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL)
            s))
         (p (s:add-new-text-paragraph))
         (t (p:add-new-text-run)))
    (set! s:anchor *page-size*)
    (set! t:font-family "Courier")
    (set! t:text (cadr block))))

(define (insert-data (slide ::org.apache.poi.xslf.usermodel.XSLFSlide)
                     contents)
  (let ((last-text-shape ::org.apache.poi.xslf.usermodel.XSLFTextShape
                         #!null))
    (do ((contents contents (cdr contents))) ((null? contents))
      (let* ((datum (car contents))
             (tag (car datum)))
        (case tag
          ((paragraph)
           (insert-paragraph slide (cdr datum) last-text-shape))
          ((plain-list)
           (let* ((s (or last-text-shape
                         (let ((s (slide:create-text-box)))
                           (s:clear-text)
                           (set! s:text-autofit
                                 org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL)
                           s))))
             (set! s:anchor *page-size*)
             (set! s:bottom-inset *text-shape-bottom-inset*)
             (set! s:left-inset *text-shape-left-inset*)
             (set! s:right-inset *text-shape-right-inset*)
             (set! s:top-inset *text-shape-top-inset*)
             (insert-list s (list datum) #t)
             (set! s:text-autofit
                   org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL)))
          ((table)
           (insert-table slide (cadr datum) (cddr datum)))
          ((example-block src-block fixed-width)
           (insert-block slide datum))
          ((center)
           (insert-center slide (cdr datum)))
          ((latex-environment latex-fragment)
           (insert-image slide (cadr datum))))
        (let* ((ss slide:shapes)
               (ln (ss:size)))
          (if (and (> ln 0)
                   (not (memq tag '(center table src-block example-block fixed-width)))
                   (org.apache.poi.xslf.usermodel.XSLFTextShape?
                    (ss (- ln 1))))
              (set! last-text-shape (ss (- ln 1)))
            (set! last-text-shape #!null)))))))

(define (create-title-only-slide title content)
  (let* ((layout ((->org.apache.poi.xslf.usermodel.XSLFSlideMaster
                   (pptx:slide-masters 0)):get-layout TITLE_ONLY))
         (slide (pptx:create-slide layout)))
    (set-slide-title! slide title)
    (insert-data slide content)
    (do ((i ::int 1 (+ i 1))) ((= i (slide:shapes:size)))
      (when (? s::org.apache.poi.xslf.usermodel.XSLFTextShape
               (slide:shapes i))
        (s:resize-to-fit-text)))
    (let repeat ((repeat-count 0))
      (let loop ((i ::int 1) (n ::int (slide:shapes:size))
                 (h ::double 0))
        (if (= i n)
            (let ((d (/ (- *page-size*:height h -20) 2)))
              (if (< d -0.5)
                  (let ((scale (/ *page-size*:height (- h 20)))
                        (y *page-size*:y))
                    (do ((i 1 (+ i 1))) ((= i n))
                      (let* ((s (->org.apache.poi.xslf.usermodel.XSLFShape
                                 (slide:shapes i)))
                             (a s:anchor))
                        (cond ((? ts::org.apache.poi.xslf.usermodel.XSLFTextShape s)
                               (set! ts:text-autofit
                                     org.apache.poi.sl.usermodel.TextShape$TextAutofit:NORMAL)
                               (a:set-rect a:x y a:width (* scale a:height))
                               (set! ts:anchor a)
                               (set! y (+ y ts:anchor:height)))
                              ((? ps::org.apache.poi.xslf.usermodel.XSLFPictureShape s)
                               (let ((scale (- 1 (* 0.25 (- 1 scale)))))
                                 (if (> (* scale a:height)
                                        (/ *page-size*:height 2))
                                     (a:set-rect
                                      (+ a:x (/ (* (- 1 scale) a:width) 2))
                                      y
                                      (* scale a:width)
                                      (* scale a:height))
                                     (a:set-rect a:x y a:width a:height)))
                               (set! ps:anchor a)
                               (set! y (+ y a:height)))
                              ((? tbl::org.apache.poi.xslf.usermodel.XSLFTable s)
                               (do ((r 0 (+ r 1))) ((= r tbl:number-of-rows))
                                 (do ((c 0 (+ c 1))) ((= c tbl:number-of-columns))
                                   (let ((cell (tbl:get-cell r c)))
                                     (for-each
                                      (lambda (paragraph ::org.apache.poi.xslf.usermodel.XSLFTextParagraph)
                                        (for-each
                                         (lambda (run ::org.apache.poi.xslf.usermodel.XSLFTextRun)
                                           (run:set-font-size
                                            (java.lang.Math:max 1.0d0 (* run:font-size scale))))
                                         paragraph:text-runs))
                                      cell:text-paragraphs))))
                               (tbl:update-cell-anchor)
                               (optimize-table-columns tbl)
                               (set! y (+ y tbl:anchor:height)))
                              (else
                               (a:set-rect a:x y a:width (* scale a:height))
                               (set! s:anchor a)
                               (set! y (+ y a:height))))))
                    (when (< repeat-count 10)
                      (repeat (+ repeat-count 1))))
                  (do ((i 1 (+ i 1))) ((= i n))
                    (let* ((s (->org.apache.poi.xslf.usermodel.XSLFShape
                               (slide:shapes i)))
                           (a s:anchor))
                      (a:set-rect a:x (+ *page-size*:y a:y d) a:width a:height)
                      (set! s:anchor a)))))
            (let ((a (slide:shapes i):anchor))
              (a:set-rect a:x h a:width a:height)
              (set! (slide:shapes i):anchor a)
              (loop (+ i 1) n (+ h 20 a:height))))))))


(define (create-slide title content)
  ((if (and (= 1 (length content))
            (eq? 'plain-list (caar content)))
       create-list-slide
       create-title-only-slide)
   title content))


;;; File contents start here.
%c
;;; File contents end here.

(let* ((pptxfile ::String (or (java.lang.System:get-property "pptxfile")
                              "file.pptx"))
       (out ::java.io.FileOutputStream (java.io.FileOutputStream pptxfile)))
  (format #t "creating ~a...~%%" pptxfile)
  (pptx:write out)
  (out:close)
  (format #t "Done.~%%"))
