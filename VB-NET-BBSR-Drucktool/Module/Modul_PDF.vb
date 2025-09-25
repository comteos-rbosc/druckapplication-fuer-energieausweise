'----------------------------------------------------------+
' DynaPDF 4.0                                              |
'----------------------------------------------------------+
' Copyright (C) 2003-2018                                  |
' Jens Boschulte, DynaForms GmbH.                          |
' All rights reserved.                                     |
'----------------------------------------------------------+
' Please report errors or other potential problems to      |
' support@dynaforms.com.                                   |
' The current version is available at www.dynaforms.com.   |
'----------------------------------------------------------+
'             Creation date: September 14, 2018            |
'----------------------------------------------------------+

Option Strict On
Option Explicit On
Imports System.Collections
Imports System.Runtime.InteropServices

Namespace DynaPDF

    Public Enum TErrMode
        emIgnoreAll = 0
        emSyntaxError = 1
        emValueError = 2
        emWarning = 4
        emFileError = 8
        emFontError = 16
        emAllErrors = 65535
        emNoFuncNames = &H10000000 ' Do not print function names in error messages
        emUseErrLog = &H20000000   ' Redirect all error messages to the error log
    End Enum

    Public Enum T3DActivationType
        at3D_AppDefault
        at3D_PageOpen    ' The annotaiton should be activated when the page is opened.
        at3D_PageVisible ' The annotaiton should be activated when the page becomes visible.
        at3D_Explicit    ' The annotation should remain inactive until explicitely activated by a script or action (default).
    End Enum

    Public Enum T3DDeActivateType
        dt3D_AppDefault
        dt3D_PageClosed    ' The annotaiton should be deactivated as soon as the page is closed.
        dt3D_PageInvisible ' The annotaiton should be deactivated as soon as the page becomes invisible (default).
        dt3D_Explicit      ' The annotation should remain active until explicitely deactivated by a script or action.
    End Enum

    Public Enum T3DDeActInstance
        di3D_AppDefault
        di3D_UnInstantiated ' The annotation will be uninstantiated (default)
        di3D_Instantiated   ' The annotation is left instantiated
        di3D_Live           ' If the 3D artwork contains an animation then it will stay live
    End Enum

    Public Enum T3DInstanceType
        it3D_AppDefault
        it3D_Instantiated ' The annotation will be instantiated but animations are disabled.
        it3D_Live         ' The annotation will be instantiated and animations are enabled (default).
    End Enum

    Public Enum T3DLightingSheme
        lsArtwork
        lsBlue
        lsCAD
        lsCube
        lsDay
        lsHard
        lsHeadlamp
        lsNight
        lsNoLights
        lsPrimary
        lsRed
        lsWhite
        lsNotSet
    End Enum

    Public Enum T3DNamedAction
        naDefault
        naFirst
        naLast
        naNext
        naPrevious
    End Enum

    Public Enum T3DProjType
        pt3DOrthographic
        pt3DPerspective
    End Enum

    Public Enum T3DRenderingMode
        rmBoundingBox
        rmHiddenWireframe
        rmIllustration
        rmShadedIllustration
        rmShadedVertices
        rmShadedWireframe
        rmSolid
        rmSolidOutline
        rmSolidWireframe
        rmTransparent
        rmTranspBBox
        rmTranspBBoxOutline
        rmTranspWireframe
        rmVertices
        rmWireframe
        rmNotSet
    End Enum

    Public Enum T3DScaleType
        st3DValue
        st3DWidth
        st3DHeight
        st3DMin
        st3DMax
    End Enum

    Public Enum TActionType
        atGoTo
        atGoToR
        atHide
        atImportData
        atJavaScript
        atLaunch
        atMovie
        atNamed
        atRendition     ' PDF 1.5
        atReset         ' ResetForm
        atSetOCGState   ' PDF 1.5
        atSound
        atSubmit        ' SubmitForm
        atThread
        atTransition
        atURI
        atGoTo3DView    ' PDF 1.6
        atGoToE         ' PDF 1.6 Like atGoToR but refers to an embedded PDF file.
        atRichMediaExec ' PDF 1.7 Extension Level 3
    End Enum

    Public Enum TAFRelationship
        arAssociated
        arData
        arSource
        arSupplement
        arAlternative
    End Enum

    Public Enum TAFDestObject
        adAnnotation
        adCatalog    ' The documents catalog is the root object
        adField
        adImage
        adPage
        adTemplate
    End Enum

    Public Enum TAnnotColor
        acBackColor
        acBorderColor
        acTextColor
    End Enum

    Public Enum TAnnotFlags
        afNone = &H0
        afInvisible = &H1
        afHidden = &H2
        afPrint = &H4
        afNoZoom = &H8
        afNoRotate = &H10
        afNoView = &H20
        afReadOnly = &H40
        afLocked = &H80
        afToggleNoView = &H100
        afLockedContents = &H200
    End Enum

    'By default all annotations which have an appearance stream and which have the print flag set are flattened.
    'All annotations are deleted when the function returns with the exception of file attachment annotations.
    'If you want to flatten the view state then set the flag affUseViewState.
    Public Enum TAnnotFlattenFlags
        affNone = &H0               ' Printable annotations independent of the type
        affUseViewState = &H1       ' If set, annotations which are visible in a viewer become flattened.
        affMarkupAnnots = &H2       ' If set, markup annotations are flattened only. Link, Sound, or FileAttach annotations are no markup annotations. These types will be left intact.
        affNonPDFA_1 = &H4          ' If set, flatten all annotations which are not supported in PDF/A 1.
        affNonPDFA_2 = &H8          ' If set, flatten all annotations which are not supported in PDF/A 2 or 3.
        affFormFields = &H10        ' If set, form fields will be flattened too.
        affUseFieldViewState = &H20 ' Meaningful only if affFormFields is set. If set, flatten the view state of form fields. Use the print state otherwise.
    End Enum

    Public Enum TAnnotIcon
        aiComment
        aiHelp
        aiInsert
        aiKey
        aiNewParagraph
        aiNote
        aiParagraph
        aiUserDefined   ' Internal, not usable
    End Enum

    Public Enum TAnnotState
        asNone
        asAccepted
        asRejected
        asCancelled
        asCompleted
        asCreateReply ' Don't add a migration state, create a reply instead. Set the contents of the reply with SetAnnotString().
    End Enum

    Public Enum TAnnotString
        asAuthor
        asContent
        asName
        asSubject
        asRichStyle ' Default style string. -> FreeText annotations only.
        asRichText  ' Rich text string.     -> Supported by markup annotations.
    End Enum

    Public Enum TAnnotType
        atCaret
        atCircle
        atFileLink    ' A Link annotation with an associated GoToR action (go to remote)
        atFreeText
        atHighlight   ' Highlight annotation
        atInk
        atLine
        atPageLink    ' A Link annotation with an associated GoTo action
        atPolygon
        atPolyLine
        atPopUp
        atSquare
        atSquiggly    ' Highlight annotation
        atStamp
        atStrikeOut   ' Highlight annotation
        atText        ' Also used as container to store the State Model
        atUnderline   ' Highlight annotation
        atWebLink     ' A Link annotation with an associated URI action
        atWidget      ' Form Fields are handled separately
        at3D          ' PDF 1.6
        atSoundAnnot  ' PDF 1.2
        atFileAttach  ' PDF 1.3
        atRedact      ' PDF 1.7
        atWatermark   ' PDF 1.6
        atUnknown     ' Unknown annotation type
        atMovieAnnot  ' PDF 1.2
        atPrinterMark ' PDF 1.4
        atProjection  ' PDF 1.7 Extension Level 3
        atRichMedia   ' PDF 1.7 Extension Level 3
        atScreen      ' PDF 1.5
        atTrapNet     ' PDF 1.3
    End Enum

    Public Enum TBaseEncoding
        beWinAnsi
        beMacRoman
        beMacExpert
        beStandard
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TBBox
        Dim x1 As Single
        Dim y1 As Single
        Dim x2 As Single
        Dim y2 As Single
    End Structure

    Public Enum TBlendMode
        bmNotSet
        bmNormal
        bmColor
        bmColorBurn
        bmColorDodge
        bmDarken
        bmDifference
        bmExclusion
        bmHardLight
        bmHue
        bmLighten
        bmLuminosity
        bmMultiply
        bmOverlay
        bmSaturation
        bmScreen
        bmSoftLight
    End Enum

    Public Enum TBorderEffect
        beSolid   ' Default
        beCloudy1 ' Circle diameter 9 units
        beCloudy2 ' Circle diameter 17 units
    End Enum

    Public Enum TBorderStyle
        bsSolid
        bsBevelled
        bsInset
        bsUnderline
        bsDashed
        bsUserDefined ' Internal
    End Enum

    Public Enum TBtnCaptionPos
        bcpCaptionOnly  ' Default
        bcpImageOnly    ' No caption; image only
        bcpCaptionBelow ' Caption below the image
        bcpCaptionAbove ' Caption above the image
        bcpCaptionRight ' Caption on the right of the image
        bcpCaptionLeft  ' Caption on the left of the image
        bcpCaptionOver  ' Caption overlaid directly on the image
    End Enum

    Public Enum TButtonState
        bsUp
        bsDown
        bsRollOver
    End Enum

    Public Enum TCheckBoxChar
        ccCheck
        ccCircle
        ccCross1
        ccCross2
        ccCross3
        ccCross4
        ccDiamond
        ccSquare
        ccStar
    End Enum

    Public Enum TCheckBoxState
        cbUnknown
        cbChecked
        cbUnChecked
    End Enum

    Public Enum TCheckOptions
        coDefault = &H10FFFF
        coEmbedSubsets = &H1
        coDeleteTransferFuncs = &H2
        coDeleteMultiMediaContents = &H4
        coDeleteActionsAndScripts = &H8
        coDeleteInvRenderingIntent = &H10
        coFlattenFormFields = &H20
        coReplaceV4ICCProfiles = &H40
        coDeleteEmbeddedFiles = &H80
        coDeleteOPIComments = &H100
        coDeleteSignatures = &H200
        coDeletePostscript = &H400        ' Delete Postscript XObjects. Rarely used and such Postscript fragments are meaningful on a Postscript device only.
        ' It is usually safe to delete such objects.
        coDeleteAlternateImages = &H800   ' Alternate images are seldom used and prohibited in PDF/A.
        coReComprJPEG2000Images = &H1000  ' Recompression results usually in larger images. It is often better to keep such files as is.
        coResolveOverprint = &H2000       ' PDF/A 2 and 3. Set the overprint mode to 0 if overprint mode = 1 and if overprinting for fill or stroke is true
        ' and if an ICCBased CMYK color space is used. Note that DeviceCMYK is treated as ICCBased color space due to implicit
        ' color conversion rules.
        coMakeLayerVisible = &H4000       ' PDF/A 2 and 3 prohibit invisible layers. Layers can also be flattened if this is no option.
        coDeleteAppEvents = &H8000        ' PDF/A 2 and 3. Application events are prohibited in PDF/A. The view state will be applied.
        coReplCCITTFaxWithFlate = &H10000 ' Replace CCITT Fax compression with Flate.
        coApplyExportState = &H20000      ' Meaningful only if coDeleteAppEvents is set. Apply the export state.
        coApplyPrintState = &H40000       ' Meaningful only if coDeleteAppEvents is set. Apply the print state.
        coDeleteReplies = &H80000         ' Delete annotation replies. If absent, replies will be converted to regular text annotations.
        coDeleteHalftones = &H100000      ' Delete halftone screens.
        coFlattenLayers = &H200000        ' Flatten layers if any.
        coDeletePresentation = &H400000   ' Presentations are prohibited in PDF/A 2 and 3.
        coCheckImages = &H800000          ' If set, images will be decompressed to identify damaged images.
        coDeleteDamagedImages = &H1000000 ' Meaningful only if coCheckImages is set.
        coRepairDamagedImages = &H2000000 ' Meaningful only if coCheckImages is set. If set, try to recompress a damaged image. The new image is maybe not complete but error free.
        coNoFontEmbedding = &H10000000    ' If set, non-embedded fonts are left as is.
        coFlushPages = &H20000000         ' Write converted pages directly into the output file to reduce the memory usage.
        coAllowDeviceSpaces = &H40000000  ' If set, device color spaces will not be replaced with ICC based color spaces. This flag is meaningful for normalization only.
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TCIDMetric
        Dim Width As Single
        Dim x As Single
        Dim y As Single
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Class CCIDMetric
        Dim Width As Single
        Dim x As Single
        Dim y As Single
        Public Shared Widening Operator CType(ByVal v As CCIDMetric) As TCIDMetric
            Dim retval As TCIDMetric
            retval.Width = v.Width
            retval.x = v.x
            retval.y = v.y
            Return retval
        End Operator
    End Class

    Public Enum TClippingMode
        cmEvenOdd
        cmWinding
    End Enum

    Public Enum TCodepage
        cp1250
        cp1251
        cp1252
        cp1253
        cp1254
        cp1255
        cp1256
        cp1257
        cp1258
        cp8859_2
        cp8859_3
        cp8859_4
        cp8859_5
        cp8859_6
        cp8859_7
        cp8859_8
        cp8859_9
        cp8859_10
        cp8859_13
        cp8859_14
        cp8859_15
        cp8859_16
        cpSymbol
        cp437
        cp737
        cp775
        cp850
        cp852
        cp855
        cp857
        cp860
        cp861
        cp862
        cp863
        cp864
        cp865
        cp866
        cp869
        cp874
        cpUnicode
        cpCJK_Big5_Uni    ' Big5 plus HKSCS extension.
        cpCJK_EUC_JP_Uni  ' EUC-JP
        cpCJK_EUC_KR_Uni  ' EUC-KR
        cpCJK_EUC_TW_Uni  ' CNS-11643-1992 (Planes 1-15).
        cpCJK_GBK_Uni     ' GBK is the Microsoft code page 936 (GB2312 EUC-CN plus GBK extension).
        cpCJK_GB12345_Uni ' GB-12345-1990 (Traditional Chinese form of GB-2312).
        cpCJK_HZ_Uni      ' Mixed ASCII / GB-2312 encoding
        cpCJK_2022_CN_Uni ' ISO-2022-CN-EXT (GB-2312 plus ISO-11643 Planes 1-7).
        cpCJK_2022_JP_Uni ' ISO-2022-JP
        cpCJK_2022_KR_Uni ' ISO-2022-KR
        cpCJK_646_CN_Uni  ' ISO-646-CN (GB-1988-80)
        cpCJK_646_JP_Uni  ' ISO-646-JP (JIS_C6220-1969-RO).
        cpCJK_IR_165_Uni  ' ISO-IR-165 (extended version of GB-2312).
        cpCJK_932_Uni     ' Microsoft extended version of SHIFT_JIS.
        cpCJK_949_Uni     ' EUC-KR extended with UHC (Unified Hangul Codes).
        cpCJK_950_Uni     ' Microsoft extended version of Big5.
        cpCJK_JOHAB_Uni   ' JOHAB
        cpShiftJIS        ' ShiftJIS charset plus code page 932 ectension.
        cpBig5            ' Big5 plus HKSCS extension.
        cpGB2312          ' GB2312 charset plus GBK and cp936 extension.
        cpWansung         ' Wansung
        cpJohab           ' JOHAB
        cpMacRoman        ' Mac Roman
        cpAdobeStd        ' This is an encoding for Type1 fonts. It should normally not be used.
        cpInternal        ' Internal -> not usable
        cpGlyphIndexes    ' Can be used with TrueType and OpenType fonts only. DynaPDF creates a reverse mapping so that copy & paste will work.
        cpPDFDocEnc       ' Internal -> not usable. Used for form fonts. This is a superset of the code page 1252 and MacRoman.
        cpExtCMap         ' Internal -> not usable. This code page is set when a font was loaded with an external cmap.
        cpDingbats        ' Internal -> Special encoding for ZapfDingbats
        cpMacExpert       ' Internal -> not usable
        cpRoman8          ' This is a standard PCL 5/6 code page
    End Enum

    ' The data for user defined columns is stored in collection items.
    Public Enum TColColumnType
        cisCreationDate ' Data comes from the embedded file
        cisDescription  ' Data comes from the embedded file
        cisFileName     ' Data comes from the embedded file
        cisModDate      ' Data comes from the embedded file
        cisSize         ' Data comes from the embedded file
        cisCustomDate   ' User defined date.
        cisCustomNumber ' User defined nummber.
        cisCustomString ' User defined string.
    End Enum

    Public Enum TColorConvFlags
        ccfBW_To_Gray = 0   ' Default, RGB Black and White set with rg or RG inline operators are converted to gray
        ccfRGB_To_Gray = 1  ' If set, inline color operators rg and RG are converted to gray
        ccfToGrayAdjust = 2 ' Converts RGB and gray inline operators to gray and allows to darken or lighten the colors
    End Enum

    Public Enum TColorMode
        cmFill
        cmStroke
        cmFillStroke
    End Enum

    Public Enum TColView
        civNotSet
        civDetails
        civTile
        civHidden
        civCustom  ' PDF 1.7 Extension Level 3, the collection view is presented by a SWF file.
    End Enum

    Public Enum TCompBBoxFlags
        cbfNone = 0
        cbfIgnoreWhiteAreas = 1 ' Ignore white vector graphics or text.
        ' Please note that images must be decompressed if one of the following flags are set. Parsing gray or color images
        ' is in most cases not useful and you should not parse such images if it is not really required.
        cbfParse1BitImages = 2  ' Find the visible area in 1 bit images. This is the most important case
        ' since scanned faxes are usually 1 bit images.
        cbfParseGrayImages = 4  ' Find the visible area in gray images.
        cbfParseColorImages = 8 ' Find the visible area in color images. This is usually not required
        ' and slow downs processing a lot.
        cbfParseAllImages = 14  ' Find the visible area in all images.
    End Enum

    Public Enum TCompressionLevel
        clNone
        clDefault
        clFastest
        clMax
    End Enum

    Public Enum TCompressionFilter
        cfFlate = 0                     ' Flate or Zip compression
        cfJPEG = 1                      ' JPEG compression
        cfCCITT3 = 2                    ' CCITT Fax G3 compression
        cfCCITT4 = 3                    ' CCITT Fax G4 compression
        cfLZW = 4                       ' TIFF or GIF output -> Very fast but less compression ratios than flate
        cfLZWBW = 5                     ' TIFF
        cfFlateBW = 6                   ' TIFF, PNG, or BMP output -> Dithered black & white output. The resulting image will be
        ' compressed with Flate or left uncompressed if the output image format is a bitmap. If
        ' you want to use CCITT Fax 4 compression (TIFF only) set the flag icUseCCITT4 in the
        ' AddImage() function call. Note that this filter is not supported for PDF creation!
        cfJP2K = 7                      ' JPEG2000 compression
        cfJBIG2 = 8                     ' PDF output only
        ' Special flags for AddRasImage(). These flags can be combined with the filters cfFlate, cfCCITT3, cfCCITT4, and LZW.
        cfDitherFloydSteinberg = &H1000 ' Floyd Steinberg dithering.
        cfConvGrayToOtsu = &H2000       ' The Otsu filter is a special filter to produce black & white images. It is very useful
        ' if an OCR scan should be applied on the resulting 1 bit image. The flag is considered
        ' in AddRasImage(), RenderPDFFile(), and RenderPageToImage() if the pixel format was set
        ' to pxfGray.
    End Enum

    Public Enum TConformanceType
        ctPDFA_1b_2005
        ctNormalize
        ctPDFA_2b
        ctPDFA_3b
        ' The following constants convert the file to PDF/A 3b and set the whished ZUGFeRD conformance level
        ' in the XMP metadata. CheckConformance() does not validate the XML invoice but it checks whether it
        ' is present. Setting the correct ZUGFeRD conformance level is very important since this value defines
        ' which fields must be present in the XML invoice.
        ctZUGFeRD_Basic    ' Set the ZUGFeRD conformance level to Basic
        ctZUGFeRD_Comfort  ' Set the ZUGFeRD conformance level to Comfort
        ctZUGFeRD_Extended ' Set the ZUGFeRD conformance level to Extended
    End Enum

    Public Enum TDateType
        dtCreationDate
        dtModeDate
    End Enum

    Public Enum TDecodeFilter
        dfNone
        dfASCII85Decode   ' No parameters
        dfASCIIHexDecode  ' No parameters
        dfCCITTFaxDecode  ' Optional Parameters
        dfDCTDecode       ' Optional Parameters
        dfFlateDecode     ' Optional Parameters
        dfJBIG2Decode     ' Optional Parameters
        dfJPXDecode       ' No parameters
        dfLZWDecode       ' Optional Parameters
        dfRunLengthDecode ' No parameters
    End Enum

    Public Enum TDecSeparator
        ' per thousand separator, decimal separator
        dsCommaDot
        dsNoneDot
        dsDotComma
        dsNoneComma
    End Enum

    Public Enum TDestType
        dtXY_Zoom ' three parameters (a, b, c) -> (X, Y, Zoom)
        dtFit ' no parameters
        dtFitH_Top ' one parameter    (a)
        dtFitV_Left ' one parameter    (a)
        dtFit_Rect ' four parameters  (a, b, c, d) -> (left, bottom, right, top)
        dtFitB ' no parameter
        dtFitBH_Top ' one parameter    (a)
        dtFitBV_Left ' one parameter    (a)
    End Enum

    Public Enum TDocumentInfo
        diAuthor
        diCreator
        diKeywords
        diProducer
        diSubject
        diTitle
        diCompany
        diPDFX_Ver     ' GetInDocInfo() or GetInDocInfoEx() only -> The PDF/X version is set by SetPDFVersion()!
        diCustom       ' User defined key
        diPDFX_Conf    ' GetInDocInfo() or GetInDocInfoEx() only. The value of the GTS_PDFXConformance key.
        diCreationDate ' GetInDocInfo() or GetInDocInfoEx() or after ImnportPDFFile() was called.
        diModDate      ' GetInDocInfo() or GetInDocInfoEx() only
    End Enum

    Public Enum TDrawDirection
        ddCounterClockwise
        ddClockwise
    End Enum

    Public Enum TDrawMode
        dmNormal
        dmStroke
        dmFillStroke
        dmInvisible
        dmFillClip
        dmStrokeClip
        dmFillStrokeClip
        dmClipping
    End Enum

    Public Enum TDuplexMode
        dpmNone          ' Default
        dpmSimplex
        dpmFlipShortEdge
        dpmFlipIntegerEdge
    End Enum

    Public Enum TEmbFileLocation
        eflChild         ' The file is an embedded file in the current document
        eflChildAnnot    ' The file is located in a file attachment annotion in the current document
        eflExternal      ' The file is an embedded file in an external document
        eflExternalAnnot ' The file is located in a file attachment annotion in an external document
        eflParent        ' The file is located in the parent document
        eflParentAnnot   ' The file is located in a file attachment annotion in the parent document
    End Enum

    Public Enum TEnumFontProcFlags
        efpAnsiPath = 0    ' Code page 1252 on Windows, UTF-8 otherwise
        efpUnicodePath = 1 ' FilePath is in Unicode format. Make a typecast to (UI16*) in this case.
        efpEmbeddable = 2  ' The font has embedding rights.
        efpEditable = 4    ' If set, the font has editing rights (important for form fields).
    End Enum

    Public Enum TExtColorSpace
        esDeviceRGB   ' Device color space
        esDeviceCMYK  ' Device color space
        esDeviceGray  ' Device color space
        esCalGray     ' CIE-based color space
        esCalRGB      ' CIE-based color space
        esLab         ' CIE-based color space
        esICCBased    ' CIE-based color space -> contains an ICC profile
        esPattern     ' Special color space
        esIndexed     ' Special color space
        esSeparation  ' Special color space
        esDeviceN     ' Special color space
        esNChannel    ' Special color space
    End Enum

    Public Enum TFieldColor
        fcBackColor
        fcBorderColor
        fcTextColor
    End Enum

    Public Enum TFieldFlags
        ffReadOnly = &H1
        ffRequired = &H2
        ffNoExport = &H4

        ffInvisible = &H8
        ffHidden = &H10
        ffPrint = &H20
        ffNoZoom = &H40
        ffNoRotate = &H80
        ffNoView = &H100

        ffMultiline = &H1000         ' Text fields only
        ffPassword = &H2000          ' Text fields only
        ffNoToggleToOff = &H4000     ' Radio buttons, check boxes
        ffRadioIsUnion = &H2000000   ' PDF-1.5 radio buttons
        ffCommitOnSelCh = &H4000000  ' PDF-1.5 combo boxes, list boxes

        ffEdit = &H40000             ' Combo boxes only
        ffSorted = &H80000           ' Combo boxes and list boxes -> sorts the choice values in ascending order
        ffFileSelect = &H100000      ' PDF 1.4 Text fields only
        ffMultiSelect = &H200000     ' PDF 1.4 List boxes only
        ffDoNotSpellCheck = &H400000 ' PDF 1.4 Text fields, combo boxes. If the field is a combo box, this flag is meaningful only if ffEdit is also set.
        ffDoNotScroll = &H800000     ' PDF 1.4 Text fields only
        ffComb = &H1000000
    End Enum

    Public Enum TFieldType
        ftButton
        ftCheckBox
        ftRadioBtn
        ftComboBox
        ftListBox
        ftText
        ftSignature
        ftGroup ' this is not a real field type, it is just an array of fields
    End Enum

    Public Enum TFileAttachIcon
        faiGraph
        faiPaperClip
        faiPushPin
        faiTag
        faiUserDefined
    End Enum

    Public Enum TFileOP
        foOpen
        foPrint
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TFltPoint
        Dim x As Single
        Dim y As Single
        Public Sub New(ByVal x_ As Single, ByVal y_ As Single)
            x = x_
            y = y_
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TI32Point
        Dim x As Integer
        Dim y As Integer
        Public Sub New(ByVal x_ As Integer, ByVal y_ As Integer)
            x = x_
            y = y_
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TFltRect
        Dim Left As Single
        Dim Bottom As Single
        Dim Right As Single
        Dim Top As Single
    End Structure

    Public Enum TFlushPageFlags
        fpfDefault = 0      ' Write anything to the file that is possible
        fpfImagesOnly = 1   ' If set, only images are written to the file. The pages are still
        ' in memory and can be modified with EditPage(). Flushed images can
        ' still be referenced in other pages. The image handles remain valid.
        fpfExclLastPage = 2 ' If set, the last page is not flushed
    End Enum

    Public Enum TFontBaseType
        fbtTrueType ' TrueType, TrueType Collections, or OpenType fonts with TrueType outlines
        fbtType1    ' Type1 font
        fbtOpenType ' OpenType font with Postscript outlines
        fbtStdFont  ' PDF Standard font
        fbtDisabled ' This value can be used to disable a specific font format. See SetFontSearchOrder() for further information.
    End Enum

    Public Enum TFontFileSubtype
        ffsType1C        ' CFF based Type1 font
        ffsCIDFontType0C ' CFF based Type1 CID font
        ffsOpenType      ' TrueType based OpenType font
        ffsOpenTypeC     ' CFF based OpenType font
        ffsCIDFontType2  ' TrueType based CID Font
        ffsReserved1
        ffsReserved2
        ffsReserved3
        ffsReserved4
    End Enum

    Public Enum TFontSelMode
        smFamilyName
        smPostScriptName
        smFullName
    End Enum

    Public Enum TFontType
        ftMMType1
        ftTrueType
        ftType0    ' Check the font file type to determine the font sub type
        ftType1
        ftType3
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TFRect
        Dim MinX As Int16
        Dim MinY As Int16
        Dim MaxX As Int16
        Dim MaxY As Int16
    End Structure

    'The font search run works as follows:

    '   - DynaPDF tries always to find the exact weight, if it cannot be found then a font with
    '     the next smaller weight is selected (if available).
    '   - Italic styles can always be emulated but it is not possible to emulate thinner weights or
    '     regular styles with an italic font.
    '   - If the specified weight is larger as the font weight the remaining weight will be emulated
    '     if the difference to the requested weight is larger than 200.
    '     With SetFontWeight() it is possible to control whether a missing weight should be emulated.
    '     If FontWeight is smaller or equal to the requested font weight then emulation will be disabled.

    'TFStyle is a bitmask that is defined as follows:

    '   - Bits 0..7   // Style bits fsItalic, fsUnderlined, fsStriked
    '   - Bits 8..19  // Width class -> Defined for future use.
    '   - Bits 20..31 // Font Weight

    '- A width class can be converted to a style constant by multiplying it with 256 (width shl 8).
    '- A font weight can be converted to a style constant by multiplying it with 1048576 (weight shl 20).
    '- Additional attributes can be added with a binary or operator (e.g. style or fsItalic).
    '- Only one width class and one font weight can be set at time.

    '- WidthFromStyle() extracts the width class.
    '- WeightFromStyle() extracts the font weight.

    'The following functions extract the width class or font weight from a style variable:

    '   widthClass  = WidthFromStyle(style);
    '   weightClass = WeightFromStyle(style);

    Public Enum TFStyle
        fsNone = &H0                  ' Regular weight (400)
        fsItalic = &H1
        fsUnderlined = &H4
        fsStriked = &H8
        fsVerticalMode = &H10         ' Not considered at this time
        ' Width class
        fsUltraCondensed = &H100      ' 1
        fsExtraCondensed = &H200      ' 2
        fsCondensed = &H300           ' 3
        fsSemiCondensed = &H400       ' 4
        fsNormal = &H500              ' 5
        fsSemiExpanded = &H600        ' 6
        fsExpanded = &H700            ' 7
        fsExtraExpanded = &H800       ' 8
        fsUltraExpanded = &H900       ' 9
        ' Weight class
        fsThin = &H6400000            ' 100
        fsExtraLight = &HC800000      ' 200
        fsLight = &H12C00000          ' 300
        fsRegular = &H19000000        ' 400 -> Same as fsNone
        fsMedium = &H1F400000         ' 500
        fsDemiBold = &H25800000       ' 600
        fsBold = &H2BC00000           ' 700 -> The old constant 2 is still supported to preserve backward compatibility
        fsExtraBold = &H32000000      ' 800
        fsBlack = &H38400000          ' 900
        fsUltraBlack = &H3E800000     ' 1000
    End Enum

    Public Enum TGStateFlags
        gfCompatible = 0         ' Compatible graphics state to earlier DynaPDF versions -> default
        gfRestorePageCoords = 1  ' Restore the coordinate system with the graphics state (the value of PageCoords, see SetPageCoords())
        gfRealTopDownCoords = 2  ' If set, the page coordinate system is not reset to bottom-up when transforming
        ' the coordinate system. However, real top-down coordinates require a large internal
        ' overhead and where never fully implemented. The usage of this flag should be avoided
        ' if possible.
        gfUseImageColorSpace = 8 ' If set, the active color space is ignored when inserting an image. The color space is taken
        ' from the image file instead.
        gfIgnoreICCProfiles = 16 ' Meaningful only if the flag gfUseImageColorSpace is set. If set, an embedded profile is not used to
        ' create an ICCBased color space for the image. The image is inserted in the corresponding device
        ' color space instead.
        gfAnsiStringIsUTF8 = 32  ' If set, single byte strings in Ansi functions are treated as UTF-8 encoded Unicode strings.
        gfRealPassThrough = 64   ' If set, JPEG images are inserted as is. JPEG images are normally rebuild, also in pass-through mode, to avoid issues
        ' with certain malformed JPEG images which cannot be displayed in Adobes Acrobat or Reader. If you know that your JPEG
        ' images work then set this flag to avoid unnecessary processing time.
        gfNoBitmapAlpha = 128    ' If set, the alpha channel in 32 bit bitmaps will be ignored. Useful for bitmaps with an invalid alpha channel.
        gfNoImageDuplCheck = 256 ' If set, no duplicate check for images will be performed. This can significantly improve processing speed.
        gfNoObjCompression = 512 ' If set, object compression will be disabled.
    End Enum

    Public Enum THashType
        htDetached ' CloseAndSignFileExt() returns the byte ranges of the finish PDF buffer to create a detached signature
        htSHA1     ' CloseAndSignFileExt() returns the SHA1 hash of the PDF file so that it can be signed
    End Enum

    Public Enum THighlightMode
        hmNone
        hmInvert
        hmOutline
        hmPush
        hmPushUpd ' Update appereance stream on changes
    End Enum

    Public Enum TICCProfileType
        ictGray
        ictRGB
        ictCMYK
        ictLab
    End Enum

    'TIFF is the only format that supports different compression filters. The Filter parameter of the function
    'AddImage() is ignored if the image format supports only one specific compression filter.
    'Note that images are automatically converted to the nearest supported color space if the image format does
    'not support the color space of the image.
    Public Enum TImageFormat
        ifmTIFF     ' DeviceRGB, DeviceCMYK, DeviceGray, Black & White -> CCITT Fax Group 3/4, JPEG, Flate, LZW.
        ifmJPEG     ' DeviceRGB, DeviceCMYK, DeviceGray    -> JPEG compression.
        ifmPNG      ' DeviceGray, DeviceRGB, Black & White -> Flate compression.
        ifmReserved ' Reserved for future extensions.
        ifmBMP      ' DeviceGray, DeviceRGB, Black & White -> Uncompressed.
        ifmJPC      ' DeviceRGB, DeviceCMYK, DeviceGray    -> JPEG2000 compression.
    End Enum

    Public Enum TImageConversionFlags
        icNone      ' Default
        icUseCCITT4 ' Use CCITT Fax 4 compression instead of Flate for dithered images.
    End Enum

    Public Enum TImportFlags
        ifImportAll = &HFFFFFFE ' default
        ifContentOnly = &H0
        ' If this flag is set, only interactive objects are imported if any, Otherwise only empty pages are imported.
        ' This flag can be used to copy an interactive form to another PDF file.
        ifNoContent = &H1
        ' The imported page is not converted to a template if ifImportAsPage is set.
        ' Note that this flag can cause resource conflicts. Use this flag carefully!
        ifImportAsPage = &H80000000
        ' base objects
        ifCatalogAction = &H2
        ifPageActions = &H4
        ifBookmarks = &H8
        ifArticles = &H10
        ifPageLabels = &H20
        ifThumbs = &H40
        ifTranspGroups = &H80        ' This flag is not Integerer considered.
        ifSeparationInfo = &H100
        ifBoxColorInfo = &H200
        ifStructureTree = &H400
        ifTransition = &H800
        ifSearchIndex = &H1000
        ifJavaScript = &H2000
        ifJSActions = &H4000
        ifDocInfo = &H8000           ' Document info entries
        ifEmbeddedFiles = &H200000   ' File attachments
        ifFileCollections = &H400000 ' File collections (PDF 1.7)
        ' Annotations -> Only the most important annotation types can be selected directly.
        ' Note that all annotation types can be deleted with DeleteAnnotation.
        ifAllAnnots = &H9F0000
        ifFreeText = &H10000
        ifTextAnnot = &H20000
        ifLink = &H40000
        ifStamp = &H80000
        if3DAnnot = &H100000
        ifOtherAnnots = &H800000
        ' Interactive Form fields are also annotations but we handle this type separately!
        ifFormFields = &H1000000
        ifPieceInfo = &H2000000      ' The PieceInfo dictionary contains arbitrary user defined data. The data in
        ' this dictionary is meaningful only for the application that created the data.

        ' -------------------- Special flags --------------------
        ifPrepareForPDFA = &H10000000 ' Replace LZW compression with Flate, set the Interpolate key of images to false, do not import embedded files.
        ifEnumFonts = &H20000000      ' Import fonts for EnumDocFonts(). The document must be deleted when this flag is set!!!
        ifAllPageObjects = &H40000000 ' Import links when using ImportPageEx() within an open page. The entire document should be imported in this case.
    End Enum

    Public Enum TImportFlags2
        if2MergeLayers = &H1          ' If set, layers with identical name are merged. If this flag is absent DynaPDF
        ' imports such layers separately so that each layer refers still to the pages
        ' where it was originally used.
        if2Normalize = &H2            ' Replace LZW compression with Flate, apply limit checks, repair errors if possible.
        if2UseProxy = &H4             ' Not meaningful for PDF files which are loaded from a memory buffer. If set, all streams are loaded from the file
        ' on demand but they are never hold in memory. This reduces drastically the memory usage and enables the processing
        ' of almost arbitrary large PDF files with minimal memory usage. The corresponding PDF file must not be deleted before
        ' CloseFile() or CloseFileEx() was called.
        if2NoMetadata = &H8           ' Ignore metadata streams which are attached to fonts, pages, images, and so on.
        if2DuplicateCheck = &H10      ' Perform a duplicate check on color spaces, fonts, images, patterns, and templates when merging PDF files.
        if2NoResNameCheck = &H20      ' Import resources as is. This flag can significantly imporove the loading time of pages with a huge resource tree.
        ' This flag should only be set in viewer applications to improve the loading time of pages.
        if2CopyEncryptDict = &H40     ' If set, the encryption settings of an encrypted PDF file are copied to the new PDF file.
        ' The flag does nothing if the file is not encrypted.
    End Enum

    Public Structure TIntRect
        Public x1 As Integer
        Public y1 As Integer
        Public x2 As Integer
        Public y2 As Integer
    End Structure

    Public Enum TKeyLen
        kl40bit    ' RC4 Encryption -> Acrobat 3 or higher
        kl128bit   ' RC4 Encryption -> Acrobat 5 or higher
        kl128bitEx ' RC4 Encryption -> Acrobat 6 or higher
        klAES128   ' AES Encryption -> Acrobat 7 or higher
        klAES256   ' AES Encryption -> Acrobat 9 or higher
        klAESRev6  ' AES Encryption -> Acrobat X or higher
    End Enum

    Public Enum TLineCapStyle
        csButtCap
        csRoundCap
        csSquareCap
    End Enum

    Public Enum TLineEndStyle
        leNone
        leButt
        leCircle
        leClosedArrow
        leDiamond
        leOpenArrow
        leRClosedArrow
        leROpenArrow
        leSlash
        leSquare
    End Enum

    Public Enum TLineCaptionPos
        cpInline ' The caption is centered inside the line
        cpTop     ' The caption is drawn on top of the line
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Class TLineAnnotParms
        Public StructSize As Integer         ' Must be set to sizeof(TLineAnnotParms)
        Public Caption As Integer            ' If true, the annotation string Content is used as caption.
        Public CaptionOffsetX As Single      ' Horizontal offset of the caption from its normal position
        Public CaptionOffsetY As Single      ' Vertical offset of the caption from its normal position
        Public CaptionPos As TLineCaptionPos ' The position where the caption should be drawn if present
        Public LeaderLineLen As Single       ' Length of the leader lines (positive or negative)
        Public LeaderLineExtend As Single    ' Optional leader line extend beyond the leader line (must be a positive value or zero)
        Public LeaderLineOffset As Single    ' Amount of space between the endpoints of the annotation and the leader lines (must be a positive value or zero)
    End Class

    Public Enum TLineJoinStyle
        jsMiterJoin
        jsRoundJoin
        jsBevelJoin
    End Enum

    Public Enum TLoadCMapFlags
        lcmDefault = 0   ' Load the cmaps in the directory now
        lcmRecursive = 1 ' Load sub directories recursively
        lcmDelayed = 2   ' Load the cmap files only when a font requires an external cmap
    End Enum

    Public Enum TMeasureNumFormat
        mnfDecimal
        mnfFractional
        mnfRound
        mnfTruncate
    End Enum

    Public Enum TMeasureLblPos
        mlpSuffix
        mlpPrefix
    End Enum

    Public Enum TMetadataObj
        mdoCatalog
        mdoFont
        mdoImage
        mdoPage
        mdoTemplate
    End Enum

    Public Enum TMetaFlags
        mfDefault = 0                   ' Default conversion
        mfDebug = 1                     ' Write debug information into the content stream
        mfShowBounds = 2                ' Show the bounding boxes of text strings
        mfNoTextScaling = 4             ' Do not scale text
        mfClipView = 8                  ' Draw the file into a clipping rectangle
        mfUseRclBounds = 16             ' Use rclBounds instead of rclFrame
        mfNoClippingRgn = 64            ' Disables SelectClippingRegion and IntersectClipRect
        mfNoFontEmbedding = 128         ' Do not embed fonts -> Fonts should be embedded!!!
        mfNoImages = 256                ' Ignore image records
        mfNoStdPatterns = 512           ' Ignore standard patterns
        mfNoBmpPatterns = 1024          ' Ignore bitmap patterns
        mfNoText = 2048                 ' Ignore text records
        mfUseUnicode = 4096             ' Ignore ANSI_CHARSET
        mfUseTextScaling = 16384        ' Scale text instead of using the intercharacter spacing array
        mfNoUnicode = 32768             ' Avoid usage of Unicode fonts -> recommended to enable PDF 1.2 compability
        mfFullScale = 65536             ' Scale coordinates to the window size. Recommended if 32 bit coordinates are used.
        mfUseRclFrame = 131072          ' This flag should be set if the rclFrame rectangle is properly set
        mfDefBkModeTransp = 262144      ' Initialize the background mode to transparent (SetBkMode() overrides this state).
        mfApplyBidiAlgo = &H80000       ' Apply the bidirectional algorithm on Unicode strings
        mfGDIFontSelection = &H100000   ' Use the GDI to select fonts
        mfRclFrameEx = &H200000         ' If set, and if the rclBounds rectangle is larger than rclFrame, the function
        ' extends the output rectangle according to rclBounds and uses the resulting
        ' bounding box to calculate the image size (rclBounds represents the unscaled
        ' image size). This is probably the correct way to calculate the image size.
        ' However, to preserve backward compatibility the default calculation cannot
        ' be changed.
        mfNoTextClipping = &H400000     ' If set, the ETO_CLIPPED flag in text records is ignored.
        mfSrcCopy_Only = &H800000       ' If set, images which use a ROP code other than SRCCOPY are ignored. This is useful when processing Excel 2007 spool files.
        mfClipRclBounds = &H1000000     ' If set, the graphic is drawn into a clipping path with the size of rclBounds.
        ' This flag is useful if the graphic contains content outside of its bounding box.
        mfDisableRasterEMF = &H2000000  ' If set, EMF files which use unsupported ROP codes are not rastered.
        mfNoBBoxCheck = &H4000000       ' If set, the rclBounds and rclFrame rectangles are used as is. DynaPDF uses normally
        ' the rclBounds rectangle to calculate the picture size if the resolution of the EMF file
        ' seems to be larger than 1800 DPI since this is mostly an indication that the rclFrame
        ' rectangle was incorrectly calculated. If you process EMF files in such a high resolution
        ' then this flag must be set. The flag can be set by default.
        mfIgnoreEmbFonts = &H8000000   ' If set, embedded fonts in GDIComment records will be ignored. This flag must be set if the fonts
        ' of an EMF spool file were pre-loaded with ConvertEMFSpool(). Spool fonts must always be loaded
        ' in a pre-processing step since required fonts are not necessarily embedded in the EMF files.

        ' Obsolete flags -> these flags are ignored, do no Integerer use them!
        mfUseSpacingArray = 32          ' enabled by default -> can be disabled with mfUseTextScaling
        mfIntersectClipRect = 8192      ' enabled by default -> can be disabled with mfNoClippingRgn
    End Enum

    Public Enum TNamedAction
        naFirstPage
        naLastPage
        naNexPage
        naPrevPage
        naGoBack
        naOpenDlg
        naPrintDlg
        naGeneralInfo
        naFontsInfo
        naSaveAs
        naSecurityInfo
        naFitPage
        naFullScreen
        naDeletePages
        naQuit
        naUserDefined ' Non predefined action
    End Enum

    Public Enum TNegativeStyle
        nsMinusBlack
        nsRed
        nsParensBlack
        nsParensRed
    End Enum

    Public Enum TNewAlign
        naUnchanged = 0
        naLeft = 1
        naCenter = 2
        naRight = 3
        naJustify = 4
    End Enum

    ' All actions which should be applied to an event except On Mouse Upend Public Enum must be a JavaScript action!
    Public Enum TObjEvent
        oeNoEvent          ' Internal use only -> DO NOT USE THIS VALUE!!!
        oeOnOpen           ' Catalog Pages
        oeOnClose          ' Pages only
        oeOnMouseUp        ' All fields page link annotations bookmarks
        oeOnMouseEnter     ' Form fields only
        oeOnMouseExit      ' Form fields only
        oeOnMouseDown      ' Form fields only
        oeOnFocus          ' Form fields only
        oeOnBlur           ' Form fields only
        oeOnKeyStroke      ' Text fields only
        oeOnFormat         ' Text fields only
        oeOnCalc           ' Text fields combo boxes list boxes
        oeOnValidate       ' All form fields except buttons
        oeOnPageVisible    ' PDF 1.5 -> Form fields only
        oeOnPageInVisible  ' PDF 1.5 -> Form fields only
        oeOnPageOpen       ' PDF 1.5 -> Form fields only
        oeOnPageClose      ' PDF 1.5 -> Form fields only
        oeOnBeforeClosing  ' PDF 1.4 -> Catalog only
        oeOnBeforeSaving   ' PDF 1.4 -> Catalog only
        oeOnAfterSaving    ' PDF 1.4 -> Catalog only
        oeOnBeforePrinting ' PDF 1.4 -> Catalog only
        oeOnAfterPrinting  ' PDF 1.4 -> Catalog only
    End Enum

    Public Enum TObjType
        otAction
        otAnnotation
        otBookmark
        otCatalog ' PDF 1.4
        otField
        otPage
        otPageLink
    End Enum

    Public Enum TOCAppEvent
        aeExport = 1
        aePrint = 2
        aeView = 4
    End Enum

    Public Enum TOCGIntent
        oiDesign = 2
        oiView = 4     ' Default
        oiAll = 8
        oiEmpty = 16   ' Internal
        ' Special flag for GetOCG().
        oiVisible = 32 ' This flag is not considered when creating a layer. It is only used in GetOCG() to determine whether a layer is visible.
    End Enum

    Public Enum TOCObject
        ooAnnotation
        ooField
        ooImage
        ooTemplate
    End Enum

    Public Enum TOCPageElement
        peBackgroundImage ' BG
        peForegroundImage ' FG
        peHeaderFooter    ' HF
        peLogo            ' L
        peNone
    End Enum

    Public Enum TOCGUsageCategory
        oucNone = 0
        oucExport = 1
        oucLanguage = 2
        oucPrint = 4
        oucUser = 8
        oucView = 16
        oucZoom = 32
    End Enum

    Public Enum TOCUserType
        utIndividual
        utOrganization
        utTitle
        utNotSet
    End Enum

    Public Enum TOCVisibility
        ovAllOff
        ovAllOn
        ovAnyOff
        ovAnyOn
        ovNotSet ' Internal
    End Enum

    Public Enum TOptimizeFlags
        ofDefault = &H0                      ' Just rebuild the content streams.
        ofInMemory = &H1                     ' Optimize the file fully in memory. Only useful for small PDF files.
        ofConvertAllColors = &H2             ' If set Separation DeviceN and NChannel color spaces will be converted to the device space.
        ofIgnoreICCBased = &H4               ' If set ICCBased color spaces will be left unchanged.
        ofScaleImages = &H8                  ' Scale images as specified in the TOptimizeParams structure.
        ofSkipMaskedImages = &H10            ' Meaningful only if ofScaleImages is set. If set don't scale images with a color mask.
        ofNewLinkNames = &H20                ' If set rename all object links to short names like F1 F2 etc.
        ofDeleteInvPaths = &H40              ' Delete invisible paths. An invisible path is a path that was finished with the no-op operator "n".
        ofFlattenLayers = &H80               ' Flatten layers if any.
        ofDeletePrivateData = &H100          ' Delete private data objects from pages templates and images.
        ofDeleteThumbnails = &H200           ' Thumbnails can be deleted since PDF viewers can create thumbnails easily on demand.
        ofDeleteAlternateImages = &H400      ' If set alternate images will be deleted.
        ofNoImageSizeCheck = &H800           ' Meaningful only if ofScaleImages is set. If set do not check whether the scaled image is smaller as the original image.
        ofIgnoreZeroLineWidth = &H1000       ' Meaningful only if the parameter MinLineWidth of the TOptimizeParams structure is greater zero.
        ' If set ignore line width operators with a value of zero (zero means one device unit).
        ofAdjZeroLineWidthOnly = &H2000      ' Meaningful only if the parameter MinLineWidth of the TOptimizeParams structure is greater zero.
        ' If set, change the line width of real hairlines only (a hairline is a one pixel width line -> LineWidth == 0).
        ofCompressWithJBIG2 = &H4000         ' If set, 1 bit images are compressed with JBIG2 if not already compressed with this filter.
        ofNoFilterCheck = &H8000             ' Meaningful only, if the flag ofCompressWithJBIG2 is set. If set, re-compress all 1 bit images, also if already compressed with JBIG2.
        ' This flag is mainly a debug flag to compare the compression ratio with other JBIG2 implementations.
        ofConvertGrayTo1Bit = &H10000        ' Useful for scanned faxes since many scanners create gray images for black & white input.
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Class TOptimizeParams
        Public StructSize As Integer               ' Must be set to sizeof(TOptimizeParams).

        Public Min1BitRes As Integer               ' Minimum resolution before scaling.
        Public MinGrayRes As Integer               ' Minimum resolution before scaling.
        Public MinColorRes As Integer              ' Minimum resolution before scaling.

        Public Res1BitImages As Integer            ' 1 bit black & white images.
        Public ResGrayImages As Integer            ' Gray images.
        Public ResColorImages As Integer           ' Multi-channel images.

        Public Filter1Bit As TCompressionFilter    ' Filter for black & white images.
        Public FilterGray As TCompressionFilter    ' Filter for gray images.
        Public FilterColor As TCompressionFilter   ' Filter for multi-channel images.
        Public JPEGQuality As Integer              ' JPEG quality.
        Public JP2KQuality As Integer              ' JPEG 2000 quality.
        Public MinLineWidth As Single              ' Zero means no hair line removal.
    End Class

    Public Enum TOrigin
        orDownLeft
        orTopLeft
    End Enum

    Public Enum TPageCoord
        pcBottomUp
        pcTopDown
    End Enum

    Public Enum TPageFormat
        pfDIN_A3
        pfDIN_A4
        pfDIN_A5
        pfDIN_B4
        pfDIN_B5
        pfDIN_B6
        pfDIN_C3
        pfDIN_C4
        pfDIN_C5
        pfDIN_C6
        pfDIN_C65
        pfDIN_DL
        pfDIN_E4
        pfDIN_E5
        pfDIN_E6
        pfDIN_E65
        pfDIN_M5
        pfDIN_M65
        pfUS_Legal
        pfUS_Letter
    End Enum

    Public Enum TPageBoundary
        pbArtBox
        pbBleedBox
        pbCropBox
        pbTrimBox
        pbMediaBox
    End Enum

    Public Enum TPageLabelFormat
        plfDecimalArabic    ' 1,2,3,4...
        plfUppercaseRoman   ' I,II,III,IV...
        plfLowercaseRoman   ' i,ii,iii,iv...
        plfUppercaseLetters ' A,B,C,D...
        plfLowercaseLetters ' a,b,c,d...
        plfNone
    End Enum

    Public Enum TPageLayout
        plSinglePage
        plOneColumn
        plTwoColumnLeft
        plTwoColumnRight
        plTwoPageLeft
        plTwoPageRight
        plDefault        ' Use viewer's default settings
    End Enum

    Public Enum TPageMode
        pmUseNone
        pmUseOutLines
        pmUseThumbs
        pmFullScreen
        pmUseOC           ' Optional Content (Layers)
        pmUseAttachments
    End Enum

    Public Enum TParseFlags
        pfNone = &H0
        pfDecomprAllImages = &H2 ' This flag causes that all image formats will be decompressed
        ' with the exception of JBIG2 compressed images. If this flag is
        ' not set, images which are already stored in a valid image file
        ' format are returned as is. This is the case for Gray and RGB JPEG
        ' images and for JPEG2000 images. If you want to extract the images
        ' of a PDF file this flag should NOT be set!

        pfNoJPXDecode = &H4      ' Meaningful only if the flag pfDecomprAllImages is set. If set,
        ' JPEG2000 images are returned as is so that you can use your own
        ' library to decompress such images since the the entire JPEG2000
        ' codec is still marked as experimental. If we find an alternative
        ' to the currently used Jasper library then we will replace it
        ' immediatly with another one...

        ' The following flags are ignored if an image is not decompressed. Note that only one flag must be set
        ' at time. If no color space conversion flag is set images are returned in their native or alternate
        ' device color space. Note that these flags do not convert colors of vector graphics and so on.
        ' Use the function ConvColor() to convert colors of the graphics state into a device space.
        pfDitherImagesToBW = &H8 ' Floyd-Steinberg dithering.
        pfConvImagesToGray = &H10
        pfConvImagesToRGB = &H20
        pfConvImagesToCMYK = &H40
        pfImageInfoOnly = &H80   ' If set, images are not decompressed. This flag is useful if you want
        ' to enumerate the images of a PDF file or if you want to determine how
        ' many images are stored in it.
        ' Note that images can be compressed with multiple filters. The member
        ' Filter of the structure TPDFImage contains only the last filter with
        ' which the image was compressed. There is no indication whether multiple
        ' decode filters are required to decompress the image buffer. So, it
        ' makes no sense to set this flag if you want to try to decompress the
        ' image buffer manually with your own decode filters.
    End Enum

    Public Enum TPathFillMode
        fmFillNoClose
        fmStrokeNoClose
        fmFillStrokeNoClose
        fmFill
        fmStroke
        fmFillStroke
        fmFillEvOdd
        fmFillStrokeEvOdd
        fmFillEvOddNoClose
        fmFillStrokeEvOddNoClose
        fmNoFill
        fmClose
    End Enum

    Public Enum TPatternType
        ptColored
        ptUnColored
        ptShadingPattern ' Cannot be created with DynaPDF
    End Enum

    ' The tags have the same meaning as the corresponding HTML tags.
    ' See PDF Reference for further information.
    Public Enum TPDFBaseTag
        btArt
        btArtifact
        btAnnot
        btBibEntry      ' BibEntry -> Bibliography entry
        btBlockQuote
        btCaption
        btCode
        btDiv
        btDocument
        btFigure
        btForm
        btFormula
        btH
        btH1
        btH2
        btH3
        btH4
        btH5
        btH6
        btIndex
        btLink
        btList          ' L
        btListElem      ' LI
        btListText      ' LBody
        btNote
        btP
        btPart
        btQuote
        btReference
        btSection       ' Sect
        btSpan
        btTable
        btTableDataCell ' TD
        btTableHeader   ' TH
        btTableRow      ' TR
        btTOC
        btTOCEntry      ' TOCI
    End Enum

    ' Notice:
    '   When using a bidirectional 8 bit code page the bidi algorithm is applied by default in Left to Right mode
    '   also if the bidi mode is set to bmNone (default). This mode produces identical results in comparison to
    '   applications like Edit or WordPad.
    '
    '   The Right to Left mode is available in applications which use Microsoft's Uniscribe, e.g. BabelPad. This
    '   mode works very well with the Reference Bidi Algorithm which is used by DynaPDF.
    '
    '   However , Uniscribe 's Left to Right mode produces different results in comparison to the Reference Bidi
    '   Algorithm. Because the bidi algorithm that is used in Uniscribe is not published it is practically
    '   impossible to get the same result in Left to Right mode without using this library.

    Public Enum TPDFBidiMode
        bmLeftToRight ' Apply the bidi algorithm on Unicode strings in Left to Right layout.
        bmRightToLeft ' Apply the bidi algorithm on Unicode strings in Right to Left layout.
        bmNone        ' Default -> not apply the bidi algorithm
    End Enum

    Public Enum TPDFColorSpace
        csDeviceRGB
        csDeviceCMYK
        csDeviceGray
    End Enum

    Public Enum TPDFDateTime
        dfMM_D
        dfM_D_YY
        dfMM_DD_YY
        dfMM_YY
        dfD_MMM
        dfD_MMM_YY
        dfDD_MMM_YY
        dfYY_MM_DD
        dfMMM_YY
        dfMMMM_YY
        dfMMM_D_YYYY
        dfMMMM_D_YYYY
        dfM_D_YY_H_MM_TT
        dfM_D_YY_HH_MM
        ' time formats
        df24HR_MM
        df12HR_MM
        df24HR_MM_SS
        df12HR_MM_SS
    End Enum

    Public Enum TPDFPrintFlags
        pffDefault = &H0 ' Gray printing
        pff1Bit = &H1 ' Black & White output
        pffColor = &H2 ' Color output
        pffAutoRotateAndCenter = &H4 ' Rotate and center the page if necessary
        pffPrintAsImage = &H8 ' Defined for future use
        pffShrinkToPrintArea = &H10 ' Scale the page so that it fits into the printable area
        pffNoStartDoc = &H20 ' If set StartDoc() of the Windows print API will not be called
        pffNoStartPage = &H40 ' If set StartPage() of the Windows print API will not be called
        pffNoEndDoc = &H80 ' If set EndDoc() of the Windows print API will not be called
        pffPrintPageAsIs = &H100 ' If set do not scale or rotate a page to fit into the printable area
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Class TPDFPrintParams
        Private StructSize As Integer            ' Must be set to sizeof(TPDFPrintParams).
        Public Compress As Integer               ' Meaningful only for image output. If true, compressed images are send to the printer.
        Public FilterGray As TCompressionFilter  ' Meaningful only for image output. Supported filters on Windows: cfFlate and cfJPEG.
        Public FilterColor As TCompressionFilter ' Meaningful only for image output. Supported filters on Windows: cfFlate and cfJPEG.
        Public JPEGQuality As Integer            ' JPEG Quality in percent. Zero == Default == 60.
        Public MaxRes As Integer                 ' Maximum print resolution. Zero == Default == 600 DPI.
        Public IgnoreDCSize As Integer           ' If true, PageSize is used to calculate the output format.
        Public PageSize As TIntRect              ' Considered only, if IgnoreDCSize is set to true.
        Public Sub New()
            MyBase.New()
            StructSize = Marshal.SizeOf(Me)
        End Sub
    End Class

    Public Enum TPDFVersion
        pvPDF_1_0 = 0
        pvPDF_1_1 = 1
        pvPDF_1_2 = 2
        pvPDF_1_3 = 3
        pvPDF_1_4 = 4
        pvPDF_1_5 = 5
        pvPDF_1_6 = 6
        pvPDF_1_7 = 7
        pvPDF_2_0 = 8      ' PDF 2.0
        pvReserved = 9     ' Reserved for future use
        pvPDFX1a_2001 = 10 ' PDF/X-1a:2001
        pvPDFX1a_2003 = 11 ' PDF/X-1a:2002
        pvPDFX3_2002 = 12  ' PDF/X-3:2002
        pvPDFX3_2003 = 13  ' PDF/X-3:2003
        pvPDFA_2005 = 14   ' PDF/A-1b 2005
        pvPDFX_4 = 15      ' PDF/X-4
        pvPDFA_1a = 16     ' PDF/A 1a 2005
        pvPDFA_2a = 17     ' PDF/A 2a
        pvPDFA_2b = 18     ' PDF/A 2b
        pvPDFA_2u = 19     ' PDF/A 2u
        pvPDFA_3a = 20     ' PDF/A 3a
        pvPDFA_3b = 21     ' PDF/A 3b
        pvPDFA_3u = 22     ' PDF/A 3u
        ' The following constants are flags which can be combined with pvPDFA_3a, pvPDFA_3b, and pvPDFA_3u.
        ' If used stand alone PDF/A 3b with the correspondig ZUGFeRD metadata will be created.
        pvZUGFeRD_Basic = &H10000
        pvZUGFeRD_Comfort = &H20000
        pvZUGFeRD_Extended = &H40000
        pvZUGFeRD_Mask = pvZUGFeRD_Basic Or pvZUGFeRD_Comfort Or pvZUGFeRD_Extended
    End Enum

    Public Enum TPrintScaling
        psAppDefault ' Default
        psNone
    End Enum

    Public Enum TProgType
        ptImportPage = 0
        ptWritePage = 1
        ptPrintPage = 2   ' Start printing the PDF file
    End Enum

    Public Enum TPwdType
        ptOpen = 0
        ptOwner = 1
        ptForceRepair = 2 ' Meaningful only when opening a PDF file with OpenImportFile() or OpenImportBuffer().
        ' If set, the PDF parser rebuilds the cross-reference table by scanning all the objects
        ' in the file. This can be useful if the cross-reference table contains damages while
        ' the top level objects are intact. Setting this flag makes only sence if the file
        ' was already previously opened in normal mode and if errors occured when importing
        ' pages of it.
        ptDontCopyBuf = 4 ' If set, OpenImportBuffer() does not copy the PDF buffer to an internal buffer. This
        ' increases the processing speed and reduces the memory usage. The PDF buffer must not
        ' be released until CloseImportFile() or CloseFile() was called.

    End Enum

    Public Enum TRawImageFlags
        rifByteAligned = &H1000
        rifRGBData = &H2000
        rifCMYKData = &H4000
    End Enum

    Public Enum TRenderingIntent
        riAbsoluteColorimetric
        riPerceptual
        riRelativeColorimetric
        riSaturation
        riNone
    End Enum

    Public Enum TReplaceImageFlags
        rifDefault = 0         ' Nothing special to do.
        rifDeleteAltImages = 1 ' Delete all alternate images that are associated with this image if any.
        rifDeleteMetadata = 2  ' Delete the meta data that was associated with the image.
        rifDeleteOCG = 4       ' Delete the Optional Content Group if any. Note that this changes the visibility state of the image. Normally, the OCG should be left as is.
        rifDeleteSoftMask = 8  ' An image can contain a soft mask that acts as an alpha channel. This mask can be deleted or left as is.
        ' The mask will always be deleted if the new image contains a soft mask or an alpha channel.
    End Enum

    Public Enum TRestrictions
        rsDenyNothing = 0
        rsDenyAll = 3900
        rsPrint = 4
        rsModify = 8
        rsCopyObj = 16
        rsAddObj = 32
        ' 128/256 bit encryption only -> these flags are ignored if 40 bit encryption is used
        rsFillInFormFields = 256
        rsExtractObj = 512
        rsAssemble = 1024
        rsPrintHighRes = 2048
        rsExlMetadata = 4096    ' PDF 1.5 Exclude metadata streams -> 128/256 bit encryption bit only.
        rsEmbFilesOnly = &H2000 ' PDF 1.6 Encrypt embedded files only -> Requires AES encryption.
    End Enum

    Public Enum TRubberStamp
        rsApproved
        rsAsIs
        rsConfidential
        rsDepartmental
        rsDraft
        rsExperimental
        rsExpired
        rsFinal
        rsForComment
        rsForPublicRelease
        rsNotApproved
        rsNotForPublicRelease
        rsSold
        rsTopSecret
    End Enum

    Public Enum TShadingType
        stUnknown       ' cannot occur -> internal use
        stFunctionBased
        stAxial
        stRadial
        stFreeFormGouraud
        stLatticeFormGouraud
        stCoonsPatch
        stTensorProduct
    End Enum

    Public Enum TSoftMaskType
        smtAlpha
        smtLuminosity
    End Enum

    Public Enum TTilingType
        ttConstSpacing
        ttNoDistortion
        ttFastConstSpacing
    End Enum

    Public Enum TTextAlign
        taLeft = 0
        taCenter = 1
        taRight = 2
        taJustify = 3
        taPlainText = &H10000000 ' If this flag is set alignment and command tags are interpreted as plain text.
        ' See WriteFText() in the help file for further information.
    End Enum

    Public Enum TSpoolConvFlags
        spcDefault = 0
        spcIgnorePaperFormat = 1  ' If set, the current page format is used as is for the entire spool file.
        spcDontAddMargins = 2     ' If set, the page format is calculated from the EMF files as is. The current page format is not used to calculate
        ' margins which are maybe required. Note that the parameters LeftMargin and TopMargin will still be considered.
        spcLoadSpoolFontsOnly = 4 ' If set, only embedded fonts will be loaded. The EMF files must be converted with the flag mfIgnoreEmbFonts in this
        ' case. This flag can be useful if you want to use your own code to convert the EMF files of the spool file.
        spcFlushPages = 8         ' If set, the function writes every finish page directly to the output file to reduce the memory usage. This flag
        ' is meaningful only if the PDF file is not created in memory. Note also that it is not possible to access already
        ' flushed pages again with EditPage().
    End Enum

    Public Enum TStdPattern
        spHorizontal ' -----
        spVertical ' |||||
        spRDiagonal ' \\\\\
        spLDiagonal ' /////
        spCross ' +++++
        spDiaCross ' xxxxx
    End Enum

    Public Enum TSubmitFlags
        sfNone = &H0
        sfExlude = &H1 ' If set, fields in a submit or reset form action are excluded
        sfInclNoValFields = &H2
        sfHTML = &H4
        sfGetMethod = &H8
        sfSubmCoords = &H10
        sfXML = &H24
        sfInclAppSaves = &H40
        sfInclAnnots = &H80
        sfPDF = &H100
        sfCanonicalFormat = &H200
        sfExlNonUserAnnots = &H400
        sfExlFKey = &H800
        sfEmbedForm = &H2000 ' PDF 1.5 embed the entire form into a file stream inside the FDF file -> requires
        ' the full version of Adobe's Acrobat
    End Enum

    Public Enum TTextExtractionFlags
        tefDefault = 0               ' Create text lines in the original order.
        tefSortTextX = 1             ' Sort text records in x-direction.
        tefSortTextY = 2             ' Sort text records in y-direction.
        tefSortTextXY = tefSortTextX Or tefSortTextY
        tefDeleteOverlappingText = 4 ' Text extraction only.
    End Enum

    Public Enum TUnicodeRange1
        urBasicLatin = &H1                         ' 0000-007F
        urLatin1Supplement = &H2                   ' 0080-00FF
        urLatinExtendedA = &H4                     ' 0100-017F
        urLatinExtendedB = &H8                     ' 0180-024F
        urIPAExtensions = &H10                     ' 0250-02AF, 1D00-1D7F, 1D80-1DBF
        urSpacingModifierLetters = &H20            ' 02B0-02FF, A700-A71F
        urCombiningDiacriticalMarks = &H40         ' 0300-036F, 1DC0-1DFF
        urGreekandCoptic = &H80                    ' 0370-03FF
        urCoptic = &H100                           ' 2C80-2CFF
        urCyrillic = &H200                         ' 0400-04FF, 0500-052F, 2DE0-2DFF, A640-A69F
        urArmenian = &H400                         ' 0530-058F
        urHebrew = &H800                           ' 0590-05FF
        urVai = &H1000                             ' A500-A63F
        urArabic = &H2000                          ' 0600-06FF, 0750-077F
        urNKo = &H4000                             ' 07C0-07FF
        urDevanagari = &H8000                      ' 0900-097F
        urBengali = &H10000                        ' 0980-09FF
        urGurmukhi = &H20000                       ' 0A00-0A7F
        urGujarati = &H40000                       ' 0A80-0AFF
        urOriya = &H80000                          ' 0B00-0B7F
        urTamil = &H100000                         ' 0B80-0BFF
        urTelugu = &H200000                        ' 0C00-0C7F
        urKannada = &H400000                       ' 0C80-0CFF
        urMalayalam = &H800000                     ' 0D00-0D7F
        urThai = &H1000000                         ' 0E00-0E7F
        urLao = &H2000000                          ' 0E80-0EFF
        urGeorgian = &H4000000                     ' 10A0-10FF, 2D00-2D2F
        urBalinese = &H8000000                     ' 1B00-1B7F
        urHangulJamo = &H10000000                  ' 1100-11FF
        urLatinExtendedAdditional = &H20000000     ' 1E00-1EFF, 2C60-2C7F, A720-A7FF
        urGreekExtended = &H40000000               ' 1F00-1FFF
        urGeneralPunctuation = &H80000000          ' 2000-206F, 2E00-2E7F
    End Enum

    Public Enum TUnicodeRange2
        urSuperscriptsAndSubscripts = &H1          ' 2070-209F
        urCurrencySymbols = &H2                    ' 20A0-20CF
        urCombDiacritMarksForSymbols = &H4         ' 20D0-20FF
        urLetterlikeSymbols = &H8                  ' 2100-214F
        urNumberForms = &H10                       ' 2150-218F
        urArrows = &H20                            ' 2190-21FF, 27F0-27FF, 2900-297F, 2B00-2BFF
        urMathematicalOperators = &H40             ' 2200-22FF, 2A00-2AFF, 27C0-27EF, 2980-29FF
        urMiscellaneousTechnical = &H80            ' 2300-23FF
        urControlPictures = &H100                  ' 2400-243F
        urOpticalCharacterRecognition = &H200      ' 2440-245F
        urEnclosedAlphanumerics = &H400            ' 2460-24FF
        urBoxDrawing = &H800                       ' 2500-257F
        urBlockElements = &H1000                   ' 2580-259F
        urGeometricShapes = &H2000                 ' 25A0-25FF
        urMiscellaneousSymbols = &H4000            ' 2600-26FF
        urDingbats = &H8000                        ' 2700-27BF
        urCJKSymbolsAndPunctuation = &H10000       ' 3000-303F
        urHiragana = &H20000                       ' 3040-309F
        urKatakana = &H40000                       ' 30A0-30FF, 31F0-31FF
        urBopomofo = &H80000                       ' 3100-312F, 31A0-31BF
        urHangulCompatibilityJamo = &H100000       ' 3130-318F
        urPhagsPa = &H200000                       ' A840-A87F
        urEnclosedCJKLettersAndMonths = &H400000   ' 3200-32FF
        urCJKCompatibility = &H800000              ' 3300-33FF
        urHangulSyllables = &H1000000              ' AC00-D7AF
        urNonPlane0 = &H2000000                    ' D800-DFFF
        urPhoenician = &H4000000                   ' 10900-1091F
        urCJKUnifiedIdeographs = &H8000000         ' 4E00-9FFF, 2E80-2EFF, 2F00-2FDF, 2FF0-2FFF, 3400-4DBF, 20000-2A6DF, 3190-319F
        urPrivateUseAreaPlane0 = &H10000000        ' E000-F8FF
        urCJKStrokes = &H20000000                  ' 31C0-31EF, F900-FAFF, 2F800-2FA1F
        urAlphabeticPresentationForms = &H40000000 ' FB00-FB4F
        urArabicPresentationFormsA = &H80000000    ' FB50-FDFF
    End Enum

    Public Enum TUnicodeRange3
        urCombiningHalfMarks = &H1                 ' FE20-FE2F
        urVerticalForms = &H2                      ' FE10-FE1F, FE30-FE4F
        urSmallFormVariants = &H4                  ' FE50-FE6F
        urArabicPresentationFormsB = &H8           ' FE70-FEFF
        urHalfwidthAndFullwidthForms = &H10        ' FF00-FFEF
        urSpecials = &H20                          ' FFF0-FFFF
        urTibetan = &H40                           ' 0F00-0FFF
        urSyriac = &H80                            ' 0700-074F
        urThaana = &H100                           ' 0780-07BF
        urSinhala = &H200                          ' 0D80-0DFF
        urMyanmar = &H400                          ' 1000-109F
        urEthiopic = &H800                         ' 1200-137F, 1380-139F, 2D80-2DDF
        urCherokee = &H1000                        ' 13A0-13FF
        urUnifiedCanadianAboriginal = &H2000       ' 1400-167F
        urOgham = &H4000                           ' 1680-169F
        urRunic = &H8000                           ' 16A0-16FF
        urKhmer = &H10000                          ' 1780-17FF, 19E0-19FF
        urMongolian = &H20000                      ' 1800-18AF
        urBraillePatterns = &H40000                ' 2800-28FF
        urYiSyllables = &H80000                    ' A000-A48F, A490-A4CF
        urTagalog = &H100000                       ' 1700-171F, 1720-173F, 1740-175F, 1760-177F
        urOldItalic = &H200000                     ' 10300-1032F
        urGothic = &H400000                        ' 10330-1034F
        urDeseret = &H800000                       ' 10400-1044F
        urMusicalSymbols = &H1000000               ' 1D000-1D0FF, 1D100-1D1FF, 1D200-1D24F
        urMathematicalAlphanumeric = &H2000000     ' 1D400-1D7FF
        urPrivateUsePlane15 = &H4000000            ' FF000-FFFFD, 100000-10FFFD
        urVariationSelectors = &H8000000           ' FE00-FE0F, E0100-E01EF
        urTags = &H10000000                        ' E0000-E007F
        urLimbu = &H20000000                       ' 1900-194F
        urTaiLe = &H40000000                       ' 1950-197F
        urNewTaiLue = &H80000000                   ' 1980-19DF
    End Enum

    Public Enum TUnicodeRange4
        urBuginese = &H1                           ' 1A00-1A1F
        urGlagolitic = &H2                         ' 2C00-2C5F
        urTifinagh = &H4                           ' 2D30-2D7F
        urYijingHexagramSymbols = &H8              ' 4DC0-4DFF
        urSylotiNagri = &H10                       ' A800-A82F
        urLinearBSyllabary = &H20                  ' 10000-1007F, 10080-100FF, 10100-1013F
        urAncientGreekNumbers = &H40               ' 10140-1018F
        urUgaritic = &H80                          ' 10380-1039F
        urOldPersian = &H100                       ' 103A0-103DF
        urShavian = &H200                          ' 10450-1047F
        urOsmanya = &H400                          ' 10480-104AF
        urCypriotSyllabary = &H800                 ' 10800-1083F
        urKharoshthi = &H1000                      ' 10A00-10A5F
        urTaiXuanJingSymbols = &H2000              ' 1D300-1D35F
        urCuneiform = &H4000                       ' 12000-123FF, 12400-1247F
        urCountingRodNumerals = &H8000             ' 1D360-1D37F
        urSundanese = &H10000                      ' 1B80-1BBF
        urLepcha = &H20000                         ' 1C00-1C4F
        urOlChiki = &H40000                        ' 1C50-1C7F
        urSaurashtra = &H80000                     ' A880-A8DF
        urKayahLi = &H100000                       ' A900-A92F
        urRejang = &H200000                        ' A930-A95F
        urCham = &H400000                          ' AA00-AA5F
        urAncientSymbols = &H800000                ' 10190-101CF
        urPhaistosDisc = &H1000000                 ' 101D0-101FF
        urCarian = &H2000000                       ' 102A0-102DF, 10280-1029F, 10920-1093F
        urDominoTiles = &H4000000                  ' 1F030-1F09F, 1F000-1F02F
    End Enum

    Public Enum TViewerPreference
        vpUseNone = &H0
        vpHideToolBar = &H1
        vpHideMenuBar = &H2
        vpHideWindowUI = &H4
        vpFitWindow = &H8
        vpCenterWindow = &H10
        vpDisplayDocTitle = &H20
        vpNonFullScrPageMode = &H40
        vpDirection = &H80
        vpViewArea = &H100
        vpViewClip = &H200
        vpPrintArea = &H400
        vpPrintClip = &H800
    End Enum

    Public Enum TViewPrefAddVal
        avNone = &H0
        avNonFullScrUseNone = &H1
        avNonFullScrUseOutlines = &H2
        avNonFullScrUseThumbs = &H4
        avNonFullScrUseOC = &H400 ' PDF 1.6
        avDirection2R = &H8
        avDirectionR2 = &H10
        avViewPrintArtBox = &H20
        avViewPrintBleedBox = &H40
        avViewPrintCropBox = &H80
        avViewPrintMediaBox = &H100
        avViewPrintTrimBox = &H200
    End Enum

    Public Structure TPDFBarcode
        Dim Caption As String      ' Optional
        Dim ECC As Single          ' 0..8 for PDF417, or 0..3 for QRCode
        Dim Height As Single       ' Height in inches
        Dim nCodeWordCol As Single ' Required for PDF417. The number of codewords per barcode coloumn.
        Dim nCodeWordRow As Single ' Required for PDF417. The number of codewords per barcode row.
        Dim Resolution As Integer  ' Required -> Should be 300
        Dim Symbology As String    ' PDF417, QRCode, or DataMatrix.
        Dim Version As Single      ' Should be 1
        Dim Width As Single        ' Width in inches
        Dim XSymHeight As Single   ' Only needed for PDF417. The vertical distance between two barcode modules,
        ' measured in pixels. The ratio XSymHeight/XSymWidth shall be an integer
        ' value. For PDF417, the acceptable ratio range is from 1 to 4. For QRCode
        ' and DataMatrix, this ratio shall always be 1.
        Dim XSymWidth As Single    ' Required -> The horizontal distance, in pixels, between two barcode modules.
    End Structure

    Public Structure TPDFBitmap
        Dim StructSize As Integer ' Must be set to sizeof(TPDFBitmap)
        Dim Buffer As IntPtr      ' Image buffer
        Dim BufSize As Integer    ' Buffer size in bytes
        Dim DestX As Integer      ' Destination x-coordinate on the main image (the rendered page)
        Dim DestY As Integer      ' Destination y-coordinate on the main image (the rendered page)
        Dim Height As Integer     ' Image height in pixels
        Dim Stride As Integer     ' Scanline length in bytes
        Dim Width As Integer      ' Image width
    End Structure

    Public Enum TBmkStyle
        bmsNormal = 0
        bmsItalic = 1
        bmsBold = 2
    End Enum

    Public Structure TBookmark
        Dim Color As Integer
        Dim DestPage As Integer
        Dim DestPos As TPDFRect
        Dim DestType As TDestType
        Dim DoOpen As Boolean
        Dim Parent As Integer
        Dim Style As TBmkStyle
        Dim Title As String
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TCTM
        Dim a As Double
        Dim b As Double
        Dim c As Double
        Dim d As Double
        Dim x As Double
        Dim y As Double
    End Structure

    Public Structure TDeviceNAttributes
        Dim IProcessColorSpace As IntPtr  ' Pointer to process color space or NULL -> GetColorSpaceEx().
        Dim ProcessColorants() As String  ' Process colorant names
        Dim Separations() As IntPtr       ' Optional pointers to Separation color spaces -> GetColorSpaceEx().
        Dim IMixingHints As IntPtr        ' Optional pointer to mixing hints. There is no API function at this time to access mixing hints.
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TRectL
        Dim rLeft As Integer
        Dim rTop As Integer
        Dim rRight As Integer
        Dim rBottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TTextRecordA
        Dim Advance As Single  ' Negative values move the cursor to right, positive to left. The value is measured in text space!
        Dim Text As IntPtr     ' Source string (not null-terminated)
        Dim Length As Integer  ' Length in characters
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TTextRecordW
        Dim Advance As Single  ' Negative values move the cursor to right, positive to left. The value is measured in text space!
        Dim Text As IntPtr     ' Already translated Unicode string (not null-terminated)
        Dim Length As Integer  ' Length in characters
        Dim Width As Single    ' String width measured in text space
    End Structure

    Public Structure TPDFAnnotation
        Dim AnnotType As TAnnotType
        Dim Deleted As Boolean
        Dim BBox As TPDFRect
        Dim BorderWidth As Double
        Dim BorderColor As Integer
        Dim BorderStyle As TBorderStyle
        Dim BackColor As Integer
        Dim Handle As Integer
        Dim Author As String
        Dim Content As String
        Dim Name As String
        Dim Subject As String
        Dim PageNum As Integer
        Dim HighlightMode As THighlightMode
    End Structure

    Public Structure TPDFAnnotationEx
        Dim AnnotType As TAnnotType
        Dim Deleted As Boolean
        Dim BBox As TPDFRect
        Dim BorderWidth As Single
        Dim BorderColor As Integer
        Dim BorderStyle As TBorderStyle
        Dim BackColor As Integer
        Dim Handle As Integer
        Dim Author As String
        Dim Content As String
        Dim Name As String
        Dim Subject As String
        Dim PageNum As Integer
        Dim HighlightMode As THighlightMode
        ' Page link annotations only
        Dim DestPage As Integer
        Dim DestPos As TPDFRect
        Dim DestType As TDestType
        Dim DestFile As String         ' File link or web link annotations
        Dim Icon As Integer            ' The Icon type depends on the annotation type. If the annotation type is atText then the Icon
        ' is of type TAnnotIcon. If the annotation type is atFileAttach then it is of type
        ' TFileAttachIcon. If the annotation type is atStamp then the Icon is the stamp type (TRubberStamp).
        ' For any other annotation type this value is not set (-1).
        Dim StampName As String        ' Set only, if Icon == rsUserDefined
        Dim AnnotFlags As Integer      ' See TAnnotFlags for available flags
        Dim CreateDate As String       ' Creation Date -> Optional
        Dim ModDate As String          ' Modification Date -> Optional
        Dim Grouped As Boolean         ' (Reply type) Meaningful only if Parent != -1 and Type != atPopUp. If true,
        ' the annotation is part of an annotation group. Properties like Content, CreateDate,
        ' ModDate, BackColor, Subject, and Open must be taken from the parent annotation.
        Dim Open As Boolean            ' Meaningful only for annotations which have a corresponding PopUp annotation.
        Dim Parent As Integer          ' Parent annotation handle of a PopUp Annotation or the parent annotation if
        ' this annotation represents a state of a base annotation. In this case,
        ' the annotation type is always atText and only the following members should
        ' be considered:
        '    State      // The current state
        '    StateModel // Marked, Review, and so on
        '    CreateDate // Creation Date
        '    ModDate    // Modification Date
        '    Author     // The user who has set the state
        '    Content    // Not displayed in Adobe's Acrobat...
        '    Subject    // Not displayed in Adobe's Acrobat...
        ' The PopUp annotation of a text annotation which represent an Annotation State
        ' must be ignored.
        Dim PopUp As Integer           ' Handle of the corresponding PopUp annotation if any.
        Dim State As String            ' The state of the annotation.
        Dim StateModel As String       ' The state model (Marked, Review, and so on).
        Dim EmbeddedFile As Integer    ' FileAttach annotations only. A handle of an embedded file -> GetEmbeddedFile().
        Dim Subtype As String          ' Set only, if Type = atUnknownAnnot
        Dim PageIndex As Integer       ' The page index is used to sort form fields. See SortFieldsByIndex().
        Dim MarkupAnnot As Boolean     ' If true, the annotation is a markup annotation. Markup annotations can be flattened
        ' separately, see FlattenAnnots().
        Dim Opacity As Single          ' Opacity = 1.0 = Opaque, Opacity < 1.0 = Transparent, Markup annotations only
        Dim QuadPoints() As Single     ' Highlight, Link, and Redact annotations only. The array contains the raw floating point values.
        ' Since a quadpoint requires always four coordinate pairs, the number of QuadPoints is QuadPointsCount divided by 8.

        Dim DashPattern() As Single    ' Only present if BorderStyle == bsDashed

        Dim Intent As String           ' Markup annotations only. The intent allows to distinguish between different uses of an annotation.
        ' For example, line annotations have two intents: LineArrow and LineDimension.
        Dim LE1 As TLineEndStyle       ' Line end style of the start point -> Line and PolyLine annotations only
        Dim LE2 As TLineEndStyle       ' Line end style of the end point -> Line and PolyLine annotations only
        Dim Vertices() As Single       ' Line, PolyLine, and Polygon annotations only. The array contains the raw floating point values.
        ' Since a vertice requires always two coordinate pairs, the number of vertices
        ' or points is VerticeCount divided by 2.

        ' Line annotations only. These properties should only be considered if the member Intent is set to the string LineDimension.
        Dim Caption As Boolean            ' If true, the annotation string Content is used as caption.
        Dim CaptionOffsetX As Single      ' Horizontal offset of the caption from its normal position
        Dim CaptionOffsetY As Single      ' Vertical offset of the caption from its normal position
        Dim CaptionPos As TLineCaptionPos ' The position where the caption should be drawn if present
        Dim LeaderLineLen As Single       ' Length of the leader lines (positive or negative)
        Dim LeaderLineExtend As Single    ' Optional leader line extend beyond the leader line (must be a positive value or zero)
        Dim LeaderLineOffset As Single    ' Amount of space between the endpoints of the annotation and the leader lines (must be a positive value or zero)

        Dim BorderEffect As TBorderEffect ' Circle, Square, FreeText, and Polygon annotations.
        Dim InkList() As IntPtr           ' Ink annotations only. Array of array. The sub arrays can be accessed with GetInkList().
        Dim RichStyle As String           ' Optional default style string.      -> FreeText annotations only.
        Dim RichText As String            ' Optional rich text string (RC key). -> Markup annotations only.
        Dim OC As Integer                 ' Handle of an OCG or OCMD or -1 if not set. See help file for further information.

        Dim RD() As Single                ' Caret, Circle, Square, and FreeText annotations.
        Dim Rotate As Integer             ' Caret annotations only. Must be zero or a multiple of 90. This key is not documented in the specs.
    End Structure

    Public Structure TPDFChoiceValue
        Dim ExpValue As String
        Dim Value As String
        Dim Selected As Boolean
    End Structure

    ' The structure contains several duplicate fields because CMap files contain usually a DSC comment
    ' section which provides Postscript specific initialization code. With exception of DSCResName the
    ' strings in the DSC section should not differ from their CMap counterparts. The Identity mapping
    ' of a character collection should contain the DSC comment "%%BeginResource: CMap (Identity)".
    ' Otherwise the string should be set to the CMap name.

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFCMap
        Dim StructSize As Integer     ' Must be set to sizeof(TPDFCMap) before calling GetCMap()!
        <MarshalAs(UnmanagedType.LPStr)>
        Dim BaseCMap As String        ' If set, this base cmap is required when loading the cmap.
        Dim CIDCount As Integer       ' 0 if not set.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim CMapName As String        ' The CMap name.
        Dim CMapType As Integer       ' Should be 1!
        Dim CMapVersion As Single     ' The CMap version.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim DSCBaseCMap As String     ' DSC comment.
        Dim DSCCMapVersion As Single  ' DSC comment.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim DSCResName As String      ' DSC comment. If the CMap uses an Identity mapping this string should be set to Identity.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim DSCTitle As String        ' DSC comment -> DSC CMap name + Registry + Ordering + Supplement, e.g. "GB-EUC-H Adobe GB1 0"
        <MarshalAs(UnmanagedType.LPStr)>
        Dim FileNameA As String       ' The file name and CMap name should be identical!
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim FileNameW As String       ' The file name and CMap name should be identical!
        <MarshalAs(UnmanagedType.LPStr)>
        Dim FilePathA As String       ' The Ansi string is set if the Ansi version of SetCMapDir() was used.
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim FilePathW As String       ' The Unicode string is set if the Unicode version of SetCMapDir() was used.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim Ordering As String        ' CIDSystemInfo -> The Character Collection, e.g. Japan1.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim Registry As String        ' CIDSystemInfo -> The registrant of the Character Collection is usually Adobe.
        Dim Supplement As Integer     ' CIDSystemInfo -> The Supplement number should be supported in the used PDF Version.
        Dim WritingMode As Integer    ' 0 == Horizontal, 1 == Vertical
    End Structure

    Public Structure TPDFColorSpaceObj
        Dim ColorSpaceType As TExtColorSpace
        Dim Alternate As TExtColorSpace ' Alternate or base color space of an Indexed or Pattern color space.
        Dim IAlternate As IntPtr             ' Only set if the color space contains an alternate or base color space -> GetColorSpaceObjEx().
        Dim Buffer As IntPtr                 ' Contains either an ICC profile or the color table of an Indexed color space.
        Dim BufSize As Integer               ' Buffer length in bytes.
        Dim BlackPoint() As Single           ' CIE blackpoint. If set, the array contains exactly 3 values.
        Dim WhitePoint() As Single           ' CIE whitepoint. If set, the array contains exactly 3 values.
        Dim Gamma() As Single                ' If set, one value per component.
        Dim Range() As Single                ' The allowed range of input values (min/max for each component).
        Dim Matrix() As Single               ' XYZ matrix. If set, the array contains exactly 9 values.
        Dim NumInComponents As Integer       ' Number of input components.
        Dim NumOutComponents As Integer      ' Number of output components.
        Dim NumColors As Integer             ' HiVal + 1 as specified in the color space. Indexed color space only.
        Dim Colorants() As String            ' Colorant names (Separation, DeviceN, and NChannel only).
        Dim Metadata As IntPtr               ' Optional XMP metadata stream -> ICCBased only.
        Dim MetadataSize As Integer          ' Metadata length in bytes.
        Dim IFunction As IntPtr              ' Pointer to function object -> Separation, DeviceN, and NChannel only.
        Dim IAttributes As IntPtr            ' Optional attributes of DeviceN or NChannel color spaces -> GetNChannelAttributes().
        Dim IColorSpaceObj As IntPtr         ' Pointer of the corresponding color space object
    End Structure

    Public Structure TPDFEmbFileNode
        Dim Name As String     ' UTF-8 encoded name. This key contains usually a 7 bit ASCII string.
        Dim EF As TPDFFileSpec ' Embedded file.
        Dim NextNode As IntPtr ' Next node if any.
    End Structure

    Public Structure TPDFError
        Dim Message As String  ' The error message
        Dim ObjNum As Integer  ' -1 if not available
        Dim Offset As Integer  ' -1 if not available
        Dim SrcFile As String  ' Source file
        Dim SrcLine As Integer ' Source line
    End Structure

    ' It is not possible to set all available graphic state parameters with DynaPDF, such as black generation functions
    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFExtGState
        Dim AutoStrokeAdjust As Integer              ' PDF_MAX_INT if not set
        Dim BlendMode As TBlendMode             ' Default bmNotSet
        Dim FlatnessTol As Single                    ' -1.0 if not set
        Dim OverPrintFill As Integer                 ' PDF_MAX_INT if not set
        Dim OverPrintStroke As Integer               ' PDF_MAX_INT if not set
        Dim OverPrintMode As Integer                 ' PDF_MAX_INT if not set
        Dim RenderingIntent As TRenderingIntent ' riNone if not set
        Dim SmoothnessTol As Single                  ' -1.0 if not set
        Dim FillAlpha As Single                      ' -1.0 if not set
        Dim StrokeAlpha As Single                    ' -1.0 if not set
        Dim AlphaIsShape As Integer                  ' PDF_MAX_INT if not set
        Dim TextKnockout As Integer                  ' PDF_MAX_INT if not set
        Dim SoftMaskNone As Integer                  ' Can be set to true to disable the active soft mask
        Dim SoftMask As IntPtr                       ' Soft mask pointer or NULL. See CreateSoftMask() for further information.
        Private Reserved1 As IntPtr
        Private Reserved2 As IntPtr
        Private Reserved3 As IntPtr
        Private Reserved4 As IntPtr
        Private Reserved5 As IntPtr
        Private Reserved6 As IntPtr
        Private Reserved7 As IntPtr
    End Structure

    ' Extended graphics state dictionary
    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFExtGState2
        Dim AlphaIsShape As Integer       ' PDF_MAX_INT if not set
        Dim AutoStrokeAdjust As Integer   ' PDF_MAX_INT if not set
        Dim BlackGen As IntPtr            ' Function handle or NULL -> GetFunction()
        Dim BlackGen2 As IntPtr           ' Function handle or NULL -> GetFunction()
        Dim BlendMode As IntPtr           ' Array of blend modes
        Dim BlendModeCount As Integer     ' Number of blend modes
        Dim FillAlpha As Single           ' -1.0 if not set
        Dim FlatnessTol As Single         ' -1.0 if not set
        Dim Halftone As IntPtr            ' Halftone handle or NULL -> GetHalftoneDict()
        Dim OverPrintFill As Integer      ' PDF_MAX_INT if not set
        Dim OverPrintStroke As Integer    ' PDF_MAX_INT if not set
        Dim OverPrintMode As Integer      ' PDF_MAX_INT if not set
        Dim RenderingIntent As TRenderingIntent ' riNone if not set
        Dim SmoothnessTol As Single       ' -1.0 if not set
        Dim SoftMask As IntPtr            ' Soft mask handle or NULL
        Dim StrokeAlpha As Single         ' -1.0 if not set
        Dim TextKnockout As Integer       ' PDF_MAX_INT if not set
        Dim TransferFunc As IntPtr        ' Array of functions -> GetFunction()
        Dim TransferFuncCount As Integer  ' Number of transfer functions
        Dim TransferFunc2 As IntPtr       ' Array of functions -> GetFunction()
        Dim TransferFunc2Count As Integer ' Number of transfer functions
        Dim UnderColorRem As IntPtr       ' Function handle or NULL -> GetFunction()
        Dim UnderColorRem2 As IntPtr      ' Function handle or NULL -> GetFunction()
        Private Reserved1 As IntPtr
        Private Reserved2 As IntPtr
        Private Reserved3 As IntPtr
        Private Reserved4 As IntPtr
        Public Function GetBlendMode() As TBlendMode
            If BlendModeCount = 0 Then
                Return TBlendMode.bmNormal
            Else
                Dim modes() As TBlendMode
                ReDim modes(BlendModeCount - 1)
                Marshal.Copy(BlendMode, modes, 0, BlendModeCount)
                Return modes(0)
            End If
        End Function
    End Structure

    Public Structure TPDFField
        Dim BBox As TPDFRect
        Dim FieldType As TFieldType
        Dim Deleted As Boolean
        Dim Handle As Integer
        Dim FieldName As String
        Dim BackCS As TPDFColorSpace
        Dim TextCS As TPDFColorSpace
        Dim BackColor As Integer
        Dim BorderColor As Integer
        Dim TextColor As Integer
        Dim Checked As Boolean
        Dim Parent As Integer
        Dim KidCount As Integer
        Dim FontName As String
        Dim FontSize As Double
        Dim Value As String
        Dim ToolTip As String
    End Structure

    Public Structure TPDFFieldEx
        Dim Deleted As Boolean               ' If true, the field was marked as deleted by DeleteField()
        Dim BBox As TPDFRect                 ' Bounding box of the field in bottom-up coordinates
        Dim FieldType As TFieldType          ' Field type
        Dim GroupType As TFieldType          ' If GroupType != FieldType the field is a terminal field of a field group
        Dim Handle As Integer                ' Field handle
        Dim BackColor As Integer             ' Background color
        Dim BackColorSP As TExtColorSpace    ' Color space of the background color
        Dim BorderColor As Integer           ' Border color
        Dim BorderColorSP As TExtColorSpace  ' Color space of the border color
        Dim BorderStyle As TBorderStyle      ' Border style
        Dim BorderWidth As Single            ' Border width
        Dim CharSpacing As Single            ' Text fields only
        Dim Checked As Boolean               ' Check boxes only
        Dim CheckBoxChar As Integer          ' ZapfDingbats character that is used to display the on state
        Dim DefState As TCheckBoxState       ' Check boxes only
        Dim DefValue As String               ' Optional default value
        Dim IEditFont As IntPtr              ' Pointer to default editing font
        Dim EditFont As String               ' Postscript name of the editing font
        Dim ExpValCount As Integer           ' Combo and list boxes only. The values can be accessed with GetFieldExpValueEx()
        Dim ExpValue As String               ' Check boxes only
        Dim FieldFlags As TFieldFlags        ' Field flags
        Dim IFieldFont As IntPtr             ' Pointer to the font that is used by the field
        Dim FieldFont As String              ' Postscript name of the font
        Dim FontSize As Double               ' Font size. 0.0 means auto font size
        Dim FieldName As String              ' Note that the children of a field group or radio button have no name
        Dim HighlightMode As THighlightMode  ' Highlight mode
        Dim IsCalcField As Boolean           ' If true, the OnCalc event of the field is connected with a JavaScript action
        Dim MapName As String                ' Optional unique mapping name of the field
        Dim MaxLen As Integer                ' Text fields only -> zero means not restricted
        Dim Kids() As IntPtr                 ' Array of child fields -> GetFieldEx2()
        Dim Parent As IntPtr                 ' Pointer to parent field or NULL
        Dim PageNum As Integer               ' Page on which the field is used or -1
        Dim Rotate As Integer                ' Rotation angle in degrees
        Dim TextAlign As TTextAlign          ' Text fields only
        Dim TextColor As Integer             ' Text color
        Dim TextColorSP As TExtColorSpace    ' Color space of the field's text
        Dim TextScaling As Single            ' Text fields only
        Dim ToolTip As String                ' Optional tool tip
        Dim UniqueName As String             ' Optional unique name (NM key)
        Dim Value As String                  ' Field value
        Dim WordSpacing As Single            ' Text fields only
        Dim PageIndex As Integer             ' Array index to change the tab order, see SortFieldsByIndex().
        Dim IBarcode As IntPtr               ' If present, this field is a barcode field. The field type is set to ftText
        ' since barcode fields are extended text fields. -> GetBarcodeDict().
        Dim ISignature As IntPtr             ' Signature fields only. Present only for imported signature fields which
        ' which have a value. That means the file was digitally signed. -> GetSigDict().
        ' Signed signature fields are always marked as deleted!
        Dim ModDate As String                ' Last modification date (optional)

        ' Push buttons only. The down and roll over states are optional. If not present, then all states use the up state.
        ' The handles of the up, down, and roll over states are template handles! The templates can be opened for editing
        ' with EditTemplate2() and parsed with ParseContent().
        Dim CaptionPos As TBtnCaptionPos     ' Where to position the caption relative to its image
        Dim DownCaption As String            ' Caption of the down state
        Dim DownImage As Integer             ' Image of the down state
        Dim RollCaption As String            ' Caption of the roll over state
        Dim RollImage As Integer             ' Image of the roll over state
        Dim UpCaption As String              ' Caption of the up state
        Dim UpImage As Integer               ' Image of the up state -> if > -1, the button is an image button
        Dim OC As Integer                    ' Handle of an OCG or OCMD or -1 if not set. See help file for further information.
        Dim Action As Integer                ' Action handle or -1 if not set. This action is executed when the field is activated.
        Dim ActionType As TActionType        ' Meaningful only, if Action >= 0.
        Dim Events As Integer                ' See GetObjEvent() if set.
    End Structure

    Public Structure TPDFFileSpec
        Dim Buffer() As Byte       ' Buffer of an embedded file
        Dim Compressed As Boolean  ' Should be false if Decompress was true in the GetEmbeddedFile() call, otherwise usually true.
        ' DynaPDF decompresses Flate encoded streams only. Other filters can occur but this is very unusual.
        Dim ColItem As IntPtr      ' If != NULL the embedded file contains a collection item with user defined data. This entry
        ' can occur in PDF Collections only (PDF 1.7). See "PDF Collections" in the help file for further
        ' information.
        Dim Name As String         ' Name of the file specification in the name tree. This value is always present.
        Dim FileName As String     ' File name as 7 bit ASCII string.
        Dim IsURL As Boolean       ' If true, FileName contains a URL.
        Dim UF As String           ' PDF 1.7. Same as FileName but Unicode is allowed.
        Dim Desc As String         ' Description
        Dim FileSize As Integer    ' Size of the decompressed stream or zero if not known. Note: this is either the Size key of
        ' the Params dictionary if present or the DL key in the file stream. Whether this value is
        ' correct depends on the file creator! The parameter is definitely correct if the file was
        ' decompressed.
        Dim MIMEType As String     ' MIME media type name as defined in Internet RFC 2046.
        Dim CreateDate As String   ' Creation date as string. See help file "The standard date format".
        Dim ModDate As String      ' Modification date as string. See help file "The standard date format".
        Dim CheckSum() As Byte     ' 16 byte MD5 digest. Note that this is a binary string. It is exactly 16 bytes long if set!
    End Structure

    Public Structure TPDFFileSpecEx
        Dim AFRelationship As String ' PDF 2.0
        Dim ColItem As IntPtr        ' If != NULL the embedded file contains a collection item with user defined data. This entry can
        ' occur in PDF Collections only (PDF 1.7). See "PDF Collections" in the help file for further information.
        Dim Description As String    ' Optional description string.
        Dim DOS As String            ' Optional DOS file name.
        Dim EmbFileNode As IntPtr    ' GetEmbeddedFileNode().
        Dim FileName As String       ' File name as 7 bit ASCII string.
        Dim FileNameIsURL As Boolean ' If true, FileName contains a URL.
        Dim ID1() As Byte            ' Optional file ID. Meaningful only if FileName points to a PDF file.
        Dim ID2() As Byte            ' Optional file ID. Meaningful only if FileName points to a PDF file.
        Dim IsVolatile As Boolean    ' If true, the file changes frequently with time.
        Dim Mac As String            ' Optional Mac file name.
        Dim Unix As String           ' Optional Unix file name.
        Dim RelFileNode As IntPtr    ' Optional related files array. -> GetRelFileNode().
        Dim Thumb As IntPtr          ' Optional thumb nail image. -> GetImageObjEx().
        Dim UFileName As String      ' PDF 1.7. Same as FileName but Unicode is allowed.
    End Structure

    Public Structure TPDFFontInfo
        Dim Ascent As Single                 ' Ascent (optional).
        Dim AvgWidth As Single               ' Average character width (optional).
        Dim BaseEncoding As TBaseEncoding    ' Valid only if HaveEncoding is true.
        Dim BaseFont As String               ' PostScript Name or Family Name.
        Dim CapHeight As Single              ' Cap height (optional).
        Dim CharSet As String                ' The charset describes which glyphs are present in the font.
        Dim CIDOrdering As String            ' SystemInfo -> A string that uniquely names the character collection within the specified registry.
        Dim CIDRegistry As String            ' SystemInfo -> Issuer of the character collection.
        Dim CIDSet() As Byte                 ' CID fonts only. This is a table of bits indexed by CIDs.
        Dim CIDSupplement As Integer         ' CIDSystemInfo -> The supllement number of the character collection.
        Dim CIDToGIDMap() As Byte            ' Allowed for embedded TrueType based CID fonts only.
        Dim CMapBuf() As Byte                ' Only available if the CMap was embedded.
        Dim CMapName As String               ' CID fonts only (this is the encoding if the CMap is not embedded).
        Dim Descent As Single                ' Descent (optional).
        Dim Encoding As String               ' Unicode mapping 0..255 -> not available for CID fonts.
        Dim FirstChar As Integer             ' First char (simple fonts only).
        Dim Flags As Integer                 ' The font flags describe various characteristics of the font. See help file for further information.
        Dim FontBBox As TBBox                ' This is the size of the largest glyph in this font. The bounding box is important for text selection.
        Dim FontBuffer() As Byte             ' The font buffer is present if the font was embedded or if it was loaded from a file buffer.
        Dim FontFamily As String             ' Optional Font Family (Family Name, always available for system fonts).
        Dim FontFilePath As String           ' Only available for system fonts.
        Dim FontFileType As TFontFileSubtype ' See description in the help file for further information.
        Dim FontName As String               ' Font name (should be the same as BaseFont).
        Dim FontStretch As String            ' Optional -> UltraCondensed, ExtraCondensed, Condensed, and so on.
        Dim FontType As TFontType            ' If ftType0 the font is a CID font. The Encoding is not set in this case.
        Dim FontWeight As Single             ' Font weight (optional).
        Dim FullName As String               ' System fonts only.
        Dim HaveEncoding As Boolean          ' If true, BaseEncoding was set from the font's encoding.
        Dim HorzWidths() As Single           ' Horizontal glyph widths.
        Dim Imported As Boolean              ' If true, the font was imported from an external PDF file.
        Dim ItalicAngle As Single            ' Italic angle
        Dim Lang As String                   ' Optional language code defined by BCP 47.
        Dim LastChar As Integer              ' Last char (simple fonts only).
        Dim Leading As Single                ' Leading (optional).
        Dim Length1 As Integer               ' Length of the clear text portion of a Type1 font.
        Dim Length2 As Integer               ' Length of the encrypted portion of a Type1 font program (Type1 fonts only).
        Dim Length3 As Integer               ' Length of the fixed-content portion of a Type1 font program or zero if not present.
        Dim MaxWidth As Single               ' Maximum glyph width (optional).
        Dim Metadata() As Byte               ' Optional XMP stream that contains metadata about the font file.
        Dim MisWidth As Single               ' Missing width (default = 0.0).
        Dim Panose() As Byte                 ' CID fonts only -> Optional 12 bytes long Panose string as described in Microsoft’s TrueType 1.0 Font Files Technical Specification.
        Dim PostScriptName As String         ' System fonts only.
        Dim SpaceWidth As Single             ' Space width in font units. A default value is set if the font contains no space character.
        Dim StemH As Single                  ' The thickness, measured vertically, of the dominant horizontal stems of glyphs in the font.
        Dim StemV As Single                  ' The thickness, measured horizontally, of the dominant vertical stems of glyphs in the font.
        Dim ToUnicode() As Byte              ' Only available for imported fonts. This is an embedded CMap that translates PDF strings to Unicode.
        Dim VertDefPos As TFltPoint          ' Default vertical displacement vector.
        Dim VertWidths() As TCIDMetric       ' Vertical glyph widths -> 0..VertWidthsCount -1.
        Dim WMode As Integer                 ' Writing Mode -> 0 == Horizontal, 1 == Vertical.
        Dim XHeight As Single                ' The height of lowercase letters (like the letter x), measured from the baseline, in fonts that have Latin characters.
    End Structure

    Public Structure TPDFFontObj
        Dim Ascent As Single      ' Ascent
        Dim BaseFont As String    ' PostScript Name or Family Name
        Dim CapHeight As Single   ' Cap height
        Dim Descent As Single     ' Descent
        Dim Encoding As Char()    ' Unicode mapping 0..255 -> not set if a CID font is selected
        Dim FirstChar As Integer  ' First char
        Dim Flags As Integer      ' Font flags -> font descriptor
        Dim FontFamily As String  ' Optional Font Family (Family Name)
        Dim FontName As String    ' Font name -> font descriptor
        Dim FontType As TFontType ' If ftType0 the font is a CID font. The Encoding is not set in this case.
        Dim ItalicAngle As Single ' Italic angle
        Dim LastChar As Integer   ' Last char
        Dim SpaceWidth As Single  ' Space width in font units. A default value is set if the font contains no space character.
        Dim Widths As Single()    ' Glyph widths -> 0..WidthsCount -1
        Dim XHeight As Single     ' x-height
        Dim DefWidth As Single    ' Default character widths -> CID fonts only
        Dim FontFile As IntPtr    ' Font file buffer -> only imported fonts are returned.
        Dim Length1 As Integer    ' Length of the clear text portion of the Type1 font, or the length of the entire font program if FontType != ffType1.
        Dim Length2 As Integer    ' Length of the encrypted portion of the Type1 font program (Type1 fonts only).
        Dim Length3 As Integer    ' Length of the fixed-content portion of the Type1 font program or zero if not present.
        Dim FontFileType As TFontFileSubtype  ' See description in the help file for further information.
    End Structure

    Public Structure TPDFGlyphOutline
        Dim AdvanceX As Single
        Dim AdvanceY As Single
        Dim OriginX As Single
        Dim OriginY As Single
        Dim Lsb As Int16
        Dim Tsb As Int16
        Dim HaveBBox As Boolean
        Dim BBox As TFRect
        Dim Outline() As TI32Point
    End Structure

    Public Structure TPDFGoToAction
        Dim DestPage As Integer           ' Destination page (the first page is denoted by 1).
        Dim DestPos() As Single           ' Destination position -> Array of 4 floating point values if set.
        Dim DestType As TDestType         ' Destination type.
        ' GoToR (GoTo Remote) actions only:
        Dim DestFile As IntPtr            ' see GetFileSpec().
        Dim DestName As String            ' Optional named destination that shall be loaded when opening the file.
        Dim NewWindow As Integer          ' Meaningful only if the destination file points to a PDF file.
        ' -1 = viewer default, 0 = false, 1 = true.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    Public Structure TPDFHideAction
        Dim Fields() As IntPtr            ' Array of field pointers -> GetFieldEx2().
        Dim Hide As Boolean               ' A flag indicating whether to hide or show the fields in the array.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFImage
        Dim Buffer As IntPtr                  ' Each scanline is aligned to a full byte.
        Dim BufSize As Integer                ' The size of the image buffer in bytes.
        Dim Filter As TDecodeFilter      ' Required decode filter if the image is compressed.
        ' Possible values are dfDCTDecode (JPEG), dfJPXDecode (JPEG2000),
        ' and dfJBIG2Decode. Other filters are always removed by DynaPDF since
        ' a conversion to a native file format is then always required.
        Dim OrgFilter As TDecodeFilter   ' The image was compressed with this filter in the PDF file. This info is
        ' useful to determine which compression filter should be used when creating
        ' a new image file from the image buffer.
        Dim JBIG2Globals As IntPtr            ' Optional global page 0 segment (dfJBIG2Decode filter only).
        Dim JBIG2GlobalsSize As Integer       ' The size of the bit stream in bytes.
        Dim BitsPerPixel As Integer           ' Bit depth of the image buffer. Possible values are 1, 2, 4, 8, 16, 24, and 32.
        Dim ColorSpace As TExtColorSpace ' The color space refers either to the image buffer or to the color table if set.
        ' Note that 1 bit images can occur with and without a color table.
        Dim NumComponents As Integer          ' The number of components stored in the image buffer.
        Dim MinIsWhite As Integer             ' If true, the colors of 1 bit images are reversed.
        Dim IColorSpaceObj As IntPtr          ' Pointer to the original color space.
        Dim ColorTable As IntPtr              ' The color table or NULL.
        Dim ColorCount As Integer             ' The number of colors in the color table.
        Dim Width As Integer                  ' Image width in pixel.
        Dim Height As Integer                 ' Image height in pixel.
        Dim ScanLineLength As Integer         ' The length of a scanline in bytes.
        Dim InlineImage As Integer            ' Boolean -> If true, the image is an inline image.
        Dim Interpolate As Integer            ' Boolean -> If true, image interpolation should be performed.
        Dim Transparent As Integer            ' Boolean -> The meaning is different depending on the bit depth and whether a color
        ' table is available. If the image is a 1 bit image and if no color table is available,
        ' black pixels must be drawn with the current fill color.
        ' If the image contains a color table ColorMask contains the range of indexes
        ' in the form min/max index which appears transparent. If no color table is
        ' present ColorMask contains the transparent ranges in the form min/max for
        ' each component.
        Dim ColorMask As IntPtr               ' The array contains ranges in the form min/max (2 values per component) for each
        ' component before decoding (data type Byte).
        Dim IMaskImage As IntPtr              ' If set, a 1 bit image is used as a transparency mask. Call GetImageObjEx() to decode the image.
        Dim ISoftMask As IntPtr               ' If set, a grayscale image is used as alpha channel. Call GetImageObjEx() to decode the image.
        Dim Decode As IntPtr                  ' If set, samples must be decoded. The array contains 2 * NumComponents values (data type Single).
        ' The decode array is never set if the image is returned decompressed since
        ' it is already applied during decompression.
        Dim Intent As TRenderingIntent   ' Default riNone.
        Dim SMaskInData As Integer            ' JPXDecode only, PDF_MAX_INT if not set. See PDF Reference for further information.
        Dim OC As IntPtr                      ' Pointer to Optional Content Group if any.
        Dim Metadata As IntPtr                ' Optional XML Metadata stream.
        Dim MetadataSize As Integer           ' Length of Metadata in bytes.
        Dim ObjectPtr As IntPtr               ' Internal pointer to the image class.
        Dim ResolutionX As Single             ' Image resolution on the x-axis.
        Dim ResolutionY As Single             ' Image resolution on the y-axis.
        Dim Measure As IntPtr                 ' Optional measure dictionary -> GetMeasureObj().
        Dim PtData As IntPtr                  ' Pointer of a Point Data dictionary. The value can be accessed with GetPtDataObj().
        ' Reserved fields for future extensions
        Private Reserved1 As Integer
        Private Reserved2 As Integer
        Private Reserved3 As Integer
        Private Reserved4 As Integer
        Private Reserved5 As Integer
        Private Reserved6 As Integer
        Private Reserved7 As Integer
        Private Reserved8 As Integer
        Private Reserved9 As Integer
        Private Reserved10 As Integer
        Private Reserved11 As Integer
        Private Reserved12 As Integer
        Private Reserved13 As Integer
        Private Reserved14 As Integer
        Private Reserved15 As Integer
        Private Reserved16 As Integer
    End Structure

    Public Structure TPDFImportDataAction
        Dim Data As TPDFFileSpecEx        ' The data or file to be loaded.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    Public Structure TPDFJavaScriptAction
        Dim Script As String              ' The script
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any.
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    Public Structure TPDFLaunchAction
        Dim AppName As String             ' Optional. The name of the application that should be launched.
        Dim DefDir As String              ' Optional default directory.
        Dim File As IntPtr                ' see GetFileSpec().
        Dim NewWindow As Integer          ' -1 = viewer default, 0 = false, 1 = true.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        Dim Operation As String           ' Optional string specifying the operation to perform (open or print).
        Dim Parameter As String           ' Optional parameter string that shall be passed to the application (AppName).
    End Structure

    Public Structure TPDFMeasure
        Dim IsRectilinear As Boolean ' If true, the members of the rectilinear measure dictionary are set. The geospatial members otherwise.
        ' --- Rectilinear measure dictionary ---
        Dim Angles() As IntPtr       ' Number format array to measure angles -> GetNumberFormatObj()
        Dim Area() As IntPtr         ' Number format array to measure areas -> GetNumberFormatObj()
        Dim CXY As Single            ' Optional, meaningful only when Y is present.
        Dim Distance() As IntPtr     ' Number format array to measure distances -> GetNumberFormatObj()
        Dim OriginX As Single        ' Origin of the measurement coordinate system.
        Dim OriginY As Single        ' Origin of the measurement coordinate system.
        Dim R As String              ' A text string expressing the scale ratio of the drawing.
        Dim Slope() As IntPtr        ' Number format array for measurement of the slope of a line -> GetNumberFormatObj()
        Dim x() As IntPtr            ' Number format array for measurement of change along the x-axis and, if Y is not present, along the y-axis as well.
        Dim y() As IntPtr            ' Number format array for measurement of change along the y-axis.

        ' --- Geospatial measure dictionary ---
        Dim Bounds() As Single       ' Array of numbers taken pairwise to describe the bounds for which geospatial transforms are valid.

        ' The DCS coordinate system is optional.
        Dim DCS_IsSet As Boolean     ' If true, the DCS members are set.
        Dim DCS_Projected As Boolean ' If true, the DCS values contains a pojected coordinate system.
        Dim DCS_EPSG As Integer      ' Optional, either EPSG or WKT is set.
        Dim DCS_WKT As String        ' Optional ASCII string

        ' The GCS coordinate system is required and should be present.
        Dim GCS_Projected As Boolean ' If true, the GCS values contains a pojected coordinate system.
        Dim GCS_EPSG As Integer      ' Optional, either EPSG or WKT is set.
        Dim GCS_WKT As String        ' Optional ASCII string

        Dim GPTS() As Single         ' Required, an array of numbers that shall be taken pairwise, defining points in geographic space as degrees of latitude and longitude, respectively.
        Dim LPTS() As Single         ' Optional, an array of numbers that shall be taken pairwise to define points in a 2D unit square.

        Dim PDU1 As String           ' Optional preferred linear display units.
        Dim PDU2 As String           ' Optional preferred area display units.
        Dim PDU3 As String           ' Optional preferred angular display units.
    End Structure

    Public Structure TPDFMovieAction
        Dim Annot As Integer              ' Optional. The movie annotation handle identifying the movie that shall be played.
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)> Dim FWPosition() As Single
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)> Dim FWScale() As Integer
        Dim Mode As String                ' Mode
        Dim Operation As String           ' Operation
        Dim Rate As Single                ' Rate
        Dim ShowControls As Boolean       ' ShowControls
        Dim Synchronous As Boolean        ' Synchronous
        Dim Title As String               ' The title of a movie annotation that shall be played. Either Annot or Title should be set, but not both.
        Dim Volume As Single              ' Volume
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    Public Structure TPDFNamedAction
        Dim Name As String                ' Only set if Type == naUserDefined.
        Dim NewWindow As Integer          ' -1 = viewer default, 0 = false, 1 = true.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any.
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        Dim Type As TNamedAction          ' Known pre-defined actions.
    End Structure

    Public Structure TPDFNamedDest
        Dim Name As String
        Dim DestFile As String  ' If set, the destination is located in this PDF file
        Dim DestPage As Integer
        Dim DestPos As TPDFRect
        Dim DestType As TDestType
    End Structure

    Public Structure TPDFNumberFormat
        Dim C As Single
        Dim D As Integer
        Dim f As TMeasureNumFormat
        Dim FD As Boolean
        Dim O As TMeasureLblPos
        Dim PS As String
        Dim RD As String
        Dim RT As String
        Dim SS As String
        Dim U As String
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFObjActions
        Dim StructSize As Integer     ' Must be set to sizeof(TPDFObjActions).
        Dim Action As Integer         ' Action handle or -1 if not set.
        Dim ActionType As TActionType ' The type of the action if Action >= 0.
        Dim Events As IntPtr          ' Additional events if any. -> GetObjEvent().
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFObjEvent
        Dim StructSize As Integer     ' Must be set to sizeof(TPDFObjEvent).
        Dim Action As Integer         ' Action to be executed.
        Dim ActionType As TActionType ' The type of the action.
        Dim ObjEvent As TObjEvent     ' The event when the action should be executed.
        Dim NextEvent As Integer      ' Pointer to the next event if any.
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFOCG
        Dim StructSize As Integer    ' Must be set to sizeof(TPDFOCG)
        Dim Handle As Integer        ' Handle or array index
        Dim Intent As TOCGIntent     ' Bitmask -> TOCGIntent
        <MarshalAs(UnmanagedType.LPStr)>
        Dim NameA As String          ' Layer name
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim NameW As String          ' Layer name
        Dim HaveContUsage As Integer ' If true, the layer contains a Content Usage dictionary. -> GetOCGContUsage().
        ' The following two members can only be set if HaveContUsage is true.
        Dim AppEvents As Integer     ' Bitmask -> see TOCAppEvent. If non-zero, the layer is included in one or more app events which control the layer state.
        Dim Categories As Integer    ' Bitmask -> see TOCGUsageCategory. The Usage Categories which control the layer state.
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFOCGContUsage
        Dim StructSize As Integer         ' Must be set to sizeof(TPDFOCGContUsage)
        Dim ExportState As Integer        ' 0 = Off, 1 = On, PDF_MAX_INT = not set.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim InfoCreatorA As String        ' CreatorInfo -> The application that created the group
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim InfoCreatorW As String        ' CreatorInfo -> The application that created the group
        <MarshalAs(UnmanagedType.LPStr)>
        Dim InfoSubtype As String         ' CreatorInfo -> A name defining the type of content, e.g. Artwork, Technical etc.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim LanguageA As String           ' A language code as described at SetLanguage() in the help file.
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim LanguageW As String           ' A language code as described at SetLanguage() in the help file.
        Dim LangPreferred As Integer      ' 0 = Off, 1 = On, PDF_MAX_INT = not set. The preffered state if there is a partial but no exact match of the language identifier.
        Dim PageElement As TOCPageElement ' If the group contains a pagination artefact.
        Dim PrintState As Integer         ' 0 = Off, 1 = On, PDF_MAX_INT = not set.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim PrintSubtype As String        ' The type of content that is controlled by the OCG, e.g. Trapping, PrintersMarks or Watermark.
        Dim UserNamesCount As Integer     ' The user names (if any) can be accessed with GetOCGUsageUserName().
        Dim UserType As TOCUserType       ' The user for whom this optional content group is primarily intendet.
        Dim ViewState As Integer          ' 0 = Off, 1 = On, PDF_MAX_INT = not set.
        Dim ZoomMin As Single             ' The minimum magnification factor at which the group should be On. -1 if not set.
        Dim ZoomMax As Single             ' The maximum magnification factor at which the group should be On. -1 if not set.
    End Structure

    Public Structure TPDFOCLayerConfig
        Dim Intent As TOCGIntent ' Possible values oiDesign, oiView, or oiAll.
        Dim IsDefault As Boolean ' If true, this is the default configuration.
        Dim Name As String       ' Optional configuration name. The default config has usually no name but all others should have one.
    End Structure

    Public Class TPDFOCUINode
        Public Label As String     ' Optional label.
        Public NextChild As IntPtr ' If set, the next child node that must be loaded.
        Public NewNode As Boolean  ' If true, a new child node must be created.
        Public OCG As Integer      ' Optional OCG handle. -1 if not set -> GetOCG().
    End Class

    Public Structure TPDFOutputIntent
        Dim Buffer() As Byte
        Dim Info As String
        Dim NumComponents As Integer
        Dim OutputCondition As String
        Dim OutputConditionID As String
        Dim RegistryName As String
        Dim SubType As String
    End Structure

    Public Structure TPDFPageLabel
        Dim StartRange As Integer      ' Number of the first page in the range. If no further label follows, the last
        ' page in the range is pdfGetPageCount(). The first page is denoted by 1.
        Dim Format As TPageLabelFormat ' Number format to be used.
        Dim FirstPageNum As Integer    ' First page number to be displayed in the page label. Subsequent pages are
        ' numbered sequentially from this value.
        Dim Prefix As String           ' Optional prefix
    End Structure

    Public Structure TPDFPrintSettings
        Dim DuplexMode As TDuplexMode
        Dim NumCopies As Integer               ' -1 means not set. Values larger than 5 are ignored in viewer applications.
        Dim PickTrayByPDFSize As Integer       ' -1 means not set. 0 == false, 1 == true.
        Dim PrintRanges() As Integer           ' If set, the array contains PrintRangesCount * 2 values. Each pair consists
        ' of the first and last pages in the sub-range. The first page in the PDF file
        ' is denoted by 0.
        Dim PrintScaling As TPrintScaling ' dpmNone means not set.
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFRawImage
        Dim StructSize As Integer       ' Must be set to sizeof(TPDFRawImage)
        Dim Buffer As IntPtr            ' Image buffer
        Dim BufSize As Integer          ' Buffer size
        Dim BitsPerComponent As Integer ' Bits per component
        Dim NumComponents As Integer    ' Number of components (max 32)
        Dim CS As TExtColorSpace        ' Image color space
        Dim CSHandle As Integer         ' Color space handle (non-device color spaces only)
        Dim Stride As Integer           ' Scanline length in bytes -> If negative, the image is defined in bottom up coordinates, top down otherwise
        Dim HasAlpha As Integer         ' If true, the last component is an alpha channel
        Dim IsBGR As Integer            ' esDeviceRGB only -> If true, the image components are defined as BGR instead of RGB
        Dim MinIsWhite As Integer       ' 1 bit images only -> If true, zero pixel values must be treated as white instead of black
        Dim Width As Integer            ' Width in pixels (must be greater zero)
        Dim Height As Integer           ' Height in pixels (must be greater zero)
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFRect
        Dim Left As Double
        Dim Bottom As Double
        Dim Right As Double
        Dim Top As Double
    End Structure

    Public Structure TPDFRelFileNode
        Dim Name As String     ' Name of this file spcification.
        Dim EF As TPDFFileSpec ' Embedded file.
        Dim NextNode As IntPtr ' Next node if any.
    End Structure

    Public Structure TPDFResetFormAction
        Dim Fields() As IntPtr            ' Array of field pointers -> GetFieldEx2().
        Dim Include As Boolean            ' If true, the fields in the Fields array must be reset. If false, these fields must be excluded.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    Public Structure TPDFSigDict
        Dim ByteRange() As Integer  ' ByteRange -> Byte offset followed by the corresponding length.
        ' The byte ranges are required to create the digest. The values
        ' are returned as is. So, you must check whether the offsets and
        ' length values are valid. There are normally at least two ranges.
        ' Overlapping ranges are not allowed! Any error breaks processing
        ' and the signature should be considered as invalid.
        Dim Cert() As Byte          ' X.509 Certificate when SubFilter is adbe.x509.rsa_sha1.
        Dim Changes() As Integer    ' If set, an array of three integers that specify changes to the
        ' document that have been made between the previous signature and
        ' this signature in this order: the number of pages altered, the
        ' number of fields altered, and the number of fields filled in.
        Dim ContactInfo As String   ' Optional contact info string, e.g. an email address
        Dim Contents() As Byte      ' The signature. This is either a DER encoded PKCS#1 binary data
        ' object or a DER-encoded PKCS#7 binary data object depending on
        ' the used SubFilter.
        Dim Filter As String        ' The name of the security handler, usually Adobe.PPKLite.
        Dim Location As String      ' Optional location of the signer
        Dim SignTime As String      ' Date/Time string
        Dim Name As String          ' Optional signers name
        Dim PropAuthTime As Integer ' Optional -> The number of seconds since the signer was last authenticated.
        Dim PropAuthType As String  ' Optional -> The method that shall be used to authenticate the signer.
        ' Valid values are PIN, Password, and Fingerprint.
        Dim Reason As String        ' Optional reason
        Dim Revision As Integer     ' Optional -> The version of the signature handler that was used to create
        ' the signature.
        Dim SubFilter As String     ' A name that describes the encoding of the signature value. Should be
        ' adbe.x509.rsa_sha1, adbe.pkcs7.detached, or adbe.pkcs7.sha1.
        Dim Version As Integer      ' The version of the signature dictionary format.
    End Structure

    Public Structure TPDFSigParms
        Dim PKCS7ObjLen As Integer    ' The maximum length of the signed PKCS#7 object
        Dim HashType As THashType     ' If set to htDetached, the bytes ranges of the PDF file will be returned.
        Dim Range1() As Byte          ' Out -> Contains either the hash or the first byte range to create a detached signature
        Dim Range2() As Byte          ' Out -> Set only if HashType == htDetached
        Dim ContactInfo As String     ' Optional, e.g. an email address_
        Dim Location As String        ' Optional location of the signer
        Dim Reason As String          ' Optional reason why the file was signed
        Dim Signer As String          ' Optional, the issuer of the certificate takes precedence
        Dim Encrypt As Boolean        ' If true, the file will be encrypted
        ' These members will be ignored if Encrypt is set to false
        Dim OpenPwd As String         ' Open password
        Dim OwnerPwd As String        ' Owner password to change the security settings
        Dim KeyLen As TKeyLen         ' Key length to be used to encrypt the file
        Dim Restrict As TRestrictions ' What should be restricted?
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFStack
        Dim ctm As TCTM                ' Pre-multiplied global transformation matrix
        Dim tm As TCTM                 ' Pre-multiplied text transformation matrix
        Dim x As Double                ' Unused -> always 0.0
        Dim y As Double                ' Unused -> always 0.0
        Dim FontSize As Double         ' Font size
        Dim CharSP As Double           ' Character spacing
        Dim WordSP As Double           ' Word spacing
        Dim HScale As Double           ' Horizontal text scaling
        Dim TextRise As Double         ' Text rise -> always 0.0 because it is already included in the text transformation matrix
        Dim Leading As Double          ' Leading
        Dim LineWidth As Double        ' Line width
        Dim DrawMode As TDrawMode      ' Text draw mode
        Dim FillCS As TPDFColorSpace   ' Fill color space
        Dim StrokeCS As TPDFColorSpace ' Stroke color space
        Dim FillColor As Integer       ' Fill color
        Dim StrokeColor As Integer     ' Stroke color
        Dim BaseObject As IntPtr       ' Internal
        Dim CIDFont As Integer         ' If true, ReplacePageText() can only be used to delete a string
        Dim Text As IntPtr             ' Raw text without kerning space
        Dim TextLen As Integer         ' Raw text length
        Dim RawKern As IntPtr          ' Raw kerning array
        Dim Kerning As IntPtr          ' Already translated Unicode kerning array
        Dim KerningCount As Integer    ' Number of kerning records
        Dim TextWidth As Single        ' The width of the entire text record measured in text space
        Dim IFont As IntPtr            ' Font object used to print the string -> fntGetFont() can be used to return the font properties
        Dim Embedded As Integer        ' If true, the font is embedded
        Dim SpaceWidth As Single       ' Measured in text space
        'These members can be modified after the structure has been initialized with InitStack().
        'If the destination color space should be DeviceCMYK initialize FillColor and StrokeColor
        'with PDF_CMYK(0,0,0,255); which represents black.
        Dim ConvColors As Integer       ' If set to true (default), all colors are converted to the specified destination color space
        Dim DestSpace As TPDFColorSpace ' Destination color space -> default == csDeviceRGB

        ' This member can be used in combination with ReplacePageText() to preserve a number
        ' of kerning records from deletion. All records above this value will be deleted.
        ' Take a look into the file examples/util/pdf_edit_text.vb to determine how this member
        ' can be used.
        Dim DeleteKerningAt As Integer
        Dim FontFlags As Integer        ' PDF font flags

        ' ------------------------------- Reserved fields -------------------------------
        Private Reserved1 As Integer
        Private Reserved2 As Integer
        Private Reserved3 As Integer
        Private Reserved4 As Integer
        Private Reserved5 As Integer
        Private Reserved6 As Integer
        Private Reserved7 As Integer
        Private Reserved8 As Integer
        Private Reserved9 As Integer
        Private Reserved10 As Integer
        Private Reserved11 As Integer
        Private ContentPtr As IntPtr
    End Structure

    Public Structure TPDFSubmitFormAction
        Dim CharSet As String             ' Optional charset in which the form should be submitted.
        Dim Fields() As IntPtr            ' Array of field pointers -> GetFieldEx2().
        Dim Flags As TSubmitFlags         ' Various flags, see CreateSubmitAction() for further information.
        Dim URL As String                 ' The URL of the script at the Web server that will process the submission.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    Public Structure TPDFSysFont
        Dim BaseType As TFontBaseType       ' Font type
        Dim CIDOrdering As String           ' OpenType CID fonts only
        Dim CIDRegistry As String           ' OpenType CID fonts only
        Dim CIDSupplement As Integer        ' OpenType CID fonts only
        Dim DataOffset As Integer           ' Data offset
        Dim FamilyName As String            ' Family name
        Dim FilePath As String              ' Font file path
        Dim FileSize As Integer             ' File size in bytes
        Dim Flags As TEnumFontProcFlags     ' Bitmask
        Dim FullName As String              ' Full name
        Dim Length1 As Integer              ' Length of the clear text portion of a Type1 font
        Dim Length2 As Integer              ' Length of the eexec encrypted binary portion of a Type1 font
        Dim PostScriptName As String        ' Postscript mame
        Dim Index As Integer                ' Zero based font index if the font is stored in a TrueType collection
        Dim IsFixedPitch As Boolean         ' If true, the font is a fixed pitch font. A proprtional font otherwise.
        Dim Style As TFStyle                ' Font style
        Dim UnicodeRange1 As TUnicodeRange1 ' Bitmask
        Dim UnicodeRange2 As TUnicodeRange2 ' Bitmask
        Dim UnicodeRange3 As TUnicodeRange3 ' Bitmask
        Dim UnicodeRange4 As TUnicodeRange4 ' Bitmask
    End Structure

    Public Structure TPDFURIAction
        Dim BaseURL As String             ' Optional, if defined in the Catalog object.
        Dim IsMap As Boolean              ' A flag specifying whether to track the mouse position when the URI is resolved: e.g. http:'test.org?50,70.
        Dim URI As String                 ' Uniform Resource Identifier.
        Dim NextAction As Integer         ' -1 or next action handle to be executed if any.
        Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
    End Structure

    Public Structure TPDFViewport
        Dim BBox As TFltRect   ' Bounding box
        Dim Measure As IntPtr  ' Optional pointer of a measure dictionary -> GetMeasureObj()
        Dim Name As String     ' Optional name
        Dim PtData As IntPtr   ' Optional pointer of a Point Data dictionary. The value can be accessed with GetPtDataObj().
    End Structure

    Public Class TPDFXFAStream
        Public Buffer() As Byte
        Public Name As String
    End Class

    ' ---------------------------------------------- Parser Interface ----------------------------------------------

    Public Delegate Function TApplyPattern(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Type As TPatternType, ByVal Pattern As IntPtr) As Integer
    Public Delegate Function TBeginPattern(ByVal Data As IntPtr, ByVal Fill As Integer, ByVal Handle As Integer, ByVal Type As TPatternType, ByRef BBox As TPDFRect, ByVal Matrix As IntPtr, ByVal XStep As IntPtr, ByVal YStep As IntPtr) As Integer
    Public Delegate Function TBeginTemplate(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Handle As Integer, ByRef BBox As TPDFRect, ByVal Matrix As IntPtr) As Integer
    Public Delegate Function TBezierTo1(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x3 As Double, ByVal y3 As Double) As Integer
    Public Delegate Function TBezierTo2(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Integer
    Public Delegate Function TBezierTo3(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Integer
    Public Delegate Function TClipPath(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal EvenOdd As Integer, ByVal Mode As TPathFillMode) As Integer
    Public Delegate Function TClosePath(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Mode As TPathFillMode) As Integer
    Public Delegate Function TDrawShading(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Type As TShadingType, ByVal Shading As IntPtr) As Integer
    Public Delegate Sub TEndPattern(ByVal Data As IntPtr)
    Public Delegate Sub TEndTemplate(ByVal Data As IntPtr)
    Public Delegate Function TInsertImage(ByVal Data As IntPtr, ByRef Image As TPDFImage) As Integer
    Public Delegate Function TLineTo(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal x As Double, ByVal y As Double) As Integer
    Public Delegate Function TMoveTo(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal x As Double, ByVal y As Double) As Integer
    Public Delegate Sub TMulMatrix(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByRef M As TCTM)
    Public Delegate Function TRectangle(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double) As Integer
    Public Delegate Function TRestoreGraphicState(ByVal Data As IntPtr) As Integer
    Public Delegate Function TSaveGraphicState(ByVal Data As IntPtr) As Integer
    Public Delegate Sub TSetCharSpacing(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Value As Double)
    Public Delegate Sub TSetExtGState(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByRef GS As TPDFExtGState2)
    Public Delegate Sub TSetFillColor(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Color As IntPtr, ByVal NumComps As Integer, ByVal CS As TExtColorSpace, ByVal IColorSpace As IntPtr)
    Public Delegate Sub TSetFont(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Type As TFontType, ByVal Embedded As Integer, ByVal FontName As IntPtr, ByVal Style As TFStyle, ByVal FontSize As Double, ByVal IFont As IntPtr)
    Public Delegate Sub TSetLeading(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Value As Double)
    Public Delegate Sub TSetLineCapStyle(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Style As TLineCapStyle)
    Public Delegate Sub TSetLineDashPattern(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Dash As IntPtr, ByVal NumValues As Integer, ByVal Phase As Integer)
    Public Delegate Sub TSetLineJoinStyle(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Style As TLineJoinStyle)
    Public Delegate Sub TSetLineWidth(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Value As Double)
    Public Delegate Sub TSetMiterLimit(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Value As Double)
    Public Delegate Sub TSetStrokeColor(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Color As IntPtr, ByVal NumComps As Integer, ByVal CS As TExtColorSpace, ByVal IColorSpace As IntPtr)
    Public Delegate Sub TSetTextDrawMode(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Mode As TDrawMode)
    Public Delegate Sub TSetTextScale(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Value As Double)
    Public Delegate Sub TSetWordSpacing(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByVal Value As Double)
    ' It is not required to declare the marshalling attributes in the real callback function.
    Public Delegate Function TShowTextArrayA(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByRef Matrix As TCTM, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> ByVal Source() As TTextRecordA, ByVal Count As Integer, ByVal Width As Double) As Integer
    Public Delegate Function TShowTextArrayW(ByVal Data As IntPtr, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> ByVal Source() As TTextRecordA, ByRef Matrix As TCTM, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=4)> ByVal Kerning() As TTextRecordW, ByVal Count As Integer, ByVal Width As Double, ByVal Decoded As Integer) As Integer
    'Public Delegate Function TShowTextArrayA(ByVal Data As IntPtr, ByVal Obj As IntPtr, ByRef Matrix As TCTM, ByVal Source As IntPtr, ByVal Count As Integer, ByVal Width As Double) As Integer
    'Public Delegate Function TShowTextArrayW(ByVal Data As IntPtr, ByVal Source As IntPtr, ByRef Matrix As TCTM, ByVal Kerning As IntPtr, ByVal Count As Integer, ByVal Width As Double, ByVal Decoded As Integer) As Integer

    ' Unnecessary callback functions can be set to null -> already default when creating a new instance of the structure.
    Public Structure TPDFParseInterface
        Dim ApplyPattern As TApplyPattern
        Dim BeginPattern As TBeginPattern
        Dim BeginTemplate As TBeginTemplate
        Dim BezierTo1 As TBezierTo1
        Dim BezierTo2 As TBezierTo2
        Dim BezierTo3 As TBezierTo3
        Dim ClipPath As TClipPath
        Dim ClosePath As TClosePath
        Dim DrawShading As TDrawShading
        Dim EndPattern As TEndPattern
        Dim EndTemplate As TEndTemplate
        Dim LineTo As TLineTo
        Dim MoveTo As TMoveTo
        Dim MulMatrix As TMulMatrix
        Dim Rectangle As TRectangle
        Dim RestoreGraphicState As TRestoreGraphicState
        Dim SaveGraphicState As TSaveGraphicState
        Dim SetCharSpacing As TSetCharSpacing
        Dim SetExtGState As TSetExtGState
        Dim SetFillColor As TSetFillColor
        Dim SetFont As TSetFont
        Dim SetLeading As TSetLeading
        Dim SetLineCapStyle As TSetLineCapStyle
        Dim SetLineDashPattern As TSetLineDashPattern
        Dim SetLineJoinStyle As TSetLineJoinStyle
        Dim SetLineWidth As TSetLineWidth
        Dim SetMiterLimit As TSetMiterLimit
        Dim SetStrokeColor As TSetStrokeColor
        Dim SetTextDrawMode As TSetTextDrawMode
        Dim SetTextScale As TSetTextScale
        Dim SetWordSpacing As TSetWordSpacing
        Private Reserved001 As IntPtr
        Private Reserved002 As IntPtr
        Dim ShowTextArrayW As TShowTextArrayW
        Dim InsertImage As TInsertImage
        Dim ShowTextArrayA As TShowTextArrayA
        Private Reserved01 As IntPtr
        Private Reserved02 As IntPtr
        Private Reserved03 As IntPtr
        Private Reserved04 As IntPtr
        Private Reserved05 As IntPtr
        Private Reserved06 As IntPtr
        Private Reserved07 As IntPtr
        Private Reserved08 As IntPtr
        Private Reserved09 As IntPtr
        Private Reserved10 As IntPtr
        Private Reserved11 As IntPtr
        Private Reserved12 As IntPtr
        Private Reserved13 As IntPtr
        Private Reserved14 As IntPtr
        Private Reserved15 As IntPtr
        Private Reserved16 As IntPtr
        Private Reserved17 As IntPtr
        Private Reserved18 As IntPtr
        Private Reserved19 As IntPtr
        Private Reserved20 As IntPtr
        Private Reserved21 As IntPtr
        Private Reserved22 As IntPtr
        Private Reserved23 As IntPtr
        Private Reserved24 As IntPtr
        Private Reserved25 As IntPtr
        Private Reserved26 As IntPtr
        Private Reserved27 As IntPtr
    End Structure


    ' ----------------------------------- Rendering API -----------------------------------

    Public Enum TPDFPixFormat
        pxf1Bit
        pxfGray
        pxfRGB
        pxfBGR
        pxfRGBA
        pxfBGRA
        pxfARGB
        pxfABGR
        pxfGrayA
        pxfCMYK
        pxfCMYKA
    End Enum

    Public Enum TPDFPageScale
        psFitWidth  ' Scale the page to the width of the image buffer
        psFitHeight ' Scale the page to the height of the image buffer
        psFitBest   ' Scale the page so that it fits fully into the image buffer
        psFitZoom   ' This mode should be used if the scaling factors of the transformation matrix are <> 1.0
    End Enum

    Public Enum TRasterFlags
        rfDefault = &H0                ' Render the page as usual
        rfScaleToMediaBox = &H1        ' Render the real paper format. Contents outside the crop box is clipped
        rfIgnoreCropBox = &H2          ' Ignore the crop box and render anything inside the media box without clipping
        ' Only one of these flags must be set at time!
        rfClipToArtBox = &H4           ' Clip the page to the art box if any
        rfClipToBleedBox = &H8         ' Clip the page to the bleed box if any
        rfClipToTrimBox = &H10         ' Clip the page to the trim box if any
        rfExclAnnotations = &H20       ' Don't render annotations
        rfExclFormFields = &H40        ' Don't render form fields
        rfSkipUpdateBG = &H80          ' Don't generate an update event arfer initializing the background to white
        rfRotate90 = &H100             ' Rotate the page 90 degress
        rfRotate180 = &H200            ' Rotate the page 180 degress
        rfRotate270 = &H400            ' Rotate the page 270 degress
        rfInitBlack = &H800            ' Initialize the image buffer to black before rendering (RGBA or GrayA must be initialized to black)
        rfCompositeWhite = &H1000      ' Composite pixel formats with an alpha channel finally with a white background. The alpha channel is
        ' 255 everywhere arfer composition. This flag is mainly provided for debug purposes but it can also be
        ' useful if the image must be copied on screen with a function that doesn't support alpha blending.
        rfExclPageContent = &H2000     ' If set, only annotations and form fields will be rendered (if any).

        ' If you want to render specific field types with RenderAnnotOrField() then use the following flags to exclude these fields.
        ' If all fields should be skipped then set the flag rfExclFormFields instead.
        rfExclButtons = &H4000
        rfExclCheckBoxes = &H8000
        rfExclComboBoxes = &H10000
        rfExclListBoxes = &H20000
        rfExclTextFields = &H40000
        rfExclSigFields = &H80000

        ' ---------------------------------
        rfScaleToBBox = &H100000             ' Considered only if the flag rfClipToArtBox rfClipToBleedBox or rfClipToTrimBox is set.
        ' If set the picture size is set to the size of the whished bounding box.
        rfDisableAAClipping = &H200000       ' Disable Anti-Aliasing for clipping paths. This flag is the most important one since clipping paths
        ' cause often visible artefacts in PDF files with flattened transparency.
        rfDisableAAText = &H400000           ' Disable Anti-Aliasing for text.
        rfDisableAAVector = &H800000         ' Disable Anti-Aliasing for vector graphics.
        rfDisableAntiAliasing = rfDisableAAClipping Or rfDisableAAText Or rfDisableAAVector ' Fully disable Anti-Aliasing.
        rfDisableBiLinearFilter = &H1000000  ' Disable the BiLevel filter for images. Sometetimes useful if sharp images are needed e.g. for barcodes.
        rfRenderInvisibleText = &H2000000    ' If set, treat text rendering mode Invisible as Normal.
    End Enum

    Public Delegate Function TOnUpdateWindow(ByVal Data As IntPtr, ByRef Area As TIntRect) As Integer

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFRasterImage
        Dim StructSize As Integer             ' Must be set to sizeof(TPDFRasterImage)
        Dim Flags As TRasterFlags             ' This is a bit mask. Flags can be combined with a binary or operator
        Dim DefScale As TPDFPageScale         ' Specifies how the page should be scaled into the image buffer.

        Dim InitWhite As Integer              ' If true, the image buffer is initialized to white before rendering.
        ' When a clipping rectangle is set, only the area inside the clipping
        ' rectangle is initialized to white.

        Dim ClipRect As TIntRect              ' Optional clipping rectangle defined in device coordinates (Pixels), default 0,0,0,0 (no clipping)
        Dim Matrix As TCTM                    ' Optional transformation matrix. Initialize the variable to the identity matrix (1,0,0,1,0,0)
        ' if you don't need it. The matrix can be used to move and scale the picture inside the image.
        Dim PageSpace As TCTM                 ' Out -> This matrix represents the mapping from page space to device space. The matrix
        ' is required when further objects should be drawn on the page, e.g. the bounding boxes.

        Dim DrawFrameRect As Integer          ' If true, the area outside the page's bounding box is filled with the
        ' frame color. InitWhite can still be used, with or without a clipping
        ' rectangle.
        Dim FrameColor As Integer             ' Must be defined in the color space of the pixel format but in the natural
        ' component order, e.g. RGB.

        Dim OnUpdateWindow As TOnUpdateWindow ' Optional, UpdateOnPathCount and UpdateOnImageCoverage define when the function should be called
        Dim OnInitDecoder As IntPtr           ' Not yet defined
        Dim OnDecodeLine As IntPtr            ' Not yet defined
        Dim UserData As IntPtr                ' Arbitrary pointer that should be passed to the callback functions

        Dim UpdateOnPathCount As Integer      ' Optional -> Call OnUpdateWindow when the limit was reached.
        ' Clipping paths increment the number too.
        ' Only full paths are considered, independent of the number of vertices
        ' they contain. The value should be larger than 50 and smaller than 10000.

        Dim UpdateOnImageCoverage As Single   ' Optional -> DynaPDF multiplies the output image width and height with this
        ' factor to calculate the coverage limit. When an image is inserted the unscaled
        ' width and height is added to the current coverage value. When the number
        ' reaches the limit the OnUpdateWindow event is raised.
        ' The factor should be around 0.5 through 5.0. Larger values cause less
        ' frequently update events.
        ' Statistics...
        Dim NumAnnots As Integer              ' Out -> Number of rendered annotations (excluding invisible annotation but annotations with no appearance are included)
        Dim NumBezierCurves As Integer        ' Out -> Number of bezier curves which where rendered. Glyph outlines are not taken into account.
        Dim NumClipPaths As Integer           ' Out -> Number of clipping paths used in the page. Should be small as possible!
        Dim NumFormFields As Integer          ' Out -> Number of rendered form fields (excluding invisible fields but fields with no appearance are included)
        Dim NumGlyphs As Integer              ' Out -> When the number of glyphs equals NumTextRecords then there is probably some room for optimization...
        Dim NumImages As Integer              ' Out -> Number of images that were rendered
        Dim NumLineTo As Integer              ' Out -> Number of LineTo operators
        Dim NumPaths As Integer               ' Out -> Number of paths which were processed
        Dim NumPatterns As Integer            ' Out -> Number of pattern which were processed
        Dim NumRectangles As Integer          ' Out -> Number of rectangle operators
        Dim NumRestoreGState As Integer       ' Out -> Should be equal to NumSaveGState
        Dim NumSaveGState As Integer          ' Out -> The number of save graphics state operators
        Dim NumShadings As Integer            ' Out -> Number shadings which were processed
        Dim NumSoftMasks As Integer           ' Out -> Number of soft masks that were processed. Alpha channels of images are not taken into account.
        Dim NumTextRecords As Integer         ' Out -> Number of independent text records which were rendered
    End Structure

    ' ------------------------------------------------- Page cache -------------------------------------------------

    '   The path names can be set in Ansi (code page 1252 on Windows) or Unicode format. The Ansi version accepts
    '   UTF-8 strings on non-Windows operating systems. UTF-16 Unicode strings are converted to UTF-8 on non-Windows
    '   operating systems.
    '
    '   In general, the DefInXXX profiles are used if no other profile is available for the color space. Possible
    '   sources are DefaultGray, DefaultRGB, DefaultCMYK, and the Rendering Intents.
    '
    '   The SoftProof profile emulates the output device. This is typically a printer profile or a default CMYK
    '   profile. If no profile is set then no device will be emulated. What you see is maybe not what you get on
    '   a printer.
    '
    '   To disable color management set the parameter Profiles of rasInitColormanagement() to NULL.

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFColorProfiles
        Dim StructSize As Integer    ' Must be set to sizeof(TPDFColorProfile)
        <MarshalAs(UnmanagedType.LPStr)>
        Dim DefInGrayA As String     ' Optional
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim DefInGrayW As String     ' Optional
        <MarshalAs(UnmanagedType.LPStr)>
        Dim DefInRGBA As String      ' Optional, sRGB is the default. The "A" stands for Ansi string and not for Alpha...
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim DefInRGBW As String      ' Optional
        <MarshalAs(UnmanagedType.LPStr)>
        Dim DefInCMYKA As String     ' Optional, CMYK colors are the problematic ones. The other profiles can be created on demand
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim DefInCMYKW As String     ' but this is not possible with a CMYK profile. So, this is the most important input profile.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim DeviceProfileA As String ' Optional, the output profile must be compatible with the output color space.
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim DeviceProfileW As String ' At this time only Gray or RGB profiles are supported. This is the monitor profile! Default is sRGB.
        <MarshalAs(UnmanagedType.LPStr)>
        Dim SoftProofA As String     ' Optional but very important. This profile emulates the output device.
        <MarshalAs(UnmanagedType.LPWStr)>
        Dim SoftProofW As String     ' Optional.
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0)>
    Public Structure TPDFColorProfilesEx
        Dim StructSize As Integer       ' Must be set to sizeof(TPDFColorProfileEx)
        Dim DefInGray As IntPtr         ' Optional
        Dim DefInGrayLen As Integer     ' Optional
        Dim DefInRGB As IntPtr          ' Optional
        Dim DefInRGBLen As Integer      ' Optional
        Dim DefInCMYK As IntPtr         ' Optional, CMYK colors are the problematic ones. The other profiles can be created on demand
        Dim DefInCMYKLen As Integer     ' but this is not possible with a CMYK profile. So, this is the most important input profile.
        Dim DeviceProfile As IntPtr     ' Optional, the output profile must be compatible with the output color space.
        Dim DeviceProfileLen As Integer ' Gray, RGB, or CMYK profiles are supported.
        Dim SoftProof As IntPtr         ' Optional but very important. This profile emulates the output device.
        Dim SoftProofLen As Integer     ' Optional.
    End Structure

    Public Enum TPDFInitCMFlags
        icmDefault = 0         ' Default rules.
        icmBPCompensation = 1  ' Black point compensation preserves the black point when converting CMYK colors to different color spaces.
        icmCheckBlackPoint = 2 ' If set, soft proofing will be disabled if the black point of the output intent is probably invalid.
    End Enum

    Public Enum TPDFCursor
        pcrHandNormal = 0
        pcrHandClosed = 1
        pcrHandPoint = 2
        pcrIBeam = 3
    End Enum

    Public Enum TInitCacheFlags
        icfDefault = 0
        icfIgnoreOpenAction = 1
        icfIgnorePageLayout = 2
    End Enum

    Public Enum TPDFThreadPriority
        ttpLowest
        ttpIdle
        ttpBelowNormal  ' This is the default value. Normal can be used too but scrolling is smoother in this mode.
        ttpNormal
        ttpAboveNormal
        ttpHighest      ' Not really useful...
        ttpTimeCritical ' Don't do that!
    End Enum

    Public Enum TUpdBmkAction
        ubaDoNothing = 0      ' Nothing to do
        ubaOpenPage = 1       ' Jump to the new page. This flag is set if the bookmark contained a destination or go to action.
        ubaPageScale = 2      ' Update the page scale with SetPageScale().
        ubaZoom = 4           ' Zoom into the page, update the scroll ranges, and set the scroll positions.
        ubaUpdScrollBars = 8  ' This flag is always set if ubaZoom is set.
        ubaExecAction = 16    ' Check the parameter Action to execute further code. This flag can occur with or without ubaOpenPage.
    End Enum

    Public Enum TUpdScrollbar
        usbNoUpdate = 0           ' Nothing to do
        usbVertRange = 1          ' Update the vertical scroll range
        usbVertScrollPos = 2      ' Update the vertical scroll position
        usbHorzRange = 4          ' Update the horizontal scroll range
        usbHorzScrollPos = 8      ' Update the horizontal scroll position
        usbUpdateAll = 15         ' Update both scroll ranges and the scroll positions
        ' The cursor constants are set by MouseMove. Since we have only one cursor there is never more than one constant set.
        usbCursorHandNormal = 16  ' This is the default if the left mouse button is not pressed and if we are not over an action field
        usbCursorHandClosed = 32  ' Occurs when the cursor leaves an action field and if the left mouse button is pressed
        usbCursorHandPoint = 64   ' Occurs when we enter link or button field
        usbCursorIBeam = 128      ' Occurs when we enter an action field that accepts text input
        usbCursorMask = 240       ' Bitmask to mask out the cursor constants
    End Enum

    ' Callback function protoypes. These prototypes are requires to enable event support in VB .Net.
    ' Data contains the current instance pointer of the event interface.
    Public Delegate Function TEnumFontsProc(ByVal Data As IntPtr, ByVal FamilyName As IntPtr, ByVal PostScriptName As IntPtr, ByVal Style As Integer) As Integer
    Public Delegate Function TEnumFontsProc2(ByVal Data As IntPtr, ByVal PDFFont As IntPtr, ByVal FontType As Integer, ByVal BaseFont As IntPtr, ByVal FontName As IntPtr, ByVal Embedded As Integer, ByVal IsFormFont As Integer, ByVal Flags As Integer) As Integer
    Public Delegate Function TEnumFontsProcEx(ByVal Data As IntPtr, ByVal FamilyName As IntPtr, ByVal PostScriptName As IntPtr, ByVal Style As Integer, ByVal BaseType As TFontBaseType, ByVal Flags As TEnumFontProcFlags, ByVal FilePath As IntPtr) As Integer
    Public Delegate Function TErrorProc(ByVal Data As IntPtr, ByVal ErrCode As Integer, ByVal ErrMessage As IntPtr, ByVal ErrType As Integer) As Integer
    Public Delegate Function TOnPageBreakProc(ByVal Data As IntPtr, ByVal LastPosX As Double, ByVal LastPosY As Double, ByVal PageBreak As Integer) As Integer
    Public Delegate Function TOnFontNotFoundProc(ByVal Data As IntPtr, ByVal PDFFont As IntPtr, ByVal FontName As IntPtr, ByVal Style As TFStyle, ByVal StdFontIndex As Integer, ByVal IsSymbolFont As Integer) As Integer
    Public Delegate Function TOnReplaceICCProfile(ByVal Data As IntPtr, ByVal Type As TICCProfileType, ByVal ColorSpace As Integer) As Integer
    Public Delegate Sub TInitProgress(ByVal Data As IntPtr, ByVal ProgType As Integer, ByVal MaxCount As Integer)
    Public Delegate Function TProgress(ByVal Data As IntPtr, ByVal ActivePage As Integer) As Integer

    Public Class Modul_PDF

        ' Event procedures
        Public Event PDFEnumDocFont(ByVal PDFFont As IntPtr, ByVal FontType As TFontType, ByVal BaseFont As String, ByVal FontName As String, ByVal Embedded As Boolean, ByVal IsFormFont As Boolean, ByVal Flags As Integer, ByRef DoBreak As Boolean)
        Public Event PDFEnumFont(ByVal FamilyName As String, ByRef PostScriptName As String, ByVal Style As Integer, ByRef DoBreak As Boolean)
        Public Event PDFEnumFontEx(ByVal FamilyName As String, ByVal PostScriptName As String, ByVal Style As Integer, ByVal BaseType As TFontBaseType, ByVal Embeddable As Boolean, ByVal FilePath As String, ByRef DoBreak As Boolean)
        Public Event PDFError(ByVal Description As String, ByVal ErrType As Integer, ByRef DoBreak As Boolean)
        Public Event PDFInitProgress(ByVal ProgType As Integer, ByVal MaxCount As Integer)
        Public Event PDFPageBreak(ByVal LastPosX As Double, ByVal LastPosY As Double, ByVal PageBreak As Integer, ByRef NewAlign As TNewAlign, ByRef DoBreak As Boolean)
        Public Event PDFProgress(ByVal ActivePage As Integer, ByRef DoBreak As Boolean)

        ' ------------------------------------------------- Default types -----------------------------------------------

        ' Error types
        Public Const E_WARNING As Integer = &H2000000
        Public Const E_SYNTAX_ERROR As Integer = &H4000000
        Public Const E_VALUE_ERROR As Integer = &H8000000
        Public Const E_FONT_ERROR As Integer = &H10000000
        Public Const E_FATAL_ERROR As Integer = &H20000000
        Public Const E_FILE_ERROR As Integer = &H40000000

        ' Specific error codes to determine whether the supplied to decrypt the input file password was wrong
        Public Const ENEED_PWD As Integer = -CLng(&HB2S Or E_FILE_ERROR)
        Public Const EWRONG_OPEN_PWD As Integer = -CLng(&HB3S Or E_FILE_ERROR)
        Public Const EWRONG_OWNER_PWD As Integer = -CLng(&HB4S Or E_FILE_ERROR)
        Public Const EWRONG_PWD As Integer = -CLng(&HB5S Or E_FILE_ERROR)

        Public Const PDF_MAX_INT As Integer = &H7FFFFFFF
        ' Basic RGB colors
        Public Const PDF_AQUA As Integer = &HFFFF00
        Public Const PDF_BLACK As Integer = &H0
        Public Const PDF_BLUE As Integer = &HFF0000
        Public Const PDF_CREAM As Integer = &HF0FBFF
        Public Const PDF_DKGRAY As Integer = &H808080
        Public Const PDF_FUCHSIA As Integer = &HFF00FF
        Public Const PDF_GRAY As Integer = &H808080
        Public Const PDF_GREEN As Integer = &H8000
        Public Const PDF_LIME As Integer = &HFF00
        Public Const PDF_LTGRAY As Integer = &HC0C0C0
        Public Const PDF_MAROON As Integer = &H80
        Public Const PDF_MEDGRAY As Integer = &HA4A0A0
        Public Const PDF_MOGREEN As Integer = &HC0DCC0
        Public Const PDF_NAVY As Integer = &H800000
        Public Const PDF_OLIVE As Integer = &H8080
        Public Const PDF_PURPLE As Integer = &H800080
        Public Const PDF_RED As Integer = &HFF
        Public Const PDF_SILVER As Integer = &HC0C0C0
        Public Const PDF_SKYBLUE As Integer = &HF0CAA6
        Public Const PDF_WHITE As Integer = &HFFFFFF
        Public Const PDF_TEAL As Integer = &H808000
        Public Const PDF_YELLOW As Integer = &HFFFF
        Public Const NO_COLOR As Integer = &HFFFFFFF1 ' Transparent color used by annotaions and form fields

        ' Specific return values of the OnPageBreak callback function
        Public Const NEW_ALIGN_LEFT As Integer = 1
        Public Const NEW_ALIGN_RIGHT As Integer = 2
        Public Const NEW_ALIGN_CENTER As Integer = 3
        Public Const NEW_ALIGN_JUSTIFY As Integer = 4

        ' Use these masks to determine which viewer preference values are defined.
        Public Const AV_NON_FULL_SRC_MASK As Integer = &H5S
        Public Const AV_DIRECTION_MASK As Integer = &H18S
        Public Const AV_VIEW_PRINT_MASK As Integer = &H3E0S

        Public Const PDF_ANNOT_INDEX As Integer = &H40000000 ' Special flag for GetPageFieldEx() to indicate that an annotation index
        ' was passed to the function. See GetPageFieldEx() for further information.
        ' This flag can be combined with the annotation handle in Set3DAnnotProps().
        ' 3D Annotations with a transparent background are supported since PDF 1.7, Extension Level 3
        Const TRANSP_3D_ANNOT As Integer = &H40000000

        Private m_Instance As IntPtr
        Private m_AddrEnumFonts As TEnumFontsProc
        Private m_AddrEnumFontsEx As TEnumFontsProcEx
        Private m_AddrEnumDocFonts As TEnumFontsProc2
        Private m_AddrErrorProc As TErrorProc
        Private m_AddrInitProgress As TInitProgress
        Private m_AddrOnFontNoFound As TOnFontNotFoundProc
        Private m_AddrOnReplaceICCProfile As TOnReplaceICCProfile
        Private m_AddrOnPageBreak As TOnPageBreakProc
        Private m_AddrProgress As TProgress

        Public Sub New()
            MyBase.New()
            m_Instance = pdfNewPDF()
            If IntPtr.Zero.Equals(m_Instance) Then Throw New System.Exception("Out of Memory")

            m_AddrEnumFonts = AddressOf pdf_EnumFontsProc
            m_AddrEnumFontsEx = AddressOf pdf_EnumFontsProcEx
            m_AddrEnumDocFonts = AddressOf pdf_EnumDocFontsProc
            m_AddrErrorProc = AddressOf pdf_ErrorProc
            m_AddrInitProgress = AddressOf pdf_InitProgress
            m_AddrProgress = AddressOf pdf_Progress
            m_AddrOnPageBreak = AddressOf pdf_OnPageBreakProc
            pdfSetOnErrorProc(m_Instance, Nothing, m_AddrErrorProc)
            pdfSetProgressProc(m_Instance, Nothing, m_AddrInitProgress, m_AddrProgress)
        End Sub

        Protected Overrides Sub Finalize()
            pdfDeletePDF(m_Instance)
            MyBase.Finalize()
        End Sub

        ' The following callback functions are used to raise events.
        ' FamilyName is a Unicode string, the PostScriptName is an ANSI string!
        Private Function pdf_EnumFontsProc(ByVal Data As IntPtr, ByVal FamilyName As IntPtr, ByVal PostScriptName As IntPtr, ByVal Style As Integer) As Integer
            On Error Resume Next
            Dim doBreak As Boolean
            Dim famName As String
            Dim postName As String
            doBreak = False

            famName = ToString(FamilyName, True)
            postName = ToString(PostScriptName, False)
            RaiseEvent PDFEnumFont(famName, postName, Style, doBreak)
            Return CInt(doBreak = True)
        End Function

        Private Function pdf_EnumFontsProcEx(ByVal Data As IntPtr, ByVal FamilyName As IntPtr, ByVal PostScriptName As IntPtr, ByVal Style As Integer, ByVal BaseType As TFontBaseType, ByVal Flags As TEnumFontProcFlags, ByVal FilePath As IntPtr) As Integer
            Dim doBreak As Boolean
            Dim famName As String
            Dim postName As String
            Dim fPath As String
            Dim uni As Boolean
            Dim embeddable As Boolean
            uni = (Flags And TEnumFontProcFlags.efpUnicodePath) <> 0
            embeddable = (Flags And TEnumFontProcFlags.efpEmbeddable) <> 0
            doBreak = False
            famName = ToString(FamilyName, True)
            postName = ToString(PostScriptName, False)
            fPath = ToString(FilePath, uni)
            RaiseEvent PDFEnumFontEx(famName, postName, Style, BaseType, embeddable, fPath, doBreak)
            Return CInt(doBreak = True)
        End Function

        Private Function pdf_EnumDocFontsProc(ByVal Data As IntPtr, ByVal PDFFont As IntPtr, ByVal FontType As Integer, ByVal BaseFont As IntPtr, ByVal FontName As IntPtr, ByVal Embedded As Integer, ByVal IsFormFont As Integer, ByVal Flags As Integer) As Integer
            Dim doBreak As Boolean
            Dim bFont As String
            Dim fName As String
            bFont = ToString(BaseFont, False)
            fName = ToString(FontName, False)
            RaiseEvent PDFEnumDocFont(PDFFont, CType(FontType, TFontType), bFont, fName, CBool(Embedded), CBool(IsFormFont), Flags, doBreak)
            Return CInt(doBreak = True)
        End Function

        Private Function pdf_ErrorProc(ByVal Data As IntPtr, ByVal ErrCode As Integer, ByVal ErrMessage As IntPtr, ByVal ErrType As Integer) As Integer
            On Error Resume Next
            Dim doBreak As Boolean
            doBreak = False
            RaiseEvent PDFError(ToString(ErrMessage, False), ErrType, doBreak)
            Return CInt(doBreak = True)
        End Function

        Private Function pdf_OnPageBreakProc(ByVal Data As IntPtr, ByVal LastPosX As Double, ByVal LastPosY As Double, ByVal PageBreak As Integer) As Integer
            On Error Resume Next
            Dim doBreak As Boolean
            Dim newAlign As TNewAlign
            doBreak = False
            newAlign = TNewAlign.naUnchanged
            RaiseEvent PDFPageBreak(LastPosX, LastPosY, PageBreak, newAlign, doBreak)
            If doBreak Then
                Return -1
            Else
                Return newAlign
            End If
        End Function

        ' pdf_InitProgress is called before pdf_Progress is called the first time
        Private Sub pdf_InitProgress(ByVal Data As IntPtr, ByVal ProgType As Integer, ByVal MaxCount As Integer)
            On Error Resume Next
            RaiseEvent PDFInitProgress(ProgType, MaxCount)
        End Sub

        Private Function pdf_Progress(ByVal Data As IntPtr, ByVal ActivePage As Integer) As Integer
            On Error Resume Next
            Dim doBreak As Boolean
            doBreak = False
            RaiseEvent PDFProgress(ActivePage, doBreak)
            Return CInt(doBreak = True)
        End Function

        Public Sub Abort(ByVal RasPtr As IntPtr)
            rasAbort(RasPtr)
        End Sub

        Public Function AddActionToObj(ByVal ObjType As TObjType, ByVal ObjEvent As TObjEvent, ByVal ActHandle As Integer, ByVal ObjHandle As Integer) As Integer
            Return pdfAddActionToObj(m_Instance, ObjType, ObjEvent, ActHandle, ObjHandle)
        End Function

        Public Function AddAnnotToPage(ByVal PageNum As Integer, ByVal Handle As Integer) As Integer
            Return pdfAddAnnotToPage(m_Instance, PageNum, Handle)
        End Function

        Public Function AddArticle(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfAddArticle(m_Instance, PosX, PosY, Width, Height))
        End Function

        Public Function AddBookmark(ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean) As Integer
            Return pdfAddBookmarkW(m_Instance, Title, Parent, DestPage, CInt(DoOpen))
        End Function

        Public Function AddBookmarkA(ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean) As Integer
            Return pdfAddBookmarkA(m_Instance, Title, Parent, DestPage, CInt(DoOpen))
        End Function

        Public Function AddBookmarkW(ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean) As Integer
            Return pdfAddBookmarkW(m_Instance, Title, Parent, DestPage, CInt(DoOpen))
        End Function

        Public Function AddBookmarkEx(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Integer) As Integer
            Return pdfAddBookmarkExW(m_Instance, Title, Parent, NamedDest, CInt(DoOpen))
        End Function

        Public Function AddBookmarkExA(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Integer) As Integer
            Return pdfAddBookmarkExA(m_Instance, Title, Parent, NamedDest, CInt(DoOpen))
        End Function

        Public Function AddBookmarkExW(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Integer) As Integer
            Return pdfAddBookmarkExW(m_Instance, Title, Parent, NamedDest, CInt(DoOpen))
        End Function

        Public Function AddBookmarkEx2(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As String, ByVal Unicode As Boolean, ByVal DoOpen As Boolean) As Integer
            If Unicode Then
                Return pdfAddBookmarkEx2WW(m_Instance, Title, Parent, NamedDest, 1, CInt(DoOpen))
            Else
                Return pdfAddBookmarkEx2WA(m_Instance, Title, Parent, NamedDest, 0, CInt(DoOpen))
            End If
        End Function

        Public Function AddBookmarkEx2A(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As String, ByVal Unicode As Boolean, ByVal DoOpen As Boolean) As Integer
            If Unicode Then
                Return pdfAddBookmarkEx2AW(m_Instance, Title, Parent, NamedDest, 1, CInt(DoOpen))
            Else
                Return pdfAddBookmarkEx2AA(m_Instance, Title, Parent, NamedDest, 0, CInt(DoOpen))
            End If
        End Function

        Public Function AddBookmarkEx2W(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As String, ByVal Unicode As Boolean, ByVal DoOpen As Boolean) As Integer
            If Unicode Then
                Return pdfAddBookmarkEx2WW(m_Instance, Title, Parent, NamedDest, 1, CInt(DoOpen))
            Else
                Return pdfAddBookmarkEx2WA(m_Instance, Title, Parent, NamedDest, 0, CInt(DoOpen))
            End If
        End Function

        Public Function AddButtonImage(ByVal BtnHandle As Integer, ByVal State As TButtonState, ByVal Caption As String, ByVal ImgFile As String) As Boolean
            Return CBool(pdfAddButtonImage(m_Instance, BtnHandle, State, Caption, ImgFile))
        End Function

        Public Function AddButtonImageEx(ByVal BtnHandle As Integer, ByVal State As TButtonState, ByVal Caption As String, ByVal hBitmap As IntPtr) As Boolean
            Return CBool(pdfAddButtonImageEx(m_Instance, BtnHandle, State, Caption, hBitmap))
        End Function

        Public Function AddContinueText(ByVal AText As String) As Boolean
            Return CBool(pdfAddContinueTextW(m_Instance, AText))
        End Function

        Public Function AddContinueTextA(ByVal AText As String) As Boolean
            Return CBool(pdfAddContinueTextA(m_Instance, AText))
        End Function

        Public Function AddContinueTextW(ByVal AText As String) As Boolean
            Return CBool(pdfAddContinueTextW(m_Instance, AText))
        End Function

        Public Function AddDeviceNProcessColorants(ByVal DeviceNCS As Integer, ByVal Colorants() As String, ByVal NumColorants As Integer, ByVal ProcessCS As TExtColorSpace, ByVal Handle As Integer) As Boolean
            Return CBool(pdfAddDeviceNProcessColorants(m_Instance, DeviceNCS, Colorants, NumColorants, ProcessCS, Handle))
        End Function

        Public Function AddDeviceNSeparations(ByVal DeviceNCS As Integer, ByVal Colorants() As String, ByVal SeparationCS() As Integer, ByVal NumColorants As Integer) As Boolean
            Return CBool(pdfAddDeviceNSeparations(m_Instance, DeviceNCS, Colorants, SeparationCS, NumColorants))
        End Function

        Public Function AddFieldToFormAction(ByVal Action As Integer, ByVal Field As Integer, ByVal DoInclude As Boolean) As Boolean
            Return CBool(pdfAddFieldToFormAction(m_Instance, Action, Field, CInt(DoInclude)))
        End Function

        Public Function AddFieldToHideAction(ByVal HideAct As Integer, ByVal Field As Integer) As Boolean
            Return CBool(pdfAddFieldToHideAction(m_Instance, HideAct, Field))
        End Function

        Public Function AddFileCommentA(ByVal AText As String) As Boolean
            Return CBool(pdfAddFileCommentA(m_Instance, AText))
        End Function

        Public Function AddFileCommentW(ByVal AText As String) As Boolean
            Return CBool(pdfAddFileCommentW(m_Instance, AText))
        End Function

        Public Function AddFontSearchPath(ByVal APath As String, ByVal Recursive As Integer) As Integer
            Return pdfAddFontSearchPathW(m_Instance, APath, Recursive)
        End Function

        Public Function AddFontSearchPathA(ByVal APath As String, ByVal Recursive As Integer) As Integer
            Return pdfAddFontSearchPathA(m_Instance, APath, Recursive)
        End Function

        Public Function AddFontSearchPathW(ByVal APath As String, ByVal Recursive As Integer) As Integer
            Return pdfAddFontSearchPathW(m_Instance, APath, Recursive)
        End Function

        Public Function AddImage(ByVal Filter As TCompressionFilter, ByVal Flags As TImageConversionFlags, ByRef Image As TPDFImage) As Boolean
            Return CBool(pdfAddImage(m_Instance, Filter, Flags, Image))
        End Function

        Public Function AddInkList(ByVal InkAnnot As Integer, ByRef Points() As TFltPoint) As Integer
            Return pdfAddInkList(m_Instance, InkAnnot, Points, Points.Length)
        End Function

        Public Function AddJavaScript(ByVal Name As String, ByVal Script As String) As Integer
            Return pdfAddJavaScriptW(m_Instance, Name, Script)
        End Function

        Public Function AddJavaScriptA(ByVal Name As String, ByVal Script As String) As Integer
            Return pdfAddJavaScriptA(m_Instance, Name, Script)
        End Function

        Public Function AddJavaScriptW(ByVal Name As String, ByVal Script As String) As Integer
            Return pdfAddJavaScriptW(m_Instance, Name, Script)
        End Function

        Public Function AddLayerToDisplTree(ByVal Parent As IntPtr, ByVal Layer As Integer, ByVal Title As String) As IntPtr
            Return pdfAddLayerToDisplTreeW(m_Instance, Parent, Layer, Title)
        End Function

        Public Function AddLayerToDisplTreeA(ByVal Parent As IntPtr, ByVal Layer As Integer, ByVal Title As String) As IntPtr
            Return pdfAddLayerToDisplTreeA(m_Instance, Parent, Layer, Title)
        End Function

        Public Function AddLayerToDisplTreeW(ByVal Parent As IntPtr, ByVal Layer As Integer, ByVal Title As String) As IntPtr
            Return pdfAddLayerToDisplTreeW(m_Instance, Parent, Layer, Title)
        End Function

        Public Function AddMaskImage(ByVal BaseImage As Integer, ByRef Buffer() As Byte, ByVal BufSize As Integer, ByVal Stride As Integer, ByVal BitsPerPixel As Integer, ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Return CBool(pdfAddMaskImage(m_Instance, BaseImage, Buffer, BufSize, Stride, BitsPerPixel, Width, Height))
        End Function

        Public Function AddMaskImage(ByVal BaseImage As Integer, ByVal Buffer As IntPtr, ByVal BufSize As Integer, ByVal Stride As Integer, ByVal BitsPerPixel As Integer, ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Return CBool(pdfAddMaskImage(m_Instance, BaseImage, Buffer, BufSize, Stride, BitsPerPixel, Width, Height))
        End Function

        Public Function AddObjectToLayer(ByVal OCG As Integer, ByVal ObjType As TOCObject, ByVal Handle As Integer) As Boolean
            Return CBool(pdfAddObjectToLayer(m_Instance, OCG, ObjType, Handle))
        End Function

        Public Function AddOCGToAppEvent(ByVal Handle As Integer, ByVal Events As TOCAppEvent, ByVal Categories As TOCGUsageCategory) As Boolean
            Return CBool(pdfAddOCGToAppEvent(m_Instance, Handle, Events, Categories))
        End Function

        Public Function AddOutputIntent(ByVal ICCFile As String) As Integer
            Return pdfAddOutputIntentW(m_Instance, ICCFile)
        End Function

        Public Function AddOutputIntentA(ByVal ICCFile As String) As Integer
            Return pdfAddOutputIntentA(m_Instance, ICCFile)
        End Function

        Public Function AddOutputIntentW(ByVal ICCFile As String) As Integer
            Return pdfAddOutputIntentW(m_Instance, ICCFile)
        End Function

        Public Function AddOutputIntentEx(ByRef Buffer() As Byte) As Integer
            Return pdfAddOutputIntentEx(m_Instance, Buffer, Buffer.Length)
        End Function

        Public Function AddPageLabel(ByVal StartRange As Integer, ByVal Format As TPageLabelFormat, ByVal Prefix As String, ByVal AddNum As Integer) As Integer
            Return pdfAddPageLabelW(m_Instance, StartRange, Format, Prefix, AddNum)
        End Function

        Public Function AddPageLabelA(ByVal StartRange As Integer, ByVal Format As TPageLabelFormat, ByVal Prefix As String, ByVal AddNum As Integer) As Integer
            Return pdfAddPageLabelA(m_Instance, StartRange, Format, Prefix, AddNum)
        End Function

        Public Function AddPageLabelW(ByVal StartRange As Integer, ByVal Format As TPageLabelFormat, ByVal Prefix As String, ByVal AddNum As Integer) As Integer
            Return pdfAddPageLabelW(m_Instance, StartRange, Format, Prefix, AddNum)
        End Function

        Public Function AddRasImage(ByVal RasPtr As IntPtr, ByVal Filter As TCompressionFilter) As Boolean
            Return CBool(pdfAddRasImage(m_Instance, RasPtr, Filter))
        End Function

        ' These functions were incorrectly named. Please use AddOutputIntent() instead.
        Public Function AddRenderingIntent(ByVal ICCFile As String) As Integer
            Return pdfAddRenderingIntentW(m_Instance, ICCFile)
        End Function

        Public Function AddRenderingIntentA(ByVal ICCFile As String) As Integer
            Return pdfAddRenderingIntentA(m_Instance, ICCFile)
        End Function

        Public Function AddRenderingIntentW(ByVal ICCFile As String) As Integer
            Return pdfAddRenderingIntentW(m_Instance, ICCFile)
        End Function

        Public Function AddRenderingIntentEx(ByRef Buffer() As Byte) As Integer
            Return pdfAddRenderingIntentEx(m_Instance, Buffer, Buffer.Length)
        End Function
        ' -----------------------------------------------------------------------------

        Public Function AddValToChoiceField(ByVal Field As Integer, ByVal ExpValue As String, ByVal Value As String, ByVal Selected As Boolean) As Boolean
            Return CBool(pdfAddValToChoiceFieldW(m_Instance, Field, ExpValue, Value, CInt(Selected)))
        End Function

        Public Function AddValToChoiceFieldA(ByVal Field As Integer, ByVal ExpValue As String, ByVal Value As String, ByVal Selected As Boolean) As Boolean
            Return CBool(pdfAddValToChoiceFieldA(m_Instance, Field, ExpValue, Value, CInt(Selected)))
        End Function

        Public Function AddValToChoiceFieldW(ByVal Field As Integer, ByVal ExpValue As String, ByVal Value As String, ByVal Selected As Boolean) As Boolean
            Return CBool(pdfAddValToChoiceFieldW(m_Instance, Field, ExpValue, Value, CInt(Selected)))
        End Function

        Public Function Append() As Boolean
            Return CBool(pdfAppend(m_Instance))
        End Function

        Public Function ApplyAppEvent(ByVal AppEvent As TOCAppEvent, ByVal SaveResult As Boolean) As Boolean
            Return CBool(pdfApplyAppEvent(m_Instance, AppEvent, Convert.ToInt32(SaveResult)))
        End Function

        Public Function ApplyPattern(ByVal PattHandle As Integer, ByVal ColorMode As TColorMode, ByVal Color As Integer) As Boolean
            Return CBool(pdfApplyPattern(m_Instance, PattHandle, ColorMode, Color))
        End Function

        Public Function ApplyShading(ByVal ShadHandle As Integer) As Boolean
            Return CBool(pdfApplyShading(m_Instance, ShadHandle))
        End Function

        Public Function AssociateEmbFile(ByVal DestObject As TAFDestObject, ByVal DestHandle As Integer, ByVal Relationship As TAFRelationship, ByVal EmbFile As Integer) As Boolean
            Return CBool(pdfAssociateEmbFile(m_Instance, DestObject, DestHandle, Relationship, EmbFile))
        End Function

        Public Function AttachFile(ByVal FilePath As String, ByVal Description As String, ByVal Compress As Boolean) As Integer
            Return pdfAttachFileW(m_Instance, FilePath, Description, CInt(Compress))
        End Function

        Public Function AttachFileA(ByVal FilePath As String, ByVal Description As String, ByVal Compress As Boolean) As Integer
            Return pdfAttachFileA(m_Instance, FilePath, Description, CInt(Compress))
        End Function

        Public Function AttachFileW(ByVal FilePath As String, ByVal Description As String, ByVal Compress As Boolean) As Integer
            Return pdfAttachFileW(m_Instance, FilePath, Description, CInt(Compress))
        End Function

        Public Function AttachFileEx(ByRef Buffer() As Byte, ByVal FileName As String, ByVal Description As String, ByVal Compress As Boolean) As Integer
            Return pdfAttachFileExW(m_Instance, Buffer, Buffer.Length(), FileName, Description, CInt(Compress))
        End Function

        Public Function AttachFileExA(ByRef Buffer() As Byte, ByVal FileName As String, ByVal Description As String, ByVal Compress As Boolean) As Integer
            Return pdfAttachFileExA(m_Instance, Buffer, Buffer.Length(), FileName, Description, CInt(Compress))
        End Function

        Public Function AttachFileExW(ByRef Buffer() As Byte, ByVal FileName As String, ByVal Description As String, ByVal Compress As Boolean) As Integer
            Return pdfAttachFileExW(m_Instance, Buffer, Buffer.Length(), FileName, Description, CInt(Compress))
        End Function

        Public Function AttachImageBuffer(ByVal RasPtr As IntPtr, ByVal Rows As IntPtr, ByVal Buffer As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal ScanlineLen As Integer) As Boolean
            Return CBool(rasAttachImageBuffer(RasPtr, Rows, Buffer, Width, Height, ScanlineLen))
        End Function

        Public Function AutoTemplate(ByVal Templ As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfAutoTemplate(m_Instance, Templ, PosX, PosY, Width, Height))
        End Function

        Public Function BeginClipPath() As Boolean
            Return True ' Obsolete function! Nothing to do.
        End Function

        Public Function BeginContinueText(ByVal PosX As Double, ByVal PosY As Double) As Boolean
            Return CBool(pdfBeginContinueText(m_Instance, PosX, PosY))
        End Function

        Public Function BeginLayer(ByVal OCG As Integer) As Integer
            Return pdfBeginLayer(m_Instance, OCG)
        End Function

        Public Function BeginPageTemplate(ByVal Name As String, ByVal UseAutoTemplates As Boolean) As Boolean
            Return CBool(pdfBeginPageTemplate(m_Instance, Name, Convert.ToInt32(UseAutoTemplates)))
        End Function

        Public Function BeginPattern(ByVal PatternType As TPatternType, ByVal TilingType As TTilingType, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfBeginPattern(m_Instance, PatternType, TilingType, Width, Height)
        End Function

        Public Function BeginTemplate(ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfBeginTemplate(m_Instance, Width, Height)
        End Function

        Public Function BeginTransparencyGroup(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal Isolated As Integer, ByVal Knockout As Integer, ByVal CS As TExtColorSpace, ByVal CSHandle As Integer) As Integer
            Return pdfBeginTransparencyGroup(m_Instance, x1, y1, x2, y2, Isolated, Knockout, CS, CSHandle)
        End Function

        Public Function Bezier_1_2_3(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Boolean
            Return CBool(pdfBezier_1_2_3(m_Instance, x1, y1, x2, y2, x3, y3))
        End Function

        Public Function Bezier_1_3(ByVal x1 As Double, ByVal y1 As Double, ByVal x3 As Double, ByVal y3 As Double) As Boolean
            Return CBool(pdfBezier_1_3(m_Instance, x1, y1, x3, y3))
        End Function

        Public Function BuildFamilyNameAndStyle(ByVal IFont As IntPtr, ByVal Name As System.Text.StringBuilder, ByVal Style As TFStyle) As Boolean
            If Name.Capacity < 128 Then Name.Capacity = 128
            Name.Length = 0
            Return CBool(fntBuildFamilyNameAndStyle(IFont, Name, Style))
        End Function

        Public Function Bezier_2_3(ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Boolean
            Return CBool(pdfBezier_2_3(m_Instance, x2, y2, x3, y3))
        End Function

        Public Sub CalcPagePixelSize(ByVal PagePtr As IntPtr, ByVal DefScale As TPDFPageScale, ByVal ScaleFactor As Single, ByVal FrameWidth As Integer, ByVal FrameHeight As Integer, ByVal Flags As TRasterFlags, ByRef Width As Integer, ByRef Height As Integer)
            rasCalcPagePixelSize(PagePtr, DefScale, ScaleFactor, FrameWidth, FrameHeight, Flags, Width, Height)
        End Sub

        Public Function CalcWidthHeight(ByVal OrgWidth As Double, ByVal OrgHeight As Double, ByVal ScaledWidth As Double, ByVal ScaledHeight As Double) As Double
            Return pdfCalcWidthHeight(m_Instance, OrgWidth, OrgHeight, ScaledWidth, ScaledHeight)
        End Function

        Public Function CaretAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfCaretAnnotW(m_Instance, PosX, PosY, Width, Height, Color, CS, Author, Subject, Content)
        End Function

        Public Function CaretAnnotA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfCaretAnnotA(m_Instance, PosX, PosY, Width, Height, Color, CS, Author, Subject, Content)
        End Function

        Public Function CaretAnnotW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfCaretAnnotW(m_Instance, PosX, PosY, Width, Height, Color, CS, Author, Subject, Content)
        End Function

        Public Function ChangeAnnotName(ByVal Handle As Integer, ByVal Name As String) As Boolean
            Return CBool(pdfChangeAnnotNameW(m_Instance, Handle, Name))
        End Function

        Public Function ChangeAnnotNameA(ByVal Handle As Integer, ByVal Name As String) As Boolean
            Return CBool(pdfChangeAnnotNameA(m_Instance, Handle, Name))
        End Function

        Public Function ChangeAnnotNameW(ByVal Handle As Integer, ByVal Name As String) As Boolean
            Return CBool(pdfChangeAnnotNameW(m_Instance, Handle, Name))
        End Function

        Public Function ChangeAnnotPos(ByVal Handle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfChangeAnnotPos(m_Instance, Handle, PosX, PosY, Width, Height))
        End Function

        Public Function ChangeBookmark(ByVal ABmk As Integer, ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean) As Boolean
            Return CBool(pdfChangeBookmarkW(m_Instance, ABmk, Title, Parent, DestPage, CInt(DoOpen)))
        End Function

        Public Function ChangeBookmarkA(ByVal ABmk As Integer, ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean) As Boolean
            Return CBool(pdfChangeBookmarkA(m_Instance, ABmk, Title, Parent, DestPage, CInt(DoOpen)))
        End Function

        Public Function ChangeBookmarkW(ByVal ABmk As Integer, ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean) As Boolean
            Return CBool(pdfChangeBookmarkW(m_Instance, ABmk, Title, Parent, DestPage, CInt(DoOpen)))
        End Function

        Public Function ChangeFont(ByVal AHandle As Integer) As Boolean
            Return CBool(pdfChangeFont(m_Instance, AHandle))
        End Function

        Public Function ChangeFontSize(ByVal Size As Double) As Boolean
            Return CBool(pdfChangeFontSize(m_Instance, Size))
        End Function

        Public Function ChangeFontStyle(ByVal Style As TFStyle) As Boolean
            Return CBool(pdfChangeFontStyle(m_Instance, Style))
        End Function

        Public Function ChangeFontStyleEx(ByVal Style As TFStyle) As Boolean
            Return CBool(pdfChangeFontStyleEx(m_Instance, Style))
        End Function

        Public Function ChangeJavaScript(ByVal AHandle As Integer, ByVal NewScript As String) As Boolean
            Return CBool(pdfChangeJavaScriptW(m_Instance, AHandle, NewScript))
        End Function

        Public Function ChangeJavaScriptA(ByVal AHandle As Integer, ByVal NewScript As String) As Boolean
            Return CBool(pdfChangeJavaScriptA(m_Instance, AHandle, NewScript))
        End Function

        Public Function ChangeJavaScriptW(ByVal AHandle As Integer, ByVal NewScript As String) As Boolean
            Return CBool(pdfChangeJavaScriptW(m_Instance, AHandle, NewScript))
        End Function

        Public Function ChangeJavaScriptAction(ByVal AHandle As Integer, ByVal NewScript As String) As Boolean
            Return CBool(pdfChangeJavaScriptActionW(m_Instance, AHandle, NewScript))
        End Function

        Public Function ChangeJavaScriptActionA(ByVal AHandle As Integer, ByVal NewScript As String) As Boolean
            Return CBool(pdfChangeJavaScriptActionA(m_Instance, AHandle, NewScript))
        End Function

        Public Function ChangeJavaScriptActionW(ByVal AHandle As Integer, ByVal NewScript As String) As Boolean
            Return CBool(pdfChangeJavaScriptActionW(m_Instance, AHandle, NewScript))
        End Function

        Public Function ChangeJavaScriptName(ByVal AHandle As Integer, ByVal Name As String) As Boolean
            Return CBool(pdfChangeJavaScriptNameW(m_Instance, AHandle, Name))
        End Function

        Public Function ChangeJavaScriptNameA(ByVal AHandle As Integer, ByVal Name As String) As Boolean
            Return CBool(pdfChangeJavaScriptNameA(m_Instance, AHandle, Name))
        End Function

        Public Function ChangeJavaScriptNameW(ByVal AHandle As Integer, ByVal Name As String) As Boolean
            Return CBool(pdfChangeJavaScriptNameW(m_Instance, AHandle, Name))
        End Function

        Public Function ChangeLinkAnnot(ByVal Handle As Integer, ByVal URL As String) As Boolean
            Return CBool(pdfChangeLinkAnnot(m_Instance, Handle, URL))
        End Function

        Public Function ChangeSeparationColor(ByVal CSHandle As Integer, ByVal NewColor As Integer, ByVal Alternate As TExtColorSpace, ByVal AltHandle As Integer) As Boolean
            Return CBool(pdfChangeSeparationColor(m_Instance, CSHandle, NewColor, Alternate, AltHandle))
        End Function

        Public Function CheckCollection() As Boolean
            Return CBool(pdfCheckCollection(m_Instance))
        End Function

        Public Function CheckFieldNames() As Integer
            Return pdfCheckFieldNames(m_Instance)
        End Function

        Public Function CheckConformance(ByVal ConfType As TConformanceType, ByVal Options As TCheckOptions, ByVal UserData As IntPtr, ByVal OnFontNotFound As TOnFontNotFoundProc, ByVal OnReplaceICCProfile As TOnReplaceICCProfile) As Integer
            m_AddrOnFontNoFound = OnFontNotFound
            m_AddrOnReplaceICCProfile = OnReplaceICCProfile
            Return pdfCheckConformance(m_Instance, ConfType, Options, UserData, OnFontNotFound, OnReplaceICCProfile)
        End Function

        Public Function CircleAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfCircleAnnotW(m_Instance, PosX, PosY, Width, Height, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function CircleAnnotA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfCircleAnnotA(m_Instance, PosX, PosY, Width, Height, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function CircleAnnotW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfCircleAnnotW(m_Instance, PosX, PosY, Width, Height, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function ClearAutoTemplates() As Boolean
            Return CBool(pdfClearAutoTemplates(m_Instance))
        End Function

        Public Sub ClearErrorLog()
            pdfClearErrorLog(m_Instance)
        End Sub

        Public Function ClearHostFonts() As Boolean
            Return CBool(pdfClearHostFonts(m_Instance))
        End Function

        Public Function ClipPath(ByVal ClipMode As TClippingMode, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfClipPath(m_Instance, ClipMode, FillMode))
        End Function

        Public Function CloseAndSignFile(ByVal CertFile As String, ByVal Password As String, ByVal Reason As String, ByVal Location As String) As Boolean
            Return CBool(pdfCloseAndSignFile(m_Instance, CertFile, Password, Reason, Location))
        End Function

        Public Function CloseAndSignFileEx(ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As TKeyLen, ByVal Restrict As TRestrictions, ByVal CertFile As String, ByVal Password As String, ByVal Reason As String, ByVal Location As String) As Boolean
            Return CBool(pdfCloseAndSignFileEx(m_Instance, OpenPwd, OwnerPwd, KeyLen, Restrict, CertFile, Password, Reason, Location))
        End Function

        Public Function CloseAndSignFileExt(ByRef SigParms As TPDFSigParms) As Boolean
            Dim p As TPDFSigParms_I = New TPDFSigParms_I()
            p.StructSize = Marshal.SizeOf(p)
            p.ContactInfoW = SigParms.ContactInfo
            p.Encrypt = Convert.ToInt32(SigParms.Encrypt)
            p.HashType = SigParms.HashType
            p.KeyLen = SigParms.KeyLen
            p.LocationW = SigParms.Location
            p.OpenPwd = SigParms.OpenPwd
            p.OwnerPwd = SigParms.OwnerPwd
            p.PKCS7ObjLen = SigParms.PKCS7ObjLen
            p.ReasonW = SigParms.Reason
            p.SignerW = SigParms.Signer
            p.Restrict = SigParms.Restrict
            If (pdfCloseAndSignFileExt(m_Instance, p) = 0) Then Return False
            ReDim SigParms.Range1(p.Range1Len - 1)
            Marshal.Copy(p.Range1, SigParms.Range1, 0, p.Range1Len)
            If (p.Range2Len > 0) Then
                ReDim SigParms.Range2(p.Range2Len - 1)
                Marshal.Copy(p.Range2, SigParms.Range2, 0, p.Range2Len)
            Else
                SigParms.Range2 = Nothing
            End If
            Return True
        End Function

        Public Function CloseFile() As Boolean
            Return CBool(pdfCloseFile(m_Instance))
        End Function

        Public Function CloseFileEx(ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Boolean
            Return CBool(pdfCloseFileEx(m_Instance, OpenPwd, OwnerPwd, KeyLen, Restrict))
        End Function

        Public Function CloseImage() As Boolean
            Return CBool(pdfCloseImage(m_Instance))
        End Function

        Public Function CloseImportFile() As Boolean
            Return CBool(pdfCloseImportFile(m_Instance))
        End Function

        Public Function CloseImportFileEx(ByVal Handle As Integer) As Boolean
            Return CBool(pdfCloseImportFileEx(m_Instance, Handle))
        End Function

        Public Function ClosePath(ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfClosePath(m_Instance, FillMode))
        End Function

        Public Function CloseTag() As Boolean
            Return CBool(pdfCloseTag(m_Instance))
        End Function

        Public Function CMYK(ByVal C As Byte, ByVal M As Byte, ByVal Y As Byte, ByVal K As Byte) As Integer
            If C >= &H80S Then
                Return (K Or (Y * &H100) Or (M * &H10000) Or (CShort(C And &H7FS) * &H1000000) Or &H80000000)
            Else
                Return (K Or (Y * &H100) Or (M * &H10000) Or (C * &H1000000))
            End If
        End Function

        Public Function ComputeBBox(ByRef BBox As TPDFRect, ByVal Flags As TCompBBoxFlags) As Boolean
            Return CBool(pdfComputeBBox(m_Instance, BBox, Flags))
        End Function

        Public Sub ConnectPageBreakEvent(ByVal Connect As Boolean)
            If Connect Then
                pdfSetOnPageBreakProc(m_Instance, Nothing, m_AddrOnPageBreak)
            Else
                pdfSetOnPageBreakProc(m_Instance, Nothing, Nothing)
            End If
        End Sub

        Public Function ConvColor(ByVal Color As IntPtr, ByVal NumComps As Integer, ByVal SourceCS As TExtColorSpace, ByVal IColorSpace As IntPtr, ByVal DestCS As TExtColorSpace) As Integer
            Return pdfConvColor(Color, NumComps, SourceCS, IColorSpace, DestCS)
        End Function

        Public Function ConvertColors(ByVal Flags As TColorConvFlags, Optional ByVal Add As Single = 0.0F) As Boolean
            Dim tmp(0) As Single
            tmp(0) = Add
            Return CBool(pdfConvertColors(m_Instance, Flags, tmp))
        End Function

        Public Function ConvertEMFSpool(ByVal SpoolFile As String, ByVal LeftMargin As Double, ByVal TopMargin As Double, ByVal Flags As TSpoolConvFlags) As Integer
            Return pdfConvertEMFSpool(m_Instance, SpoolFile, LeftMargin, TopMargin, Flags)
        End Function

        Public Function ConvToUnicode(ByVal AString As String, ByVal CP As TCodepage) As String
            Dim retval As String
            Dim pBuffer As IntPtr

            pBuffer = pdfConvToUnicode(m_Instance, AString, CP)
            retval = Marshal.PtrToStringUni(pBuffer)
            Return retval
        End Function

        Public Function CopyChoiceValues(ByVal Source As Integer, ByVal Dest As Integer, ByVal Share As Boolean) As Boolean
            Return CBool(pdfCopyChoiceValues(m_Instance, Source, Dest, Convert.ToInt32(Share)))
        End Function

        Public Function CopyIntPtrArray(ByVal Source As IntPtr, ByVal Count As Integer) As IntPtr()
            If Count < 1 Then Return Nothing
            Dim retval(Count - 1) As IntPtr
            pdfCopyMemIntPtr(Source, retval, Count * IntPtr.Size)
            Return retval
        End Function

        Public Function Create3DAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Name As String, ByVal U3DFile As String, ByVal ImageFile As String) As Integer
            Return pdfCreate3DAnnot(m_Instance, PosX, PosY, Width, Height, Author, Name, U3DFile, ImageFile)
        End Function

        Public Function Create3DBackground(ByVal IView As IntPtr, ByVal BackColor As Integer) As Boolean
            Return CBool(pdfCreate3DBackground(m_Instance, IView, BackColor))
        End Function

        Public Function Create3DGotoViewAction(ByVal Base3DAnnot As Integer, ByVal IView As IntPtr, ByVal Named As T3DNamedAction) As Integer
            Return pdfCreate3DGotoViewAction(m_Instance, Base3DAnnot, IView, Named)
        End Function

        Public Function Create3DProjection(ByVal IView As IntPtr, ByVal ProjType As T3DProjType, ByVal ScaleType As T3DScaleType, ByVal Diameter As Double, ByVal FOV As Double) As Boolean
            Return CBool(pdfCreate3DProjection(m_Instance, IView, ProjType, ScaleType, Diameter, FOV))
        End Function

        Public Function Create3DView(ByVal Base3DAnnot As Integer, ByVal Name As String, ByVal SetAsDefault As Boolean, ByRef Matrix() As Double, ByVal CamDistance As Double, ByVal RM As T3DRenderingMode, ByVal LS As T3DLightingSheme) As IntPtr
            Try
                ' A 3D Matrix contains exactly 12 elements
                If UBound(Matrix) < 11 Then
                    Return pdfCreate3DView(m_Instance, Base3DAnnot, Name, CInt(SetAsDefault), Nothing, CamDistance, RM, LS)
                Else
                    Return pdfCreate3DView(m_Instance, Base3DAnnot, Name, CInt(SetAsDefault), Matrix, CamDistance, RM, LS)
                End If
            Catch
                Return pdfCreate3DView(m_Instance, Base3DAnnot, Name, CInt(SetAsDefault), Nothing, CamDistance, RM, LS)
            End Try
        End Function

        Public Function CreateAnnotAP(ByVal Annot As Integer) As Integer
            Return pdfCreateAnnotAP(m_Instance, Annot)
        End Function

        Public Function CreateArticleThread(ByVal ThreadName As String) As Integer
            Return pdfCreateArticleThreadW(m_Instance, ThreadName)
        End Function

        Public Function CreateArticleThreadA(ByVal ThreadName As String) As Integer
            Return pdfCreateArticleThreadA(m_Instance, ThreadName)
        End Function

        Public Function CreateArticleThreadW(ByVal ThreadName As String) As Integer
            Return pdfCreateArticleThreadW(m_Instance, ThreadName)
        End Function

        Public Function CreateAxialShading(ByVal sx As Double, ByVal sy As Double, ByVal eX As Double, ByVal eY As Double, ByVal SCenter As Double, ByVal SColor As Integer, ByVal EColor As Integer, ByVal Extend1 As Integer, ByVal Extend2 As Integer) As Integer
            Return pdfCreateAxialShading(m_Instance, sx, sy, eX, eY, SCenter, SColor, EColor, Extend1, Extend2)
        End Function

        Public Function CreateBarcodeField(ByVal Name As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByRef Barcode As TPDFBarcode) As Integer
            Dim bc As TPDFBarcode_I
            Dim retval As Integer
            bc.StructSize = Marshal.SizeOf(bc)
            bc.CaptionW = Marshal.StringToHGlobalUni(Barcode.Caption)
            bc.ECC = Barcode.ECC
            bc.Height = Barcode.Height
            bc.nCodeWordCol = Barcode.nCodeWordCol
            bc.nCodeWordRow = Barcode.nCodeWordRow
            bc.Resolution = Barcode.Resolution
            bc.Symbology = Marshal.StringToHGlobalAnsi(Barcode.Symbology)
            bc.Version = Barcode.Version
            bc.Width = Barcode.Width
            bc.XSymHeight = Barcode.XSymHeight
            bc.XSymWidth = Barcode.XSymWidth
            retval = pdfCreateBarcodeField(m_Instance, Name, Parent, PosX, PosY, Width, Height, bc)
            Marshal.FreeHGlobal(bc.CaptionW)
            Marshal.FreeHGlobal(bc.Symbology)
            Return retval
        End Function

        Public Function CreateButton(ByVal Name As String, ByVal Caption As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateButtonW(m_Instance, Name, Caption, Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateButtonA(ByVal Name As String, ByVal Caption As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateButtonA(m_Instance, Name, Caption, Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateButtonW(ByVal Name As String, ByVal Caption As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateButtonW(m_Instance, Name, Caption, Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateCheckBox(ByVal Name As String, ByVal ExpValue As String, ByVal Checked As Boolean, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateCheckBox(m_Instance, Name, ExpValue, CInt(Checked), Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateCIEColorSpace(ByVal Base As TExtColorSpace, ByRef WhitePoint() As Single, ByRef BlackPoint() As Single, ByRef Gamma() As Single, ByRef Matrix() As Single) As Integer
            Return pdfCreateCIEColorSpace(m_Instance, Base, WhitePoint, BlackPoint, Gamma, Matrix)
        End Function

        Public Function CreateCollection(ByVal View As TColView) As Boolean
            Return CBool(pdfCreateCollection(m_Instance, View))
        End Function

        Public Function CreateCollectionField(ByVal ColType As TColColumnType, ByVal Column As Integer, ByVal Name As String, ByVal Key As String, ByVal Visible As Boolean, ByVal Editable As Boolean) As Integer
            Return pdfCreateCollectionFieldW(m_Instance, ColType, Column, Name, Key, CInt(Visible), CInt(Editable))
        End Function

        Public Function CreateCollectionFieldA(ByVal ColType As TColColumnType, ByVal Column As Integer, ByVal Name As String, ByVal Key As String, ByVal Visible As Boolean, ByVal Editable As Boolean) As Integer
            Return pdfCreateCollectionFieldA(m_Instance, ColType, Column, Name, Key, CInt(Visible), CInt(Editable))
        End Function

        Public Function CreateCollectionFieldW(ByVal ColType As TColColumnType, ByVal Column As Integer, ByVal Name As String, ByVal Key As String, ByVal Visible As Boolean, ByVal Editable As Boolean) As Integer
            Return pdfCreateCollectionFieldW(m_Instance, ColType, Column, Name, Key, CInt(Visible), CInt(Editable))
        End Function
        ' 32 Bit Implementation due to the stupid data type handling of .Net!
        Public Function CreateColItemDate(ByVal EmbFile As Integer, ByVal Key As String, ByVal DateVal As Integer, ByVal Prefix As String) As Boolean
            Return CBool(pdfCreateColItemDate(m_Instance, EmbFile, Key, DateVal, Prefix))
        End Function

        Public Function CreateColItemNumber(ByVal EmbFile As Integer, ByVal Key As String, ByVal Value As Double, ByVal Prefix As String) As Boolean
            Return CBool(pdfCreateColItemNumber(m_Instance, EmbFile, Key, Value, Prefix))
        End Function

        Public Function CreateColItemString(ByVal EmbFile As Integer, ByVal Key As String, ByVal Value As String, ByVal Prefix As String) As Boolean
            Return CBool(pdfCreateColItemStringW(m_Instance, EmbFile, Key, Value, Prefix))
        End Function

        Public Function CreateColItemStringA(ByVal EmbFile As Integer, ByVal Key As String, ByVal Value As String, ByVal Prefix As String) As Boolean
            Return CBool(pdfCreateColItemStringA(m_Instance, EmbFile, Key, Value, Prefix))
        End Function

        Public Function CreateColItemStringW(ByVal EmbFile As Integer, ByVal Key As String, ByVal Value As String, ByVal Prefix As String) As Boolean
            Return CBool(pdfCreateColItemStringW(m_Instance, EmbFile, Key, Value, Prefix))
        End Function

        Public Function CreateComboBox(ByVal Name As String, ByVal Sort As Integer, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateComboBox(m_Instance, Name, Sort, Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateDeviceNColorSpace(ByRef Colorants() As String, ByVal PostScriptFunc As String, ByVal Alternate As TExtColorSpace, ByVal Handle As Integer) As Integer
            Return pdfCreateDeviceNColorSpace(m_Instance, Colorants, Colorants.Length, PostScriptFunc, Alternate, Handle)
        End Function

        Public Function CreateExtGState(ByRef GS As TPDFExtGState) As Integer
            Return pdfCreateExtGState(m_Instance, GS)
        End Function

        Public Function CreateGoToAction(ByVal DestType As TDestType, ByVal PageNum As Integer, ByVal a As Double, ByVal b As Double, ByVal C As Double, ByVal d As Double) As Integer
            Return pdfCreateGoToAction(m_Instance, DestType, PageNum, a, b, C, d)
        End Function

        Public Function CreateGoToActionEx(ByVal NamedDest As Integer) As Integer
            Return pdfCreateGoToActionEx(m_Instance, NamedDest)
        End Function

        Public Function CreateGoToEAction(ByVal Location As TEmbFileLocation, ByVal Source As String, ByVal SrcPage As Integer, ByVal Target As String, ByVal DestName As String, ByVal DestPage As Integer, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToEActionW(m_Instance, Location, Source, SrcPage, Target, DestName, DestPage, Convert.ToInt32(NewWindow))
        End Function

        Public Function CreateGoToEActionA(ByVal Location As TEmbFileLocation, ByVal Source As String, ByVal SrcPage As Integer, ByVal Target As String, ByVal DestName As String, ByVal DestPage As Integer, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToEActionA(m_Instance, Location, Source, SrcPage, Target, DestName, DestPage, Convert.ToInt32(NewWindow))
        End Function

        Public Function CreateGoToEActionW(ByVal Location As TEmbFileLocation, ByVal Source As String, ByVal SrcPage As Integer, ByVal Target As String, ByVal DestName As String, ByVal DestPage As Integer, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToEActionW(m_Instance, Location, Source, SrcPage, Target, DestName, DestPage, Convert.ToInt32(NewWindow))
        End Function

        Public Function CreateGoToRAction(ByVal FileName As String, ByVal PageNum As Integer) As Integer
            Return pdfCreateGoToRActionW(m_Instance, FileName, PageNum)
        End Function

        Public Function CreateGoToRActionA(ByVal FileName As String, ByVal PageNum As Integer) As Integer
            Return pdfCreateGoToRActionA(m_Instance, FileName, PageNum)
        End Function

        Public Function CreateGoToRActionW(ByVal FileName As String, ByVal PageNum As Integer) As Integer
            Return pdfCreateGoToRActionW(m_Instance, FileName, PageNum)
        End Function

        Public Function CreateGoToRActionEx(ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToRActionExW(m_Instance, FileName, DestName, CInt(NewWindow))
        End Function

        Public Function CreateGoToRActionExA(ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToRActionExA(m_Instance, FileName, DestName, CInt(NewWindow))
        End Function

        Public Function CreateGoToRActionExW(ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToRActionExW(m_Instance, FileName, DestName, CInt(NewWindow))
        End Function

        Public Function CreateGoToRActionExU(ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToRActionExUW(m_Instance, FileName, DestName, CInt(NewWindow))
        End Function

        Public Function CreateGoToRActionExUA(ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToRActionExUA(m_Instance, FileName, DestName, CInt(NewWindow))
        End Function

        Public Function CreateGoToRActionExUW(ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateGoToRActionExUW(m_Instance, FileName, DestName, CInt(NewWindow))
        End Function

        Public Function CreateGroupField(ByVal Name As String, ByVal Parent As Integer) As Integer
            Return pdfCreateGroupField(m_Instance, Name, Parent)
        End Function

        Public Function CreateHideAction(ByVal Field As Integer, ByVal Hide As Integer) As Integer
            Return pdfCreateHideAction(m_Instance, Field, Hide)
        End Function

        Public Function CreateICCBasedColorSpace(ByVal ICCProfile As String) As Integer
            Return pdfCreateICCBasedColorSpaceW(m_Instance, ICCProfile)
        End Function

        Public Function CreateICCBasedColorSpaceA(ByVal ICCProfile As String) As Integer
            Return pdfCreateICCBasedColorSpaceA(m_Instance, ICCProfile)
        End Function

        Public Function CreateICCBasedColorSpaceW(ByVal ICCProfile As String) As Integer
            Return pdfCreateICCBasedColorSpaceW(m_Instance, ICCProfile)
        End Function

        Public Function CreateICCBasedColorSpaceEx(ByRef Buffer() As Byte) As Integer
            Return pdfCreateICCBasedColorSpaceEx(m_Instance, Buffer, Buffer.Length)
        End Function

        Public Function CreateImage(ByVal FileName As String, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfCreateImageW(m_Instance, FileName, Format))
        End Function

        Public Function CreateImageA(ByVal FileName As String, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfCreateImageA(m_Instance, FileName, Format))
        End Function

        Public Function CreateImageW(ByVal FileName As String, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfCreateImageW(m_Instance, FileName, Format))
        End Function

        Public Function CreateImpDataAction(ByVal DataFile As String) As Integer
            Return pdfCreateImpDataActionW(m_Instance, DataFile)
        End Function

        Public Function CreateImpDataActionA(ByVal DataFile As String) As Integer
            Return pdfCreateImpDataActionA(m_Instance, DataFile)
        End Function

        Public Function CreateImpDataActionW(ByVal DataFile As String) As Integer
            Return pdfCreateImpDataActionW(m_Instance, DataFile)
        End Function

        Public Function CreateIndexedColorSpace(ByVal Base As TExtColorSpace, ByVal Handle As Integer, ByRef ColorTable() As Byte, ByVal NumColors As Integer) As Integer
            CreateIndexedColorSpace = pdfCreateIndexedColorSpace(m_Instance, Base, Handle, ColorTable, NumColors)
        End Function

        Public Function CreateJSAction(ByVal Script As String) As Integer
            Return pdfCreateJSActionW(m_Instance, Script)
        End Function

        Public Function CreateJSActionA(ByVal Script As String) As Integer
            Return pdfCreateJSActionA(m_Instance, Script)
        End Function

        Public Function CreateJSActionW(ByVal Script As String) As Integer
            Return pdfCreateJSActionW(m_Instance, Script)
        End Function

        Public Function CreateLaunchAction(ByVal OP As TFileOP, ByVal FileName As String, ByVal DefDir As String, ByVal Param As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateLaunchAction(m_Instance, OP, FileName, DefDir, Param, Convert.ToInt32(NewWindow))
        End Function

        Public Function CreateLaunchActionEx(ByVal FileName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateLaunchActionExW(m_Instance, FileName, Convert.ToInt32(NewWindow))
        End Function

        Public Function CreateLaunchActionExA(ByVal FileName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateLaunchActionExA(m_Instance, FileName, Convert.ToInt32(NewWindow))
        End Function

        Public Function CreateLaunchActionExW(ByVal FileName As String, ByVal NewWindow As Boolean) As Integer
            Return pdfCreateLaunchActionExW(m_Instance, FileName, Convert.ToInt32(NewWindow))
        End Function

        Public Function CreateListBox(ByVal Name As String, ByVal Sort As Integer, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateListBox(m_Instance, Name, Sort, Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateNamedAction(ByVal Action As TNamedAction) As Integer
            Return pdfCreateNamedAction(m_Instance, Action)
        End Function

        Public Function CreateNamedDest(ByVal Name As String, ByVal DestPage As Integer, ByVal DestType As TDestType, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Integer
            Return pdfCreateNamedDestW(m_Instance, Name, DestPage, DestType, a, b, c, d)
        End Function

        Public Function CreateNamedDestA(ByVal Name As String, ByVal DestPage As Integer, ByVal DestType As TDestType, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Integer
            Return pdfCreateNamedDestA(m_Instance, Name, DestPage, DestType, a, b, c, d)
        End Function

        Public Function CreateNamedDestW(ByVal Name As String, ByVal DestPage As Integer, ByVal DestType As TDestType, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Integer
            Return pdfCreateNamedDestW(m_Instance, Name, DestPage, DestType, a, b, c, d)
        End Function

        Public Function CreateNewPDF(ByVal OutPDF As String) As Boolean
            Return CBool(pdfCreateNewPDFW(m_Instance, OutPDF))
        End Function

        Public Function CreateNewPDFA(ByVal OutPDF As String) As Boolean
            Return CBool(pdfCreateNewPDFA(m_Instance, OutPDF))
        End Function

        Public Function CreateNewPDFW(ByVal OutPDF As String) As Boolean
            Return CBool(pdfCreateNewPDFW(m_Instance, OutPDF))
        End Function

        Public Function CreateOCG(ByVal Name As String, ByVal DisplayInUI As Boolean, ByVal Visible As Boolean, ByVal Intent As TOCGIntent) As Integer
            Return pdfCreateOCGW(m_Instance, Name, Convert.ToInt32(DisplayInUI), Convert.ToInt32(Visible), Intent)
        End Function

        Public Function CreateOCGA(ByVal Name As String, ByVal DisplayInUI As Boolean, ByVal Visible As Boolean, ByVal Intent As TOCGIntent) As Integer
            Return pdfCreateOCGA(m_Instance, Name, Convert.ToInt32(DisplayInUI), Convert.ToInt32(Visible), Intent)
        End Function

        Public Function CreateOCGW(ByVal Name As String, ByVal DisplayInUI As Boolean, ByVal Visible As Boolean, ByVal Intent As TOCGIntent) As Integer
            Return pdfCreateOCGW(m_Instance, Name, Convert.ToInt32(DisplayInUI), Convert.ToInt32(Visible), Intent)
        End Function

        Public Function CreateOCMD(ByVal Visibility As TOCVisibility, ByVal OCGs() As Integer) As Integer
            Return pdfCreateOCMD(m_Instance, Visibility, OCGs, OCGs.Length)
        End Function

        Public Function CreateRadialShading(ByVal sx As Double, ByVal sy As Double, ByVal R1 As Double, ByVal eX As Double, ByVal eY As Double, ByVal R2 As Double, ByVal SCenter As Double, ByVal SColor As Integer, ByVal EColor As Integer, ByVal Extend1 As Integer, ByVal Extend2 As Integer) As Integer
            Return pdfCreateRadialShading(m_Instance, sx, sy, R1, eX, eY, R2, SCenter, SColor, EColor, Extend1, Extend2)
        End Function

        Public Function CreateRadioButton(ByVal Name As String, ByVal ExpValue As String, ByVal Checked As Integer, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateRadioButton(m_Instance, Name, ExpValue, Checked, Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateRasterizer(ByVal Rows As IntPtr, ByVal Buffer As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal ScanlineLen As Integer, ByVal PixFmt As TPDFPixFormat) As IntPtr
            Return rasCreateRasterizer(m_Instance, Rows, Buffer, Width, Height, ScanlineLen, PixFmt)
        End Function

        Public Function CreateRasterizerEx(ByVal DC As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal PixFmt As TPDFPixFormat) As IntPtr
            Return rasCreateRasterizerEx(m_Instance, DC, Width, Height, PixFmt)
        End Function

        Public Function CreateResetAction() As Integer
            Return pdfCreateResetAction(m_Instance)
        End Function

        Public Function CreateSeparationCS(ByVal Colorant As String, ByVal Alternate As TExtColorSpace, ByVal Handle As Integer, ByVal Color As Integer) As Integer
            Return pdfCreateSeparationCS(m_Instance, Colorant, Alternate, Handle, Color)
        End Function

        Public Function CreateSetOCGStateAction(ByVal _On() As Integer, ByVal Off() As Integer, ByVal Toggle() As Integer, ByVal PreserveRB As Boolean) As Integer
            Dim numOn As Integer
            Dim numOff As Integer
            Dim numToggle As Integer
            If Not IsNothing(_On) Then
                numOn = _On.Length
            End If
            If Not IsNothing(Off) Then
                numOff = Off.Length
            End If
            If Not IsNothing(Toggle) Then
                numToggle = Toggle.Length
            End If
            Return pdfCreateSetOCGStateAction(m_Instance, _On, numOn, Off, numOff, Toggle, numToggle, Convert.ToInt32(PreserveRB))
        End Function

        Public Function CreateSigField(ByVal Name As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateSigField(m_Instance, Name, Parent, PosX, PosY, Width, Height)
        End Function

        Public Function CreateSigFieldAP(ByVal SigField As Integer) As Integer
            Return pdfCreateSigFieldAP(m_Instance, SigField)
        End Function

        Public Function CreateSoftMask(ByVal TranspGroup As Integer, ByVal MaskType As TSoftMaskType, ByVal BackColor As Integer) As Integer
            Return pdfCreateSoftMask(m_Instance, TranspGroup, MaskType, BackColor)
        End Function

        Public Function CreateStdPattern(ByVal Pattern As TStdPattern, ByVal LineWidth As Double, ByVal Distance As Double, ByVal LineColor As Integer, ByVal BackColor As Integer) As Integer
            Return pdfCreateStdPattern(m_Instance, Pattern, LineWidth, Distance, LineColor, BackColor)
        End Function

        Public Function CreateStructureTree() As Boolean
            Return CBool(pdfCreateStructureTree(m_Instance))
        End Function

        Public Function CreateSubmitAction(ByVal Flags As TSubmitFlags, ByVal URL As String) As Integer
            Return pdfCreateSubmitAction(m_Instance, Flags, URL)
        End Function

        Public Function CreateTextField(ByVal Name As String, ByVal Parent As Integer, ByVal Multiline As Integer, ByVal MaxLen As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfCreateTextField(m_Instance, Name, Parent, Multiline, MaxLen, PosX, PosY, Width, Height)
        End Function

        Public Function CreateURIAction(ByVal URL As String) As Integer
            Return pdfCreateURIAction(m_Instance, URL)
        End Function

        Public Function DecryptPDF(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal Password As String) As Integer
            Return pdfDecryptPDFW(m_Instance, FileName, PwdType, Password)
        End Function

        Public Function DecryptPDFA(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal Password As String) As Integer
            Return pdfDecryptPDFA(m_Instance, FileName, PwdType, Password)
        End Function

        Public Function DecryptPDFW(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal Password As String) As Integer
            Return pdfDecryptPDFW(m_Instance, FileName, PwdType, Password)
        End Function

        Public Sub DeleteAcroForm()
            pdfDeleteAcroForm(m_Instance)
        End Sub

        Public Function DeleteActionFromObj(ByVal ObjType As TObjType, ByVal ActHandle As Integer, ByVal ObjHandle As Integer) As Boolean
            Return CBool(pdfDeleteActionFromObj(m_Instance, ObjType, ActHandle, ObjHandle))
        End Function

        Public Function DeleteActionFromObjEx(ByVal ObjType As TObjType, ByVal ObjHandle As Integer, ByVal ActIndex As Integer) As Boolean
            Return CBool(pdfDeleteActionFromObjEx(m_Instance, ObjType, ObjHandle, ActIndex))
        End Function

        Public Function DeleteAnnotation(ByVal Handle As Integer) As Boolean
            Return CBool(pdfDeleteAnnotation(m_Instance, Handle))
        End Function

        Public Function DeleteAnnotationFromPage(ByVal PageNum As Integer, ByVal Handle As Integer) As Boolean
            Return CBool(pdfDeleteAnnotationFromPage(m_Instance, PageNum, Handle))
        End Function

        Public Function DeleteAppEvents(ByVal ApplyEvent As Boolean, ByVal AppEvent As TOCAppEvent) As Integer
            Return pdfDeleteAppEvents(m_Instance, Convert.ToInt32(ApplyEvent), AppEvent)
        End Function

        Public Function DeleteBookmark(ByVal ABmk As Integer) As Integer
            Return pdfDeleteBookmark(m_Instance, ABmk)
        End Function

        Public Function DeleteEmbeddedFile(ByVal Handle As Integer) As Boolean
            Return CBool(pdfDeleteEmbeddedFile(m_Instance, Handle))
        End Function

        Public Function DeleteField(ByVal Field As Integer) As Boolean
            Return CBool(pdfDeleteField(m_Instance, Field))
        End Function

        Public Function DeleteFieldEx(ByVal Name As String) As Boolean
            Return CBool(pdfDeleteFieldEx(m_Instance, Name))
        End Function

        Public Sub DeleteJavaScripts(ByVal DelJavaScriptActions As Boolean)
            pdfDeleteJavaScripts(m_Instance, CInt(DelJavaScriptActions))
        End Sub

        Public Function DeleteOCGFromAppEvent(ByVal Handle As Integer, ByVal Events As TOCAppEvent, ByVal Categories As TOCGUsageCategory, ByVal DelCategoryOnly As Boolean) As Boolean
            Return CBool(pdfDeleteOCGFromAppEvent(m_Instance, Handle, Events, Categories, Convert.ToInt32(DelCategoryOnly)))
        End Function

        Public Function DeleteOutputIntent(ByVal Index As Integer) As Integer
            Return pdfDeleteOutputIntent(m_Instance, Index)
        End Function

        Public Function DeletePage(ByVal PageNum As Integer) As Integer
            Return pdfDeletePage(m_Instance, PageNum)
        End Function

        Public Sub DeletePageLabels()
            pdfDeletePageLabels(m_Instance)
        End Sub

        Public Sub DeleteRasterizer(ByRef RasPtr As IntPtr)
            rasDeleteRasterizer(RasPtr)
        End Sub

        Public Function DeleteSeparationInfo(ByVal AllPages As Boolean) As Boolean
            Return CBool(pdfDeleteSeparationInfo(m_Instance, CInt(AllPages)))
        End Function

        Public Function DeleteTemplate(ByVal Handle As Integer) As Boolean
            Return CBool(pdfDeleteTemplate(m_Instance, Handle))
        End Function

        Public Function DeleteTemplateEx(ByVal Index As Integer) As Boolean
            Return CBool(pdfDeleteTemplateEx(m_Instance, Index))
        End Function

        Public Sub DeleteXFAForm()
            pdfDeleteXFAForm(m_Instance)
        End Sub

        Public Function DrawArc(ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal StartAngle As Double, ByVal EndAngle As Double) As Boolean
            Return CBool(pdfDrawArc(m_Instance, PosX, PosY, Radius, StartAngle, EndAngle))
        End Function

        Public Function DrawArcEx(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal StartAngle As Double, ByVal EndAngle As Double) As Boolean
            Return CBool(pdfDrawArcEx(m_Instance, PosX, PosY, Width, Height, StartAngle, EndAngle))
        End Function

        Public Function DrawChord(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal StartAngle As Double, ByVal EndAngle As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfDrawChord(m_Instance, PosX, PosY, Width, Height, StartAngle, EndAngle, FillMode))
        End Function

        Public Function DrawCircle(ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfDrawCircle(m_Instance, PosX, PosY, Radius, FillMode))
        End Function

        Public Function DrawPie(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal StartAngle As Double, ByVal EndAngle As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfDrawPie(m_Instance, PosX, PosY, Width, Height, StartAngle, EndAngle, FillMode))
        End Function

        Public Function EditPage(ByVal PageNum As Integer) As Boolean
            Return CBool(pdfEditPage(m_Instance, PageNum))
        End Function

        Public Function EditTemplate(ByVal Index As Integer) As Boolean
            Return CBool(pdfEditTemplate(m_Instance, Index))
        End Function

        Public Function EditTemplate2(ByVal Handle As Integer) As Boolean
            Return CBool(pdfEditTemplate2(m_Instance, Handle))
        End Function

        Public Function Ellipse(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfEllipse(m_Instance, PosX, PosY, Width, Height, FillMode))
        End Function

        Public Function EncryptPDF(ByVal FileName As String, ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Integer
            Return pdfEncryptPDFW(m_Instance, FileName, OpenPwd, OwnerPwd, KeyLen, Restrict)
        End Function

        Public Function EncryptPDFA(ByVal FileName As String, ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Integer
            Return pdfEncryptPDFA(m_Instance, FileName, OpenPwd, OwnerPwd, KeyLen, Restrict)
        End Function

        Public Function EncryptPDFW(ByVal FileName As String, ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Integer
            Return pdfEncryptPDFW(m_Instance, FileName, OpenPwd, OwnerPwd, KeyLen, Restrict)
        End Function

        Public Function EndContinueText() As Boolean
            Return CBool(pdfEndContinueText(m_Instance))
        End Function

        Public Function EndPage() As Boolean
            Return CBool(pdfEndPage(m_Instance))
        End Function

        Public Function EndLayer() As Boolean
            EndLayer = CBool(pdfEndLayer(m_Instance))
        End Function

        Public Function EndPattern() As Boolean
            Return CBool(pdfEndPattern(m_Instance))
        End Function

        Public Function EndTemplate() As Boolean
            Return CBool(pdfEndTemplate(m_Instance))
        End Function

        Public Function EnumDocFonts() As Integer ' With events
            EnumDocFonts = pdfEnumDocFonts(m_Instance, IntPtr.Zero, m_AddrEnumDocFonts)
        End Function

        Public Function EnumDocFonts(ByVal EnumProc As TEnumFontsProc2) As Integer ' With callback function
            m_AddrEnumDocFonts = EnumProc
            EnumDocFonts = pdfEnumDocFonts(m_Instance, IntPtr.Zero, m_AddrEnumDocFonts)
        End Function

        Public Function EnumHostFonts() As Integer ' With events
            Return pdfEnumHostFonts(m_Instance, IntPtr.Zero, m_AddrEnumFonts)
        End Function

        Public Function EnumHostFonts(ByVal EnumProc As TEnumFontsProc) As Integer ' With callback function
            m_AddrEnumFonts = EnumProc
            Return pdfEnumHostFonts(m_Instance, IntPtr.Zero, m_AddrEnumFonts)
        End Function

        Public Function EnumHostFontsEx() As Integer ' With events
            Return pdfEnumHostFontsEx(m_Instance, IntPtr.Zero, m_AddrEnumFontsEx)
        End Function

        Public Function EnumHostFontsEx(ByVal EnumProc As TEnumFontsProcEx) As Integer ' With callback function
            m_AddrEnumFontsEx = EnumProc
            Return pdfEnumHostFontsEx(m_Instance, IntPtr.Zero, m_AddrEnumFontsEx)
        End Function

        Public Function ExchangeBookmarks(ByVal Bmk1 As Integer, ByVal Bmk2 As Integer) As Boolean
            Return CBool(pdfExchangeBookmarks(m_Instance, Bmk1, Bmk2))
        End Function

        Public Function ExchangePages(ByVal First As Integer, ByVal Second As Integer) As Boolean
            Return CBool(pdfExchangePages(m_Instance, First, Second))
        End Function

        Public Function ExtractText(ByVal PageNum As Integer, ByVal Flags As TTextExtractionFlags, ByRef Area As TFltRect, ByRef Text As String) As Boolean
            Dim txt As IntPtr, txtLen As Integer
            ExtractText = CBool(pdfExtractText(m_Instance, PageNum, Flags, Area, txt, txtLen))
            If ExtractText Then
                Text = Marshal.PtrToStringUni(txt, txtLen)
            Else
                Text = Nothing
            End If
        End Function

        Public Function FileAttachAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Icon As TFileAttachIcon, ByVal Author As String, ByVal Desc As String, ByVal AFile As String, ByVal Compress As Boolean) As Integer
            Return pdfFileAttachAnnotW(m_Instance, PosX, PosY, Icon, Author, Desc, AFile, CInt(Compress))
        End Function

        Public Function FileAttachAnnotA(ByVal PosX As Double, ByVal PosY As Double, ByVal Icon As TFileAttachIcon, ByVal Author As String, ByVal Desc As String, ByVal AFile As String, ByVal Compress As Boolean) As Integer
            Return pdfFileAttachAnnotA(m_Instance, PosX, PosY, Icon, Author, Desc, AFile, CInt(Compress))
        End Function

        Public Function FileAttachAnnotW(ByVal PosX As Double, ByVal PosY As Double, ByVal Icon As TFileAttachIcon, ByVal Author As String, ByVal Desc As String, ByVal AFile As String, ByVal Compress As Boolean) As Integer
            Return pdfFileAttachAnnotW(m_Instance, PosX, PosY, Icon, Author, Desc, AFile, CInt(Compress))
        End Function

        Public Function FileAttachAnnotEx(ByVal PosX As Double, ByVal PosY As Double, ByVal Icon As TFileAttachIcon, ByVal FileName As String, ByVal Author As String, ByVal Desc As String, ByRef Buffer() As Byte, ByVal Compress As Boolean) As Integer
            Return pdfFileAttachAnnotEx(m_Instance, PosX, PosY, Icon, FileName, Author, Desc, Buffer, Buffer.Length, CInt(Compress))
        End Function

        Public Function FileLink(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal AFilePath As String) As Integer
            Return pdfFileLinkW(m_Instance, PosX, PosY, Width, Height, AFilePath)
        End Function

        Public Function FileLinkA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal AFilePath As String) As Integer
            Return pdfFileLinkA(m_Instance, PosX, PosY, Width, Height, AFilePath)
        End Function

        Public Function FileLinkW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal AFilePath As String) As Integer
            Return pdfFileLinkW(m_Instance, PosX, PosY, Width, Height, AFilePath)
        End Function

        Public Function FindBookmark(ByVal DestPage As Integer, ByVal Title As String) As Integer
            Return pdfFindBookmarkW(m_Instance, DestPage, Title)
        End Function

        Public Function FindBookmarkA(ByVal DestPage As Integer, ByVal Title As String) As Integer
            Return pdfFindBookmarkA(m_Instance, DestPage, Title)
        End Function

        Public Function FindBookmarkW(ByVal DestPage As Integer, ByVal Title As String) As Integer
            Return pdfFindBookmarkW(m_Instance, DestPage, Title)
        End Function

        Public Function pdfFindEmbeddedFile(ByVal Name As String) As Integer
            Return pdfFindEmbeddedFile(m_Instance, Name)
        End Function

        Public Function FindField(ByVal Name As String) As Integer
            Return pdfFindFieldW(m_Instance, Name)
        End Function

        Public Function FindFieldA(ByVal Name As String) As Integer
            Return pdfFindFieldA(m_Instance, Name)
        End Function

        Public Function FindFieldW(ByVal Name As String) As Integer
            Return pdfFindFieldW(m_Instance, Name)
        End Function

        Public Function FindLinkAnnot(ByVal URL As String) As Integer
            Return pdfFindLinkAnnot(m_Instance, URL)
        End Function

        Public Function FindNextBookmark() As Integer
            Return pdfFindNextBookmark(m_Instance)
        End Function

        Public Function FinishSignature(ByVal PKCS7Obj() As Byte) As Boolean
            Return pdfFinishSignature(m_Instance, PKCS7Obj, PKCS7Obj.Length) <> 0
        End Function

        Public Function FlattenAnnots(ByVal Flags As TAnnotFlattenFlags) As Integer
            Return pdfFlattenAnnots(m_Instance, Flags)
        End Function

        Public Function FlattenForm() As Boolean
            Return CBool(pdfFlattenForm(m_Instance))
        End Function

        Public Function FlushPageContent(ByRef Stack As TPDFStack) As Boolean
            Return CBool(pdfFlushPageContent(m_Instance, Stack))
        End Function

        Public Function FlushPages(ByVal Flags As TFlushPageFlags) As Boolean
            Return CBool(pdfFlushPages(m_Instance, Flags))
        End Function

        Public Sub FreeImageBuffer()
            pdfFreeImageBuffer(m_Instance)
        End Sub

        Public Function FreeImageObj(ByVal Handle As Integer) As Boolean
            Return CBool(pdfFreeImageObj(m_Instance, Handle))
        End Function

        Public Function FreeImageObjEx(ByVal ImagePtr As IntPtr) As Boolean
            Return CBool(pdfFreeImageObjEx(m_Instance, ImagePtr))
        End Function

        Public Function FreePDF() As Boolean
            Return CBool(pdfFreePDF(m_Instance))
        End Function

        Public Function FreeTextAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal AText As String, ByVal Align As TTextAlign) As Integer
            Return pdfFreeTextAnnotW(m_Instance, PosX, PosY, Width, Height, Author, AText, Align)
        End Function

        Public Function FreeTextAnnotA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal AText As String, ByVal Align As TTextAlign) As Integer
            Return pdfFreeTextAnnotA(m_Instance, PosX, PosY, Width, Height, Author, AText, Align)
        End Function

        Public Function FreeTextAnnotW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal AText As String, ByVal Align As TTextAlign) As Integer
            Return pdfFreeTextAnnotW(m_Instance, PosX, PosY, Width, Height, Author, AText, Align)
        End Function

        Public Function FreeUniBuf() As Boolean
            Return CBool(pdfFreeUniBuf(m_Instance))
        End Function

        Public Function Get3DAnnotStream(ByVal Annot As Integer, ByRef Data() As Byte, ByRef SubType As String) As Boolean
            Dim pdata As IntPtr
            Dim size As Integer
            Data = Nothing
            SubType = Nothing
            If CBool(pdfGet3DAnnotStream(m_Instance, Annot, pdata, size, SubType)) Then
                If size > 0 Then
                    ReDim Data(size - 1)
                    Marshal.Copy(pdata, Data, 0, size)
                End If
                Return True
            End If
            Return False
        End Function

        Public Function GetActionCount() As Integer
            Return pdfGetActionCount(m_Instance)
        End Function

        Public Function GetActionHandle(ByVal ObjType As TObjType, ByVal ObjHandle As Integer, ByVal ActIndex As Integer) As Integer
            Return pdfGetActionHandle(m_Instance, ObjType, ObjHandle, ActIndex)
        End Function

        Public Function GetActionType(ByVal ActHandle As Integer) As Integer
            Return pdfGetActionType(m_Instance, ActHandle)
        End Function

        Public Function GetActionTypeEx(ByVal ObjType As TObjType, ByVal ObjHandle As Integer, ByVal ActIndex As Integer) As Integer
            Return pdfGetActionTypeEx(m_Instance, ObjType, ObjHandle, ActIndex)
        End Function

        Public Function GetActiveFont() As Integer
            Return pdfGetActiveFont(m_Instance)
        End Function

        Public Function GetAllocBy() As Integer
            Return pdfGetAllocBy(m_Instance)
        End Function

        Public Function GetAnnot(ByVal Handle As Integer, ByRef Annot As TPDFAnnotation) As Boolean
            Dim retval As TPDFAnnotation_I
            If pdfGetAnnot(m_Instance, Handle, retval) = 0 Then
                Annot = Nothing
                Return False
            End If
            GetIntAnnot(retval, Annot)
            Return True
        End Function

        Public Function GetAnnotBBox(ByVal Handle As Integer, ByRef BBox As TPDFRect) As Boolean
            Return CBool(pdfGetAnnotBBox(m_Instance, Handle, BBox))
        End Function

        Public Function GetAnnotCount() As Integer
            Return pdfGetAnnotCount(m_Instance)
        End Function

        Public Function GetAnnotEx(ByVal Handle As Integer, ByRef Annot As TPDFAnnotationEx) As Boolean
            Dim retval As TPDFAnnotationEx_I
            If pdfGetAnnotEx(m_Instance, Handle, retval) = 0 Then
                Annot = Nothing
                Return False
            End If
            GetIntAnnotEx(retval, Annot)
            Return True
        End Function

        Public Function GetAnnotFlags() As Integer
            Return pdfGetAnnotFlags(m_Instance)
        End Function

        Public Function GetAnnotLink(ByVal Handle As Integer) As String
            Dim pBuffer As IntPtr
            pBuffer = pdfGetAnnotLink(m_Instance, Handle)
            Return ToString(pBuffer, False)
        End Function

        Public Function GetAnnotType(ByVal Handle As Integer) As Integer
            Return pdfGetAnnotType(m_Instance, Handle)
        End Function

        Public Function GetAscent() As Double
            Return pdfGetAscent(m_Instance)
        End Function

        Public Function GetBarcodeDict(ByVal IBarcode As IntPtr, ByRef Barcode As TPDFBarcode) As Boolean
            Dim bc As TPDFBarcode_I
            bc.StructSize = Marshal.SizeOf(bc)
            If pdfGetBarcodeDict(IBarcode, bc) = 0 Then Return False
            GetIntBarcode(bc, Barcode)
            Return True
        End Function

        Public Function GetBBox(ByVal Boundary As TPageBoundary, ByRef BBox As TPDFRect) As Integer
            Return pdfGetBBox(m_Instance, Boundary, BBox)
        End Function

        Public Function GetBidiMode() As TPDFBidiMode
            Return CType(pdfGetBidiMode(m_Instance), TPDFBidiMode)
        End Function

        Public Function GetBookmark(ByVal AHandle As Integer, ByRef Bmk As TBookmark) As Boolean
            Dim b As TBookmark_I
            If pdfGetBookmark(m_Instance, AHandle, b) = 0 Then
                Bmk = Nothing
                Return False
            End If
            Bmk.Color = b.Color
            Bmk.DestPage = b.DestPage
            Bmk.DestPos = b.DestPos
            Bmk.DestType = b.DestType
            Bmk.DoOpen = b.DoOpen <> 0
            Bmk.Parent = b.Parent
            Bmk.Style = b.Style
            Bmk.Title = ToString(b.Title, b.TitleLen, b.bUnicode <> 0)
            Return True
        End Function

        Public Function GetBookmarkCount() As Integer
            Return pdfGetBookmarkCount(m_Instance)
        End Function

        Public Function GetBorderStyle() As Integer
            Return pdfGetBorderStyle(m_Instance)
        End Function

        Public Function GetBuffer() As Byte()
            Dim pBuffer As IntPtr
            Dim BufSize As Integer
            Dim retval() As Byte

            pBuffer = pdfGetBuffer(m_Instance, BufSize)
            If BufSize = 0 Or IntPtr.Zero.Equals(pBuffer) Then Return Nothing
            ReDim retval(BufSize - 1)
            Marshal.Copy(pBuffer, retval, 0, BufSize)
            pdfFreePDF(m_Instance)
            Return retval
        End Function

        Public Function GetCapHeight() As Double
            Return pdfGetCapHeight(m_Instance)
        End Function

        Public Function GetCharacterSpacing() As Double
            Return pdfGetCharacterSpacing(m_Instance)
        End Function

        Public Function GetCheckBoxChar() As Integer
            Return pdfGetCheckBoxChar(m_Instance)
        End Function

        Public Function GetCheckBoxCharEx(ByVal Field As Integer) As Integer
            Return pdfGetCheckBoxCharEx(m_Instance, Field)
        End Function

        Public Function GetCheckBoxDefState(ByVal Field As Integer) As Integer
            Return pdfGetCheckBoxDefState(m_Instance, Field)
        End Function

        Public Function GetCMap(ByVal Index As Integer, ByRef CMap As TPDFCMap) As Boolean
            CMap.StructSize = Marshal.SizeOf(CMap)
            Return CBool(pdfGetCMap(m_Instance, Index, CMap))
        End Function

        Public Function GetCMapCount() As Integer
            Return pdfGetCMapCount(m_Instance)
        End Function

        Public Function GetColorSpace() As Integer
            Return pdfGetColorSpace(m_Instance)
        End Function

        Public Function GetColorSpaceCount() As Integer
            GetColorSpaceCount = pdfGetColorSpaceCount(m_Instance)
        End Function

        Public Function GetColorSpaceObj(ByVal Handle As Integer, ByRef cs As TPDFColorSpaceObj) As Boolean
            Dim ics As New TPDFColorSpaceObj_I()
            If pdfGetColorSpaceObj(m_Instance, Handle, ics) = 0 Then Return False
            Return GetIntColorSpaceObj(ics, cs)
        End Function

        Public Function GetColorSpaceObjEx(ByVal IColorSpace As IntPtr, ByRef cs As TPDFColorSpaceObj) As Boolean
            Dim ics As New TPDFColorSpaceObj_I()
            If pdfGetColorSpaceObjEx(IColorSpace, ics) = 0 Then Return False
            Return GetIntColorSpaceObj(ics, cs)
        End Function

        Public Function GetCompressionFilter() As Integer
            Return pdfGetCompressionFilter(m_Instance)
        End Function

        Public Function GetCompressionLevel() As Integer
            Return pdfGetCompressionLevel(m_Instance)
        End Function

        Public Function GetContent(ByRef Buffer() As Byte) As Boolean
            Dim BufSize As Integer
            Dim pBuffer As IntPtr

            BufSize = pdfGetContent(m_Instance, pBuffer)
            If BufSize <= 0 Then
                Buffer = Nothing
                Return False
            End If
            ReDim Buffer(BufSize - 1)
            Marshal.Copy(pBuffer, Buffer, 0, BufSize)
            Return True
        End Function

        Public Function GetDefBitsPerPixel() As Integer
            Return pdfGetDefBitsPerPixel(m_Instance)
        End Function

        Public Function GetDescent() As Double
            Return pdfGetDescent(m_Instance)
        End Function

        Public Function GetDeviceNAttributes(ByVal IAttributes As IntPtr, ByRef Attributes As TDeviceNAttributes) As Boolean
            Dim i As Integer
            Dim attr As New TDeviceNAttributes_I()
            If pdfGetDeviceNAttributes(IAttributes, attr) = 0 Then Return False
            Attributes.IProcessColorSpace = attr.IProcessColorSpace
            If attr.ProcessColorantsCount > 0 Then
                ReDim Attributes.ProcessColorants(attr.ProcessColorantsCount - 1)
                For i = 0 To attr.ProcessColorantsCount - 1
                    Attributes.ProcessColorants(i) = ToString(attr.ProcessColorants(i), False)
                Next i
            Else
                Attributes.ProcessColorants = Nothing
            End If
            If attr.SeparationsCount > 0 Then
                Dim bytes() As Byte
                ReDim Attributes.Separations(attr.SeparationsCount - 1)
                For i = 0 To attr.SeparationsCount - 1
                    bytes = System.Text.Encoding.Unicode.GetBytes(Marshal.PtrToStringAnsi(attr.ProcessColorants(i)))
                    Attributes.ProcessColorants(i) = System.Text.Encoding.UTF8.GetString(bytes)
                Next i
            Else
                Attributes.Separations = Nothing
            End If
            Attributes.IMixingHints = attr.IMixingHints
            Return True
        End Function

        Public Function GetDocInfo(ByVal DInfo As TDocumentInfo, ByRef Value As String) As Boolean
            Dim pBuffer As IntPtr
            If pdfGetDocInfo(m_Instance, DInfo, pBuffer) <= 0 Then
                Value = Nothing
                Return False
            End If
            Value = ToString(pBuffer, True)
            Return True
        End Function

        Public Function GetDocInfoCount() As Integer
            Return pdfGetDocInfoCount(m_Instance)
        End Function

        Public Function GetDocInfoEx(ByVal Index As Integer, ByRef DInfo As TDocumentInfo, ByRef Key As String, ByRef Value As String) As Boolean
            Dim valName As IntPtr
            Dim keyName As IntPtr
            Dim valLen As Integer
            Dim uni As Integer
            Dim info As Integer
            valLen = pdfGetDocInfoEx(m_Instance, Index, info, keyName, valName, uni)
            If valLen <= 0 Then
                Key = Nothing
                Value = Nothing
                Return False
            End If
            DInfo = CType(info, TDocumentInfo)
            Key = ToString(keyName, False)
            Value = ToString(valName, valLen, uni <> 0)
            Return True
        End Function

        Public Function GetDocUsesTransparency(ByVal Flags As Integer) As Boolean
            Return CBool(pdfGetDocUsesTransparency(m_Instance, Flags))
        End Function

        Public Function GetDrawDirection() As Integer
            Return pdfGetDrawDirection(m_Instance)
        End Function

        Public Function GetDynaPDFVersion() As String
            Return ToString(pdfGetDynaPDFVersion(), False)
        End Function

        Public Function GetDynaPDFVersionInt() As Integer
            Dim retval As Integer = 0, value As Integer = 0, p As Integer = 0, len As Integer
            Dim ver As IntPtr = pdfGetDynaPDFVersion()
            Dim str As String = Marshal.PtrToStringAnsi(ver)
            Dim bval As Byte() = System.Text.Encoding.ASCII.GetBytes(str)
            len = ParseInt(bval, p, value)
            If len = 0 Then Return 0
            p = len + 3
            retval += value * 10000000
            len = ParseInt(bval, p, value)
            If len = 0 Then Return 0
            p += len + 1
            retval += value * 10000
            len = ParseInt(bval, p, value)
            If len = 0 Then Return 0
            retval += value
            Return retval
        End Function

        Public Function GetEmbeddedFile(ByVal Handle As Integer, ByRef FileSpec As TPDFFileSpec, ByVal Decompress As Boolean) As Boolean
            Dim fSpec As TPDFFileSpec_I
            If pdfGetEmbeddedFile(m_Instance, Handle, fSpec, CInt(Decompress)) = 0 Then
                FileSpec = Nothing
                Return False
            End If
            If Not IntPtr.Zero.Equals(fSpec.Buffer) Then
                ReDim FileSpec.Buffer(fSpec.BufSize - 1)
                Marshal.Copy(fSpec.Buffer, FileSpec.Buffer, 0, fSpec.BufSize)
            Else
                FileSpec.Buffer = Nothing
            End If
            FileSpec.ColItem = fSpec.ColItem
            FileSpec.Compressed = fSpec.Compressed <> 0
            FileSpec.FileSize = fSpec.FileSize
            FileSpec.IsURL = fSpec.IsURL <> 0
            If Not IntPtr.Zero.Equals(fSpec.CheckSum) Then
                ReDim FileSpec.CheckSum(15)
                Marshal.Copy(fSpec.CheckSum, FileSpec.CheckSum, 0, 16)
            End If
            FileSpec.CreateDate = ToString(fSpec.CreateDate, False)
            FileSpec.Desc = ToString(fSpec.Desc, fSpec.DescUnicode <> 0)
            FileSpec.FileName = ToString(fSpec.FileName, False)
            FileSpec.MIMEType = ToString(fSpec.MIMEType, False)
            FileSpec.ModDate = ToString(fSpec.ModDate, False)
            FileSpec.Name = ToString(fSpec.Name, fSpec.NameUnicode <> 0)
            FileSpec.UF = ToString(fSpec.UF, fSpec.UFUnicode <> 0)
            Return True
        End Function

        Public Function GetEmbeddedFileCount() As Integer
            Return pdfGetEmbeddedFileCount(m_Instance)
        End Function

        Public Function GetEmbeddedFileNode(ByVal IEF As IntPtr, ByRef F As TPDFEmbFileNode, ByVal Decompress As Boolean) As Boolean
            Dim node As TPDFEmbFileNode_I
            node.StructSize = Marshal.SizeOf(node)
            If pdfGetEmbeddedFileNode(IEF, node, Convert.ToInt32(Decompress)) = 0 Then Return False
            Dim bytes() As Byte
            bytes = System.Text.Encoding.Unicode.GetBytes(Marshal.PtrToStringAnsi(node.Name))
            F.Name = System.Text.Encoding.UTF8.GetString(bytes)
            F.NextNode = node.NextNode
            If node.EF.BufSize > 0 Then
                ReDim F.EF.Buffer(node.EF.BufSize - 1)
                Marshal.Copy(node.EF.Buffer, F.EF.Buffer, 0, node.EF.BufSize)
            Else
                Erase F.EF.Buffer
            End If
            F.EF.ColItem = node.EF.ColItem
            F.EF.Compressed = node.EF.Compressed <> 0
            F.EF.FileSize = node.EF.FileSize
            F.EF.IsURL = node.EF.IsURL <> 0
            If node.EF.CheckSum <> IntPtr.Zero Then
                ReDim F.EF.CheckSum(15)
                Marshal.Copy(node.EF.CheckSum, F.EF.CheckSum, 0, 16)
            Else
                Erase F.EF.CheckSum
            End If
            F.EF.CreateDate = ToString(node.EF.CreateDate, False)
            F.EF.Desc = ToString(node.EF.Desc, node.EF.DescUnicode <> 0)
            F.EF.FileName = ToString(node.EF.FileName, False)
            F.EF.MIMEType = ToString(node.EF.MIMEType, False)
            F.EF.ModDate = ToString(node.EF.ModDate, False)
            F.EF.Name = ToString(node.EF.Name, node.EF.NameUnicode <> 0)
            F.EF.UF = ToString(node.EF.UF, node.EF.UFUnicode <> 0)
            Return True
        End Function

        Public Function GetErrorMessage() As String
            Return ToString(pdfGetErrorMessage(m_Instance), False)
        End Function

        Public Function GetEMFPatternDistance() As Double
            Return pdfGetEMFPatternDistance(m_Instance)
        End Function

        Public Function GetErrLogMessage(ByVal Index As Integer, ByRef Err As TPDFError) As Boolean
            Dim e As TPDFError_I
            e.StructSize = Marshal.SizeOf(e)
            If pdfGetErrLogMessage(m_Instance, Index, e) = 0 Then Return False
            GetIntError(e, Err)
            Return True
        End Function

        Public Function GetErrLogMessageCount() As Integer
            Return pdfGetErrLogMessageCount(m_Instance)
        End Function

        Public Function GetErrorMode() As Integer
            Return pdfGetErrorMode(m_Instance)
        End Function

        Public Function GetField(ByVal AHandle As Integer, ByRef Field As TPDFField) As Boolean
            Dim f As TPDFField_I
            If pdfGetField(m_Instance, AHandle, f) < 0 Then
                Field = Nothing
                Return False
            End If
            GetIntField(f, Field)
            Return True
        End Function

        Public Function GetFieldBackColor() As Integer
            Return pdfGetFieldBackColor(m_Instance)
        End Function

        Public Function GetFieldBorderColor() As Integer
            Return pdfGetFieldBorderColor(m_Instance)
        End Function

        Public Function GetFieldBorderStyle(ByVal Field As Integer) As Integer
            Return pdfGetFieldBorderStyle(m_Instance, Field)
        End Function

        Public Function GetFieldBorderWidth(ByVal Field As Integer) As Double
            Return pdfGetFieldBorderWidth(m_Instance, Field)
        End Function

        Public Function GetFieldChoiceValue(ByVal Field As Integer, ByVal ValIndex As Integer, ByRef Value As TPDFChoiceValue) As Boolean
            Dim v As TPDFChoiceValue_I
            v.StructSize = Marshal.SizeOf(v)
            If pdfGetFieldChoiceValue(m_Instance, Field, ValIndex, v) = 0 Then Return False
            Value.ExpValue = ToString(v.ExpValueA, v.ExpValueW, v.ExpValueLen)
            Value.Value = ToString(v.ValueA, v.ValueW, v.ValueLen)
            Value.Selected = CBool(v.Selected)
            Return True
        End Function

        Public Function GetFieldColor(ByVal Field As Integer, ByVal ColorType As TFieldColor, ByRef ColorSpace As TPDFColorSpace, ByRef Color As Integer) As Boolean
            Dim cs As Integer
            If pdfGetFieldColor(m_Instance, Field, ColorType, cs, Color) = 0 Then
                Color = 0
                ColorSpace = TPDFColorSpace.csDeviceRGB
                Return False
            End If
            ColorSpace = CType(cs, TPDFColorSpace)
            Return True
        End Function

        Public Function GetFieldCount() As Integer
            Return pdfGetFieldCount(m_Instance)
        End Function

        Public Function GetFieldEx(ByVal Handle As Integer, ByRef Field As TPDFFieldEx) As Boolean
            Dim f As TPDFFieldEx_I
            f.StructSize = Marshal.SizeOf(f)
            If pdfGetFieldEx(m_Instance, Handle, f) = 0 Then Return False
            GetIntFieldEx(f, Field)
            Return True
        End Function

        Public Function GetFieldEx2(ByVal IField As IntPtr, ByRef Field As TPDFFieldEx) As Boolean
            Dim f As TPDFFieldEx_I
            f.StructSize = Marshal.SizeOf(f)
            If pdfGetFieldEx2(IField, f) = 0 Then Return False
            GetIntFieldEx(f, Field)
            Return True
        End Function

        Public Function GetFieldExpValCount(ByVal Field As Integer) As Integer
            Return pdfGetFieldExpValCount(m_Instance, Field)
        End Function

        Public Function GetFieldExpValue(ByVal Field As Integer, ByRef Value As String) As Boolean
            Dim pBuffer As IntPtr
            Dim Len As Integer
            Len = pdfGetFieldExpValue(m_Instance, Field, pBuffer)
            Value = ToString(pBuffer, False)
            Return (Len > 0)
        End Function

        Public Function GetFieldExpValueEx(ByVal Field As Integer, ByVal ValIndex As Integer, ByRef Value As String, ByRef ExpValue As String, ByRef bSelected As Integer) As Boolean
            Dim valPtr As IntPtr
            Dim expPtr As IntPtr
            If pdfGetFieldExpValueEx(m_Instance, Field, ValIndex, valPtr, expPtr, bSelected) = 0 Then
                Value = Nothing
                ExpValue = Nothing
                Return False
            End If
            Value = ToString(valPtr, False)
            ExpValue = ToString(expPtr, False)
            Return True
        End Function

        Public Function GetFieldFlags(ByVal Field As Integer) As Integer
            Return pdfGetFieldFlags(m_Instance, Field)
        End Function

        Public Function GetFieldGroupType(ByVal Field As Integer) As Integer
            Return pdfGetFieldGroupType(m_Instance, Field)
        End Function

        Public Function GetFieldHighlightMode(ByVal Field As Integer) As Integer
            Return pdfGetFieldHighlightMode(m_Instance, Field)
        End Function

        Public Function GetFieldIndex(ByVal Field As Integer) As Integer
            Return pdfGetFieldIndex(m_Instance, Field)
        End Function

        Public Function GetFieldMapName(ByVal Field As Integer, ByRef Value As String) As Boolean
            Dim uni As Integer
            Dim pBuffer As IntPtr
            If pdfGetFieldMapName(m_Instance, Field, pBuffer, uni) < 0 Then
                Value = Nothing
                Return False
            End If
            Value = ToString(pBuffer, uni <> 0)
            Return True
        End Function

        Public Function GetFieldName(ByVal Field As Integer, ByRef Name As String) As Boolean
            Dim pBuffer As IntPtr
            Dim asLen As Integer
            asLen = pdfGetFieldName(m_Instance, Field, pBuffer)
            If asLen <= 0 Then
                Name = Nothing
                Return False
            End If
            Dim nameLen As Integer
            nameLen = pdfStrLenA(pBuffer)
            Name = ToString(pBuffer, asLen, asLen <> nameLen)
            Return True
        End Function

        Public Function GetFieldOrientation(ByVal Field As Integer) As Integer
            Return pdfGetFieldOrientation(m_Instance, Field)
        End Function

        Public Function GetFieldTextAlign(ByVal Field As Integer) As TTextAlign
            Dim retval As Integer
            retval = pdfGetFieldTextAlign(m_Instance, Field)
            If retval < 0 Then
                Return TTextAlign.taLeft
            Else
                Return CType(retval, TTextAlign)
            End If
        End Function

        Public Function GetFieldTextColor() As Integer
            Return pdfGetFieldTextColor(m_Instance)
        End Function

        Public Function GetFieldToolTip(ByVal Field As Integer, ByRef Value As String) As Boolean
            Dim uni As Integer
            Dim Len As Integer
            Dim pBuffer As IntPtr
            Len = pdfGetFieldToolTip(m_Instance, Field, pBuffer, uni)
            If Len < 0 Then
                Value = Nothing
                Return False
            End If
            Value = ToString(pBuffer, Len, uni <> 0)
            Return True
        End Function

        Public Function GetFieldType(ByVal Field As Integer, ByRef Value As TFieldType) As Boolean
            Dim retval As Integer
            retval = pdfGetFieldType(m_Instance, Field)
            If retval < 0 Then
                Return False
            Else
                Value = CType(retval, TFieldType)
                Return True
            End If
        End Function

        Public Function GetFileSpec(ByVal IFS As IntPtr, ByRef F As TPDFFileSpecEx) As Boolean
            Dim fs As TPDFFileSpecEx_I
            fs.StructSize = Marshal.SizeOf(fs)
            If pdfGetFileSpec(IFS, fs) = 0 Then Return False
            CopyFileSpecEx(fs, F)
            Return True
        End Function

        Public Function GetFillColor() As Integer
            Return pdfGetFillColor(m_Instance)
        End Function

        Public Function GetFont(ByVal IFont As IntPtr, ByRef F As TPDFFontObj) As Boolean
            Dim font As TPDFFontObj_I = New TPDFFontObj_I()
            If fntGetFont(IFont, font) = 0 Then
                F = Nothing
                Return False
            End If
            GetIntFont(font, F)
            Return True
        End Function

        Public Function GetFontCount() As Integer
            Return pdfGetFontCount(m_Instance)
        End Function

        Public Function GetFontEx(ByVal Handle As Integer, ByRef F As TPDFFontObj) As Boolean
            Dim font As TPDFFontObj_I = New TPDFFontObj_I()
            If pdfGetFontEx(m_Instance, Handle, font) = 0 Then
                F = Nothing
                Return False
            End If
            GetIntFont(font, F)
            Return True
        End Function

        Public Function GetFontInfo(ByVal IFont As IntPtr, ByRef Font As TPDFFontInfo) As Boolean
            Dim fnt As TPDFFontInfo_I
            fnt.StructSize = Marshal.SizeOf(fnt)
            If fntGetFontInfo(IFont, fnt) = 0 Then Return False
            GetIntFontInfo(fnt, Font)
            Return True
        End Function

        Public Function GetFontInfoEx(ByVal Handle As Integer, ByRef Font As TPDFFontInfo) As Boolean
            Dim fnt As TPDFFontInfo_I
            fnt.StructSize = Marshal.SizeOf(fnt)
            If pdfGetFontInfoEx(m_Instance, Handle, fnt) = 0 Then Return False
            GetIntFontInfo(fnt, Font)
            Return True
        End Function

        Public Function GetFontOrigin() As TOrigin
            Return CType(pdfGetFontOrigin(m_Instance), TOrigin)
        End Function

        Public Sub GetFontSearchOrder(ByRef Order() As TFontBaseType)
            pdfGetFontSearchOrder(m_Instance, Order)
        End Sub

        Public Function GetFontSelMode() As TFontSelMode
            Return CType(pdfGetFontSelMode(m_Instance), TFontSelMode)
        End Function

        Public Function GetFontWeight() As Integer
            Return pdfGetFontWeight(m_Instance)
        End Function

        Public Function GetFTextHeight(ByVal Align As TTextAlign, ByVal AText As String) As Double
            Return pdfGetFTextHeightW(m_Instance, Align, AText)
        End Function

        Public Function GetFTextHeightA(ByVal Align As TTextAlign, ByVal AText As String) As Double
            Return pdfGetFTextHeightA(m_Instance, Align, AText)
        End Function

        Public Function GetFTextHeightW(ByVal Align As TTextAlign, ByVal AText As String) As Double
            Return pdfGetFTextHeightW(m_Instance, Align, AText)
        End Function

        Public Function GetFTextHeightEx(ByVal Width As Double, ByVal Align As TTextAlign, ByVal AText As String) As Double
            Return pdfGetFTextHeightExW(m_Instance, Width, Align, AText)
        End Function

        Public Function GetFTextHeightExA(ByVal Width As Double, ByVal Align As TTextAlign, ByVal AText As String) As Double
            Return pdfGetFTextHeightExA(m_Instance, Width, Align, AText)
        End Function

        Public Function GetFTextHeightExW(ByVal Width As Double, ByVal Align As TTextAlign, ByVal AText As String) As Double
            Return pdfGetFTextHeightExW(m_Instance, Width, Align, AText)
        End Function

        Public Function GetGlyphIndex(ByVal Index As Integer) As Integer
            Return pdfGetGlyphIndex(m_Instance, Index)
        End Function

        Public Function GetGlyphOutline(ByVal Index As Integer, ByRef Outline As TPDFGlyphOutline) As Integer
            Dim retval As Integer
            Dim g As New TPDFGlyphOutline_I
            Erase Outline.Outline
            retval = pdfGetGlyphOutline(m_Instance, Index, Nothing)
            If retval > 0 Then
                Dim ptr As IntPtr
                Dim i As Integer, p As Long
                Dim dummy As New TI32Point
                g.Outline = Marshal.AllocHGlobal(retval * Marshal.SizeOf(dummy))
                retval = pdfGetGlyphOutline(m_Instance, Index, g)
                ReDim Outline.Outline(retval - 1)
                ptr = g.Outline
                p = ptr.ToInt64()
                For i = 0 To retval - 1
                    Outline.Outline(i) = CType(Marshal.PtrToStructure(ptr, GetType(TI32Point)), TI32Point)
                    p += Marshal.SizeOf(dummy)
                    ptr = New IntPtr(p)
                Next
                Marshal.FreeHGlobal(g.Outline)
            End If
            Outline.AdvanceX = g.AdvanceX
            Outline.AdvanceY = g.AdvanceY
            Outline.OriginX = g.OriginX
            Outline.OriginY = g.OriginY
            Outline.Lsb = g.Lsb
            Outline.Tsb = g.Tsb
            Outline.HaveBBox = g.HaveBBox <> 0
            Outline.BBox = g.BBox
            Return retval
        End Function

        Public Function GetGoToAction(ByVal Handle As Integer, ByRef Action As TPDFGoToAction) As Boolean
            Dim a As TPDFGoToAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetGoToAction(m_Instance, Handle, a) = 0 Then Return False
            Action.DestFile = a.DestFile
            Action.DestPage = a.DestPage
            Action.DestName = ToString(a.DestNameA, a.DestNameW)
            Action.DestType = a.DestType
            Action.NewWindow = a.NewWindow
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            If a.DestPos <> IntPtr.Zero Then
                ReDim Action.DestPos(3)
                Marshal.Copy(a.DestPos, Action.DestPos, 0, 4)
            Else
                Erase Action.DestPos
            End If
            Return True
        End Function

        Public Function GetGoToRAction(ByVal Handle As Integer, ByRef Action As TPDFGoToAction) As Boolean
            Return GetGoToAction(Handle, Action)
        End Function

        Public Function GetGStateFlags() As Integer
            Return pdfGetGStateFlags(m_Instance)
        End Function

        Public Function GetHideAction(ByVal Handle As Integer, ByRef Action As TPDFHideAction) As Boolean
            Dim a As TPDFHideAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetHideAction(m_Instance, Handle, a) = 0 Then Return False
            Action.Hide = CBool(a.Hide)
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            If a.FieldsCount > 0 Then
                ReDim Action.Fields(a.FieldsCount - 1)
                Marshal.Copy(a.Fields, Action.Fields, 0, a.FieldsCount)
            Else
                Erase Action.Fields
            End If
            Return True
        End Function

        Public Function GetIconColor() As Integer
            Return pdfGetIconColor(m_Instance)
        End Function

        Public Function GetImageBuffer(ByRef BufSize As Integer) As IntPtr
            Return pdfGetImageBuffer(m_Instance, BufSize)
        End Function

        Public Function GetImageCount(ByVal FileName As String) As Integer
            Return pdfGetImageCount(m_Instance, FileName)
        End Function

        Public Function GetImageHeight(ByVal AHandle As Integer) As Integer
            Return pdfGetImageHeight(m_Instance, AHandle)
        End Function

        Public Function GetImageObj(ByVal Handle As Integer, ByVal Flags As TParseFlags, ByRef Image As TPDFImage) As Boolean
            Return CBool(pdfGetImageObj(m_Instance, Handle, Flags, Image))
        End Function

        Public Function GetImageObjCount() As Integer
            Return pdfGetImageObjCount(m_Instance)
        End Function

        Public Function GetImageObjEx(ByVal ImagePtr As IntPtr, ByVal Flags As TParseFlags, ByRef Image As TPDFImage) As Boolean
            Return CBool(pdfGetImageObjEx(m_Instance, ImagePtr, Flags, Image))
        End Function

        Public Function GetImageWidth(ByVal AHandle As Integer) As Integer
            Return pdfGetImageWidth(m_Instance, AHandle)
        End Function

        Public Function GetImportDataAction(ByVal Handle As Integer, ByRef Action As TPDFImportDataAction) As Boolean
            Dim a As TPDFImportDataAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetImportDataAction(m_Instance, Handle, a) = 0 Then Return False
            CopyFileSpecEx(a.Data, Action.Data)
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            Return True
        End Function

        Public Function GetImportFlags() As Integer
            Return pdfGetImportFlags(m_Instance)
        End Function

        Public Function GetImportFlags2() As Integer
            Return pdfGetImportFlags2(m_Instance)
        End Function

        Public Function GetInBBox(ByVal PageNum As Integer, ByVal Boundary As TPageBoundary, ByRef BBox As TPDFRect) As Boolean
            Return CBool(pdfGetInBBox(m_Instance, PageNum, Boundary, BBox))
        End Function

        Public Function GetInDocInfo(ByVal DInfo As TDocumentInfo, ByRef Value As String) As Boolean
            Dim pBuffer As IntPtr
            Dim Len As Integer
            Len = pdfGetInDocInfo(m_Instance, DInfo, pBuffer)
            Value = ToString(pBuffer, Len, True)
            Return (Len >= 0)
        End Function

        Public Function GetInDocInfoCount() As Integer
            Return pdfGetInDocInfoCount(m_Instance)
        End Function

        Public Function GetInDocInfoEx(ByVal Index As Integer, ByRef DInfo As TDocumentInfo, ByRef Key As String, ByRef Value As String) As Boolean
            Dim info As Integer
            Dim uni As Integer
            Dim keyBuf As IntPtr
            Dim valBuf As IntPtr
            Dim Len As Integer
            Len = pdfGetInDocInfoEx(m_Instance, Index, info, keyBuf, valBuf, uni)
            If Len <= 0 Then
                Key = Nothing
                Value = Nothing
                Return (Len = 0)
            End If
            DInfo = CType(info, TDocumentInfo)
            Key = ToString(keyBuf, False)
            Value = ToString(valBuf, Len, uni <> 0)
            Return True
        End Function

        Public Function GetInEncryptionFlags() As Integer
            Return pdfGetInEncryptionFlags(m_Instance)
        End Function

        Public Function GetInFieldCount() As Integer
            Return pdfGetInFieldCount(m_Instance)
        End Function

        Public Function GetInIsCollection() As Boolean
            Dim retval As Integer
            retval = pdfGetInIsCollection(m_Instance)
            If retval < 0 Then Return False
            Return retval <> 0
        End Function

        Public Function GetInIsEncrypted() As Boolean
            Dim retval As Integer
            retval = pdfGetInIsEncrypted(m_Instance)
            If retval < 0 Then Return False
            Return retval <> 0
        End Function

        Public Function GetInIsSigned() As Boolean
            Dim retval As Integer
            retval = pdfGetInIsSigned(m_Instance)
            If retval < 0 Then Return False
            Return retval <> 0
        End Function

        Public Function GetInIsTrapped() As Boolean
            Dim retval As Integer
            retval = pdfGetInIsTrapped(m_Instance)
            If retval < 0 Then Return False
            Return retval <> 0
        End Function

        Public Function GetInIsXFAForm() As Boolean
            Dim retval As Integer
            retval = pdfGetInIsXFAForm(m_Instance)
            If retval < 0 Then Return False
            Return retval <> 0
        End Function

        Public Function GetInkList(ByVal List As IntPtr, ByRef Points() As Single) As Boolean
            Dim pts As IntPtr, cnt As Integer
            If pdfGetInkList(List, pts, cnt) <> 0 Then
                If cnt > 0 Then
                    ReDim Points(cnt - 1)
                    Marshal.Copy(pts, Points, 0, cnt)
                Else
                    Erase Points
                End If
                Return True
            Else
                Erase Points
                Return False
            End If
        End Function

        Public Function GetInMetadata(ByVal PageNum As Integer, ByRef Buffer() As Byte) As Boolean
            Dim pbuf As IntPtr
            Dim BufSize As Integer
            If pdfGetInMetadata(m_Instance, PageNum, pbuf, BufSize) <> 0 Then
                If BufSize > 0 Then
                    ReDim Buffer(BufSize - 1)
                    Marshal.Copy(pbuf, Buffer, 0, BufSize)
                Else
                    Erase Buffer
                End If
                Return True
            Else
                Erase Buffer
                Return False
            End If
        End Function

        Public Function GetInOrientation(ByVal PageNum As Integer) As Integer
            Return pdfGetInOrientation(m_Instance, PageNum)
        End Function

        Public Function GetInPageCount() As Integer
            Return pdfGetInPageCount(m_Instance)
        End Function

        Public Function GetInPDFVersion() As Integer
            Return pdfGetInPDFVersion(m_Instance)
        End Function

        Public Function GetInPrintSettings(ByRef Settings As TPDFPrintSettings) As Boolean
            Dim s As TPDFPrintSettings_I
            If pdfGetInPrintSettings(m_Instance, s) = 0 Then Return False
            Settings.DuplexMode = s.DuplexMode
            Settings.NumCopies = s.NumCopies
            Settings.PickTrayByPDFSize = s.PickTrayByPDFSize
            Settings.PrintScaling = s.PrintScaling
            If s.PrintRangesCount > 0 Then
                ReDim Settings.PrintRanges(s.PrintRangesCount * 2 - 1)
                Marshal.Copy(s.PrintRanges, Settings.PrintRanges, 0, s.PrintRangesCount * 2)
            Else
                Erase Settings.PrintRanges
            End If
            Return True
        End Function

        Public Function GetInRepairMode() As Integer
            Return pdfGetInRepairMode(m_Instance)
        End Function

        Public Function GetIsFixedPitch() As Boolean
            Return CBool(pdfGetIsFixedPitch(m_Instance))
        End Function

        Public Function GetIsTaggingEnabled(ByVal Flags As Integer) As Boolean
            Return CBool(pdfGetIsTaggingEnabled(m_Instance))
        End Function

        Public Function GetItalicAngle() As Double
            Return pdfGetItalicAngle(m_Instance)
        End Function

        Public Function GetJavaScript(ByVal AHandle As Integer, ByRef Script As String) As Boolean
            Dim uni As Integer
            Dim pBuffer As IntPtr
            Dim BufSize As Integer

            pBuffer = pdfGetJavaScript(m_Instance, AHandle, BufSize, uni)
            Script = ToString(pBuffer, BufSize, uni <> 0)
            Return True
        End Function

        Public Function GetJavaScriptAction(ByVal AHandle As Integer, ByRef Script As String) As Boolean
            Dim uni As Integer
            Dim pBuffer As IntPtr
            Dim BufSize As Integer

            pBuffer = pdfGetJavaScriptAction(m_Instance, AHandle, BufSize, uni)
            Script = ToString(pBuffer, BufSize, uni <> 0)
            Return True
        End Function

        Public Function GetJavaScriptAction2(ByVal ObjType As TObjType, ByVal ObjHandle As Integer, ByVal ActIndex As Integer, ByRef Script As String, ByRef ObjEvent As TObjEvent) As Boolean
            Dim uni As Integer
            Dim pBuffer As IntPtr
            Dim BufSize As Integer

            pBuffer = pdfGetJavaScriptAction2(m_Instance, ObjType, ObjHandle, ActIndex, BufSize, uni, ObjEvent)
            Script = ToString(pBuffer, BufSize, uni <> 0)
            Return True
        End Function

        Public Function GetJavaScriptActionEx(ByVal Handle As Integer, ByRef Action As TPDFJavaScriptAction) As Boolean
            Dim a As TPDFJavaScriptAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetJavaScriptActionEx(m_Instance, Handle, a) = 0 Then Return False
            Action.Script = ToString(a.ScriptA, a.ScriptW, a.ScriptLen)
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            Return True
        End Function

        Public Function GetJavaScriptCount() As Integer
            Return pdfGetJavaScriptCount(m_Instance)
        End Function

        Public Function GetJavaScriptEx(ByVal Name As String, ByRef Script As String) As Boolean
            Dim uni As Integer
            Dim pBuffer As IntPtr
            Dim BufSize As Integer

            pBuffer = pdfGetJavaScriptEx(m_Instance, Name, BufSize, uni)
            Script = ToString(pBuffer, BufSize, uni <> 0)
            Return True
        End Function

        Public Function GetJavaScriptName(ByVal Handle As Integer) As String
            Dim uni As Integer
            Dim BufSize As Integer
            Dim pBuffer As IntPtr
            pBuffer = pdfGetJavaScriptName(m_Instance, Handle, BufSize, uni)
            Return ToString(pBuffer, BufSize, uni <> 0)
        End Function

        Public Function GetJPEGQuality() As Integer
            GetJPEGQuality = pdfGetJPEGQuality(m_Instance)
        End Function

        Public Function GetLanguage() As String
            Return ToString(pdfGetLanguage(m_Instance), False)
        End Function

        Public Function GetLastTextPosX() As Double
            Return pdfGetLastTextPosX(m_Instance)
        End Function

        Public Function GetLastTextPosY() As Double
            Return pdfGetLastTextPosY(m_Instance)
        End Function

        Public Function GetLayerConfig(ByVal Index As Integer, ByRef Config As TPDFOCLayerConfig) As Boolean
            Dim cfg As TPDFOCLayerConfig_I
            cfg.StructSize = Marshal.SizeOf(cfg)
            GetLayerConfig = CBool(pdfGetLayerConfig(m_Instance, Index, cfg))
            If GetLayerConfig Then
                Config.Intent = cfg.Intent
                Config.IsDefault = cfg.IsDefault <> 0
                Config.Name = ToString(cfg.NameA, cfg.NameW, cfg.NameLen)
            End If
        End Function

        Public Function GetLayerConfigCount() As Integer
            Return pdfGetLayerConfigCount(m_Instance)
        End Function

        Public Function GetLaunchAction(ByVal Handle As Integer, ByRef Action As TPDFLaunchAction) As Boolean
            Dim a As TPDFLaunchAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetLaunchAction(m_Instance, Handle, a) = 0 Then Return False
            Action.AppName = ToString(a.AppName, False)
            Action.DefDir = ToString(a.DefDir, False)
            Action.File = a.File
            Action.NewWindow = a.NewWindow
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            Action.Operation = ToString(a.Operation, False)
            Action.Parameter = ToString(a.Parameter, False)
            Return True
        End Function

        Public Function GetLeading() As Double
            Return pdfGetLeading(m_Instance)
        End Function

        Public Function GetLineCapStyle() As TLineCapStyle
            Return CType(pdfGetLineCapStyle(m_Instance), TLineCapStyle)
        End Function

        Public Function GetLineJoinStyle() As TLineJoinStyle
            Return CType(pdfGetLineJoinStyle(m_Instance), TLineJoinStyle)
        End Function

        Public Function GetLineWidth() As Double
            Return pdfGetLineWidth(m_Instance)
        End Function

        Public Function GetLinkHighlightMode() As THighlightMode
            Return CType(pdfGetLinkHighlightMode(m_Instance), THighlightMode)
        End Function

        Public Function GetLogMetafileSize(ByVal FileName As String, ByRef ARect As TRectL) As Boolean
            Return CBool(pdfGetLogMetafileSize(m_Instance, FileName, ARect))
        End Function

        Public Function GetLogMetafileSizeEx(ByRef Buffer() As Byte, ByRef ARect As TRectL) As Boolean
            Return CBool(pdfGetLogMetafileSizeEx(m_Instance, Buffer, Buffer.Length, ARect))
        End Function

        Public Function GetMatrix(ByRef Matrix As TCTM) As Boolean
            Return CBool(pdfGetMatrix(m_Instance, Matrix))
        End Function

        Public Function GetMaxFieldLen(ByVal TxtField As Integer) As Integer
            Return pdfGetMaxFieldLen(m_Instance, TxtField)
        End Function

        Public Function GetMeasureObj(ByVal MeasurePtr As IntPtr, ByRef Value As TPDFMeasure) As Boolean
            Dim m As TPDFMeasure_I
            m.StructSize = Marshal.SizeOf(m)
            If pdfGetMeasureObj(MeasurePtr, m) = 0 Then Return False
            If m.IsRectilinear <> 0 Then
                Value.IsRectilinear = True
                Value.Angles = CopyIntPtrArray(m.Angles, m.AnglesCount)
                Value.Area = CopyIntPtrArray(m.Area, m.AreaCount)
                Value.CXY = m.CXY
                Value.Distance = CopyIntPtrArray(m.Distance, m.DistanceCount)
                Value.OriginX = m.OriginX
                Value.OriginY = m.OriginY
                Value.R = ToString(m.RA, m.RW)
                Value.Slope = CopyIntPtrArray(m.Slope, m.SlopeCount)
                Value.x = CopyIntPtrArray(m.x, m.XCount)
                Value.y = CopyIntPtrArray(m.y, m.YCount)
            Else
                Value.Bounds = CopyFloatArray(m.Bounds, m.BoundCount)
                If m.DCS_IsSet <> 0 Then
                    Value.DCS_IsSet = True
                    Value.DCS_Projected = m.DCS_Projected <> 0
                    Value.DCS_EPSG = m.DCS_EPSG
                    Value.DCS_WKT = ToString(m.DCS_WKT, False)
                End If
                Value.GCS_Projected = m.GCS_Projected <> 0
                Value.GCS_EPSG = m.GCS_EPSG
                Value.GCS_WKT = ToString(m.GCS_WKT, False)
                Value.GPTS = CopyFloatArray(m.GPTS, m.GPTSCount)
                Value.LPTS = CopyFloatArray(m.LPTS, m.LPTSCount)
                Value.PDU1 = ToString(m.PDU1, False)
                Value.PDU2 = ToString(m.PDU2, False)
                Value.PDU3 = ToString(m.PDU3, False)
            End If
            Return True
        End Function

        Public Function GetMetaConvFlags() As Integer
            Return pdfGetMetaConvFlags(m_Instance)
        End Function

        Public Function GetMetadata(ByVal ObjType As TMetadataObj, ByVal Handle As Integer, ByRef Buffer() As Byte) As Boolean
            Dim pbuf As IntPtr
            Dim BufSize As Integer
            If pdfGetMetadata(m_Instance, ObjType, Handle, pbuf, BufSize) <> 0 Then
                If BufSize > 0 Then
                    ReDim Buffer(BufSize - 1)
                    Marshal.Copy(pbuf, Buffer, 0, BufSize)
                Else
                    Erase Buffer
                End If
                Return True
            Else
                Erase Buffer
                Return False
            End If
        End Function

        Public Function GetMissingGlyphs() As System.UInt32()
            Dim count As Integer
            Dim ptr As IntPtr = pdfGetMissingGlyphs(m_Instance, count)
            If (IntPtr.Zero.Equals(ptr)) Then
                Return Nothing
            Else
                Dim retval() As System.UInt32
                ReDim retval(count - 1)
                ' Why does the marshaller not provide a copy function for uint[]?
                pdfCopyMemUInt(ptr, retval, count * 4)
                Return retval
            End If
        End Function

        Public Function GetMiterLimit() As Double
            Return pdfGetMiterLimit(m_Instance)
        End Function

        Public Function GetMovieAction(ByVal Handle As Integer, ByRef Action As TPDFMovieAction) As Boolean
            Dim a As New TPDFMovieAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetMovieAction(m_Instance, Handle, a) = 0 Then Return False
            Action.Annot = a.Annot
            Action.FWPosition(0) = a.FWPosition(0)
            Action.FWPosition(1) = a.FWPosition(1)
            Action.FWScale(0) = a.FWScale(0)
            Action.FWScale(1) = a.FWScale(1)
            Action.Mode = ToString(a.Mode, False)
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            Action.Operation = ToString(a.Operation, False)
            Action.Rate = a.Rate
            Action.ShowControls = a.ShowControls <> 0
            Action.Synchronous = a.Synchronous <> 0
            Action.Title = ToString(a.TitleA, a.TitleW)
            Action.Volume = a.Volume
            Return True
        End Function

        Public Function GetNamedAction(ByVal Handle As Integer, ByRef Action As TPDFNamedAction) As Boolean
            Dim a As TPDFNamedAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetNamedAction(m_Instance, Handle, a) = 0 Then Return False
            Action.Name = ToString(a.Name, False)
            Action.NewWindow = a.NewWindow
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            Action.Type = a.Type
            Return True
        End Function

        Public Function GetNamedDest(ByVal Index As Integer, ByRef Dest As TPDFNamedDest) As Boolean
            Dim d As TPDFNamedDest_I
            If pdfGetNamedDest(m_Instance, Index, d) = 0 Then Return False
            Dest.Name = ToString(d.NameA, d.NameW, d.NameLen)
            Dest.DestFile = ToString(d.DestFileA, d.DestFileW, d.DestFileLen)
            Dest.DestPage = d.DestPage
            Dest.DestPos = d.DestPos
            Dest.DestType = d.DestType
            Return True
        End Function

        Public Function GetNamedDestCount() As Integer
            Return pdfGetNamedDestCount(m_Instance)
        End Function

        Public Function GetNeedAppearance() As Boolean
            Return CBool(pdfGetNeedAppearance(m_Instance))
        End Function

        Public Function GetNumberFormatObj(ByVal NumFmtPtr As IntPtr, ByRef Value As TPDFNumberFormat) As Boolean
            Dim nf As TPDFNumberFormat_I
            nf.StructSize = Marshal.SizeOf(nf)
            If pdfGetNumberFormatObj(NumFmtPtr, nf) = 0 Then Return False
            Value.C = nf.C
            Value.D = nf.D
            Value.f = nf.f
            Value.FD = nf.FD <> 0
            Value.O = nf.O
            Value.PS = ToString(nf.PSA, nf.PSW)
            Value.RD = ToString(nf.RDA, nf.RDW)
            Value.RT = ToString(nf.RTA, nf.RTW)
            Value.SS = ToString(nf.SSA, nf.SSW)
            Value.U = ToString(nf.UA, nf.UW)
            Return True
        End Function

        Public Function GetObjActionCount(ByVal ObjType As TObjType, ByVal ObjHandle As Integer) As Integer
            Return pdfGetObjActionCount(m_Instance, ObjType, ObjHandle)
        End Function

        Public Function GetObjActions(ByVal ObjType As TObjType, ByVal ObjHandle As Integer, ByRef Actions As TPDFObjActions) As Integer
            Actions.StructSize = Marshal.SizeOf(Actions)
            Return pdfGetObjActions(m_Instance, ObjType, ObjHandle, Actions)
        End Function

        Public Function GetObjEvent(ByVal IEvent As IntPtr, ByRef ObjEvent As TPDFObjEvent) As Boolean
            ObjEvent.StructSize = Marshal.SizeOf(ObjEvent)
            Return CBool(pdfGetObjEvent(IEvent, ObjEvent))
        End Function

        Public Function GetOCG(ByVal Handle As Integer, ByRef Value As TPDFOCG) As Boolean
            Value.StructSize = Marshal.SizeOf(Value)
            Return CBool(pdfGetOCG(m_Instance, Handle, Value))
        End Function

        Public Function GetOCGContUsage(ByVal Handle As Integer, ByRef Value As TPDFOCGContUsage) As Boolean
            Value.StructSize = Marshal.SizeOf(Value)
            Return CBool(pdfGetOCGContUsage(m_Instance, Handle, Value))
        End Function

        Public Function GetOCGCount() As Integer
            Return pdfGetOCGCount(m_Instance)
        End Function

        Public Function GetOCGUsageUserName(ByVal Handle As Integer, ByVal Index As Integer, ByRef Name As String) As Boolean
            Dim nmeA As IntPtr, nmeW As IntPtr
            If pdfGetOCGUsageUserName(m_Instance, Handle, Index, nmeA, nmeW) <> 0 Then
                Name = ToString(nmeA, nmeW)
                Return True
            Else
                Name = Nothing
                Return False
            End If
        End Function

        Public Function GetOCHandle(ByVal OC As IntPtr) As Integer
            Return pdfGetOCHandle(OC)
        End Function

        Public Function GetOCUINode(ByVal Node As IntPtr, ByVal OutNode As TPDFOCUINode) As IntPtr
            If Node.Equals(IntPtr.Zero) Then
                Return pdfGetOCUINode(m_Instance, Node, Nothing)
            Else
                Dim n As TPDFOCUINode_I = New TPDFOCUINode_I()
                n.StructSize = Marshal.SizeOf(n)
                GetOCUINode = pdfGetOCUINode(m_Instance, Node, n)
                OutNode.Label = ToString(n.LabelA, n.LabelW, n.LabelLength)
                OutNode.NewNode = CBool(n.NewNode)
                OutNode.NextChild = n.NextChild
                OutNode.OCG = n.OCG
            End If
        End Function

        Public Function GetOpacity() As Double
            Return pdfGetOpacity(m_Instance)
        End Function

        Public Function GetOrientation() As Integer
            Return pdfGetOrientation(m_Instance)
        End Function

        Public Function GetOutputIntent(ByVal Index As Integer, ByRef Intent As TPDFOutputIntent) As Boolean
            Dim retval As TPDFOutputIntent_I
            retval.StructSize = Marshal.SizeOf(retval)
            If pdfGetOutputIntent(m_Instance, Index, retval) = 0 Then Return False
            GetIntOutputIntent(retval, Intent)
            Return True
        End Function

        Public Function GetOutputIntentCount() As Integer
            Return pdfGetOutputIntentCount(m_Instance)
        End Function

        Public Function GetPageAnnot(ByVal Index As Integer, ByRef Annot As TPDFAnnotation) As Boolean
            Dim retval As TPDFAnnotation_I
            If pdfGetPageAnnot(m_Instance, Index, retval) = 0 Then
                Annot = Nothing
                Return False
            End If
            GetIntAnnot(retval, Annot)
            Return True
        End Function

        Public Function GetPageAnnotCount() As Integer
            Return pdfGetPageAnnotCount(m_Instance)
        End Function

        Public Function GetPageAnnotEx(ByVal Index As Integer, ByRef Annot As TPDFAnnotationEx) As Boolean
            Dim retval As TPDFAnnotationEx_I
            If pdfGetPageAnnotEx(m_Instance, Index, retval) = 0 Then
                Annot = Nothing
                Return False
            End If
            GetIntAnnotEx(retval, Annot)
            Return True
        End Function

        Public Function GetPageBBox(ByVal PagePtr As IntPtr, ByVal Boundary As TPageBoundary, ByRef BBox As TFltRect) As Boolean
            Return CBool(pdfGetPageBBox(PagePtr, Boundary, BBox))
        End Function

        Public Function GetPageCoords() As TPageCoord
            Return CType(pdfGetPageCoords(m_Instance), TPageCoord)
        End Function

        Public Function GetPageCount() As Integer
            Return pdfGetPageCount(m_Instance)
        End Function

        Public Function GetPageField(ByVal Index As Integer, ByRef Field As TPDFField) As Boolean
            Dim f As TPDFField_I
            If pdfGetPageField(m_Instance, Index, f) < 0 Then
                Field = Nothing
                Return False
            End If
            GetIntField(f, Field)
            Return True
        End Function

        Public Function GetPageFieldCount() As Integer
            Return pdfGetPageFieldCount(m_Instance)
        End Function

        Public Function GetPageFieldEx(ByVal Index As Integer, ByRef Field As TPDFFieldEx) As Boolean
            Dim f As TPDFFieldEx_I
            f.StructSize = Marshal.SizeOf(f)
            If pdfGetPageFieldEx(m_Instance, Index, f) = 0 Then Return False
            GetIntFieldEx(f, Field)
            Return True
        End Function

        Public Function GetPageHeight() As Double
            Return pdfGetPageHeight(m_Instance)
        End Function

        Public Function GetPageLabel(ByVal Index As Integer, ByRef Label As TPDFPageLabel) As Boolean
            Dim lbl As TPDFPageLabel_I
            If pdfGetPageLabel(m_Instance, Index, lbl) = 0 Then Return False
            Label.StartRange = lbl.StartRange
            Label.Format = lbl.Format
            Label.FirstPageNum = lbl.FirstPageNum
            Label.Prefix = ToString(lbl.Prefix, lbl.PrefixLen, lbl.PrefixUni <> 0)
            Return True
        End Function

        Public Function GetPageLabelCount() As Integer
            Return pdfGetPageLabelCount(m_Instance)
        End Function

        Public Function GetPageLayout() As TPageLayout
            Return CType(pdfGetPageLayout(m_Instance), TPageLayout)
        End Function

        Public Function GetPageMode() As TPageMode
            Return CType(pdfGetPageMode(m_Instance), TPageMode)
        End Function

        Public Function GetPageNum() As Integer
            Return pdfGetPageNum(m_Instance)
        End Function

        Public Function GetPageObject(ByVal PageNum As Integer) As IntPtr
            Return pdfGetPageObject(m_Instance, PageNum)
        End Function

        Public Function GetPageOrientation(ByVal PagePtr As IntPtr) As Integer
            Return pdfGetPageOrientation(PagePtr)
        End Function

        Public Function GetPageText(ByRef Stack As TPDFStack) As Boolean
            If pdfGetPageText(m_Instance, Stack) = 0 Then Return False
            Return True
        End Function

        Public Function GetPageWidth() As Double
            Return pdfGetPageWidth(m_Instance)
        End Function

        Public Function GetPDFInstance() As IntPtr
            Return m_Instance
        End Function

        Public Function GetPDFVersion() As Integer
            Return pdfGetPDFVersion(m_Instance)
        End Function

        Public Function GetPrintSettings(ByRef Settings As TPDFPrintSettings) As Boolean
            Dim s As TPDFPrintSettings_I
            If pdfGetPrintSettings(m_Instance, s) = 0 Then Return False
            Settings.DuplexMode = s.DuplexMode
            Settings.NumCopies = s.NumCopies
            Settings.PickTrayByPDFSize = s.PickTrayByPDFSize
            Settings.PrintScaling = s.PrintScaling
            If s.PrintRangesCount > 0 Then
                ReDim Settings.PrintRanges(s.PrintRangesCount * 2 - 1)
                Marshal.Copy(s.PrintRanges, Settings.PrintRanges, 0, s.PrintRangesCount * 2)
            End If
            Return True
        End Function

        Public Function GetPtDataArray(ByVal PtData As IntPtr, ByVal Index As Integer, ByRef DataType As String, ByRef Values() As Single) As Boolean
            Dim dtype As IntPtr, val As IntPtr, valCount As Integer
            DataType = Nothing
            Erase Values
            If pdfGetPtDataArray(PtData, Index, dtype, val, valCount) = 0 Then Return False
            If valCount < 2 Then Return False
            DataType = ToString(dtype, False)
            ReDim Values(valCount - 1)
            Marshal.Copy(val, Values, 0, valCount)
            Return True
        End Function

        Public Function GetPtDataObj(ByVal PtData As IntPtr, ByRef Subtype As String, ByRef NumArrays As Integer) As Boolean
            Dim st As IntPtr
            Subtype = Nothing
            NumArrays = 0
            If pdfGetPtDataObj(PtData, st, NumArrays) = 0 Then Return False
            Subtype = ToString(st, False)
            Return True
        End Function

        Public Function GetRelFileNode(ByVal IRF As IntPtr, ByRef F As TPDFRelFileNode, ByVal Decompress As Boolean) As Boolean
            Dim node As TPDFRelFileNode_I
            node.StructSize = Marshal.SizeOf(node)
            If pdfGetRelFileNode(IRF, node, Convert.ToInt32(Decompress)) = 0 Then Return False

            If Not IntPtr.Zero.Equals(node.EF.Buffer) Then
                ReDim F.EF.Buffer(node.EF.BufSize - 1)
                Marshal.Copy(node.EF.Buffer, F.EF.Buffer, 0, node.EF.BufSize)
            Else
                F.EF.Buffer = Nothing
            End If
            F.EF.ColItem = node.EF.ColItem
            F.EF.Compressed = node.EF.Compressed <> 0
            F.EF.FileSize = node.EF.FileSize
            F.EF.IsURL = node.EF.IsURL <> 0
            If Not IntPtr.Zero.Equals(node.EF.CheckSum) Then
                ReDim F.EF.CheckSum(15)
                Marshal.Copy(node.EF.CheckSum, F.EF.CheckSum, 0, 16)
            End If
            F.EF.CreateDate = ToString(node.EF.CreateDate, False)
            F.EF.Desc = ToString(node.EF.Desc, node.EF.DescUnicode <> 0)
            F.EF.FileName = ToString(node.EF.FileName, False)
            F.EF.MIMEType = ToString(node.EF.MIMEType, False)
            F.EF.ModDate = ToString(node.EF.ModDate, False)
            F.EF.Name = ToString(node.EF.Name, node.EF.NameUnicode <> 0)
            F.EF.UF = ToString(node.EF.UF, node.EF.UFUnicode <> 0)



            F.Name = ToString(node.NameA, node.NameW)
            F.NextNode = node.NextNode
            Return True
        End Function

        Public Function GetResetAction(ByVal Handle As Integer, ByRef Value As TPDFResetFormAction) As Boolean
            Dim v As TPDFResetFormAction_I
            v.StructSize = Marshal.SizeOf(v)
            If pdfGetResetAction(m_Instance, Handle, v) = 0 Then Return False
            If v.FieldsCount > 0 Then
                ReDim Value.Fields(v.FieldsCount - 1)
                Marshal.Copy(v.Fields, Value.Fields, 0, v.FieldsCount)
            Else
                Erase Value.Fields
            End If
            Value.Include = v.Include <> 0
            Value.NextAction = v.NextAction
            Value.NextActionType = v.NextActionType
            Return True
        End Function

        Public Function GetResolution() As Integer
            Return pdfGetResolution(m_Instance)
        End Function

        Public Function GetSaveNewImageFormat() As Boolean
            Return CBool(pdfGetSaveNewImageFormat(m_Instance))
        End Function

        Public Function GetSeparationInfo(ByRef Colorant As String, ByRef CS As TExtColorSpace) As Boolean
            Dim clt As IntPtr
            If pdfGetSeparationInfo(m_Instance, clt, CS) = 0 Then Return False
            Colorant = ToString(clt, False)
            Return True
        End Function

        Public Function GetSpaceWidth(ByVal IFont As IntPtr, ByVal FontSize As Double) As Double
            Return fntGetSpaceWidth(IFont, FontSize)
        End Function

        Public Function GetSigDict(ByVal ISignature As IntPtr, ByRef SigDict As TPDFSigDict) As Boolean
            Dim sd As TPDFSigDict_I
            sd.StructSize = Marshal.SizeOf(sd)
            If pdfGetSigDict(ISignature, sd) = 0 Then Return False
            GetIntSigDict(sd, SigDict)
            Return True
        End Function

        Public Function GetStrokeColor() As Integer
            Return pdfGetStrokeColor(m_Instance)
        End Function

        Public Function GetSubmitAction(ByVal Handle As Integer, ByRef Value As TPDFSubmitFormAction) As Boolean
            Dim v As TPDFSubmitFormAction_I
            v.StructSize = Marshal.SizeOf(v)
            If pdfGetSubmitAction(m_Instance, Handle, v) = 0 Then Return False
            Value.CharSet = ToString(v.CharSet, False)
            If v.FieldsCount > 0 Then
                ReDim Value.Fields(v.FieldsCount - 1)
                Marshal.Copy(v.Fields, Value.Fields, 0, v.FieldsCount)
            Else
                Erase Value.Fields
            End If
            Value.Flags = v.Flags
            Value.URL = ToString(v.URL, False)
            Value.NextAction = v.NextAction
            Value.NextActionType = v.NextActionType
            Return True
        End Function

        Public Function GetSysFontInfo(ByVal Handle As Integer, ByRef Value As TPDFSysFont) As Integer
            Dim retval As Integer
            Dim f As TPDFSysFont_I
            f.StructSize = Marshal.SizeOf(f)
            retval = pdfGetSysFontInfo(m_Instance, Handle, f)
            If retval < 0 Then Return retval
            If retval = 0 Then
                If IntPtr.Zero.Equals(f.FamilyName) Then
                    Value = New TPDFSysFont()
                    Return 0
                End If
            End If
            Value.BaseType = f.BaseType
            Value.CIDOrdering = ToString(f.CIDOrdering, False)
            Value.CIDRegistry = ToString(f.CIDRegistry, False)
            Value.CIDSupplement = f.CIDSupplement
            Value.DataOffset = f.DataOffset
            Value.FamilyName = ToString(f.FamilyName, True)
            Value.FilePath = ToString(f.FilePathA, f.FilePathW)
            Value.FileSize = f.FileSize
            Value.Flags = f.Flags
            Value.FullName = ToString(f.FullName, True)
            Value.Index = f.Index
            Value.IsFixedPitch = f.IsFixedPitch <> 0
            Value.Length1 = f.Length1
            Value.Length2 = f.Length2
            Value.PostScriptName = ToString(f.PostScriptNameA, f.PostScriptNameW)
            Value.Style = f.Style
            Value.UnicodeRange1 = f.UnicodeRange1
            Value.UnicodeRange2 = f.UnicodeRange2
            Value.UnicodeRange3 = f.UnicodeRange3
            Value.UnicodeRange4 = f.UnicodeRange4
            Return retval
        End Function

        Public Function GetTabLen() As Integer
            Return pdfGetTabLen(m_Instance)
        End Function

        Public Function GetTemplCount() As Integer
            Return pdfGetTemplCount(m_Instance)
        End Function

        Public Function GetTemplHandle() As Integer
            Return pdfGetTemplHandle(m_Instance)
        End Function

        Public Function GetTemplHeight(ByVal Handle As Integer) As Double
            Return pdfGetTemplHeight(m_Instance, Handle)
        End Function

        Public Function GetTemplWidth(ByVal Handle As Integer) As Double
            Return pdfGetTemplWidth(m_Instance, Handle)
        End Function

        Public Function GetTextDrawMode() As TDrawMode
            Return CType(pdfGetTextDrawMode(m_Instance), TDrawMode)
        End Function

        Public Function GetTextFieldValue(ByVal AField As Integer, ByRef Value As String, ByRef DefValue As String) As Boolean
            Dim valUni As Integer
            Dim defValUni As Integer
            Dim val As IntPtr
            Dim defVal As IntPtr
            If pdfGetTextFieldValue(m_Instance, AField, val, valUni, defVal, defValUni) = 0 Then
                Value = ""
                DefValue = ""
                Return False
            End If
            Value = ToString(val, valUni <> 0)
            DefValue = ToString(defVal, defValUni <> 0)
            Return True
        End Function

        Public Function GetTextRect(ByRef PosX As Double, ByRef PosY As Double, ByRef Width As Double, ByRef Height As Double) As Boolean
            Return CBool(pdfGetTextRect(m_Instance, PosX, PosY, Width, Height))
        End Function

        Public Function GetTextRise() As Double
            Return pdfGetTextRise(m_Instance)
        End Function

        Public Function GetTextScaling() As Double
            Return pdfGetTextScaling(m_Instance)
        End Function

        Public Function GetTextWidth(ByVal IFont As IntPtr, ByVal Text As IntPtr, ByVal Len As Integer, ByVal CharSpacing As Single, ByVal WordSpacing As Single, ByVal TextScale As Single) As Double
            Return fntGetTextWidth(IFont, Text, Len, CharSpacing, WordSpacing, TextScale)
        End Function

        Public Function GetTextWidth(ByVal AText As String) As Double
            Return pdfGetTextWidthW(m_Instance, AText)
        End Function

        Public Function GetTextWidthA(ByVal AText As String) As Double
            Return pdfGetTextWidthA(m_Instance, AText)
        End Function

        Public Function GetTextWidthW(ByVal AText As String) As Double
            Return pdfGetTextWidthW(m_Instance, AText)
        End Function

        Public Function GetTextWidthEx(ByVal AText As String, ByVal Len As Integer) As Double
            Return pdfGetTextWidthExW(m_Instance, AText, Len)
        End Function

        Public Function GetTextWidthExA(ByVal AText As String, ByVal Len As Integer) As Double
            Return pdfGetTextWidthExA(m_Instance, AText, Len)
        End Function

        Public Function GetTextWidthExW(ByVal AText As String, ByVal Len As Integer) As Double
            Return pdfGetTextWidthExW(m_Instance, AText, Len)
        End Function

        Public Function GetTransparentColor() As Integer
            Return pdfGetTransparentColor(m_Instance)
        End Function

        Public Function GetTrapped() As Boolean
            Return CBool(pdfGetTrapped(m_Instance))
        End Function

        Public Function GetURIAction(ByVal Handle As Integer, ByRef Action As TPDFURIAction) As Boolean
            Dim a As TPDFURIAction_I
            a.StructSize = Marshal.SizeOf(a)
            If pdfGetURIAction(m_Instance, Handle, a) = 0 Then Return False
            Action.BaseURL = ToString(a.BaseURL, False)
            Action.IsMap = a.IsMap <> 0
            Action.NextAction = a.NextAction
            Action.NextActionType = a.NextActionType
            Action.URI = ToString(a.URI, False)
            Return True
        End Function

        Public Function GetUseExactPwd() As Boolean
            Return CBool(pdfGetUseExactPwd(m_Instance))
        End Function

        Public Function GetUseGlobalImpFiles() As Boolean
            Return CBool(pdfGetUseGlobalImpFiles(m_Instance))
        End Function

        Public Function GetUserRights() As Integer
            Return pdfGetUserRights(m_Instance)
        End Function

        Public Function GetUserUnit() As Single
            Return pdfGetUserUnit(m_Instance)
        End Function

        Public Function GetUseStdFonts() As Boolean
            Return CBool(pdfGetUseStdFonts(m_Instance))
        End Function

        Public Function GetUseSystemFonts() As Boolean
            Return CBool(pdfGetUseSystemFonts(m_Instance))
        End Function

        Public Function GetUsesTransparency(ByVal PageNum As Integer) As Integer
            GetUsesTransparency = pdfGetUsesTransparency(m_Instance, PageNum)
        End Function

        Public Function GetUseTransparency() As Boolean
            Return CBool(pdfGetUseTransparency(m_Instance))
        End Function

        Public Function GetUseVisibleCoords() As Boolean
            Return CBool(pdfGetUseVisibleCoords(m_Instance))
        End Function

        Public Function GetViewerPrefrences(ByRef Preference As Integer, ByRef AddVal As Integer) As Boolean
            Return CBool(pdfGetViewerPrefrences(m_Instance, Preference, AddVal))
        End Function

        Public Function GetViewport(ByVal PageNum As Integer, ByVal Index As Integer, ByRef Value As TPDFViewport) As Boolean
            Dim vp As TPDFViewport_I
            vp.StructSize = Marshal.SizeOf(vp)
            If pdfGetViewport(m_Instance, PageNum, Index, vp) = 0 Then Return False
            Value.BBox = vp.BBox
            Value.Measure = vp.Measure
            Value.Name = ToString(vp.NameA, vp.NameW)
            Value.PtData = vp.PtData
            Return True
        End Function

        Public Function GetViewportCount(ByVal PageNum As Integer) As Integer
            Return pdfGetViewportCount(m_Instance, PageNum)
        End Function

        Public Function GetXFAStream(ByVal Index As Integer, ByVal Out As TPDFXFAStream) As Boolean
            Dim strm As New TPDFXFAStream_I()
            strm.StructSize = Marshal.SizeOf(strm)
            If IsNothing(Out) Then
                Return pdfGetXFAStream(m_Instance, Index, Nothing) <> 0
            ElseIf pdfGetXFAStream(m_Instance, Index, strm) <> 0 Then
                ReDim Out.Buffer(strm.BufSize - 1)
                Marshal.Copy(strm.Buffer, Out.Buffer, 0, strm.BufSize)
                Out.Name = ToString(strm.NameA, strm.NameW)
                Return True
            Else
                Out.Buffer = Nothing
                Out.Name = ""
                Return False
            End If
        End Function

        Public Function GetXFAStreamCount() As Integer
            Return pdfGetXFAStreamCount(m_Instance)
        End Function

        Public Function GetWidthHeight(ByVal PagePtr As IntPtr, ByVal Flags As TRasterFlags, ByRef Width As Single, ByRef Height As Single, ByVal Rotate As Integer, ByRef BBox As TFltRect) As Boolean
            Dim v As IntPtr
            If rasGetWidthHeight(PagePtr, Flags, Width, Height, Rotate, v) <> 0 Then
                BBox = CType(Marshal.PtrToStructure(v, GetType(TFltRect)), TFltRect)
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetWMFDefExtent(ByRef Width As Integer, ByRef Height As Integer) As Boolean
            Return CBool(pdfGetWMFDefExtent(m_Instance, Width, Height))
        End Function

        Public Function GetWMFPixelPerInch() As Integer
            Return pdfGetWMFPixelPerInch(m_Instance)
        End Function

        Public Function GetWordSpacing() As Double
            Return pdfGetWordSpacing(m_Instance)
        End Function

        Public Function HaveOpenDoc() As Boolean
            Return CBool(pdfHaveOpenDoc(m_Instance))
        End Function

        Public Function HaveOpenPage() As Boolean
            Return CBool(pdfHaveOpenPage(m_Instance))
        End Function

        Public Function HighlightAnnot(ByVal SubType As TAnnotType, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfHighlightAnnotW(m_Instance, SubType, PosX, PosY, Width, Height, Color, Author, Subject, Comment)
        End Function

        Public Function HighlightAnnotA(ByVal SubType As TAnnotType, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfHighlightAnnotA(m_Instance, SubType, PosX, PosY, Width, Height, Color, Author, Subject, Comment)
        End Function

        Public Function HighlightAnnotW(ByVal SubType As TAnnotType, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfHighlightAnnotW(m_Instance, SubType, PosX, PosY, Width, Height, Color, Author, Subject, Comment)
        End Function

        Public Function ImportBookmarks() As Integer
            Return pdfImportBookmarks(m_Instance)
        End Function

        Public Function ImportCatalogObjects() As Boolean
            Return CBool(pdfImportCatalogObjects(m_Instance))
        End Function

        Public Function ImportDocInfo() As Boolean
            Return CBool(pdfImportDocInfo(m_Instance))
        End Function

        Public Function ImportEncryptionSettings() As Boolean
            Return CBool(pdfImportEncryptionSettings(m_Instance))
        End Function

        Public Function ImportOCProperties() As Boolean
            Return CBool(pdfImportOCProperties(m_Instance))
        End Function

        Public Function ImportPage(ByVal PageNum As Integer) As Integer
            Return pdfImportPage(m_Instance, PageNum)
        End Function

        Public Function ImportPageEx(ByVal PageNum As Integer, ByVal ScaleX As Double, ByVal ScaleY As Double) As Integer
            Return pdfImportPageEx(m_Instance, PageNum, ScaleX, ScaleY)
        End Function

        Public Function ImportPDFFile(ByVal DestPage As Integer, ByVal ScaleX As Double, ByVal ScaleY As Double) As Integer
            Return pdfImportPDFFile(m_Instance, DestPage, ScaleX, ScaleY)
        End Function

        Public Function InitColorManagement(ByRef Profiles As TPDFColorProfiles, ByVal DestSpace As TPDFColorSpace, ByVal Flags As TPDFInitCMFlags) As Boolean
            Return CBool(pdfInitColorManagement(m_Instance, Profiles, DestSpace, Flags))
        End Function

        Public Function InitColorManagementEx(ByRef Profiles As TPDFColorProfilesEx, ByVal DestSpace As TPDFColorSpace, ByVal Flags As TPDFInitCMFlags) As Boolean
            Return CBool(pdfInitColorManagementEx(m_Instance, Profiles, DestSpace, Flags))
        End Function

        Public Function InitExtGState(ByRef GS As TPDFExtGState) As Boolean
            Return CBool(pdfInitExtGState(GS))
        End Function

        Public Function InitOCGContUsage(ByRef Value As TPDFOCGContUsage) As Boolean
            Value.StructSize = Marshal.SizeOf(Value)
            Return CBool(pdfInitOCGContUsage(Value))
        End Function

        Public Function InitStack(ByRef Stack As TPDFStack) As Boolean
            Return pdfInitStack(m_Instance, Stack) <> 0
        End Function

        Public Function InkAnnot(ByRef Points() As TFltPoint, ByVal LineWidth As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfInkAnnotW(m_Instance, Points, Points.Length, LineWidth, Color, CS, Author, Subject, Content)
        End Function

        Public Function InkAnnotA(ByRef Points() As TFltPoint, ByVal LineWidth As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfInkAnnotA(m_Instance, Points, Points.Length, LineWidth, Color, CS, Author, Subject, Content)
        End Function

        Public Function InkAnnotW(ByRef Points() As TFltPoint, ByVal LineWidth As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfInkAnnotW(m_Instance, Points, Points.Length, LineWidth, Color, CS, Author, Subject, Content)
        End Function

        Public Function InsertBMPFromBuffer(ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByRef Buffer() As Byte) As Integer
            Return pdfInsertBMPFromBuffer(m_Instance, PosX, PosY, ScaleWidth, ScaleHeight, Buffer)
        End Function

        Public Function InsertBMPFromHandle(ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal hBitmap As IntPtr) As Integer
            Return pdfInsertBMPFromHandle(m_Instance, PosX, PosY, ScaleWidth, ScaleHeight, hBitmap)
        End Function

        Public Function InsertBookmark(ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean, ByVal AddChildren As Boolean) As Integer
            Return pdfInsertBookmarkW(m_Instance, Title, Parent, DestPage, Convert.ToInt32(DoOpen), Convert.ToInt32(AddChildren))
        End Function

        Public Function InsertBookmarkA(ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean, ByVal AddChildren As Boolean) As Integer
            Return pdfInsertBookmarkA(m_Instance, Title, Parent, DestPage, Convert.ToInt32(DoOpen), Convert.ToInt32(AddChildren))
        End Function

        Public Function InsertBookmarkW(ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Boolean, ByVal AddChildren As Boolean) As Integer
            Return pdfInsertBookmarkW(m_Instance, Title, Parent, DestPage, Convert.ToInt32(DoOpen), Convert.ToInt32(AddChildren))
        End Function

        Public Function InsertBookmarkEx(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Boolean, ByVal AddChildren As Boolean) As Integer
            Return pdfInsertBookmarkExW(m_Instance, Title, Parent, NamedDest, Convert.ToInt32(DoOpen), Convert.ToInt32(AddChildren))
        End Function

        Public Function InsertBookmarkExA(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Boolean, ByVal AddChildren As Boolean) As Integer
            Return pdfInsertBookmarkExA(m_Instance, Title, Parent, NamedDest, Convert.ToInt32(DoOpen), Convert.ToInt32(AddChildren))
        End Function

        Public Function InsertBookmarkExW(ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Boolean, ByVal AddChildren As Boolean) As Integer
            Return pdfInsertBookmarkExW(m_Instance, Title, Parent, NamedDest, Convert.ToInt32(DoOpen), Convert.ToInt32(AddChildren))
        End Function

        Public Function InsertImage(ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal AFile As String) As Integer
            Return pdfInsertImage(m_Instance, PosX, PosY, ScaleWidth, ScaleHeight, AFile)
        End Function

        Public Function InsertImageEx(ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal AFile As String, ByVal Index As Integer) As Integer
            Return pdfInsertImageEx(m_Instance, PosX, PosY, ScaleWidth, ScaleHeight, AFile, Index)
        End Function

        Public Function InsertImageFromBuffer(ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByRef Buffer() As Byte, ByVal Index As Integer) As Integer
            Return pdfInsertImageFromBuffer(m_Instance, PosX, PosY, ScaleWidth, ScaleHeight, Buffer, Buffer.Length, Index)
        End Function

        Public Function InsertImageFromBuffer(ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal Buffer As IntPtr, ByVal Length As Integer, ByVal Index As Integer) As Integer
            Return pdfInsertImageFromBuffer2(m_Instance, PosX, PosY, ScaleWidth, ScaleHeight, Buffer, Length, Index)
        End Function

        Public Function InsertMetafile(ByVal FileName As String, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfInsertMetafile(m_Instance, FileName, PosX, PosY, Width, Height))
        End Function

        Public Function InsertMetafileEx(ByRef Buffer() As Byte, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfInsertMetafileEx(m_Instance, Buffer, Buffer.Length, PosX, PosY, Width, Height))
        End Function

        Public Function InsertMetafileExt(ByVal FileName As String, ByRef View As TRectL, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfInsertMetafileExt(m_Instance, FileName, View, PosX, PosY, Width, Height))
        End Function

        Public Function InsertMetafileExtEx(ByRef Buffer() As Byte, ByRef View As TRectL, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfInsertMetafileExtEx(m_Instance, Buffer, Buffer.Length, View, PosX, PosY, Width, Height))
        End Function

        Public Function InsertMetafileFromHandle(ByVal hEnhMetafile As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfInsertMetafileFromHandle(m_Instance, hEnhMetafile, PosX, PosY, Width, Height))
        End Function

        Public Function InsertMetafileFromHandleEx(ByVal hEnhMetafile As IntPtr, ByRef View As TRectL, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfInsertMetafileFromHandleEx(m_Instance, hEnhMetafile, View, PosX, PosY, Width, Height))
        End Function

        Public Function InsertRawImage(ByRef Buffer() As Byte, ByVal BitsPerPixel As Integer, ByVal ColorCount As Integer, ByVal ImgWidth As Integer, ByVal ImgHeight As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Integer
            Return pdfInsertRawImage(m_Instance, Buffer, BitsPerPixel, ColorCount, ImgWidth, ImgHeight, PosX, PosY, ScaleWidth, ScaleHeight)
        End Function

        Public Function InsertRawImage(ByVal Buffer As IntPtr, ByVal BitsPerPixel As Integer, ByVal ColorCount As Integer, ByVal ImgWidth As Integer, ByVal ImgHeight As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Integer
            Return pdfInsertRawImage(m_Instance, Buffer, BitsPerPixel, ColorCount, ImgWidth, ImgHeight, PosX, PosY, ScaleWidth, ScaleHeight)
        End Function

        Public Function InsertRawImageEx(ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByRef Image As TPDFRawImage) As Integer
            Image.StructSize = Marshal.SizeOf(Image)
            Return pdfInsertRawImageEx(m_Instance, PosX, PosY, ScaleWidth, ScaleHeight, Image)
        End Function

        Public Function IsBidiText(ByVal AText As String) As Integer
            Return pdfIsBidiText(m_Instance, AText)
        End Function

        Public Function IsColorPage(ByVal GrayIsColor As Boolean) As Integer
            Return pdfIsColorPage(m_Instance, CInt(GrayIsColor))
        End Function

        Public Function IsEmptyPage() As Integer
            Return pdfIsEmptyPage(m_Instance)
        End Function

        Public Function IsWrongPwd(ByVal ErrCode As Integer) As Boolean
            Return (((ErrCode = ENEED_PWD) Or (ErrCode = EWRONG_OPEN_PWD) Or (ErrCode = EWRONG_OWNER_PWD) Or (ErrCode = EWRONG_PWD)))
        End Function

        Public Function LineAnnot(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfLineAnnotW(m_Instance, x1, y1, x2, y2, LineWidth, lStart, lEnd, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function LineAnnotA(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfLineAnnotA(m_Instance, x1, y1, x2, y2, LineWidth, lStart, lEnd, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function LineAnnotW(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfLineAnnotW(m_Instance, x1, y1, x2, y2, LineWidth, lStart, lEnd, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function LineTo(ByVal PosX As Double, ByVal PosY As Double) As Boolean
            Return CBool(pdfLineTo(m_Instance, PosX, PosY))
        End Function

        Public Function LoadCMap(ByVal CMapName As String, ByVal Embed As Boolean) As Boolean
            Return CBool(pdfLoadCMap(m_Instance, CMapName, CInt(Embed)))
        End Function

        Public Function LoadFDFData(ByVal FileName As String, ByVal Password As String, ByVal Flags As Integer) As Boolean
            Return CBool(pdfLoadFDFDataW(m_Instance, FileName, Password, Flags))
        End Function

        Public Function LoadFDFDataA(ByVal FileName As String, ByVal Password As String, ByVal Flags As Integer) As Boolean
            Return CBool(pdfLoadFDFDataA(m_Instance, FileName, Password, Flags))
        End Function

        Public Function LoadFDFDataW(ByVal FileName As String, ByVal Password As String, ByVal Flags As Integer) As Boolean
            Return CBool(pdfLoadFDFDataW(m_Instance, FileName, Password, Flags))
        End Function

        Public Function LoadFDFDataEx(ByRef Buffer() As Byte, ByVal Password As String, ByVal Flags As Integer) As Boolean
            Return CBool(pdfLoadFDFDataEx(m_Instance, Buffer, Buffer.Length, Password, Flags))
        End Function

        Public Function LoadFont(ByRef Buffer() As Byte, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfLoadFont(m_Instance, Buffer, Buffer.Length, Style, Size, CInt(Embed), CP)
        End Function

        Public Function LoadFontEx(ByVal FontFile As String, ByVal Index As Integer, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfLoadFontEx(m_Instance, FontFile, Index, Style, Size, CInt(Embed), CP)
        End Function

        Public Function LoadLayerConfig(ByVal Index As Integer) As Boolean
            Return CBool(pdfLoadLayerConfig(m_Instance, Index))
        End Function

        Public Function LockLayer(ByVal Layer As Integer) As Boolean
            Return CBool(pdfLockLayer(m_Instance, Layer))
        End Function

        Public Function MovePage(ByVal Source As Integer, ByVal Dest As Integer) As Boolean
            Return CBool(pdfMovePage(m_Instance, Source, Dest))
        End Function

        Public Function MoveTo(ByVal PosX As Double, ByVal PosY As Double) As Boolean
            Return CBool(pdfMoveTo(m_Instance, PosX, PosY))
        End Function

        Public Sub MuliplyMatrix(ByRef M1 As TCTM, ByRef M2 As TCTM, ByRef NewMatrix As TCTM)
            NewMatrix.a = M2.a * M1.a + M2.b * M1.c
            NewMatrix.b = M2.a * M1.b + M2.b * M1.d
            NewMatrix.c = M2.c * M1.a + M2.d * M1.c
            NewMatrix.d = M2.c * M1.b + M2.d * M1.d
            NewMatrix.x = M2.x * M1.a + M2.y * M1.c + M1.x
            NewMatrix.y = M2.x * M1.b + M2.y * M1.d + M1.y
        End Sub

        Public Function OpenImportBuffer(ByRef Buffer() As Byte, ByVal PwdType As TPwdType, ByVal Password As String) As Integer
            Return pdfOpenImportBuffer(m_Instance, Buffer, Buffer.Length, PwdType, Password)
        End Function

        Public Function OpenImportFile(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal Password As String) As Integer
            Return pdfOpenImportFileW(m_Instance, FileName, PwdType, Password)
        End Function

        Public Function OpenImportFileA(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal Password As String) As Integer
            Return pdfOpenImportFileA(m_Instance, FileName, PwdType, Password)
        End Function

        Public Function OpenImportFileW(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal Password As String) As Integer
            Return pdfOpenImportFileW(m_Instance, FileName, PwdType, Password)
        End Function

        Public Function OpenOutputFile(ByVal OutPDF As String) As Boolean
            Return CBool(pdfOpenOutputFileW(m_Instance, OutPDF))
        End Function

        Public Function OpenOutputFileA(ByVal OutPDF As String) As Boolean
            Return CBool(pdfOpenOutputFileA(m_Instance, OutPDF))
        End Function

        Public Function OpenOutputFileW(ByVal OutPDF As String) As Boolean
            Return CBool(pdfOpenOutputFileW(m_Instance, OutPDF))
        End Function

        Public Function OpenOutputFileEncrypted(ByVal OutPDF As String, ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Boolean
            Return CBool(pdfOpenOutputFileEncrypted(m_Instance, OutPDF, OpenPwd, OwnerPwd, KeyLen, Restrict))
        End Function

        Public Function OpenTag(ByVal Tag As TPDFBaseTag, ByVal Lang As String, ByVal AltText As String, ByVal Expansion As String) As Boolean
            Return CBool(pdfOpenTagW(m_Instance, Tag, Lang, AltText, Expansion))
        End Function

        Public Function OpenTagA(ByVal Tag As TPDFBaseTag, ByVal Lang As String, ByVal AltText As String, ByVal Expansion As String) As Boolean
            Return CBool(pdfOpenTagA(m_Instance, Tag, Lang, AltText, Expansion))
        End Function

        Public Function OpenTagW(ByVal Tag As TPDFBaseTag, ByVal Lang As String, ByVal AltText As String, ByVal Expansion As String) As Boolean
            Return CBool(pdfOpenTagW(m_Instance, Tag, Lang, AltText, Expansion))
        End Function

        Public Function Optimize(ByVal Flags As TOptimizeFlags, ByVal Parms As TOptimizeParams) As Boolean
            If Not Parms Is Nothing Then Parms.StructSize = Marshal.SizeOf(Parms)
            Return CBool(pdfOptimize(m_Instance, Flags, Parms))
        End Function

        Public Function PageLink(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal DestPage As Integer) As Integer
            Return pdfPageLink(m_Instance, PosX, PosY, Width, Height, DestPage)
        End Function

        Public Function PageLink2(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal NamedDest As Integer) As Integer
            Return pdfPageLink2(m_Instance, PosX, PosY, Width, Height, NamedDest)
        End Function

        Public Function PageLink3(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal NamedDest As String) As Integer
            Return pdfPageLink3W(m_Instance, PosX, PosY, Width, Height, NamedDest)
        End Function

        Public Function PageLink3A(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal NamedDest As String) As Integer
            Return pdfPageLink3A(m_Instance, PosX, PosY, Width, Height, NamedDest)
        End Function

        Public Function PageLink3W(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal NamedDest As String) As Integer
            Return pdfPageLink3W(m_Instance, PosX, PosY, Width, Height, NamedDest)
        End Function

        Public Function PageLinkEx(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal DestType As TDestType, ByVal DestPage As Integer, ByVal a As Double, ByVal b As Double, ByVal C As Double, ByVal d As Double) As Integer
            Return pdfPageLinkEx(m_Instance, PosX, PosY, Width, Height, DestType, DestPage, a, b, C, d)
        End Function

        Public Function ParseContent(ByRef Stack As TPDFParseInterface, ByVal Flags As TParseFlags) As Boolean
            Return CBool(pdfParseContent(m_Instance, IntPtr.Zero, Stack, Flags))
        End Function

        Public Function PlaceImage(ByVal ImgHandle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Boolean
            Return CBool(pdfPlaceImage(m_Instance, ImgHandle, PosX, PosY, ScaleWidth, ScaleHeight))
        End Function

        Public Function PlaceSigFieldValidateIcon(ByVal SigField As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfPlaceSigFieldValidateIcon(m_Instance, SigField, PosX, PosY, Width, Height))
        End Function

        Public Function PlaceTemplate(ByVal TmplHandle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Boolean
            Return CBool(pdfPlaceTemplate(m_Instance, TmplHandle, PosX, PosY, ScaleWidth, ScaleHeight))
        End Function

        Public Function PlaceTemplateEx(ByVal TmplHandle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Boolean
            Return CBool(pdfPlaceTemplateEx(m_Instance, TmplHandle, PosX, PosY, ScaleWidth, ScaleHeight))
        End Function

        Public Function PolygonAnnot(ByVal Vertices() As TFltPoint, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfPolygonAnnotW(m_Instance, Vertices, Vertices.Length, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Content)
        End Function

        Public Function PolygonAnnotA(ByVal Vertices() As TFltPoint, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfPolygonAnnotA(m_Instance, Vertices, Vertices.Length, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Content)
        End Function

        Public Function PolygonAnnotW(ByVal Vertices() As TFltPoint, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfPolygonAnnotW(m_Instance, Vertices, Vertices.Length, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Content)
        End Function

        Public Function PolyLineAnnot(ByVal Vertices() As TFltPoint, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfPolyLineAnnotW(m_Instance, Vertices, Vertices.Length, LineWidth, lStart, lEnd, FillColor, StrokeColor, CS, Author, Subject, Content)
        End Function

        Public Function PolyLineAnnotA(ByVal Vertices() As TFltPoint, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfPolyLineAnnotA(m_Instance, Vertices, Vertices.Length, LineWidth, lStart, lEnd, FillColor, StrokeColor, CS, Author, Subject, Content)
        End Function

        Public Function PolyLineAnnotW(ByVal Vertices() As TFltPoint, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
            Return pdfPolyLineAnnotW(m_Instance, Vertices, Vertices.Length, LineWidth, lStart, lEnd, FillColor, StrokeColor, CS, Author, Subject, Content)
        End Function

        Public Function PrintPage(ByVal PageNum As Integer, ByVal DocName As String, ByVal DC As IntPtr, ByVal Flags As TPDFPrintFlags, ByRef Margin As TRectL) As Boolean
            Return CBool(pdfPrintPage(m_Instance, PageNum, DocName, DC, Flags, Margin, Nothing))
        End Function

        Public Function PrintPage(ByVal PageNum As Integer, ByVal DocName As String, ByVal DC As IntPtr, ByVal Flags As TPDFPrintFlags, ByRef Margin As TRectL, ByVal Parms As TPDFPrintParams) As Boolean
            Return CBool(pdfPrintPage(m_Instance, PageNum, DocName, DC, Flags, Margin, Parms))
        End Function

        Public Function PrintPDFFile(ByVal TmpDir As String, ByVal DocName As String, ByVal DC As IntPtr, ByVal Flags As TPDFPrintFlags, ByRef Margin As TRectL) As Boolean
            Return CBool(pdfPrintPDFFile(m_Instance, TmpDir, DocName, DC, Flags, Margin, Nothing))
        End Function

        Public Function PrintPDFFile(ByVal TmpDir As String, ByVal DocName As String, ByVal DC As IntPtr, ByVal Flags As TPDFPrintFlags, ByRef Margin As TRectL, ByVal Parms As TPDFPrintParams) As Boolean
            Return CBool(pdfPrintPDFFile(m_Instance, TmpDir, DocName, DC, Flags, Margin, Parms))
        End Function

        Public Function ReadImageFormat(ByVal FileName As String, ByRef Width As Integer, ByRef Height As Integer, ByRef BitsPerPixel As Integer, ByRef UseZip As Integer) As Boolean
            Return CBool(pdfReadImageFormat(m_Instance, FileName, Width, Height, BitsPerPixel, UseZip))
        End Function

        Public Function ReadImageFormat2(ByVal FileName As String, ByVal Index As Integer, ByRef Width As Integer, ByRef Height As Integer, ByRef BitsPerPixel As Integer, ByRef UseZip As Integer) As Boolean
            Return CBool(pdfReadImageFormat2(m_Instance, FileName, Index, Width, Height, BitsPerPixel, UseZip))
        End Function

        Public Function ReadImageFormatEx(ByVal hBitmap As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal BitsPerPixel As Integer, ByVal UseZip As Integer) As Boolean
            Return CBool(pdfReadImageFormatEx(m_Instance, hBitmap, Width, Height, BitsPerPixel, UseZip))
        End Function

        Public Function ReadImageFormatFromBuffer(ByRef Buffer() As Byte, ByVal Index As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal BitsPerPixel As Integer, ByVal UseZip As Integer) As Boolean
            Return CBool(pdfReadImageFormatFromBuffer(m_Instance, Buffer, Buffer.Length, Index, Width, Height, BitsPerPixel, UseZip))
        End Function

        Public Function ReadImageResolution(ByVal FileName As String, ByVal Index As Integer, ByRef ResX As Integer, ByRef ResY As Integer) As Boolean
            Return CBool(pdfReadImageResolution(m_Instance, FileName, Index, ResX, ResY))
        End Function

        Public Function ReadImageResolutionEx(ByRef Buffer() As Byte, ByVal Index As Integer, ByRef ResX As Integer, ByRef ResY As Integer) As Boolean
            Return CBool(pdfReadImageResolutionEx(m_Instance, Buffer, Buffer.Length, Index, ResX, ResY))
        End Function

        Public Function Rectangle(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfRectangle(m_Instance, PosX, PosY, Width, Height, FillMode))
        End Function

        Public Sub Redraw(ByVal RasPtr As IntPtr, ByVal DC As IntPtr, ByVal DestX As Integer, ByVal DestY As Integer)
            rasRedraw(RasPtr, DC, DestX, DestY)
        End Sub

        Public Function ReEncryptPDF(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal InPwd As String, ByVal NewOpenPwd As String, ByVal NewOwnerPwd As String, ByVal NewKeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Integer
            Return pdfReEncryptPDFW(m_Instance, FileName, PwdType, InPwd, NewOpenPwd, NewOwnerPwd, NewKeyLen, Restrict)
        End Function

        Public Function ReEncryptPDFA(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal InPwd As String, ByVal NewOpenPwd As String, ByVal NewOwnerPwd As String, ByVal NewKeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Integer
            Return pdfReEncryptPDFA(m_Instance, FileName, PwdType, InPwd, NewOpenPwd, NewOwnerPwd, NewKeyLen, Restrict)
        End Function

        Public Function ReEncryptPDFW(ByVal FileName As String, ByVal PwdType As TPwdType, ByVal InPwd As String, ByVal NewOpenPwd As String, ByVal NewOwnerPwd As String, ByVal NewKeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Integer
            Return pdfReEncryptPDFW(m_Instance, FileName, PwdType, InPwd, NewOpenPwd, NewOwnerPwd, NewKeyLen, Restrict)
        End Function

        Public Function RenameSpotColor(ByVal Colorant As String, ByVal NewName As String) As Integer
            Return pdfRenameSpotColor(m_Instance, Colorant, NewName)
        End Function

        Public Function RenderAnnotOrField(ByVal Handle As Integer, ByVal IsAnnot As Boolean, ByVal State As TButtonState, ByRef Matrix As TCTM, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByRef Out As TPDFBitmap) As Integer
            Out.StructSize = Marshal.SizeOf(Out)
            Return pdfRenderAnnotOrField(m_Instance, Handle, CInt(IsAnnot), State, Matrix, Flags, PixFmt, Filter, Out)
        End Function

        Public Function RenderPage(ByVal PagePtr As IntPtr, ByVal RasPtr As IntPtr, ByRef Img As TPDFRasterImage) As Boolean
            Return CBool(pdfRenderPage(m_Instance, PagePtr, RasPtr, Img))
        End Function

        Public Function RenderPageEx(ByVal DC As IntPtr, ByRef DestX As Integer, ByRef DestY As Integer, ByVal PagePtr As IntPtr, ByVal RasPtr As IntPtr, ByRef Img As TPDFRasterImage) As Boolean
            Return CBool(pdfRenderPageEx(m_Instance, DC, DestX, DestY, PagePtr, RasPtr, Img))
        End Function

        Public Function RenderPageToImage(ByVal PageNum As Integer, ByVal OutFile As String, ByVal Resolution As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfRenderPageToImage(m_Instance, PageNum, OutFile, Resolution, Width, Height, Flags, PixFmt, Filter, Format))
        End Function

        Public Function RenderPDFFile(ByVal OutFile As String, ByVal Resolution As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfRenderPDFFile(m_Instance, OutFile, Resolution, Flags, PixFmt, Filter, Format))
        End Function

        Public Function RenderPDFFileA(ByVal OutFile As String, ByVal Resolution As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfRenderPDFFileA(m_Instance, OutFile, Resolution, Flags, PixFmt, Filter, Format))
        End Function

        Public Function RenderPDFFileW(ByVal OutFile As String, ByVal Resolution As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfRenderPDFFileW(m_Instance, OutFile, Resolution, Flags, PixFmt, Filter, Format))
        End Function

        Public Function RenderPDFFileEx(ByVal OutFile As String, ByVal Resolution As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Boolean
            Return CBool(pdfRenderPDFFileEx(m_Instance, OutFile, Resolution, Width, Height, Flags, PixFmt, Filter, Format))
        End Function

        Public Function ReOpenImportFile(ByVal Handle As Integer) As Boolean
            Return CBool(pdfReOpenImportFile(m_Instance, Handle))
        End Function

        Public Function ReplaceFont(ByVal PDFFont As IntPtr, ByVal Name As String, ByVal Style As TFStyle, ByVal NameIsFamilyName As Boolean) As Integer
            Return pdfReplaceFontW(m_Instance, PDFFont, Name, Style, CInt(NameIsFamilyName))
        End Function

        Public Function ReplaceFontA(ByVal PDFFont As IntPtr, ByVal Name As String, ByVal Style As TFStyle, ByVal NameIsFamilyName As Boolean) As Integer
            Return pdfReplaceFontA(m_Instance, PDFFont, Name, Style, CInt(NameIsFamilyName))
        End Function

        Public Function ReplaceFontW(ByVal PDFFont As IntPtr, ByVal Name As String, ByVal Style As TFStyle, ByVal NameIsFamilyName As Boolean) As Integer
            Return pdfReplaceFontW(m_Instance, PDFFont, Name, Style, CInt(NameIsFamilyName))
        End Function

        Public Function ReplaceFontEx(ByVal PDFFont As IntPtr, ByVal FontFile As String, ByVal Embed As Boolean) As Integer
            Return pdfReplaceFontExW(m_Instance, PDFFont, FontFile, CInt(Embed))
        End Function

        Public Function ReplaceFontExA(ByVal PDFFont As IntPtr, ByVal FontFile As String, ByVal Embed As Boolean) As Integer
            Return pdfReplaceFontExA(m_Instance, PDFFont, FontFile, CInt(Embed))
        End Function

        Public Function ReplaceFontExW(ByVal PDFFont As IntPtr, ByVal FontFile As String, ByVal Embed As Boolean) As Integer
            Return pdfReplaceFontExW(m_Instance, PDFFont, FontFile, CInt(Embed))
        End Function

        Public Function ReplaceICCProfile(ByVal ColorSpace As Integer, ByVal ICCFile As String) As Integer
            Return pdfReplaceICCProfileW(m_Instance, ColorSpace, ICCFile)
        End Function

        Public Function ReplaceICCProfileA(ByVal ColorSpace As Integer, ByVal ICCFile As String) As Integer
            Return pdfReplaceICCProfileA(m_Instance, ColorSpace, ICCFile)
        End Function

        Public Function ReplaceICCProfileW(ByVal ColorSpace As Integer, ByVal ICCFile As String) As Integer
            Return pdfReplaceICCProfileW(m_Instance, ColorSpace, ICCFile)
        End Function

        Public Function ReplaceICCProfileEx(ByVal ColorSpace As Integer, ByVal Buffer As Byte()) As Integer
            Return pdfReplaceICCProfileEx(m_Instance, ColorSpace, Buffer, Buffer.Length)
        End Function

        Public Function ReplaceImage(ByVal Source As IntPtr, ByVal Image As String, ByVal Index As Integer, ByVal CS As TExtColorSpace, ByVal CSHandle As Integer, ByVal Flags As TReplaceImageFlags) As Boolean
            Return CBool(pdfReplaceImage(m_Instance, Source, Image, Index, CS, CSHandle, Flags))
        End Function

        Public Function ReplaceImageEx(ByVal Source As IntPtr, ByRef Buffer() As Byte, ByVal Index As Integer, ByVal CS As TExtColorSpace, ByVal CSHandle As Integer, ByVal Flags As TReplaceImageFlags) As Boolean
            Return CBool(pdfReplaceImageEx(m_Instance, Source, Buffer, Buffer.Length, Index, CS, CSHandle, Flags))
        End Function

        Public Function ReplacePageTextA(ByVal NewText As String, ByRef Stack As TPDFStack) As Boolean
            Return CBool(pdfReplacePageTextA(m_Instance, NewText, Stack))
        End Function

        Public Function ReplacePageTextEx(ByVal NewText As String, ByRef Stack As TPDFStack) As Boolean
            Return CBool(pdfReplacePageTextExW(m_Instance, NewText, Stack))
        End Function

        Public Function ReplacePageTextExA(ByVal NewText As String, ByRef Stack As TPDFStack) As Boolean
            Return CBool(pdfReplacePageTextExA(m_Instance, NewText, Stack))
        End Function

        Public Function ReplacePageTextExW(ByVal NewText As String, ByRef Stack As TPDFStack) As Boolean
            Return CBool(pdfReplacePageTextExW(m_Instance, NewText, Stack))
        End Function

        Public Function ResetEncryptionSettings() As Boolean
            Return CBool(pdfResetEncryptionSettings(m_Instance))
        End Function

        Public Function ResetLineDashPattern() As Boolean
            Return CBool(pdfResetLineDashPattern(m_Instance))
        End Function

        Public Function ResizeBitmap(ByVal RasPtr As IntPtr, ByVal DC As IntPtr, ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Return CBool(rasResizeBitmap(RasPtr, DC, Width, Height))
        End Function

        Public Function RestoreGraphicState() As Boolean
            Return CBool(pdfRestoreGraphicState(m_Instance))
        End Function

        Public Function RotateCoords(ByVal alpha As Double, ByVal OriginX As Double, ByVal OriginY As Double) As Integer
            Return pdfRotateCoords(m_Instance, alpha, OriginX, OriginY)
        End Function

        Public Function RoundRect(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Radius As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfRoundRect(m_Instance, PosX, PosY, Width, Height, Radius, FillMode))
        End Function

        Public Function RoundRectEx(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal rWidth As Double, ByVal rHeight As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfRoundRectEx(m_Instance, PosX, PosY, Width, Height, rWidth, rHeight, FillMode))
        End Function

        Public Function SaveGraphicState() As Boolean
            Return CBool(pdfSaveGraphicState(m_Instance))
        End Function

        Public Function ScaleCoords(ByVal sx As Double, ByVal sy As Double) As Integer
            Return pdfScaleCoords(m_Instance, sx, sy)
        End Function

        Public Function SelfTest() As Boolean
            Return CBool(pdfSelfTest(m_Instance))
        End Function

        Public Function Set3DAnnotProps(ByVal Annot As Integer, ByVal ActType As T3DActivationType, ByVal DeActType As T3DDeActivateType, ByVal InstType As T3DInstanceType, ByVal DeInstType As T3DDeActInstance, ByVal DisplToolbar As Boolean, ByVal DisplModelTree As Boolean) As Boolean
            Return pdfSet3DAnnotProps(m_Instance, Annot, ActType, DeActType, InstType, DeInstType, CInt(DisplToolbar), CInt(DisplModelTree)) <> 0
        End Function

        Public Function Set3DAnnotScriptA(ByVal Annot As Integer, ByVal Value As String) As Boolean
            Return pdfSet3DAnnotScriptA(m_Instance, Annot, Value, Value.Length) <> 0
        End Function

        Public Function SetAllocBy(ByVal Value As Integer) As Integer
            Return pdfSetAllocBy(m_Instance, Value)
        End Function

        Public Function SetAnnotBorderStyle(ByVal Handle As Integer, ByVal Style As TBorderStyle) As Boolean
            Return CBool(pdfSetAnnotBorderStyle(m_Instance, Handle, Style))
        End Function

        Public Function SetAnnotBorderWidth(ByVal Handle As Integer, ByVal LineWidth As Double) As Boolean
            Return CBool(pdfSetAnnotBorderWidth(m_Instance, Handle, LineWidth))
        End Function

        Public Function SetAnnotColor(ByVal Handle As Integer, ByVal ColorType As TAnnotColor, ByVal CS As TPDFColorSpace, ByVal Color As Integer) As Boolean
            Return CBool(pdfSetAnnotColor(m_Instance, Handle, ColorType, CS, Color))
        End Function

        Public Function SetAnnotFlags(ByVal Flags As TAnnotFlags) As Boolean
            Return CBool(pdfSetAnnotFlags(m_Instance, Flags))
        End Function

        Public Function SetAnnotFlagsEx(ByVal Handle As Integer, ByVal Flags As Integer) As Boolean
            Return CBool(pdfSetAnnotFlagsEx(m_Instance, Handle, Flags))
        End Function

        Public Function SetAnnotHighlightMode(ByVal Handle As Integer, ByVal Mode As THighlightMode) As Boolean
            Return CBool(pdfSetAnnotHighlightMode(m_Instance, Handle, Mode))
        End Function

        Public Function SetAnnotIcon(ByVal Handle As Integer, ByVal Icon As TAnnotIcon) As Boolean
            Return CBool(pdfSetAnnotIcon(m_Instance, Handle, Icon))
        End Function

        Public Function SetAnnotLineDashPattern(ByVal Handle As Integer, ByRef Dash() As Single, ByVal NumValues As Integer) As Boolean
            Return CBool(pdfSetAnnotLineDashPattern(m_Instance, Handle, Dash, NumValues))
        End Function

        Public Function SetAnnotLineEndStyle(ByVal Handle As Integer, ByVal StartStyle As TLineEndStyle, ByVal EndStyle As TLineEndStyle) As Boolean
            Return CBool(pdfSetAnnotLineEndStyle(m_Instance, Handle, StartStyle, EndStyle))
        End Function

        Public Function SetAnnotMigrationState(ByVal Annot As Integer, ByVal State As TAnnotState, ByVal User As String) As Integer
            Return pdfSetAnnotMigrationState(m_Instance, Annot, State, User)
        End Function

        Public Function SetAnnotOpacity(ByVal Handle As Integer, ByVal Value As Double) As Boolean
            Return CBool(pdfSetAnnotOpacity(m_Instance, Handle, Value))
        End Function

        Public Function SetAnnotOpenState(ByVal Handle As Integer, ByVal Open As Boolean) As Boolean
            Return CBool(pdfSetAnnotOpenState(m_Instance, Handle, Convert.ToInt32(Open)))
        End Function

        Public Function SetAnnotOrFieldDate(ByVal Handle As Integer, ByVal IsField As Integer, ByVal DateType As TDateType, ByVal DateTime As Integer) As Boolean
            Return CBool(pdfSetAnnotOrFieldDate(m_Instance, Handle, IsField, DateType, DateTime))
        End Function

        Public Function SetAnnotQuadPoints(ByVal Handle As Integer, ByVal Value() As TFltPoint) As Boolean
            Return pdfSetAnnotQuadPoints(m_Instance, Handle, Value, Value.Length) <> 0
        End Function

        Public Function SetAnnotString(ByVal Handle As Integer, ByVal StringType As TAnnotString, ByVal Value As String) As Boolean
            Return CBool(pdfSetAnnotStringW(m_Instance, Handle, StringType, Value))
        End Function

        Public Function SetAnnotStringA(ByVal Handle As Integer, ByVal StringType As TAnnotString, ByVal Value As String) As Boolean
            Return CBool(pdfSetAnnotStringA(m_Instance, Handle, StringType, Value))
        End Function

        Public Function SetAnnotStringW(ByVal Handle As Integer, ByVal StringType As TAnnotString, ByVal Value As String) As Boolean
            Return CBool(pdfSetAnnotStringW(m_Instance, Handle, StringType, Value))
        End Function

        Public Function SetAnnotSubject(ByVal Handle As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetAnnotSubjectW(m_Instance, Handle, Value))
        End Function

        Public Function SetAnnotSubjectA(ByVal Handle As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetAnnotSubjectA(m_Instance, Handle, Value))
        End Function

        Public Function SetAnnotSubjectW(ByVal Handle As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetAnnotSubjectW(m_Instance, Handle, Value))
        End Function

        Public Function SetBBox(ByVal Boundary As TPageBoundary, ByVal LeftX As Double, ByVal LeftY As Double, ByVal RightX As Double, ByVal RightY As Double) As Boolean
            Return CBool(pdfSetBBox(m_Instance, Boundary, LeftX, LeftY, RightX, RightY))
        End Function

        Public Function SetBidiMode(ByVal Mode As TPDFBidiMode) As Boolean
            Return pdfSetBidiMode(m_Instance, Mode) <> 0
        End Function

        Public Function SetBookmarkDest(ByVal ABmk As Integer, ByVal DestType As TDestType, ByVal a As Double, ByVal b As Double, ByVal C As Double, ByVal d As Double) As Boolean
            Return CBool(pdfSetBookmarkDest(m_Instance, ABmk, DestType, a, b, C, d))
        End Function

        Public Function SetBookmarkStyle(ByVal ABmk As Integer, ByVal Style As TFStyle, ByVal RGBColor As Integer) As Boolean
            Return CBool(pdfSetBookmarkStyle(m_Instance, ABmk, Style, RGBColor))
        End Function

        Public Function SetAnnotBorderEffect(ByVal Handle As Integer, ByVal Value As TBorderEffect) As Boolean
            Return pdfSetAnnotBorderEffect(m_Instance, Handle, Value) <> 0
        End Function

        Public Function SetBorderStyle(ByVal Style As TBorderStyle) As Boolean
            Return CBool(pdfSetBorderStyle(m_Instance, Style))
        End Function

        Public Function SetCharacterSpacing(ByVal Value As Double) As Boolean
            Return CBool(pdfSetCharacterSpacing(m_Instance, Value))
        End Function

        Public Function SetCheckBoxChar(ByVal CheckBoxChar As TCheckBoxChar) As Boolean
            Return CBool(pdfSetCheckBoxChar(m_Instance, CheckBoxChar))
        End Function

        Public Function SetCheckBoxDefState(ByVal Field As Integer, ByVal Checked As Boolean) As Boolean
            Return CBool(pdfSetCheckBoxDefState(m_Instance, Field, CInt(Checked)))
        End Function

        Public Function SetCheckBoxState(ByVal Field As Integer, ByVal Checked As Boolean) As Boolean
            Return CBool(pdfSetCheckBoxState(m_Instance, Field, CInt(Checked)))
        End Function

        Public Function SetCIDFont(ByVal CMapHandle As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean) As Integer
            Return pdfSetCIDFontW(m_Instance, CMapHandle, Name, Style, Size, CInt(Embed))
        End Function

        Public Function SetCIDFontA(ByVal CMapHandle As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean) As Integer
            Return pdfSetCIDFontA(m_Instance, CMapHandle, Name, Style, Size, CInt(Embed))
        End Function

        Public Function SetCIDFontW(ByVal CMapHandle As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean) As Integer
            Return pdfSetCIDFontW(m_Instance, CMapHandle, Name, Style, Size, CInt(Embed))
        End Function

        Public Function SetCMapDir(ByVal Path As String, ByVal Flags As TLoadCMapFlags) As Integer
            Return pdfSetCMapDirW(m_Instance, Path, Flags)
        End Function

        Public Function SetCMapDirA(ByVal Path As String, ByVal Flags As TLoadCMapFlags) As Integer
            Return pdfSetCMapDirA(m_Instance, Path, Flags)
        End Function

        Public Function SetCMapDirW(ByVal Path As String, ByVal Flags As TLoadCMapFlags) As Integer
            Return pdfSetCMapDirW(m_Instance, Path, Flags)
        End Function

        Public Function SetColDefFile(ByVal EmbFile As Integer) As Boolean
            Return CBool(pdfSetColDefFile(m_Instance, EmbFile))
        End Function

        Public Function SetColorMask(ByVal ImageHandle As Integer, ByRef Mask() As Integer) As Boolean
            Dim len As Integer = 0
            If Not Mask Is Nothing Then
                len = Mask.Length
            End If
            Return pdfSetColorMask(m_Instance, ImageHandle, Mask, len) <> 0
        End Function

        Public Function SetColors(ByVal Color As Integer) As Boolean
            Return CBool(pdfSetColors(m_Instance, Color))
        End Function

        Public Function SetColorSpace(ByVal ColorSpace As TPDFColorSpace) As Boolean
            Return CBool(pdfSetColorSpace(m_Instance, ColorSpace))
        End Function

        Public Function SetColSortField(ByVal ColField As Integer, ByVal AscendingOrder As Boolean) As Boolean
            Return CBool(pdfSetColSortField(m_Instance, ColField, CInt(AscendingOrder)))
        End Function

        Public Function SetCompressionFilter(ByVal ComprFilter As TCompressionFilter) As Boolean
            Return CBool(pdfSetCompressionFilter(m_Instance, ComprFilter))
        End Function

        Public Function SetCompressionLevel(ByVal CompressLevel As TCompressionLevel) As Boolean
            Return CBool(pdfSetCompressionLevel(m_Instance, CompressLevel))
        End Function

        Public Function SetContent(ByRef Buffer() As Byte) As Boolean
            Return CBool(pdfSetContent(m_Instance, Buffer, Buffer.Length))
        End Function

        Public Function SetDateTimeFormat(ByVal TxtField As Integer, ByVal Fmt As TPDFDateTime) As Boolean
            Return CBool(pdfSetDateTimeFormat(m_Instance, TxtField, Fmt))
        End Function

        Public Function SetDefBitsPerPixel(ByVal Value As Integer) As Boolean
            Return CBool(pdfSetDefBitsPerPixel(m_Instance, Value))
        End Function

        Public Function SetDocInfo(ByVal DInfo As TDocumentInfo, ByVal Value As String) As Boolean
            Return CBool(pdfSetDocInfoW(m_Instance, DInfo, Value))
        End Function

        Public Function SetDocInfoA(ByVal DInfo As TDocumentInfo, ByVal Value As String) As Boolean
            Return CBool(pdfSetDocInfoA(m_Instance, DInfo, Value))
        End Function

        Public Function SetDocInfoW(ByVal DInfo As TDocumentInfo, ByVal Value As String) As Boolean
            Return CBool(pdfSetDocInfoW(m_Instance, DInfo, Value))
        End Function

        Public Function SetDocInfoEx(ByVal DInfo As TDocumentInfo, ByVal Key As String, ByVal Value As String) As Boolean
            Return CBool(pdfSetDocInfoExW(m_Instance, DInfo, Key, Value))
        End Function

        Public Function SetDocInfoExA(ByVal DInfo As TDocumentInfo, ByVal Key As String, ByVal Value As String) As Boolean
            Return CBool(pdfSetDocInfoExA(m_Instance, DInfo, Key, Value))
        End Function

        Public Function SetDocInfoExW(ByVal DInfo As TDocumentInfo, ByVal Key As String, ByVal Value As String) As Boolean
            Return CBool(pdfSetDocInfoExW(m_Instance, DInfo, Key, Value))
        End Function

        Public Function SetDrawDirection(ByVal Direction As TDrawDirection) As Boolean
            Return CBool(pdfSetDrawDirection(m_Instance, Direction))
        End Function

        Public Function SetEMFFrameDPI(ByVal DPIX As Integer, ByVal DPIY As Integer) As Boolean
            Return CBool(pdfSetEMFFrameDPI(m_Instance, DPIX, DPIY))
        End Function

        Public Function SetEMFPatternDistance(ByVal Value As Double) As Boolean
            Return CBool(pdfSetEMFPatternDistance(m_Instance, Value))
        End Function

        Public Function SetErrorMode(ByVal ErrMode As TErrMode) As Boolean
            Return CBool(pdfSetErrorMode(m_Instance, ErrMode))
        End Function

        Public Function SetExtColorSpace(ByVal Handle As Integer) As Boolean
            Return CBool(pdfSetExtColorSpace(m_Instance, Handle))
        End Function

        Public Function SetExtFillColorSpace(ByVal Handle As Integer) As Boolean
            Return CBool(pdfSetExtFillColorSpace(m_Instance, Handle))
        End Function

        Public Function SetExtGState(ByVal Handle As Integer) As Boolean
            Return CBool(pdfSetExtGState(m_Instance, Handle))
        End Function

        Public Function SetExtStrokeColorSpace(ByVal Handle As Integer) As Boolean
            Return CBool(pdfSetExtStrokeColorSpace(m_Instance, Handle))
        End Function

        Public Function SetFieldBackColor(ByVal AColor As Integer) As Boolean
            Return CBool(pdfSetFieldBackColor(m_Instance, AColor))
        End Function

        Public Function SetFieldBBox(ByVal Field As Integer, ByRef BBox As TPDFRect) As Boolean
            Return CBool(pdfSetFieldBBox(m_Instance, Field, BBox))
        End Function

        Public Function SetFieldBorderColor(ByVal AColor As Integer) As Boolean
            Return CBool(pdfSetFieldBorderColor(m_Instance, AColor))
        End Function

        Public Function SetFieldBorderStyle(ByVal Field As Integer, ByVal Style As TBorderStyle) As Boolean
            Return CBool(pdfSetFieldBorderStyle(m_Instance, Field, Style))
        End Function

        Public Function SetFieldBorderWidth(ByVal Field As Integer, ByVal LineWidth As Double) As Boolean
            Return CBool(pdfSetFieldBorderWidth(m_Instance, Field, LineWidth))
        End Function

        Public Function SetFieldColor(ByVal Field As Integer, ByVal ColorType As TFieldColor, ByVal CS As TPDFColorSpace, ByVal Color As Integer) As Boolean
            Return CBool(pdfSetFieldColor(m_Instance, Field, ColorType, CS, Color))
        End Function

        Public Function SetFieldExpValue(ByVal Field As Integer, ByVal ValIndex As Integer, ByVal Value As String, ByVal ExpValue As String, ByVal Selected As Boolean) As Boolean
            Return CBool(pdfSetFieldExpValueW(m_Instance, Field, ValIndex, Value, ExpValue, CInt(Selected)))
        End Function

        Public Function SetFieldExpValueA(ByVal Field As Integer, ByVal ValIndex As Integer, ByVal Value As String, ByVal ExpValue As String, ByVal Selected As Boolean) As Boolean
            Return CBool(pdfSetFieldExpValueA(m_Instance, Field, ValIndex, Value, ExpValue, CInt(Selected)))
        End Function

        Public Function SetFieldExpValueW(ByVal Field As Integer, ByVal ValIndex As Integer, ByVal Value As String, ByVal ExpValue As String, ByVal Selected As Boolean) As Boolean
            Return CBool(pdfSetFieldExpValueW(m_Instance, Field, ValIndex, Value, ExpValue, CInt(Selected)))
        End Function

        Public Function SetFieldExpValueEx(ByVal AField As Integer, ByVal ValIndex As Integer, ByVal Selected As Boolean, ByVal DefSelected As Boolean) As Boolean
            Return CBool(pdfSetFieldExpValueEx(m_Instance, AField, ValIndex, CInt(Selected), CInt(DefSelected)))
        End Function

        Public Function SetFieldFlags(ByVal Field As Integer, ByVal Flags As TFieldFlags, ByVal DoReset As Boolean) As Boolean
            Return CBool(pdfSetFieldFlags(m_Instance, Field, Flags, CInt(DoReset)))
        End Function

        Public Function SetFieldFont(ByVal Field As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfSetFieldFont(m_Instance, Field, Name, Style, Size, CInt(Embed), CP)
        End Function

        Public Function SetFieldFontEx(ByVal Field As Integer, ByVal Handle As Integer, ByVal FontSize As Double) As Integer
            Return pdfSetFieldFontEx(m_Instance, Field, Handle, FontSize)
        End Function

        Public Function SetFieldFontSize(ByVal Field As Integer, ByVal FontSize As Double) As Boolean
            Return CBool(pdfSetFieldFontSize(m_Instance, Field, FontSize))
        End Function

        Public Function SetFieldHighlightMode(ByVal Field As Integer, ByVal Mode As THighlightMode) As Boolean
            Return CBool(pdfSetFieldHighlightMode(m_Instance, Field, Mode))
        End Function

        Public Function SetFieldIndex(ByVal Field As Integer, ByVal Index As Integer) As Boolean
            Return CBool(pdfSetFieldIndex(m_Instance, Field, Index))
        End Function

        Public Function SetFieldMapName(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetFieldMapNameW(m_Instance, Field, Value))
        End Function

        Public Function SetFieldMapNameA(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetFieldMapNameA(m_Instance, Field, Value))
        End Function

        Public Function SetFieldMapNameW(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetFieldMapNameW(m_Instance, Field, Value))
        End Function

        Public Function SetFieldName(ByVal Field As Integer, ByVal NewName As String) As Boolean
            Return CBool(pdfSetFieldName(m_Instance, Field, NewName))
        End Function

        Public Function SetFieldOrientation(ByVal Field As Integer, ByVal Value As Integer) As Boolean
            Return CBool(pdfSetFieldOrientation(m_Instance, Field, Value))
        End Function

        Public Function SetFieldTextAlign(ByVal Field As Integer, ByVal Align As TTextAlign) As Boolean
            Return CBool(pdfSetFieldTextAlign(m_Instance, Field, Align))
        End Function

        Public Function SetFieldTextColor(ByVal Color As Integer) As Boolean
            Return CBool(pdfSetFieldTextColor(m_Instance, Color))
        End Function

        Public Function SetFieldToolTip(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetFieldToolTipW(m_Instance, Field, Value))
        End Function

        Public Function SetFieldToolTipA(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetFieldToolTipA(m_Instance, Field, Value))
        End Function

        Public Function SetFieldToolTipW(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return CBool(pdfSetFieldToolTipW(m_Instance, Field, Value))
        End Function

        Public Function SetFillColor(ByVal Color As Integer) As Boolean
            Return CBool(pdfSetFillColor(m_Instance, Color))
        End Function

        Public Function SetFillColorEx(ByVal Color() As Byte, ByVal NumComponents As Integer) As Boolean
            Return CBool(pdfSetFillColorEx(m_Instance, Color, NumComponents))
        End Function

        Public Function SetFillColorF(ByVal Color() As Single, ByVal NumComponents As Integer) As Boolean
            Return CBool(pdfSetFillColorF(m_Instance, Color, NumComponents))
        End Function

        Public Function SetFloatPrecision(ByVal NumTextDecDigits As Integer, ByVal NumVectDecDigits As Integer) As Boolean
            Return pdfSetFloatPrecision(m_Instance, NumTextDecDigits, NumVectDecDigits) <> 0
        End Function

        Public Function SetFont(ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfSetFontW(m_Instance, Name, Style, Size, CInt(Embed), CP)
        End Function

        Public Function SetFontA(ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfSetFontA(m_Instance, Name, Style, Size, CInt(Embed), CP)
        End Function

        Public Function SetFontW(ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfSetFontW(m_Instance, Name, Style, Size, CInt(Embed), CP)
        End Function

        Public Function SetFontEx(ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfSetFontExW(m_Instance, Name, Style, Size, CInt(Embed), CP)
        End Function

        Public Function SetFontExA(ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfSetFontExA(m_Instance, Name, Style, Size, CInt(Embed), CP)
        End Function

        Public Function SetFontExW(ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Boolean, ByVal CP As TCodepage) As Integer
            Return pdfSetFontExW(m_Instance, Name, Style, Size, CInt(Embed), CP)
        End Function

        Public Function SetFontOrigin(ByVal Origin As TOrigin) As Boolean
            Return CBool(pdfSetFontOrigin(m_Instance, Origin))
        End Function

        Public Sub SetFontSearchOrder(ByVal Order() As TFontBaseType)
            pdfSetFontSearchOrder(m_Instance, Order)
        End Sub

        Public Sub SetFontSearchOrderEx(ByVal S1 As TFontBaseType, ByVal S2 As TFontBaseType, ByVal S3 As TFontBaseType, ByVal S4 As TFontBaseType)
            pdfSetFontSearchOrderEx(m_Instance, S1, S2, S3, S4)
        End Sub

        Public Function SetFontSelMode(ByVal Mode As TFontSelMode) As Boolean
            Return CBool(pdfSetFontSelMode(m_Instance, Mode))
        End Function

        Public Function SetFontWeight(ByVal Weight As Integer) As Boolean
            Return CBool(pdfSetFontWeight(m_Instance, Weight))
        End Function

        Public Sub SetGStateFlags(ByVal Flags As TGStateFlags, ByVal Reset As Boolean)
            pdfSetGStateFlags(m_Instance, Flags, CInt(Reset))
        End Sub

        Public Function SetIconColor(ByVal Color As Integer) As Boolean
            Return CBool(pdfSetIconColor(m_Instance, Color))
        End Function

        Public Function SetImportFlags(ByVal Flags As TImportFlags) As Boolean
            Return CBool(pdfSetImportFlags(m_Instance, Flags))
        End Function

        Public Function SetImportFlags2(ByVal Flags As TImportFlags2) As Boolean
            Return CBool(pdfSetImportFlags2(m_Instance, Flags))
        End Function

        Public Function SetItalicAngle(ByVal Angle As Double) As Boolean
            Return CBool(pdfSetItalicAngle(m_Instance, Angle))
        End Function

        Public Function SetJPEGQuality(ByVal Value As Integer) As Boolean
            Return CBool(pdfSetJPEGQuality(m_Instance, Value))
        End Function

        Public Function SetLanguage(ByVal ISOTag As String) As Boolean
            Return CBool(pdfSetLanguage(m_Instance, ISOTag))
        End Function

        Public Function SetLeading(ByVal Value As Double) As Boolean
            Return CBool(pdfSetLeading(m_Instance, Value))
        End Function

        Public Function SetLicenseKey(ByVal Value As String) As Boolean
            Return CBool(pdfSetLicenseKey(m_Instance, Value))
        End Function

        Public Function SetLineAnnotParms(ByVal Handle As Integer, ByVal FontHandle As Integer, ByVal FontSize As Double, ByVal Parms As TLineAnnotParms) As Boolean
            If Not Parms Is Nothing Then Parms.StructSize = Marshal.SizeOf(Parms)
            Return CBool(pdfSetLineAnnotParms(m_Instance, Handle, FontHandle, FontSize, Parms))
        End Function

        Public Function SetLineCapStyle(ByVal Style As TLineCapStyle) As Boolean
            Return CBool(pdfSetLineCapStyle(m_Instance, Style))
        End Function

        Public Function SetLineDashPattern(ByVal Dash As String, ByVal Phase As Integer) As Boolean
            Return CBool(pdfSetLineDashPattern(m_Instance, Dash, Phase))
        End Function

        Public Function SetLineDashPatternEx(ByRef Dash() As Double, ByVal NumValues As Integer, ByVal Phase As Integer) As Boolean
            Return CBool(pdfSetLineDashPatternEx(m_Instance, Dash, NumValues, Phase))
        End Function

        Public Function SetLineJoinStyle(ByVal Style As TLineJoinStyle) As Boolean
            Return CBool(pdfSetLineJoinStyle(m_Instance, Style))
        End Function

        Public Function SetLineWidth(ByVal Value As Double) As Boolean
            Return CBool(pdfSetLineWidth(m_Instance, Value))
        End Function

        Public Function SetLinkHighlightMode(ByVal Mode As THighlightMode) As Boolean
            Return CBool(pdfSetLinkHighlightMode(m_Instance, Mode))
        End Function

        Public Function SetListFont(ByVal Handle As Integer) As Boolean
            Return CBool(pdfSetListFont(m_Instance, Handle))
        End Function

        Public Function SetMatrix(ByRef Matrix As TCTM) As Boolean
            Return CBool(pdfSetMatrix(m_Instance, Matrix))
        End Function

        Public Sub SetMaxErrLogMsgCount(ByVal Value As Integer)
            pdfSetMaxErrLogMsgCount(m_Instance, Value)
        End Sub

        Public Function SetMaxFieldLen(ByVal TxtField As Integer, ByVal MaxLen As Integer) As Boolean
            Return CBool(pdfSetMaxFieldLen(m_Instance, TxtField, MaxLen))
        End Function

        Public Function SetMetaConvFlags(ByVal Flags As TMetaFlags) As Boolean
            Return CBool(pdfSetMetaConvFlags(m_Instance, Flags))
        End Function

        Public Function SetMetadata(ByVal ObjType As TMetadataObj, ByVal Handle As Integer, ByRef Buffer() As Byte) As Boolean
            Return pdfSetMetadata(m_Instance, ObjType, Handle, Buffer, Buffer.Length) <> 0
        End Function

        Public Function SetMiterLimit(ByVal Value As Double) As Boolean
            Return CBool(pdfSetMiterLimit(m_Instance, Value))
        End Function

        Public Function SetNeedAppearance(ByVal Value As Boolean) As Boolean
            Return pdfSetNeedAppearance(m_Instance, CInt(Value)) <> 0
        End Function

        Public Function SetNumberFormat(ByVal TxtField As Integer, ByVal Sep As TDecSeparator, ByVal DecPlaces As Integer, ByVal NegStyle As TNegativeStyle, ByVal CurrStr As String, ByVal Prepend As Integer) As Boolean
            Return CBool(pdfSetNumberFormat(m_Instance, TxtField, Sep, DecPlaces, NegStyle, CurrStr, Prepend))
        End Function

        Public Function SetOCGContUsage(ByVal Handle As Integer, ByRef Value As TPDFOCGContUsage) As Boolean
            Value.StructSize = Marshal.SizeOf(Value)
            Return CBool(pdfSetOCGContUsage(m_Instance, Handle, Value))
        End Function

        Public Function SetOCGState(ByVal Handle As Integer, ByVal Visible As Boolean, ByVal SaveState As Boolean) As Boolean
            Return CBool(pdfSetOCGState(m_Instance, Handle, CInt(Visible), CInt(SaveState)))
        End Function

        Public Function SetOnErrorProc(ByVal ErrProc As TErrorProc) As Boolean
            m_AddrErrorProc = ErrProc
            Return CBool(pdfSetOnErrorProc(m_Instance, IntPtr.Zero, m_AddrErrorProc))
        End Function

        Public Function SetOnPageBreakProc(ByVal PageBreakProc As TOnPageBreakProc) As Boolean
            m_AddrOnPageBreak = PageBreakProc
            Return CBool(pdfSetOnPageBreakProc(m_Instance, IntPtr.Zero, m_AddrOnPageBreak))
        End Function

        Public Function SetOpacity(ByVal Value As Double) As Boolean
            Return CBool(pdfSetOpacity(m_Instance, Value))
        End Function

        Public Function SetOrientation(ByVal Value As Integer) As Boolean
            Return CBool(pdfSetOrientation(m_Instance, Value))
        End Function

        Public Function SetOrientationEx(ByVal Value As Integer) As Boolean
            Return CBool(pdfSetOrientationEx(m_Instance, Value))
        End Function

        Public Function SetPageBBox(ByVal PagePtr As IntPtr, ByVal Boundary As TPageBoundary, ByRef BBox As TFltRect) As Boolean
            SetPageBBox = pdfSetPageBBox(PagePtr, Boundary, BBox) <> 0
        End Function

        Public Function SetPageCoords(ByVal PageCoords As TPageCoord) As Boolean
            Return CBool(pdfSetPageCoords(m_Instance, PageCoords))
        End Function

        Public Function SetPageFormat(ByVal Value As TPageFormat) As Boolean
            Return CBool(pdfSetPageFormat(m_Instance, Value))
        End Function

        Public Function SetPageHeight(ByVal Value As Double) As Boolean
            Return CBool(pdfSetPageHeight(m_Instance, Value))
        End Function

        Public Function SetPageLayout(ByVal Layout As TPageLayout) As Boolean
            Return CBool(pdfSetPageLayout(m_Instance, Layout))
        End Function

        Public Function SetPageMode(ByVal Mode As TPageMode) As Boolean
            Return CBool(pdfSetPageMode(m_Instance, Mode))
        End Function

        Public Function SetPageWidth(ByVal Value As Double) As Boolean
            Return CBool(pdfSetPageWidth(m_Instance, Value))
        End Function

        Public Function SetPDFVersion(ByVal Version As TPDFVersion) As Boolean
            Return CBool(pdfSetPDFVersion(m_Instance, Version))
        End Function

        Public Function SetPrintSettings(ByVal Mode As TDuplexMode, ByVal PickTrayByPDFSize As Integer, ByVal NumCopies As Integer, ByVal PrintScaling As TPrintScaling, ByRef PrintRanges() As Integer) As Boolean
            Return pdfSetPrintSettings(m_Instance, Mode, PickTrayByPDFSize, NumCopies, PrintScaling, PrintRanges, CInt(PrintRanges.Length / 2)) <> 0
        End Function

        Public Function SetProgressProc(ByVal InitProc As TInitProgress, ByVal ProgressProc As TProgress) As Integer
            m_AddrInitProgress = InitProc
            m_AddrProgress = ProgressProc
            Return pdfSetProgressProc(m_Instance, IntPtr.Zero, m_AddrInitProgress, m_AddrProgress)
        End Function

        Public Function SetResolution(ByVal Value As Integer) As Boolean
            Return CBool(pdfSetResolution(m_Instance, Value))
        End Function

        Public Function SetSaveNewImageFormat(ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetSaveNewImageFormat(m_Instance, CInt(Value)))
        End Function

        Public Function SetSeparationInfo(ByVal Handle As Integer) As Boolean
            Return CBool(pdfSetSeparationInfo(m_Instance, Handle))
        End Function

        Public Function SetStrokeColor(ByVal Color As Integer) As Boolean
            Return CBool(pdfSetStrokeColor(m_Instance, Color))
        End Function

        Public Function SetStrokeColorEx(ByVal Color() As Byte, ByVal NumComponents As Integer) As Boolean
            Return CBool(pdfSetStrokeColorEx(m_Instance, Color, NumComponents))
        End Function

        Public Function SetStrokeColorF(ByVal Color() As Single, ByVal NumComponents As Integer) As Boolean
            Return CBool(pdfSetStrokeColorF(m_Instance, Color, NumComponents))
        End Function

        Public Function SetTabLen(ByVal TabLen As Integer) As Boolean
            Return CBool(pdfSetTabLen(m_Instance, TabLen))
        End Function

        Public Function SetTextDrawMode(ByVal Mode As TDrawMode) As Boolean
            Return CBool(pdfSetTextDrawMode(m_Instance, Mode))
        End Function

        Public Function SetTextFieldValue(ByVal Field As Integer, ByVal Value As String, ByVal DefValue As String, ByVal Align As TTextAlign) As Boolean
            Return CBool(pdfSetTextFieldValueW(m_Instance, Field, Value, DefValue, Align))
        End Function

        Public Function SetTextFieldValueA(ByVal Field As Integer, ByVal Value As String, ByVal DefValue As String, ByVal Align As TTextAlign) As Boolean
            Return CBool(pdfSetTextFieldValueA(m_Instance, Field, Value, DefValue, Align))
        End Function

        Public Function SetTextFieldValueW(ByVal Field As Integer, ByVal Value As String, ByVal DefValue As String, ByVal Align As TTextAlign) As Boolean
            Return CBool(pdfSetTextFieldValueW(m_Instance, Field, Value, DefValue, Align))
        End Function

        Public Function SetTextFieldValueEx(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return pdfSetTextFieldValueExW(m_Instance, Field, Value) <> 0
        End Function

        Public Function SetTextFieldValueExA(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return pdfSetTextFieldValueExA(m_Instance, Field, Value) <> 0
        End Function

        Public Function SetTextFieldValueExW(ByVal Field As Integer, ByVal Value As String) As Boolean
            Return pdfSetTextFieldValueExW(m_Instance, Field, Value) <> 0
        End Function

        Public Function SetTextRect(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Boolean
            Return CBool(pdfSetTextRect(m_Instance, PosX, PosY, Width, Height))
        End Function

        Public Function SetTextRise(ByVal Value As Double) As Boolean
            Return CBool(pdfSetTextRise(m_Instance, Value))
        End Function

        Public Function SetTextScaling(ByVal Value As Double) As Boolean
            Return CBool(pdfSetTextScaling(m_Instance, Value))
        End Function

        Public Function SetTransparentColor(ByVal AColor As Integer) As Boolean
            Return CBool(pdfSetTransparentColor(m_Instance, AColor))
        End Function

        Public Sub SetTrapped(ByVal Value As Boolean)
            pdfSetTrapped(m_Instance, CInt(Value))
        End Sub

        Public Function SetUseExactPwd(ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseExactPwd(m_Instance, CInt(Value)))
        End Function

        Public Function SetUseGlobalImpFiles(ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseGlobalImpFiles(m_Instance, CInt(Value)))
        End Function

        Public Function SetUseImageInterpolation(ByVal Handle As Integer, ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseImageInterpolation(m_Instance, Handle, CInt(Value)))
        End Function

        Public Function SetUseImageInterpolationEx(ByVal Image As IntPtr, ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseImageInterpolationEx(Image, CInt(Value)))
        End Function

        Public Function SetUserUnit(ByVal Value As Single) As Boolean
            Return CBool(pdfSetUserUnit(m_Instance, Value))
        End Function

        Public Function SetUseStdFonts(ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseStdFonts(m_Instance, CInt(Value)))
        End Function

        Public Function SetUseSwapFile(ByVal SwapContents As Boolean, ByVal SwapLimit As Integer) As Boolean
            Return CBool(pdfSetUseSwapFile(m_Instance, CInt(SwapContents), SwapLimit))
        End Function

        Public Function SetUseSwapFileEx(ByVal SwapContents As Boolean, ByVal SwapLimit As Integer, ByVal SwapDir As String) As Boolean
            Return CBool(pdfSetUseSwapFileEx(m_Instance, CInt(SwapContents), SwapLimit, SwapDir))
        End Function

        Public Function SetUseSystemFonts(ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseSystemFonts(m_Instance, CInt(Value)))
        End Function

        Public Function SetUseTransparency(ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseTransparency(m_Instance, CInt(Value)))
        End Function

        Public Function SetUseVisibleCoords(ByVal Value As Boolean) As Boolean
            Return CBool(pdfSetUseVisibleCoords(m_Instance, CInt(Value)))
        End Function

        Public Function SetViewerPreferences(ByVal Value As TViewerPreference, ByVal AddVal As TViewPrefAddVal) As Boolean
            Return CBool(pdfSetViewerPreferences(m_Instance, Value, AddVal))
        End Function

        Public Function SetWMFDefExtent(ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Return CBool(pdfSetWMFDefExtent(m_Instance, Width, Height))
        End Function

        Public Function SetWMFPixelPerInch(ByVal Value As Integer) As Boolean
            Return CBool(pdfSetWMFPixelPerInch(m_Instance, Value))
        End Function

        Public Function SetWordSpacing(ByVal Value As Double) As Boolean
            Return CBool(pdfSetWordSpacing(m_Instance, Value))
        End Function

        Public Function SkewCoords(ByVal alpha As Double, ByVal beta As Double, ByVal OriginX As Double, ByVal OriginY As Double) As Boolean
            Return CBool(pdfSkewCoords(m_Instance, alpha, beta, OriginX, OriginY))
        End Function

        Public Function SortFieldsByIndex() As Boolean
            Return CBool(pdfSortFieldsByIndex(m_Instance))
        End Function

        Public Function SortFieldsByName() As Boolean
            Return CBool(pdfSortFieldsByName(m_Instance))
        End Function

        Public Function SquareAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfSquareAnnotW(m_Instance, PosX, PosY, Width, Height, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function SquareAnnotA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfSquareAnnotA(m_Instance, PosX, PosY, Width, Height, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function SquareAnnotW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfSquareAnnotW(m_Instance, PosX, PosY, Width, Height, LineWidth, FillColor, StrokeColor, CS, Author, Subject, Comment)
        End Function

        Public Function StampAnnot(ByVal SubType As TRubberStamp, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfStampAnnotW(m_Instance, SubType, PosX, PosY, Width, Height, Author, Subject, Comment)
        End Function

        Public Function StampAnnotA(ByVal SubType As TRubberStamp, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfStampAnnotA(m_Instance, SubType, PosX, PosY, Width, Height, Author, Subject, Comment)
        End Function

        Public Function StampAnnotW(ByVal SubType As TRubberStamp, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
            Return pdfStampAnnotW(m_Instance, SubType, PosX, PosY, Width, Height, Author, Subject, Comment)
        End Function

        Public Function StrokePath() As Boolean
            Return CBool(pdfStrokePath(m_Instance))
        End Function

        Public Function TestGlyphs(ByVal FontHandle As Integer, ByVal Text As String) As Integer
            Return pdfTestGlyphsEx(m_Instance, FontHandle, Text, Text.Length)
        End Function

        Public Function TestGlyphs(ByVal FontHandle As Integer, ByVal Text As String, ByVal Length As Integer) As Integer
            If Text.Length < Length Then Length = Text.Length
            Return pdfTestGlyphsEx(m_Instance, FontHandle, Text, Length)
        End Function

        Public Function TextAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Text As String, ByVal Icon As TAnnotIcon, ByVal DoOpen As Boolean) As Integer
            Return pdfTextAnnotW(m_Instance, PosX, PosY, Width, Height, Author, Text, Icon, CInt(DoOpen))
        End Function

        Public Function TextAnnotA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Text As String, ByVal Icon As TAnnotIcon, ByVal DoOpen As Boolean) As Integer
            Return pdfTextAnnotA(m_Instance, PosX, PosY, Width, Height, Author, Text, Icon, CInt(DoOpen))
        End Function

        Public Function TextAnnotW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Text As String, ByVal Icon As TAnnotIcon, ByVal DoOpen As Boolean) As Integer
            Return pdfTextAnnotW(m_Instance, PosX, PosY, Width, Height, Author, Text, Icon, CInt(DoOpen))
        End Function

        Public Function TranslateCoords(ByVal OriginX As Double, ByVal OriginY As Double) As Boolean
            Return CBool(pdfTranslateCoords(m_Instance, OriginX, OriginY))
        End Function

        Public Function TranslateRawCode(ByVal IFont As IntPtr, ByVal Text As IntPtr, ByVal Len As Integer, ByRef Width As Double, ByRef OutText() As Byte, ByRef OutLen As Integer, ByRef Decoded As Integer, ByVal CharSpacing As Single, ByVal WordSpacing As Single, ByVal TextScale As Single) As Integer
            TranslateRawCode = fntTranslateRawCode(IFont, Text, Len, Width, OutText, OutLen, Decoded, CharSpacing, WordSpacing, TextScale)
        End Function

        Public Function TranslateString(ByRef Stack As TPDFStack, ByRef OutText As String, ByVal Flags As Integer) As Integer
            Dim buf(CInt(Stack.TextLen * 16 / 10) * 2 + 64) As Byte
            Dim retval As Integer = fntTranslateString(Stack, buf, CInt(buf.Length / 2), Flags)
            OutText = System.Text.UnicodeEncoding.Unicode.GetString(buf, 0, retval * 2)
            Return retval
        End Function

        Public Function TranslateString2(ByVal IFont As IntPtr, ByVal Text As IntPtr, ByVal Len As Integer, ByRef OutText As String, ByVal Flags As Integer) As Integer
            Dim buf(CInt(Len * 16 / 10) * 2 + 64) As Byte
            Dim retval As Integer = fntTranslateString2(IFont, Text, Len, buf, CInt(buf.Length / 2), Flags)
            OutText = System.Text.UnicodeEncoding.Unicode.GetString(buf, 0, retval * 2)
            Return retval
        End Function

        Public Function Triangle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal FillMode As TPathFillMode) As Boolean
            Return CBool(pdfTriangle(m_Instance, x1, y1, x2, y2, x3, y3, FillMode))
        End Function

        Public Function UnLockLayer(ByVal Layer As Integer) As Boolean
            Return CBool(pdfUnLockLayer(m_Instance, Layer))
        End Function

        Public Function WatermarkAnnot(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
            Return pdfWatermarkAnnot(m_Instance, PosX, PosY, Width, Height)
        End Function

        Public Function WebLink(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal URL As String) As Integer
            Return pdfWebLinkW(m_Instance, PosX, PosY, Width, Height, URL)
        End Function

        Public Function WebLinkA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal URL As String) As Integer
            Return pdfWebLinkA(m_Instance, PosX, PosY, Width, Height, URL)
        End Function

        Public Function WebLinkW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal URL As String) As Integer
            Return pdfWebLinkW(m_Instance, PosX, PosY, Width, Height, URL)
        End Function

        Public Function WeightFromStyle(ByVal Style As TFStyle) As Integer
            Return CInt((Style And Not &HC00FFFFF) / 1048576)
        End Function

        Public Function WeightToStyle(ByVal Weight As Integer) As TFStyle
            Return CType(Weight * 1048576, TFStyle)
        End Function

        Public Function WidthFromStyle(ByVal Style As TFStyle) As Integer
            Return CInt((Style And Not &HFFF2D00F) / 256)
        End Function

        Public Function WidthToStyle(ByVal Width As Integer) As TFStyle
            Return CType(Width * 256, TFStyle)
        End Function

        Public Function WriteAngleText(ByVal AText As String, ByVal Angle As Double, ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal YOrigin As Double) As Boolean
            Return CBool(pdfWriteAngleTextW(m_Instance, AText, Angle, PosX, PosY, Radius, YOrigin))
        End Function

        Public Function WriteAngleTextA(ByVal AText As String, ByVal Angle As Double, ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal YOrigin As Double) As Boolean
            Return CBool(pdfWriteAngleTextA(m_Instance, AText, Angle, PosX, PosY, Radius, YOrigin))
        End Function

        Public Function WriteAngleTextW(ByVal AText As String, ByVal Angle As Double, ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal YOrigin As Double) As Boolean
            Return CBool(pdfWriteAngleTextW(m_Instance, AText, Angle, PosX, PosY, Radius, YOrigin))
        End Function

        Public Function WriteFText(ByVal Align As TTextAlign, ByVal AText As String) As Boolean
            Return CBool(pdfWriteFTextW(m_Instance, Align, AText))
        End Function

        Public Function WriteFTextA(ByVal Align As TTextAlign, ByVal AText As String) As Boolean
            Return CBool(pdfWriteFTextA(m_Instance, Align, AText))
        End Function

        Public Function WriteFTextW(ByVal Align As TTextAlign, ByVal AText As String) As Boolean
            Return CBool(pdfWriteFTextW(m_Instance, Align, AText))
        End Function

        Public Function WriteFTextEx(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Align As TTextAlign, ByVal AText As String) As Boolean
            Return CBool(pdfWriteFTextExW(m_Instance, PosX, PosY, Width, Height, Align, AText))
        End Function

        Public Function WriteFTextExA(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Align As TTextAlign, ByVal AText As String) As Boolean
            Return CBool(pdfWriteFTextExA(m_Instance, PosX, PosY, Width, Height, Align, AText))
        End Function

        Public Function WriteFTextExW(ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Align As TTextAlign, ByVal AText As String) As Boolean
            Return CBool(pdfWriteFTextExW(m_Instance, PosX, PosY, Width, Height, Align, AText))
        End Function

        Public Function WriteText(ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String) As Boolean
            Return CBool(pdfWriteTextW(m_Instance, PosX, PosY, AText))
        End Function

        Public Function WriteTextA(ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String) As Boolean
            Return CBool(pdfWriteTextA(m_Instance, PosX, PosY, AText))
        End Function

        Public Function WriteTextW(ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String) As Boolean
            Return CBool(pdfWriteTextW(m_Instance, PosX, PosY, AText))
        End Function

        Public Function WriteTextEx(ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String, ByVal Len As Integer) As Boolean
            Return CBool(pdfWriteTextExW(m_Instance, PosX, PosY, AText, Len))
        End Function

        Public Function WriteTextExA(ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String, ByVal Len As Integer) As Boolean
            Return CBool(pdfWriteTextExA(m_Instance, PosX, PosY, AText, Len))
        End Function

        Public Function WriteTextExW(ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String, ByVal Len As Integer) As Boolean
            Return CBool(pdfWriteTextExW(m_Instance, PosX, PosY, AText, Len))
        End Function

        Public Function WriteTextMatrix(ByRef Matrix As TCTM, ByVal AText As String) As Boolean
            Return CBool(pdfWriteTextMatrixW(m_Instance, Matrix, AText))
        End Function

        Public Function WriteTextMatrixA(ByRef Matrix As TCTM, ByVal AText As String) As Boolean
            Return CBool(pdfWriteTextMatrixA(m_Instance, Matrix, AText))
        End Function

        Public Function WriteTextMatrixW(ByRef Matrix As TCTM, ByVal AText As String) As Boolean
            Return CBool(pdfWriteTextMatrixW(m_Instance, Matrix, AText))
        End Function

        Public Function WriteTextMatrixEx(ByRef Matrix As TCTM, ByVal AText As String, ByVal Len As Integer) As Boolean
            Return CBool(pdfWriteTextMatrixExW(m_Instance, Matrix, AText, Len))
        End Function

        Public Function WriteTextMatrixExA(ByRef Matrix As TCTM, ByVal AText As String, ByVal Len As Integer) As Boolean
            Return CBool(pdfWriteTextMatrixExA(m_Instance, Matrix, AText, Len))
        End Function

        Public Function WriteTextMatrixExW(ByRef Matrix As TCTM, ByVal AText As String, ByVal Len As Integer) As Boolean
            Return CBool(pdfWriteTextMatrixExW(m_Instance, Matrix, AText, Len))
        End Function

        ' --------------------------------------------- Private Functions ---------------------------------------------

        Private Sub CopyFileSpecEx(ByRef Src As TPDFFileSpecEx_I, ByRef Dst As TPDFFileSpecEx)
            Dst.AFRelationship = ToString(Src.AFRelationship, False)
            Dst.ColItem = Src.ColItem
            Dst.Description = ToString(Src.DescriptionA, Src.DescriptionW)
            Dst.DOS = ToString(Src.DOS, False)
            Dst.EmbFileNode = Src.EmbFileNode
            Dst.FileName = ToString(Src.FileName, False)
            Dst.FileNameIsURL = Src.FileNameIsURL <> 0
            If Src.ID1Len > 0 Then
                ReDim Dst.ID1(Src.ID1Len - 1)
                Marshal.Copy(Src.ID1, Dst.ID1, 0, Src.ID1Len)
            Else
                Erase Dst.ID1
            End If
            If Src.ID2Len > 0 Then
                ReDim Dst.ID2(Src.ID2Len - 1)
                Marshal.Copy(Src.ID2, Dst.ID2, 0, Src.ID2Len)
            Else
                Erase Dst.ID2
            End If
            Dst.IsVolatile = Src.IsVolatile <> 0
            Dst.Mac = ToString(Src.Mac, False)
            Dst.Unix = ToString(Src.Unix, False)
            Dst.RelFileNode = Src.RelFileNode
            Dst.Thumb = Src.Thumb
            Dst.UFileName = ToString(Src.UFileNameA, Src.UFileNameW)
        End Sub

        Private Function CopyFloatArray(ByVal Source As IntPtr, ByVal Count As Integer) As Single()
            If Count < 1 Then Return Nothing
            Dim retval(Count - 1) As Single
            Marshal.Copy(Source, retval, 0, Count)
            Return retval
        End Function

        Private Sub GetIntAnnot(ByRef IAnnot As TPDFAnnotation_I, ByRef Annot As TPDFAnnotation)
            Annot.AnnotType = CType(IAnnot.AnnotType, TAnnotType)
            Annot.Deleted = CBool(IAnnot.Deleted)
            Annot.BBox = IAnnot.BBox
            Annot.BorderWidth = IAnnot.BorderWidth
            Annot.BorderColor = IAnnot.BorderColor
            Annot.BorderStyle = CType(IAnnot.BorderStyle, TBorderStyle)
            Annot.BackColor = IAnnot.BackColor
            Annot.Handle = IAnnot.Handle
            Annot.PageNum = IAnnot.PageNum
            Annot.HighlightMode = CType(IAnnot.HighlightMode, THighlightMode)
            Annot.Author = ToString(IAnnot.AuthorA, IAnnot.AuthorW)
            Annot.Content = ToString(IAnnot.ContentA, IAnnot.ContentW)
            Annot.Name = ToString(IAnnot.NameA, IAnnot.NameW)
            Annot.Subject = ToString(IAnnot.SubjectA, IAnnot.SubjectW)
        End Sub

        Private Sub GetIntAnnotEx(ByRef IAnnot As TPDFAnnotationEx_I, ByRef Annot As TPDFAnnotationEx)
            Annot.AnnotFlags = IAnnot.AnnotFlags
            Annot.AnnotType = CType(IAnnot.AnnotType, TAnnotType)
            Annot.Author = ToString(IAnnot.AuthorA, IAnnot.AuthorW)
            Annot.BackColor = IAnnot.BackColor
            Annot.BBox = IAnnot.BBox
            Annot.BorderColor = IAnnot.BorderColor
            Annot.BorderEffect = IAnnot.BorderEffect
            Annot.BorderStyle = CType(IAnnot.BorderStyle, TBorderStyle)
            Annot.BorderWidth = IAnnot.BorderWidth
            Annot.Content = ToString(IAnnot.ContentA, IAnnot.ContentW)
            Annot.CreateDate = ToString(IAnnot.CreateDate, False)
            Annot.Deleted = CBool(IAnnot.Deleted)
            Annot.DestFile = ToString(IAnnot.DestFile, False)
            Annot.DestPage = IAnnot.DestPage
            Annot.DestPos = IAnnot.DestPos
            Annot.DestType = IAnnot.DestType
            Annot.EmbeddedFile = IAnnot.EmbeddedFile
            Annot.Grouped = IAnnot.Grouped <> 0
            Annot.Handle = IAnnot.Handle
            Annot.HighlightMode = CType(IAnnot.HighlightMode, THighlightMode)
            Annot.Icon = IAnnot.Icon
            Annot.MarkupAnnot = IAnnot.MarkupAnnot
            Annot.ModDate = ToString(IAnnot.ModDate, False)
            Annot.Name = ToString(IAnnot.NameA, IAnnot.NameW)
            Annot.OC = IAnnot.OC
            Annot.Opacity = IAnnot.Opacity
            Annot.Open = IAnnot.Open <> 0
            Annot.PageIndex = IAnnot.PageIndex
            Annot.PageNum = IAnnot.PageNum
            Annot.Parent = IAnnot.Parent
            Annot.PopUp = IAnnot.PopUp
            Annot.RichStyle = ToString(IAnnot.RichStyle, False)
            Annot.RichText = ToString(IAnnot.RichText, False)
            Annot.Rotate = IAnnot.Rotate
            Annot.StampName = ToString(IAnnot.StampName, False)
            Annot.State = ToString(IAnnot.State, False)
            Annot.StateModel = ToString(IAnnot.StateModel, False)
            Annot.Subject = ToString(IAnnot.SubjectA, IAnnot.SubjectW)
            Annot.Subtype = ToString(IAnnot.Subtype, False)
            If IAnnot.InkListCount > 0 Then
                ReDim Annot.InkList(IAnnot.InkListCount - 1)
                Marshal.Copy(IAnnot.InkList, Annot.InkList, 0, IAnnot.InkListCount)
            Else
                Erase Annot.InkList
            End If
            If IAnnot.DashPatternCount > 0 Then
                ReDim Annot.DashPattern(IAnnot.DashPatternCount - 1)
                Marshal.Copy(IAnnot.DashPattern, Annot.DashPattern, 0, IAnnot.DashPatternCount)
            Else
                Annot.DashPattern = Nothing
            End If
            If IAnnot.RD <> IntPtr.Zero Then
                ReDim Annot.RD(3)
                Marshal.Copy(IAnnot.RD, Annot.RD, 0, 4)
            Else
                Annot.RD = Nothing
            End If
            If IAnnot.QuadPointsCount > 0 Then
                ReDim Annot.QuadPoints(IAnnot.QuadPointsCount - 1)
                Marshal.Copy(IAnnot.QuadPoints, Annot.QuadPoints, 0, IAnnot.QuadPointsCount)
            Else
                Annot.QuadPoints = Nothing
            End If
            If IAnnot.VerticesCount > 0 Then
                ReDim Annot.Vertices(IAnnot.VerticesCount - 1)
                Marshal.Copy(IAnnot.Vertices, Annot.Vertices, 0, IAnnot.VerticesCount)
            Else
                Annot.Vertices = Nothing
            End If
            Annot.Intent = ToString(IAnnot.Intent, False)
            Annot.LE1 = IAnnot.LE1
            Annot.LE2 = IAnnot.LE2
            Annot.Caption = IAnnot.Caption <> 0
            Annot.CaptionOffsetX = IAnnot.CaptionOffsetX
            Annot.CaptionOffsetY = IAnnot.CaptionOffsetY
            Annot.CaptionPos = IAnnot.CaptionPos
            Annot.LeaderLineLen = IAnnot.LeaderLineLen
            Annot.LeaderLineExtend = IAnnot.LeaderLineExtend
            Annot.LeaderLineOffset = IAnnot.LeaderLineOffset
        End Sub

        Private Sub GetIntBarcode(ByRef IBC As TPDFBarcode_I, ByRef BC As TPDFBarcode)
            BC.Caption = ToString(IBC.CaptionA, IBC.CaptionW)
            BC.ECC = IBC.ECC
            BC.Height = IBC.Height
            BC.nCodeWordCol = IBC.nCodeWordCol
            BC.nCodeWordRow = IBC.nCodeWordRow
            BC.Resolution = IBC.Resolution
            BC.Symbology = ToString(IBC.Symbology, False)
            BC.Version = IBC.Version
            BC.Width = IBC.Width
            BC.XSymHeight = IBC.XSymHeight
            BC.XSymWidth = IBC.XSymWidth
        End Sub

        Private Function GetIntColorSpaceObj(ByRef ics As TPDFColorSpaceObj_I, ByRef cs As TPDFColorSpaceObj) As Boolean
            Dim i As Integer
            cs.ColorSpaceType = ics.ColorSpaceType
            cs.Alternate = ics.Alternate
            cs.IAlternate = ics.IAlternate
            cs.Buffer = ics.Buffer
            cs.BufSize = ics.BufSize
            cs.NumInComponents = ics.NumInComponents
            cs.NumOutComponents = ics.NumOutComponents
            cs.NumColors = ics.NumColors
            cs.Metadata = ics.Metadata
            cs.MetadataSize = ics.MetadataSize
            cs.IFunction = ics.IFunction
            cs.IAttributes = ics.IAttributes
            cs.IColorSpaceObj = ics.IColorSpaceObj
            If Not IntPtr.Zero.Equals(ics.BlackPoint) Then
                ReDim cs.BlackPoint(2)
                Marshal.Copy(ics.BlackPoint, cs.BlackPoint, 0, 3)
            Else
                cs.BlackPoint = Nothing
            End If
            If Not IntPtr.Zero.Equals(ics.WhitePoint) Then
                ReDim cs.WhitePoint(2)
                Marshal.Copy(ics.WhitePoint, cs.WhitePoint, 0, 3)
            Else
                cs.WhitePoint = Nothing
            End If
            If Not IntPtr.Zero.Equals(ics.Gamma) Then
                ReDim cs.Gamma(cs.NumInComponents - 1)
                Marshal.Copy(ics.Gamma, cs.Gamma, 0, cs.NumInComponents)
            Else
                cs.Gamma = Nothing
            End If
            If Not IntPtr.Zero.Equals(ics.Range) Then
                If ics.ColorSpaceType = TExtColorSpace.esLab Then
                    ReDim cs.Range(3)
                    Marshal.Copy(ics.Range, cs.Range, 0, 4)
                Else
                    ReDim cs.Range(cs.NumInComponents * 2 - 1)
                    Marshal.Copy(ics.Range, cs.Range, 0, cs.NumInComponents * 2)
                End If
            Else
                cs.Range = Nothing
            End If
            If Not IntPtr.Zero.Equals(ics.Matrix) Then
                ReDim cs.Matrix(7)
                Marshal.Copy(ics.Matrix, cs.Matrix, 0, 8)
            Else
                cs.Matrix = Nothing
            End If
            If ics.ColorantsCount = 0 Then
                cs.Colorants = Nothing
            Else
                Dim len As Integer
                Dim bytes() As Byte
                ReDim cs.Colorants(ics.ColorantsCount - 1)
                For i = 0 To ics.ColorantsCount - 1
                    len = pdfStrLenA(ics.Colorants(i))
                    ReDim bytes(len)
                    Marshal.Copy(ics.Colorants(i), bytes, 0, len)
                    bytes = System.Text.Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.Unicode, bytes)
                    cs.Colorants(i) = System.Text.Encoding.Unicode.GetString(bytes)
                Next i
            End If
            Return True
        End Function

        Private Sub GetIntError(ByVal IErr As TPDFError_I, ByRef Err As TPDFError)
            Err.Message = ToString(IErr.Message, False)
            Err.ObjNum = IErr.ObjNum
            Err.Offset = IErr.Offset
            Err.SrcFile = ToString(IErr.SrcFile, False)
            Err.SrcLine = IErr.SrcLine
        End Sub

        Private Sub GetIntField(ByVal IField As TPDFField_I, ByRef Field As TPDFField)
            Dim nameLen As Integer
            Field.BackColor = IField.BackColor
            Field.BBox = IField.BBox
            Field.BorderColor = IField.BorderColor
            Field.Checked = IField.Checked <> 0
            Field.BackCS = CType(IField.BackCS, TPDFColorSpace)
            Field.TextCS = CType(IField.TextCS, TPDFColorSpace)
            Field.Deleted = IField.Deleted <> 0
            Field.Handle = IField.Handle
            nameLen = pdfStrLenA(IField.FieldName)
            Field.FieldName = ToString(IField.FieldName, IField.FieldNameLen, nameLen <> IField.FieldNameLen)
            Field.FieldType = CType(IField.FieldType, TFieldType)
            Field.FontName = ToString(IField.FontName, False)
            Field.FontSize = IField.FontSize
            Field.KidCount = IField.KidCount
            Field.Parent = IField.Parent
            Field.TextColor = IField.TextColor
            Field.ToolTip = ToString(IField.ToolTip, IField.ToolTipLen, IField.UniToolTip <> 0)
            Field.Value = ToString(IField.Value, IField.ValLen, IField.UniVal <> 0)
        End Sub

        Private Sub GetIntFieldEx(ByVal IField As TPDFFieldEx_I, ByRef Field As TPDFFieldEx)
            Field.BackColor = IField.BackColor
            Field.BackColorSP = IField.BackColorSP
            Field.BBox = IField.BBox
            Field.BorderColor = IField.BorderColor
            Field.BorderColorSP = IField.BorderColorSP
            Field.BorderStyle = IField.BorderStyle
            Field.BorderWidth = IField.BorderWidth
            Field.CharSpacing = IField.CharSpacing
            Field.CheckBoxChar = IField.CheckBoxChar
            Field.Checked = CBool(IField.Checked)
            Field.DefState = IField.DefState
            Field.DefValue = ToString(IField.DefValueA, IField.DefValueW)
            Field.Deleted = CBool(IField.Deleted)
            Field.EditFont = ToString(IField.EditFont, False)
            Field.ExpValCount = IField.ExpValCount
            Field.ExpValue = ToString(IField.ExpValueA, IField.ExpValueW)
            Field.FieldFlags = IField.FieldFlags
            Field.FieldFont = ToString(IField.FieldFont, False)
            Field.FieldName = ToString(IField.FieldNameA, IField.FieldNameW)
            Field.FieldType = IField.FieldType
            Field.FontSize = IField.FontSize
            Field.GroupType = IField.GroupType
            Field.Handle = IField.Handle
            Field.HighlightMode = IField.HighlightMode
            Field.IEditFont = IField.IEditFont
            Field.IFieldFont = IField.IFieldFont
            Field.IsCalcField = CBool(IField.IsCalcField)
            Field.Kids = CopyIntPtrArray(IField.Kids, IField.KidCount)
            Field.MapName = ToString(IField.MapNameA, IField.MapNameW)
            Field.MaxLen = IField.MaxLen
            Field.PageNum = IField.PageNum
            Field.OC = IField.OC
            Field.Parent = IField.Parent
            Field.Rotate = IField.Rotate
            Field.TextAlign = IField.TextAlign
            Field.TextColor = IField.TextColor
            Field.TextColorSP = IField.TextColorSP
            Field.TextScaling = IField.TextScaling
            Field.ToolTip = ToString(IField.ToolTipA, IField.ToolTipW)
            Field.UniqueName = ToString(IField.UniqueNameA, IField.UniqueNameW)
            Field.Value = ToString(IField.ValueA, IField.ValueW)
            Field.WordSpacing = IField.WordSpacing
            Field.PageIndex = IField.PageIndex
            Field.IBarcode = IField.IBarcode
            Field.ISignature = IField.ISignature
            Field.ModDate = ToString(IField.ModDate, False)
            Field.CaptionPos = IField.CaptionPos
            Field.DownCaption = ToString(IField.DownCaptionA, IField.DownCaptionW)
            Field.DownImage = IField.DownImage
            Field.RollCaption = ToString(IField.RollCaptionA, IField.RollCaptionW)
            Field.RollImage = IField.RollImage
            Field.UpCaption = ToString(IField.UpCaptionA, IField.UpCaptionW)
            Field.UpImage = IField.UpImage
            Field.Action = IField.Action
            Field.ActionType = IField.ActionType
            Field.Events = IField.Events
        End Sub

        Private Sub GetIntFont(ByRef IFont As TPDFFontObj_I, ByRef F As TPDFFontObj)
            F.Ascent = IFont.Ascent
            F.BaseFont = ToString(IFont.BaseFont, False)
            F.CapHeight = IFont.CapHeight
            F.DefWidth = IFont.DefWidth
            F.Descent = IFont.Descent
            F.FirstChar = IFont.FirstChar
            F.Flags = IFont.Flags
            F.FontFamily = ToString(IFont.FontFamily, IFont.FontFamilyUni <> 0)
            F.FontName = ToString(IFont.FontName, False)
            F.FontType = IFont.FontType
            F.ItalicAngle = IFont.ItalicAngle
            F.LastChar = IFont.LastChar
            F.SpaceWidth = IFont.SpaceWidth
            F.XHeight = IFont.XHeight
            ' We don't copy the buffer here! Copy the buffer manually with pdfCopyMem() if you need it!
            ' However, if IFont.Length1 == 0 and if Not IntPtr.Zero.Equals(IFont.FontFile) then the variable
            ' contains the file path to the used font file. This is always a null-terminated Ansi string at
            ' this time.
            F.FontFile = IFont.FontFile
            F.Length1 = IFont.Length1
            F.Length2 = IFont.Length2
            F.Length3 = IFont.Length3
            F.FontFileType = IFont.FontFileType
            If Not IntPtr.Zero.Equals(IFont.Encoding) Then
                F.Encoding = ToString(IFont.Encoding, 256, True).ToCharArray()
            Else
                F.Encoding = Nothing
            End If
            If IFont.WidthsCount > 0 Then
                ReDim F.Widths(IFont.WidthsCount - 1)
                Marshal.Copy(IFont.Widths, F.Widths, 0, IFont.WidthsCount)
            Else
                F.Widths = Nothing
            End If
        End Sub

        Private Sub GetIntFontInfo(ByRef IFont As TPDFFontInfo_I, ByRef Font As TPDFFontInfo)
            Font.Ascent = IFont.Ascent
            Font.BaseEncoding = IFont.BaseEncoding
            Font.BaseFont = ToString(IFont.BaseFont, False)
            Font.CapHeight = IFont.CapHeight
            Font.CharSet = ToString(IFont.CharSet, False)
            Font.CIDOrdering = ToString(IFont.CIDOrdering, False)
            Font.CIDRegistry = ToString(IFont.CIDRegistry, False)
            If Not IntPtr.Zero.Equals(IFont.CIDSet) Then
                ReDim Font.CIDSet(IFont.CIDSetSize - 1)
                Marshal.Copy(IFont.CIDSet, Font.CIDSet, 0, IFont.CIDSetSize)
            Else
                Font.CIDSet = Nothing
            End If
            Font.CIDSupplement = IFont.CIDSupplement
            If Not IntPtr.Zero.Equals(IFont.CIDToGIDMap) Then
                ReDim Font.CIDToGIDMap(IFont.CIDToGIDMapSize - 1)
                Marshal.Copy(IFont.CIDToGIDMap, Font.CIDToGIDMap, 0, IFont.CIDToGIDMapSize)
            Else
                Font.CIDToGIDMap = Nothing
            End If
            If Not IntPtr.Zero.Equals(IFont.CMapBuf) Then
                ReDim Font.CMapBuf(IFont.CMapBufSize - 1)
                Marshal.Copy(IFont.CMapBuf, Font.CMapBuf, 0, IFont.CMapBufSize)
            Else
                Font.CMapBuf = Nothing
            End If
            Font.CMapName = ToString(IFont.CMapName, False)
            Font.Descent = IFont.Descent
            If Not IntPtr.Zero.Equals(IFont.Encoding) Then
                Font.Encoding = ToString(IFont.Encoding, 256, True).ToCharArray()
            Else
                Font.Encoding = Nothing
            End If
            Font.FirstChar = IFont.FirstChar
            Font.Flags = IFont.Flags
            Font.FontBBox = IFont.FontBBox
            If Not IntPtr.Zero.Equals(IFont.FontBuffer) Then
                ReDim Font.FontBuffer(IFont.FontBufSize - 1)
                Marshal.Copy(IFont.FontBuffer, Font.FontBuffer, 0, IFont.FontBufSize)
            Else
                Font.FontBuffer = Nothing
            End If
            Font.FontFamily = ToString(IFont.FontFamilyA, IFont.FontFamilyW)
            Font.FontFilePath = ToString(IFont.FontFilePathA, IFont.FontFilePathW)
            Font.FontFileType = IFont.FontFileType
            Font.FontName = ToString(IFont.FontName, False)
            Font.FontStretch = ToString(IFont.FontStretch, False)
            Font.FontType = IFont.FontType
            Font.FontWeight = IFont.FontWeight
            Font.FullName = ToString(IFont.FullNameA, IFont.FullNameW)
            Font.HaveEncoding = IFont.HaveEncoding <> 0
            If IFont.HorzWidthsCount > 0 Then
                ReDim Font.HorzWidths(IFont.HorzWidthsCount - 1)
                Marshal.Copy(IFont.HorzWidths, Font.HorzWidths, 0, IFont.HorzWidthsCount)
            Else
                Font.HorzWidths = Nothing
            End If
            Font.Imported = IFont.Imported <> 0
            Font.ItalicAngle = IFont.ItalicAngle
            Font.Lang = ToString(IFont.Lang, False)
            Font.LastChar = IFont.LastChar
            Font.Leading = IFont.Leading
            Font.Length1 = IFont.Length1
            Font.Length2 = IFont.Length2
            Font.Length3 = IFont.Length3
            Font.MaxWidth = IFont.MaxWidth
            If Not IntPtr.Zero.Equals(IFont.Metadata) Then
                ReDim Font.Metadata(IFont.MetadataSize - 1)
                Marshal.Copy(IFont.Metadata, Font.Metadata, 0, IFont.MetadataSize)
            Else
                Font.Metadata = Nothing
            End If
            Font.MisWidth = IFont.MisWidth
            If Not IntPtr.Zero.Equals(IFont.Panose) Then
                ReDim Font.Panose(11)
                Marshal.Copy(IFont.Panose, Font.Panose, 0, 12)
            Else
                Font.Panose = Nothing
            End If
            Font.PostScriptName = ToString(IFont.PostScriptNameA, IFont.PostScriptNameW)
            Font.SpaceWidth = IFont.SpaceWidth
            Font.StemH = IFont.StemH
            Font.StemV = IFont.StemV
            If Not IntPtr.Zero.Equals(IFont.ToUnicode) Then
                ReDim Font.ToUnicode(IFont.ToUnicodeSize - 1)
                Marshal.Copy(IFont.ToUnicode, Font.ToUnicode, 0, IFont.ToUnicodeSize)
            Else
                Font.ToUnicode = Nothing
            End If
            Font.VertDefPos = IFont.VertDefPos
            If IFont.VertWidthsCount > 0 Then
                Dim pt As New CCIDMetric
                ReDim Font.VertWidths(IFont.VertWidthsCount - 1)
                Dim ptr As Long = IFont.VertWidths.ToInt64()
                Dim i As Integer
                For i = 0 To IFont.VertWidthsCount - 1
                    Marshal.PtrToStructure(New IntPtr(ptr), pt)
                    Font.VertWidths(i) = pt
                    ptr += 12
                Next
            Else
                Font.VertWidths = Nothing
            End If
            Font.WMode = IFont.WMode
            Font.XHeight = IFont.XHeight
        End Sub

        Private Sub GetIntOutputIntent(ByRef IIntent As TPDFOutputIntent_I, ByRef Intent As TPDFOutputIntent)
            If IIntent.BufSize > 0 Then
                ReDim Intent.Buffer(IIntent.BufSize - 1)
                Call Marshal.Copy(IIntent.Buffer, Intent.Buffer, 0, IIntent.BufSize)
            Else
                Erase Intent.Buffer
            End If
            Intent.Info = ToString(IIntent.InfoA, IIntent.InfoW)
            Intent.NumComponents = IIntent.NumComponents
            Intent.OutputCondition = ToString(IIntent.OutputConditionA, IIntent.OutputConditionW)
            Intent.OutputConditionID = ToString(IIntent.OutputConditionIDA, IIntent.OutputConditionIDW)
            Intent.RegistryName = ToString(IIntent.RegistryNameA, IIntent.RegistryNameW)
            Intent.SubType = ToString(IIntent.SubType, False)
        End Sub

        Private Sub GetIntSigDict(ByRef ISD As TPDFSigDict_I, ByRef SD As TPDFSigDict)
            If ISD.ByteRangeCount > 0 Then
                ReDim SD.ByteRange(ISD.ByteRangeCount * 2 - 1)
                Marshal.Copy(ISD.ByteRange, SD.ByteRange, 0, ISD.ByteRangeCount * 2)
            Else
                SD.ByteRange = Nothing
            End If
            If ISD.CertLen > 0 Then
                ReDim SD.Cert(ISD.CertLen - 1)
                Marshal.Copy(ISD.Cert, SD.Cert, 0, ISD.CertLen)
            Else
                SD.Cert = Nothing
            End If
            If Not IntPtr.Zero.Equals(ISD.Changes) Then
                ReDim SD.Changes(2)
                Marshal.Copy(ISD.Changes, SD.Changes, 0, 3)
            Else
                SD.Changes = Nothing
            End If
            SD.ContactInfo = ToString(ISD.ContactInfoA, ISD.ContactInfoW)
            If ISD.ContentsSize > 0 Then
                ReDim SD.Contents(ISD.ContentsSize - 1)
                Marshal.Copy(ISD.Contents, SD.Contents, 0, ISD.ContentsSize)
            Else
                SD.Contents = Nothing
            End If
            SD.Filter = ToString(ISD.Filter, False)
            SD.Location = ToString(ISD.LocationA, ISD.LocationW)
            SD.Name = ToString(ISD.NameA, ISD.NameW)
            SD.PropAuthTime = ISD.PropAuthTime
            SD.PropAuthType = ToString(ISD.PropAuthType, False)
            SD.Reason = ToString(ISD.ReasonA, ISD.ReasonW)
            SD.Revision = ISD.Revision
            SD.SignTime = ToString(ISD.SignTime, False)
            SD.SubFilter = ToString(ISD.SubFilter, False)
            SD.Version = ISD.Version
        End Sub

        Private Function IsNum(ByVal c As Integer) As Boolean
            c -= 48
            Return (c < 10 AndAlso c > -1)
        End Function

        Private Function ParseInt(ByRef Str() As Byte, ByVal Pos As Integer, ByRef Value As Integer) As Integer
            Dim isNegative As Boolean = False
            Dim c As Integer, p As Integer = Pos

            Value = 0
            If Pos >= Str.Length Then Return 0
            Select Case Str(Pos)
                Case 45
                    isNegative = True
                    Pos += 1
                Case 43
                    Pos += 1
            End Select
            While Pos < Str.Length
                c = Str(Pos)
                If Not IsNum(c) Then Exit While
                Value = Value * 10 + c - 48
                Pos += 1
            End While
            If isNegative Then Value = -Value
            Return (Pos - p)
        End Function

        Private Overloads Function ToString(ByVal Ptr As IntPtr, ByVal IsUnicode As Boolean) As String
            If IntPtr.Zero.Equals(Ptr) Then Return Nothing
            If Not IsUnicode Then
                Return Marshal.PtrToStringAnsi(Ptr)
            Else
                Return Marshal.PtrToStringUni(Ptr)
            End If
        End Function

        Private Overloads Function ToString(ByVal Ptr As IntPtr, ByVal Len As Integer, ByVal IsUnicode As Boolean) As String
            If IntPtr.Zero.Equals(Ptr) Or Len <= 0 Then Return Nothing
            If Not IsUnicode Then
                Return Marshal.PtrToStringAnsi(Ptr, Len)
            Else
                Return Marshal.PtrToStringUni(Ptr, Len)
            End If
        End Function

        Private Overloads Function ToString(ByVal PtrA As IntPtr, ByVal PtrW As IntPtr) As String
            If Not IntPtr.Zero.Equals(PtrA) Then
                Return Marshal.PtrToStringAnsi(PtrA)
            ElseIf Not IntPtr.Zero.Equals(PtrW) Then
                Return Marshal.PtrToStringUni(PtrW)
            Else
                Return Nothing
            End If
        End Function

        Private Overloads Function ToString(ByVal PtrA As IntPtr, ByVal PtrW As IntPtr, ByVal Len As Integer) As String
            If Not IntPtr.Zero.Equals(PtrA) Then
                Return Marshal.PtrToStringAnsi(PtrA, Len)
            ElseIf Not IntPtr.Zero.Equals(PtrW) Then
                Return Marshal.PtrToStringUni(PtrW, Len)
            Else
                Return Nothing
            End If
        End Function

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TBookmark_I
            Dim Color As Integer
            Dim DestPage As Integer
            Dim DestPos As TPDFRect
            Dim DestType As TDestType
            Dim DoOpen As Integer
            Dim Parent As Integer
            Dim Style As TBmkStyle
            Dim Title As IntPtr
            Dim TitleLen As Integer
            Dim bUnicode As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFAnnotation_I
            Dim AnnotType As Integer
            Dim Deleted As Integer
            Dim BBox As TPDFRect
            Dim BorderWidth As Double
            Dim BorderColor As Integer
            Dim BorderStyle As Integer
            Dim BackColor As Integer
            Dim Handle As Integer
            Dim AuthorA As IntPtr
            Dim AuthorW As IntPtr
            Dim ContentA As IntPtr
            Dim ContentW As IntPtr
            Dim NameA As IntPtr
            Dim NameW As IntPtr
            Dim SubjectA As IntPtr
            Dim SubjectW As IntPtr
            Dim PageNum As Integer
            Dim HighlightMode As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFAnnotationEx_I
            Dim AnnotType As Integer
            Dim Deleted As Integer
            Dim BBox As TPDFRect
            Dim BorderWidth As Single
            Dim BorderColor As Integer
            Dim BorderStyle As Integer
            Dim BackColor As Integer
            Dim Handle As Integer
            Dim AuthorA As IntPtr
            Dim AuthorW As IntPtr
            Dim ContentA As IntPtr
            Dim ContentW As IntPtr
            Dim NameA As IntPtr
            Dim NameW As IntPtr
            Dim SubjectA As IntPtr
            Dim SubjectW As IntPtr
            Dim PageNum As Integer
            Dim HighlightMode As Integer
            ' Page link annotations only
            Dim DestPage As Integer
            Dim DestPos As TPDFRect
            Dim DestType As TDestType
            Dim DestFile As IntPtr         ' File link or web link annotations
            Dim Icon As Integer            ' The Icon type depends on the annotation type. If the annotation type is atText then the Icon
            ' is of type TAnnotIcon. If the annotation type is atFileAttach then it is of type
            ' TFileAttachIcon. If the annotation type is atStamp then the Icon is the stamp type (TRubberStamp).
            ' For any other annotion type this value is not set (-1).
            Dim StampName As IntPtr        ' Set only, if Icon == rsUserDefined
            Dim AnnotFlags As Integer      ' See TAnnotFlags for available flags

            Dim CreateDate As IntPtr       ' Creation Date -> Optional
            Dim ModDate As IntPtr          ' Modification Date -> Optional
            Dim Grouped As Integer         ' (Reply type) Meaningful only if Parent != -1 and Type != atPopUp. If true,
            ' the annotation is part of an annotation group. Properties like Content, CreateDate,
            ' ModDate, BackColor, Subject, and Open must be taken from the parent annotation.
            Dim Open As Integer            ' Meaningful only for annotations which have a corresponding PopUp annotation.
            Dim Parent As Integer          ' Parent annotation handle of a PopUp Annotation or the parent annotation if
            ' this annotation represents a state of a base annotation. In this case,
            ' the annotation type is always atText and only the following members should
            ' be considered:
            '    State      // The current state
            '    StateModel // Marked, Review, and so on
            '    CreateDate // Creation Date
            '    ModDate    // Modification Date
            '    Author     // The user who has set the state
            '    Content    // Not displayed in Adobe's Acrobat...
            '    Subject    // Not displayed in Adobe's Acrobat...
            ' The PopUp annotation of a text annotation which represent an Annotation State
            ' must be ignored.
            Dim PopUp As Integer            ' Handle of the corresponding PopUp annotation if any.
            Dim State As IntPtr             ' The state of the annotation.
            Dim StateModel As IntPtr        ' The state model (Marked, Review, and so on).
            Dim EmbeddedFile As Integer     ' FileAttach annotations only. A handle of an embedded file -> GetEmbeddedFile().
            Dim Subtype As IntPtr           ' Set only, if Type = atUnknownAnnot
            Dim PageIndex As Integer        ' The page index is used to sort form fields. See SortFieldsByIndex().
            Dim MarkupAnnot As Boolean      ' If true, the annotation is a markup annotation. Markup annotations can be flattened
            ' separately, see FlattenAnnots().
            Dim Opacity As Single           ' Opacity = 1.0 = Opaque, Opacity < 1.0 = Transparent, Markup annotations only
            Dim QuadPoints As IntPtr        ' Highlight, Link, and Redact annotations only
            Dim QuadPointsCount As Integer  ' Highlight, Link, and Redact annotations only

            Dim DashPattern As IntPtr       ' Only present if BorderStyle == bsDashed
            Dim DashPatternCount As Integer ' Number of values in the array

            Dim Intent As IntPtr            ' Markup annotations only. The intent allows to distinguish between different uses of an annotation.
            ' For example, line annotations have two intents: LineArrow and LineDimension.
            Dim LE1 As TLineEndStyle        ' Line end style of the start point -> Line and PolyLine annotations only
            Dim LE2 As TLineEndStyle        ' Line end style of the end point -> Line and PolyLine annotations only

            Dim Vertices As IntPtr          ' Line, PolyLine, and Polygon annotations only
            Dim VerticesCount As Integer    ' Number of values in the array. This is the raw number of floating point values.
            ' Since a vertice requires always two coordinate pairs, the number of vertices
            ' or points is VerticeCount divided by 2.

            ' Line annotations only. These properties should only be considered if the member Intent is set to the string LineDimension.
            Dim Caption As Integer            ' If true, the annotation string Content is used as caption.
            Dim CaptionOffsetX As Single      ' Horizontal offset of the caption from its normal position
            Dim CaptionOffsetY As Single      ' Vertical offset of the caption from its normal position
            Dim CaptionPos As TLineCaptionPos ' The position where the caption should be drawn if present
            Dim LeaderLineLen As Single       ' Length of the leader lines (positive or negative)
            Dim LeaderLineExtend As Single    ' Optional leader line extend beyond the leader line (must be a positive value or zero)
            Dim LeaderLineOffset As Single    ' Amount of space between the endpoints of the annotation and the leader lines (must be a positive value or zero)

            Dim BorderEffect As TBorderEffect ' Circle, Square, FreeText, and Polygon annotations.
            Dim RichStyle As IntPtr           ' Optional default style string.      -> FreeText annotations only.
            Dim RichText As IntPtr            ' Optional rich text string (RC key). -> Markup annotations only.
            Dim InkList As IntPtr             ' Ink annotations only. Array of array. The sub arrays can be accessed with GetInkList().
            Dim InkListCount As Integer       ' Number of ink arrays.
            Dim OC As Integer                 ' Handle of an OCG or OCMD or -1 if not set. See help file for further information.

            Dim RD As IntPtr                  ' Caret, Circle, Square, and FreeText annotations.
            Dim Rotate As Integer             ' Caret annotations only. Must be zero or a multiple of 90. This key is not documented in the specs.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFBarcode_I
            Dim StructSize As Integer  ' Must be set to sizeof(TPDFBarcode) before calling CreateBarcodeField()!
            Dim CaptionA As IntPtr     ' Optional, the ansi string takes precedence
            Dim CaptionW As IntPtr     ' Optional
            Dim ECC As Single          ' 0..8 for PDF417, or 0..3 for QRCode
            Dim Height As Single       ' Height in inches
            Dim nCodeWordCol As Single ' Required for PDF417. The number of codewords per barcode coloumn.
            Dim nCodeWordRow As Single ' Required for PDF417. The number of codewords per barcode row.
            Dim Resolution As Integer  ' Required -> Should be 300
            Dim Symbology As IntPtr    ' PDF417, QRCode, or DataMatrix.
            Dim Version As Single      ' Should be 1
            Dim Width As Single        ' Width in inches
            Dim XSymHeight As Single   ' Only needed for PDF417. The vertical distance between two barcode modules,
            ' measured in pixels. The ratio XSymHeight/XSymWidth shall be an integer
            ' value. For PDF417, the acceptable ratio range is from 1 to 4. For QRCode
            ' and DataMatrix, this ratio shall always be 1.
            Dim XSymWidth As Single    ' Required -> The horizontal distance, in pixels, between two barcode modules.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFChoiceValue_I
            Dim StructSize As Integer ' Must be initialzed to sizeof(TPDFChoiceValue_I)
            Dim ExpValueA As IntPtr
            Dim ExpValueW As IntPtr
            Dim ExpValueLen As Integer
            Dim ValueA As IntPtr
            Dim ValueW As IntPtr
            Dim ValueLen As Integer
            Dim Selected As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFColorSpaceObj_I
            Dim ColorSpaceType As TExtColorSpace
            Dim Alternate As TExtColorSpace ' Alternate color space or base space of an Indexed or Pattern color space.
            Dim IAlternate As IntPtr             ' Only set if the color space contains an alternate or base color space -> GetColorSpaceObjEx().
            Dim Buffer As IntPtr                 ' Contains either an ICC profile or the color table of an Indexed color space.
            Dim BufSize As Integer               ' Buffer length in bytes.
            Dim BlackPoint As IntPtr             ' CIE blackpoint. If set, the array contains exactly 3 values.
            Dim WhitePoint As IntPtr             ' CIE whitepoint. If set, the array contains exactly 3 values.
            Dim Gamma As IntPtr                  ' If set, one value per component.
            Dim Range As IntPtr                  ' min/max for each component or for the .a and .b components of a Lab color space.
            Dim Matrix As IntPtr                 ' XYZ matrix. If set, the array contains exactly 9 values.
            Dim NumInComponents As Integer       ' Number of input components.
            Dim NumOutComponents As Integer      ' Number of output components.
            Dim NumColors As Integer             ' HiVal + 1 as specified in the color space. Indexed color space only.
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> Dim Colorants() As IntPtr ' Colorant names (Separation, DeviceN, and NChannel only).
            Dim ColorantsCount As Integer        ' The number of colorants in the array.
            Dim Metadata As IntPtr               ' Optional XMP metadata stream -> ICCBased only.
            Dim MetadataSize As Integer          ' Metadata length in bytes.
            Dim IFunction As IntPtr              ' IntPtr to function object -> Separation, DeviceN, and NChannel only.
            Dim IAttributes As IntPtr            ' Optional attributes of DeviceN or NChannel color spaces -> GetNChannelAttributes().
            Dim IColorSpaceObj As IntPtr         ' Pointer of the corresponding color space object
            Dim Reserved01 As IntPtr
            Dim Reserved02 As IntPtr
            Dim Reserved03 As IntPtr
            Dim Reserved04 As IntPtr
            Dim Reserved05 As IntPtr
            Dim Reserved06 As IntPtr
            Dim Reserved07 As IntPtr
            Dim Reserved08 As IntPtr
            Dim Reserved09 As IntPtr
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TDeviceNAttributes_I
            Dim IProcessColorSpace As IntPtr     ' Pointer to process color space or NULL -> GetColorSpaceEx().
            ' Does a process color space with more than 8 components exist? 6 components
            ' is the maximum so far I know. However, 8 components should be large enough for
            ' the next years...
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> Dim ProcessColorants() As IntPtr
            Dim ProcessColorantsCount As Integer ' Number of process colorants in the array or zero if not set.
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> Dim Separations() As IntPtr ' Optional pointers to Separation color spaces -> GetColorSpaceEx().
            Dim SeparationsCount As Integer      ' Number of color spaces in the array.
            Dim IMixingHints As IntPtr           ' Optional pointer to mixing hints. There is no API function at this time to access mixing hints.
            Dim Reserved01 As IntPtr
            Dim Reserved02 As IntPtr
            Dim Reserved03 As IntPtr
            Dim Reserved04 As IntPtr
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFEmbFileNode_I
            Dim StructSize As Integer ' Must be set to sizeof(TPDFEmbFileNode).
            Dim Name As IntPtr        ' UTF-8 encoded name. This key contains usually a 7 bit ASCII string.
            Dim EF As TPDFFileSpec_I  ' Embedded file
            Dim NextNode As IntPtr    ' Next node if any.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFError_I
            Dim StructSize As Integer ' Must be initialized to sizeof(TPDFError)
            Dim Message As IntPtr     ' The error message
            Dim ObjNum As Integer     ' -1 if not available
            Dim Offset As Integer     ' -1 if not available
            Dim SrcFile As IntPtr     ' Source file
            Dim SrcLine As Integer    ' Source line
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFField_I
            Dim FieldType As Integer
            Dim Deleted As Integer
            Dim BBox As TPDFRect
            Dim Handle As Integer
            Dim FieldName As IntPtr
            Dim FieldNameLen As Integer
            Dim BackCS As Integer
            Dim TextCS As Integer
            Dim BackColor As Integer
            Dim BorderColor As Integer
            Dim TextColor As Integer
            Dim Checked As Integer
            Dim Parent As Integer
            Dim KidCount As Integer
            Dim FontName As IntPtr
            Dim FontSize As Double
            Dim Value As IntPtr
            Dim UniVal As Integer
            Dim ValLen As Integer
            Dim ToolTip As IntPtr
            Dim UniToolTip As Integer
            Dim ToolTipLen As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFFieldEx_I
            Dim StructSize As Integer           ' Must be set to sizeof(TPDFFieldEx_I)
            Dim Deleted As Integer              ' If true, the field was marked as deleted by DeleteField()
            Dim BBox As TPDFRect                ' Bounding box of the field in bottom-up coordinates
            Dim FieldType As TFieldType         ' Field type
            Dim GroupType As TFieldType         ' If GroupType != FieldType the field is a terminal field of a field group
            Dim Handle As Integer               ' Field handle
            Dim BackColor As Integer            ' Background color
            Dim BackColorSP As TExtColorSpace   ' Color space of the background color
            Dim BorderColor As Integer          ' Border color
            Dim BorderColorSP As TExtColorSpace ' Color space of the border color
            Dim BorderStyle As TBorderStyle     ' Border style
            Dim BorderWidth As Single           ' Border width
            Dim CharSpacing As Single           ' Text fields only
            Dim Checked As Integer              ' Check boxes only
            Dim CheckBoxChar As Integer         ' ZapfDingbats character that is used to display the on state
            Dim DefState As TCheckBoxState      ' Check boxes only
            Dim DefValueA As IntPtr             ' Optional default value
            Dim DefValueW As IntPtr             ' Optional default value
            Dim IEditFont As IntPtr             ' Pointer to default editing font
            Dim EditFont As IntPtr              ' Postscript name of the editing font
            Dim ExpValCount As Integer          ' Combo and list boxes only. The values can be accessed with GetFieldExpValueEx()
            Dim ExpValueA As IntPtr             ' Check boxes only
            Dim ExpValueW As IntPtr             ' Check boxes only
            Dim FieldFlags As TFieldFlags       ' Field flags
            Dim IFieldFont As IntPtr            ' Pointer to the font that is used by the field
            Dim FieldFont As IntPtr             ' Postscript name of the font
            Private Reserved As Integer         ' Reserved field to avoid alignment errors
            Dim FontSize As Double              ' Font size. 0.0 means auto font size
            Dim FieldNameA As IntPtr            ' Note that children of a field group or radio button have no name
            Dim FieldNameW As IntPtr            ' Field name length in characters
            Dim HighlightMode As THighlightMode ' Highlight mode
            Dim IsCalcField As Integer          ' If true, the OnCalc event of the field is connected with a JavaScript action
            Dim MapNameA As IntPtr              ' Optional unique mapping name of the field
            Dim MapNameW As IntPtr              ' MapName length in characters
            Dim MaxLen As Integer               ' Text fields only -> zero means not restricted
            Dim Kids As IntPtr                  ' Array of child fields -> GetFieldEx2()
            Dim KidCount As Integer             ' Number of fields in the array
            Dim Parent As IntPtr                ' Pointer to parent field or NULL
            Dim PageNum As Integer              ' Page on which the field is used or -1
            Dim Rotate As Integer               ' Rotation angle in degrees
            Dim TextAlign As TTextAlign         ' Text fields only
            Dim TextColor As Integer            ' Text color
            Dim TextColorSP As TExtColorSpace   ' Color space of the field's text
            Dim TextScaling As Single           ' Text fields only
            Dim ToolTipA As IntPtr              ' Optional tool tip
            Dim ToolTipW As IntPtr              ' Optional tool tip
            Dim UniqueNameA As IntPtr           ' Optional unique name (NM key)
            Dim UniqueNameW As IntPtr           ' Optional unique name (NM key)
            Dim ValueA As IntPtr                ' Field value
            Dim ValueW As IntPtr                ' Field value
            Dim WordSpacing As Single           ' Text fields only
            Dim PageIndex As Integer            ' Array index to change the tab order, see SortFieldsByIndex().
            Dim IBarcode As IntPtr              ' If present, this field is a barcode field. The field type is set to ftText
            ' since barcode fields are extended text fields. -> GetBarcodeDict().
            Dim ISignature As IntPtr            ' Signature fields only. Present only for imported signature fields which
            ' which have a value. That means the file was digitally signed. -> GetSigDict().
            ' Signed signature fields are always marked as deleted!

            Dim ModDate As IntPtr               ' Last modification date (optional)
            ' Push buttons only. The down and roll over states are optional. If not present, then all states use the up state.
            ' The handles of the up, down, and roll over states are template handles! The templates can be opened for editing
            ' with EditTemplate2() and parsed with ParseContent().
            Dim CaptionPos As TBtnCaptionPos    ' Where to position the caption relative to its image
            Dim DownCaptionA As IntPtr          ' Caption of the down state
            Dim DownCaptionW As IntPtr          ' Caption of the down state
            Dim DownImage As Integer            ' Image of the down state
            Dim RollCaptionA As IntPtr          ' Caption of the roll over state
            Dim RollCaptionW As IntPtr          ' Caption of the roll over state
            Dim RollImage As Integer            ' Image of the roll over state
            Dim UpCaptionA As IntPtr            ' Caption of the up state
            Dim UpCaptionW As IntPtr            ' Caption of the up state
            Dim UpImage As Integer              ' Image of the up state -> if > -1, the button is an image button
            Dim OC As Integer                   ' Handle of an OCG or OCMD or -1 if not set. See help file for further information.
            Dim Action As Integer               ' Action handle or -1 if not set. This action is executed when the field is activated.
            Dim ActionType As TActionType       ' Meaningful only, if Action >= 0.
            Dim Events As Integer               ' See GetObjEvent() if set.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFFileSpec_I
            Dim Buffer As IntPtr
            Dim BufSize As Integer
            Dim Compressed As Integer
            Dim ColItem As IntPtr
            Dim Name As IntPtr
            Dim NameUnicode As Integer
            Dim FileName As IntPtr
            Dim IsURL As Integer
            Dim UF As IntPtr
            Dim UFUnicode As Integer
            Dim Desc As IntPtr
            Dim DescUnicode As Integer
            Dim FileSize As Integer
            Dim MIMEType As IntPtr
            Dim CreateDate As IntPtr
            Dim ModDate As IntPtr
            Dim CheckSum As IntPtr
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFFileSpecEx_I
            Dim StructSize As Integer    ' Must be set to sizeof(TPDFFileSpecEx).
            Dim AFRelationship As IntPtr ' PDF 2.0
            Dim ColItem As IntPtr        ' If != NULL the embedded file contains a collection item with user defined data. This entry can
            ' occur in PDF Collections only (PDF 1.7). See "PDF Collections" in the help file for further information.
            Dim DescriptionA As IntPtr   ' Optional description string.
            Dim DescriptionW As IntPtr   ' Optional description string.
            Dim DOS As IntPtr            ' Optional DOS file name.
            Dim EmbFileNode As IntPtr    ' GetEmbeddedFileNode().
            Dim FileName As IntPtr       ' File name as 7 bit ASCII string.
            Dim FileNameIsURL As Integer ' If true, FileName contains a URL.
            Dim ID1 As IntPtr            ' Optional file ID. Meaningful only if FileName points to a PDF file.
            Dim ID1Len As Integer        ' String length in bytes.
            Dim ID2 As IntPtr            ' Optional file ID. Meaningful only if FileName points to a PDF file.
            Dim ID2Len As Integer        ' String length in bytes.
            Dim IsVolatile As Integer    ' If true, the file changes frequently with time.
            Dim Mac As IntPtr            ' Optional Mac file name.
            Dim Unix As IntPtr           ' Optional Unix file name.
            Dim RelFileNode As IntPtr    ' Optional related files array. -> GetRelFileNode().
            Dim Thumb As IntPtr          ' Optional thumb nail image. -> GetImageObjEx().
            Dim UFileNameA As IntPtr     ' PDF 1.7. Same as FileName but Unicode is allowed.
            Dim UFileNameW As IntPtr     ' Either the Ansi or Unicode string is set but never both.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFFontInfo_I
            Dim StructSize As Integer            ' Must be set to sizeof(TPDFFontInfo).
            Dim Ascent As Single                 ' Ascent (optional).
            Dim AvgWidth As Single               ' Average character width (optional).
            Dim BaseEncoding As TBaseEncoding    ' Valid only if HaveEncoding is true.
            Dim BaseFont As IntPtr               ' PostScript Name or Family Name.
            Dim CapHeight As Single              ' Cap height (optional).
            Dim CharSet As IntPtr                ' The charset describes which glyphs are present in the font.
            Dim CharSetSize As Integer           ' Length of the CharSet string in bytes.
            Dim CIDOrdering As IntPtr            ' SystemInfo -> A string that uniquely names the character collection within the specified registry.
            Dim CIDRegistry As IntPtr            ' SystemInfo -> Issuer of the character collection.
            Dim CIDSet As IntPtr                 ' CID fonts only. This is a table of bits indexed by CIDs.
            Dim CIDSetSize As Integer            ' Length of the CIDSet in bytes.
            Dim CIDSupplement As Integer         ' CIDSystemInfo -> The supllement number of the character collection.
            Dim CIDToGIDMap As IntPtr            ' Allowed for embedded TrueType based CID fonts only.
            Dim CIDToGIDMapSize As Integer       ' Length of the stream in bytes.
            Dim CMapBuf As IntPtr                ' Only available if the CMap was embedded.
            Dim CMapBufSize As Integer           ' Buffer size in bytes.
            Dim CMapName As IntPtr               ' CID fonts only (this is the encoding if the CMap is not embedded).
            Dim Descent As Single                ' Descent (optional).
            Dim Encoding As IntPtr               ' Unicode mapping 0..255 -> not available for CID fonts.
            Dim FirstChar As Integer             ' First char (simple fonts only).
            Dim Flags As Integer                 ' The font flags describe various characteristics of the font. See help file for further information.
            Dim FontBBox As TBBox                ' This is the size of the largest glyph in this font. The bounding box is important for text selection.
            Dim FontBuffer As IntPtr             ' The font buffer is present if the font was embedded or if it was loaded from a file buffer.
            Dim FontBufSize As Integer           ' Font file size in bytes.
            Dim FontFamilyA As IntPtr            ' Optional Font Family (Family Name, always available for system fonts).
            Dim FontFamilyW As IntPtr            ' Optional Font Family (either the Ansi or Unicode string is set, but never both).
            Dim FontFilePathA As IntPtr          ' Only available for system fonts.
            Dim FontFilePathW As IntPtr          ' Either the Ansi or Unicode path is set, but never both.
            Dim FontFileType As TFontFileSubtype ' See description in the help file for further information.
            Dim FontName As IntPtr               ' Font name (should be the same as BaseFont).
            Dim FontStretch As IntPtr            ' Optional -> UltraCondensed, ExtraCondensed, Condensed, and so on.
            Dim FontType As TFontType            ' If ftType0 the font is a CID font. The Encoding is not set in this case.
            Dim FontWeight As Single             ' Font weight (optional).
            Dim FullNameA As IntPtr              ' System fonts only.
            Dim FullNameW As IntPtr              ' System fonts only (either the Ansi or Unicode string is set, but never both).
            Dim HaveEncoding As Integer          ' If true, BaseEncoding was set from the font's encoding.
            Dim HorzWidths As IntPtr             ' Horizontal glyph widths -> 0..HorzWidthsCount -1.
            Dim HorzWidthsCount As Integer       ' Number of horizontal widths in the array.
            Dim Imported As Integer              ' If true, the font was imported from an external PDF file.
            Dim ItalicAngle As Single            ' Italic angle
            Dim Lang As IntPtr                   ' Optional language code defined by BCP 47.
            Dim LastChar As Integer              ' Last char (simple fonts only).
            Dim Leading As Single                ' Leading (optional).
            Dim Length1 As Integer               ' Length of the clear text portion of a Type1 font.
            Dim Length2 As Integer               ' Length of the encrypted portion of a Type1 font program (Type1 fonts only).
            Dim Length3 As Integer               ' Length of the fixed-content portion of a Type1 font program or zero if not present.
            Dim MaxWidth As Single               ' Maximum glyph width (optional).
            Dim Metadata As IntPtr               ' Optional XMP stream that contains metadata about the font file.
            Dim MetadataSize As Integer          ' Buffer size in bytes.
            Dim MisWidth As Single               ' Missing width (default = 0.0).
            Dim Panose As IntPtr                 ' CID fonts only -> Optional 12 bytes long Panose string as described in Microsoft’s TrueType 1.0 Font Files Technical Specification.
            Dim PostScriptNameA As IntPtr        ' System fonts only.
            Dim PostScriptNameW As IntPtr        ' System fonts only (either the Ansi or Unicode string is set, but never both).
            Dim SpaceWidth As Single             ' Space width in font units. A default value is set if the font contains no space character.
            Dim StemH As Single                  ' The thickness, measured vertically, of the dominant horizontal stems of glyphs in the font.
            Dim StemV As Single                  ' The thickness, measured horizontally, of the dominant vertical stems of glyphs in the font.
            Dim ToUnicode As IntPtr              ' Only available for imported fonts. This is an embedded CMap that translates PDF strings to Unicode.
            Dim ToUnicodeSize As Integer         ' Buffer size in bytes.
            Dim VertDefPos As TFltPoint          ' Default vertical displacement vector.
            Dim VertWidths As IntPtr             ' Vertical glyph widths -> 0..VertWidthsCount -1.
            Dim VertWidthsCount As Integer       ' Number of vertical widths in the array.
            Dim WMode As Integer                 ' Writing Mode -> 0 == Horizontal, 1 == Vertical.
            Dim XHeight As Single                ' The height of lowercase letters (like the letter x), measured from the baseline, in fonts that have Latin characters.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFFontObj_I
            Dim Ascent As Single         ' Ascent
            Dim BaseFont As IntPtr       ' PostScript Name or Family Name
            Dim CapHeight As Single      ' Cap height
            Dim Descent As Single        ' Descent
            Dim Encoding As IntPtr       ' Unicode mapping 0..255 -> not set if a CID font is selected
            Dim FirstChar As Integer     ' First char
            Dim Flags As Integer         ' Font flags -> font descriptor
            Dim FontFamily As IntPtr     ' Optional Font Family (Family Name)
            Dim FontFamilyUni As Integer ' Is FontFamily in Unicode format?
            Dim FontName As IntPtr       ' Font name -> font descriptor
            Dim FontType As TFontType    ' If ftType0 the font is a CID font. The Encoding is not set in this case.
            Dim ItalicAngle As Single    ' Italic angle
            Dim LastChar As Integer      ' Last char
            Dim SpaceWidth As Single     ' Space width in font units. A default value is set if the font contains no space character.
            Dim Widths As IntPtr         ' Glyph widths -> 0..WidthsCount -1
            Dim WidthsCount As Integer   ' Number of widths in the array
            Dim XHeight As Single        ' x-height
            Dim DefWidth As Single       ' Default character widths -> CID fonts only
            Dim FontFile As IntPtr       ' Font file buffer -> only imported fonts are returned.
            Dim Length1 As Integer       ' Length of the clear text portion of the Type1 font, or the length of the entire font program if FontType != ffType1.
            Dim Length2 As Integer       ' Length of the encrypted portion of the Type1 font program (Type1 fonts only).
            Dim Length3 As Integer       ' Length of the fixed-content portion of the Type1 font program or zero if not present.
            Dim FontFileType As TFontFileSubtype  ' See PDF Reference for further information.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Class TPDFGlyphOutline_I
            Public AdvanceX As Single
            Public AdvanceY As Single
            Public OriginX As Single
            Public OriginY As Single
            Public Lsb As Int16
            Public Tsb As Int16
            Public HaveBBox As Integer
            Public BBox As TFRect
            Public Outline As IntPtr
            Public Size As Integer
        End Class

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFGoToAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFGoToAction).
            Dim DestPage As Integer           ' Destination page (the first page is denoted by 1).
            Dim DestPos As IntPtr             ' Destination position -> Array of 4 floating point values if set.
            Dim DestType As TDestType         ' Destination type.
            ' GoToR (GoTo Remote) actions only:
            Dim DestFile As IntPtr            ' see GetFileSpec().
            Dim DestNameA As IntPtr           ' Optional named destination that shall be loaded when opening the file.
            Dim DestNameW As IntPtr           ' Either the Ansi or Unicode string is set but never both.
            Dim NewWindow As Integer          ' Meaningful only if the destination file points to a PDF file.
            ' -1 = viewer default, 0 = false, 1 = true.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFHideAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFHideAction).
            Dim Fields As IntPtr              ' Array of field pointers -> GetFieldEx2().
            Dim FieldsCount As Integer        ' Number of fields in the array.
            Dim Hide As Integer               ' A flag indicating whether to hide or show the fields in the array.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFImportDataAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFImportDataAction).
            Dim Data As TPDFFileSpecEx_I      ' The data or file to be loaded.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFJavaScriptAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFJavaScriptAction)
            Dim ScriptA As IntPtr             ' The script
            Dim ScriptW As IntPtr             ' Either the Ansi or Unicode string is set but never both
            Dim ScriptLen As Integer          ' Script length in characters
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFLaunchAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFLaunchAction).
            Dim AppName As IntPtr             ' Optional. The name of the application that should be launched.
            Dim DefDir As IntPtr              ' Optional default directory.
            Dim File As IntPtr                ' see GetFileSpec().
            Dim NewWindow As Integer          ' -1 = viewer default, 0 = false, 1 = true.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
            Dim Operation As IntPtr           ' Optional string specifying the operation to perform (open or print).
            Dim Parameter As IntPtr           ' Optional parameter string that shall be passed to the application (AppName).
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFMeasure_I
            Dim StructSize As Integer    ' Must be set to sizeof(TPDFMeasure)
            Dim IsRectilinear As Integer ' If true, the members of the rectilinear measure dictionary are set. The geospatial members otherwise.
            ' --- Rectilinear measure dictionary ---
            Dim Angles As IntPtr         ' Number format array to measure angles -> GetNumberFormatObj()
            Dim AnglesCount As Integer   ' Number of objects in the array.
            Dim Area As IntPtr           ' Number format array to measure areas -> GetNumberFormatObj()
            Dim AreaCount As Integer     ' Number of objects in the array.
            Dim CXY As Single            ' Optional, meaningful only when Y is present.
            Dim Distance As IntPtr       ' Number format array to measure distances -> GetNumberFormatObj()
            Dim DistanceCount As Integer ' Number of objects in the array.
            Dim OriginX As Single        ' Origin of the measurement coordinate system.
            Dim OriginY As Single        ' Origin of the measurement coordinate system.
            Dim RA As IntPtr             ' A text string expressing the scale ratio of the drawing.
            Dim RW As IntPtr             ' A text string expressing the scale ratio of the drawing.
            Dim Slope As IntPtr          ' Number format array for measurement of the slope of a line -> GetNumberFormatObj()
            Dim SlopeCount As Integer    ' Number of objects in the array.
            Dim x As IntPtr              ' Number format array for measurement of change along the x-axis and, if Y is not present, along the y-axis as well.
            Dim XCount As Integer        ' Number of objects in the array.
            Dim y As IntPtr              ' Number format array for measurement of change along the y-axis.
            Dim YCount As Integer        ' Number of objects in the array.

            ' --- Geospatial measure dictionary ---
            Dim Bounds As IntPtr         ' Array of numbers taken pairwise to describe the bounds for which geospatial transforms are valid.
            Dim BoundCount As Integer    ' Number of values in the array. Should be a multiple of two.

            ' The DCS coordinate system is optional.
            Dim DCS_IsSet As Integer     ' If true, the DCS members are set.
            Dim DCS_Projected As Integer ' If true, the DCS values contains a pojected coordinate system.
            Dim DCS_EPSG As Integer      ' Optional, either EPSG or WKT is set.
            Dim DCS_WKT As IntPtr        ' Optional ASCII string

            ' The GCS coordinate system is required and should be present.
            Dim GCS_Projected As Integer ' If true, the GCS values contains a pojected coordinate system.
            Dim GCS_EPSG As Integer      ' Optional, either EPSG or WKT is set.
            Dim GCS_WKT As IntPtr        ' Optional ASCII string

            Dim GPTS As IntPtr           ' Required, an array of numbers that shall be taken pairwise, defining points in geographic space as degrees of latitude and longitude, respectively.
            Dim GPTSCount As Integer     ' Number of values in the array.
            Dim LPTS As IntPtr           ' Optional, an array of numbers that shall be taken pairwise to define points in a 2D unit square.
            Dim LPTSCount As Integer     ' Number of values in the array.

            Dim PDU1 As IntPtr           ' Optional preferred linear display units.
            Dim PDU2 As IntPtr           ' Optional preferred area display units.
            Dim PDU3 As IntPtr           ' Optional preferred angular display units.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFMovieAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFMovieAction).
            Dim Annot As Integer              ' Optional. The movie annotation handle identifying the movie that shall be played.
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)> Dim FWPosition() As Single
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)> Dim FWScale() As Integer
            Dim Mode As IntPtr                ' Mode
            Dim Operation As IntPtr           ' Operation
            Dim Rate As Single                ' Rate
            Dim ShowControls As Integer       ' ShowControls
            Dim Synchronous As Integer        ' Synchronous
            Dim TitleA As IntPtr              ' The title of a movie annotation that shall be played. Either Annot or Title should be set, but not both.
            Dim TitleW As IntPtr              ' Either the Ansi or Unicode string is set at time.
            Dim Volume As Single              ' Volume
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFNamedAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFNamedAction).
            Dim Name As IntPtr                ' Only set if Type == naUserDefined
            Dim NewWindow As Integer          ' -1 = viewer default, 0 = false, 1 = true.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
            Dim Type As TNamedAction          ' Known pre-defined actions.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFNamedDest_I
            Dim StructSize As Integer ' Must be initialized to sizeof(TPDFNamedDest)
            Dim NameA As IntPtr
            Dim NameW As IntPtr
            Dim NameLen As Integer
            Dim DestFileA As IntPtr
            Dim DestFileW As IntPtr
            Dim DestFileLen As Integer
            Dim DestPage As Integer
            Dim DestPos As TPDFRect
            Dim DestType As TDestType
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFNumberFormat_I
            Dim StructSize As Integer
            Dim C As Single
            Dim D As Integer
            Dim f As TMeasureNumFormat
            Dim FD As Integer
            Dim O As TMeasureLblPos
            Dim PSA As IntPtr
            Dim PSW As IntPtr
            Dim RDA As IntPtr
            Dim RDW As IntPtr
            Dim RTA As IntPtr
            Dim RTW As IntPtr
            Dim SSA As IntPtr
            Dim SSW As IntPtr
            Dim UA As IntPtr
            Dim UW As IntPtr
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFOCLayerConfig_I
            Dim StructSize As Integer ' Must be set to sizeof(TOCLayerConfig)
            Dim Intent As TOCGIntent  ' Possible values oiDesign, oiView, or oiAll.
            Dim IsDefault As Integer  ' If true, this is the default configuration.
            Dim NameA As IntPtr       ' Optional configuration name. The default config has usually no name but all others should have one.
            Dim NameW As IntPtr       ' Either the Ansi or Unicode string is set at time but never both.
            Dim NameLen As Integer    ' Length in characters.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Class TPDFOCUINode_I
            Public StructSize As Integer  ' Must be set to sizeof(TOCUINode)
            Public LabelA As IntPtr       ' Optional label.
            Public LabelW As IntPtr       ' Either the Ansi or Unicode string is set at time but never both.
            Public LabelLength As Integer ' Length in characters.
            Public NextChild As IntPtr    ' If set, the next child node that must be loaded.
            Public NewNode As Integer     ' If true, a new child node must be created.
            Public OCG As Integer         ' Optional OCG handle. -1 if not set -> GetOCG().
        End Class

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFOutputIntent_I
            Dim StructSize As Integer
            Dim Buffer As IntPtr
            Dim BufSize As Integer
            Dim InfoA As IntPtr
            Dim InfoW As IntPtr
            Dim NumComponents As Integer
            Dim OutputConditionA As IntPtr
            Dim OutputConditionW As IntPtr
            Dim OutputConditionIDA As IntPtr
            Dim OutputConditionIDW As IntPtr
            Dim RegistryNameA As IntPtr
            Dim RegistryNameW As IntPtr
            Dim SubType As IntPtr
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFPageLabel_I
            Dim StartRange As Integer
            Dim Format As TPageLabelFormat
            Dim FirstPageNum As Integer
            Dim Prefix As IntPtr
            Dim PrefixLen As Integer
            Dim PrefixUni As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFPrintSettings_I
            Dim DuplexMode As TDuplexMode
            Dim NumCopies As Integer               ' -1 means not set. Values larger than 5 are ignored in viewer applications.
            Dim PickTrayByPDFSize As Integer       ' -1 means not set. 0 == false, 1 == true.
            Dim PrintRanges As IntPtr              ' If set, the array contains PrintRangesCount * 2 values. Each pair consists
            ' of the first and last pages in the sub-range. The first page in the PDF file
            ' is denoted by 0.
            Dim PrintRangesCount As Integer        ' Number of ranges available in PrintRanges.
            Dim PrintScaling As TPrintScaling ' dpmNone means not set.
            ' Reserved fields for future extensions
            Dim Reserved0 As Integer
            Dim Reserved1 As Integer
            Dim Reserved2 As Integer
            Dim Reserved3 As Integer
            Dim Reserved4 As Integer
            Dim Reserved5 As Integer
            Dim Reserved6 As Integer
            Dim Reserved7 As Integer
            Dim Reserved8 As Integer
            Dim Reserved9 As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFRelFileNode_I
            Dim StructSize As Integer ' Must be set to sizeof(TPDFRelFileNode).
            Dim NameA As IntPtr       ' Name of this file spcification.
            Dim NameW As IntPtr       ' Either the Ansi or Unicode name is set but never both.
            Dim EF As TPDFFileSpec_I  ' Embedded file.
            Dim NextNode As IntPtr    ' Next node if any.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFResetFormAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TResetFormAction)
            Dim Fields As IntPtr              ' Array of field pointers -> GetFieldEx2().
            Dim FieldsCount As Integer        ' Number of fields in the array.
            Dim Include As Integer            ' If true, the fields in the Fields array must be reset. If false, these fields must be excluded.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFSigDict_I
            Dim StructSize As Integer     ' Must be set to sizeof(TPDFSigDict_I).
            Dim ByteRange As IntPtr       ' ByteRange -> Byte offset followed by the corresponding length.
            ' The byte ranges are required to create the digest. The values
            ' are returned as is. So, you must check whether the offsets and
            ' length values are valid. There are normally at least two ranges.
            ' Overlapping ranges are not allowed! Any error breaks processing
            ' and the signature should be considered as invalid.
            Dim ByteRangeCount As Integer ' The number of Offset / Length pairs. ByteRange contains 2 * ByteRangeCount values!
            Dim Cert As IntPtr            ' X.509 Certificate when SubFilter is adbe.x509.rsa_sha1.
            Dim CertLen As Integer        ' Length in bytes
            Dim Changes As IntPtr         ' If set, an array of three integers that specify changes to the
            ' document that have been made between the previous signature and
            ' this signature in this order: the number of pages altered, the
            ' number of fields altered, and the number of fields filled in.
            Dim ContactInfoA As IntPtr    ' Optional contact info string, e.g. an email address
            Dim ContactInfoW As IntPtr    ' Optional contact info string, e.g. an email address
            Dim Contents As IntPtr        ' The signature. This is either a DER encoded PKCS#1 binary data
            ' object or a DER-encoded PKCS#7 binary data object depending on
            ' the used SubFilter.
            Dim ContentsSize As Integer   ' Length in bytes.
            Dim Filter As IntPtr          ' The name of the security handler, usually Adobe.PPKLite.
            Dim LocationA As IntPtr       ' Optional location of the signer
            Dim LocationW As IntPtr       ' Optional location of the signer
            Dim SignTime As IntPtr        ' Date/Time string
            Dim NameA As IntPtr           ' Optional signers name
            Dim NameW As IntPtr           ' Optional signers name
            Dim PropAuthTime As Integer   ' Optional -> The number of seconds since the signer was last authenticated.
            Dim PropAuthType As IntPtr    ' Optional -> The method that shall be used to authenticate the signer.
            ' Valid values are PIN, Password, and Fingerprint.
            Dim ReasonA As IntPtr         ' Optional reason
            Dim ReasonW As IntPtr         ' Optional reason
            Dim Revision As Integer       ' Optional -> The version of the signature handler that was used to create
            ' the signature.
            Dim SubFilter As IntPtr       ' A name that describes the encoding of the signature value. Should be
            ' adbe.x509.rsa_sha1, adbe.pkcs7.detached, or adbe.pkcs7.sha1.
            Dim Version As Integer        ' The version of the signature dictionary format.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFSigParms_I
            Dim StructSize As Integer     ' Must be set to sizeof(TSignParms)
            Dim PKCS7ObjLen As Integer    ' The maximum length of the signed PKCS#7 object
            Dim HashType As THashType     ' If set to htDetached, the bytes ranges of the PDF file will be returned.
            Dim Range1 As IntPtr          ' Out -> Contains either the hash or the first byte range to create a detached signature
            Dim Range1Len As Integer      ' Out -> Length of the buffer
            Dim Range2 As IntPtr          ' Out -> Set only if HashType == htDetached
            Dim Range2Len As Integer      ' Out -> Length of the buffer
            <MarshalAs(UnmanagedType.LPStr)>
            Dim ContactInfoA As String    ' Optional, e.g. an email address
            <MarshalAs(UnmanagedType.LPWStr)>
            Dim ContactInfoW As String    ' Optional, e.g. an email address
            <MarshalAs(UnmanagedType.LPStr)>
            Dim LocationA As String       ' Optional location of the signer
            <MarshalAs(UnmanagedType.LPWStr)>
            Dim LocationW As String       ' Optional location of the signer
            <MarshalAs(UnmanagedType.LPStr)>
            Dim ReasonA As String         ' Optional reason why the file was signed
            <MarshalAs(UnmanagedType.LPWStr)>
            Dim ReasonW As String         ' Optional reason why the file was signed
            <MarshalAs(UnmanagedType.LPStr)>
            Dim SignerA As String         ' Optional, the issuer of the certificate takes precedence
            <MarshalAs(UnmanagedType.LPWStr)>
            Dim SignerW As String         ' Optional, the issuer of the certificate takes precedence
            Dim Encrypt As Integer        ' If true, the file will be encrypted
            ' These members will be ignored if Encrypt is set to false
            <MarshalAs(UnmanagedType.LPStr)>
            Dim OpenPwd As String         ' Open password
            <MarshalAs(UnmanagedType.LPStr)>
            Dim OwnerPwd As String        ' Owner password to change the security settings
            Dim KeyLen As TKeyLen         ' Key length to be used to encrypt the file
            Dim Restrict As TRestrictions ' What should be restricted?
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFSubmitFormAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TSubmitFormAction)
            Dim CharSet As IntPtr             ' Optional charset in which the form should be submitted.
            Dim Fields As IntPtr              ' Array of field pointers -> GetFieldEx2().
            Dim FieldsCount As Integer        ' Number of fields in the array.
            Dim Flags As TSubmitFlags         ' Various flags, see CreateSubmitAction() for further information.
            Dim URL As IntPtr                 ' The URL of the script at the Web server that will process the submission.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFSysFont_I
            Dim StructSize As Integer           ' Must be set to sizeof(TPDFSysFont)
            Dim BaseType As TFontBaseType       ' Font type
            Dim CIDOrdering As IntPtr           ' OpenType CID fonts only
            Dim CIDRegistry As IntPtr           ' OpenType CID fonts only
            Dim CIDSupplement As Integer        ' OpenType CID fonts only
            Dim DataOffset As Integer           ' Data offset
            Dim FamilyName As IntPtr            ' Family name
            Dim FilePathA As IntPtr             ' Font file path
            Dim FilePathW As IntPtr             ' Font file path
            Dim FileSize As Integer             ' File size in bytes
            Dim Flags As TEnumFontProcFlags     ' Bitmask
            Dim FullName As IntPtr              ' Full name
            Dim Length1 As Integer              ' Length of the clear text portion of a Type1 font
            Dim Length2 As Integer              ' Length of the eexec encrypted binary portion of a Type1 font
            Dim PostScriptNameA As IntPtr       ' Postscript mame
            Dim PostScriptNameW As IntPtr       ' Postscript mame
            Dim Index As Integer                ' Zero based font index if the font is stored in a TrueType collection
            Dim IsFixedPitch As Integer         ' If true, the font is a fixed pitch font. A proprtional font otherwise.
            Dim Style As TFStyle                ' Font style
            Dim UnicodeRange1 As TUnicodeRange1 ' Bitmask
            Dim UnicodeRange2 As TUnicodeRange2 ' Bitmask
            Dim UnicodeRange3 As TUnicodeRange3 ' Bitmask
            Dim UnicodeRange4 As TUnicodeRange4 ' Bitmask
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFURIAction_I
            Dim StructSize As Integer         ' Must be set to sizeof(TPDFURIAction)
            Dim BaseURL As IntPtr             ' Optional, if defined in the Catalog object.
            Dim IsMap As Integer              ' A flag specifying whether to track the mouse position when the URI is resolved: e.g. http:'test.org?50,70.
            Dim URI As IntPtr                 ' Uniform Resource Identifier.
            Dim NextAction As Integer         ' -1 or next action handle to be executed if any
            Dim NextActionType As TActionType ' Only set if NextAction is >= 0.
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Structure TPDFViewport_I
            Dim StructSize As Integer ' Must be set to sizeof(TPDFViewport)
            Dim BBox As TFltRect      ' Bounding box
            Dim Measure As IntPtr     ' Optional -> GetMeasureObj()
            Dim NameA As IntPtr       ' Optional name
            Dim NameW As IntPtr       ' Optional name
            Dim PtData As IntPtr      ' Pointer of a Point Data dictionary. The value can be accessed with GetPtDataObj().
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=0)>
        Private Class TPDFXFAStream_I
            Public StructSize As Integer ' Must be set to sizeof(TPDFXFAStream_I)
            Public Buffer As IntPtr
            Public BufSize As Integer
            Public NameA As IntPtr
            Public NameW As IntPtr
        End Class

        ' This is a special version of pdfCopyMem to enable copying of the TPDFStack structure.
        Private Declare Function pdfCopyStack Lib "dynapdf.dll" Alias "pdfCopyMem" (ByVal Source As IntPtr, ByRef Dest As TPDFStack, ByVal Len As Integer) As Integer

        Private Declare Ansi Function fntBuildFamilyNameAndStyle Lib "dynapdf.dll" (ByVal IFont As IntPtr, ByVal Name As System.Text.StringBuilder, ByRef Style As TFStyle) As Integer
        Private Declare Function fntGetFont Lib "dynapdf.dll" (ByVal IFont As IntPtr, ByRef Font As TPDFFontObj_I) As Integer
        Private Declare Function fntGetFontInfo Lib "dynapdf.dll" (ByVal IFont As IntPtr, ByRef Font As TPDFFontInfo_I) As Integer
        Private Declare Function fntGetSpaceWidth Lib "dynapdf.dll" (ByVal IFont As IntPtr, ByVal FontSize As Double) As Double
        Private Declare Function fntGetTextWidth Lib "dynapdf.dll" (ByVal IFont As IntPtr, ByVal Text As IntPtr, ByVal Len As Integer, ByVal CharSpacing As Single, ByVal WordSpacing As Single, ByVal TextScale As Single) As Double
        Private Declare Function fntTranslateRawCode Lib "dynapdf.dll" (ByVal IFont As IntPtr, ByVal Text As IntPtr, ByVal Len As Integer, ByRef Width As Double, ByVal OutText() As Byte, ByRef OutLen As Integer, ByRef Decoded As Integer, ByVal CharSpacing As Single, ByVal WordSpacing As Single, ByVal TextScale As Single) As Integer
        Private Declare Function fntTranslateString Lib "dynapdf.dll" (ByRef Stack As TPDFStack, ByVal OutText() As Byte, ByVal Size As Integer, ByVal Flags As Integer) As Integer
        Private Declare Function fntTranslateString2 Lib "dynapdf.dll" (ByVal IFont As IntPtr, ByVal Text As IntPtr, ByVal Len As Integer, ByVal OutText() As Byte, ByVal Size As Integer, ByVal Flags As Integer) As Integer

        ' ----------------------------------- Rendering API -----------------------------------

        Private Declare Sub rasAbort Lib "dynapdf.dll" (ByVal RasPtr As IntPtr)
        Private Declare Function rasAttachImageBuffer Lib "dynapdf.dll" (ByVal RasPtr As IntPtr, ByVal Rows As IntPtr, ByVal Buffer As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal ScanlineLen As Integer) As Integer
        Private Declare Sub rasCalcPagePixelSize Lib "dynapdf.dll" (ByVal PagePtr As IntPtr, ByVal DefScale As TPDFPageScale, ByVal ScaleFactor As Single, ByVal FrameWidth As Integer, ByVal FrameHeight As Integer, ByVal Flags As TRasterFlags, ByRef Width As Integer, ByRef Height As Integer)
        Private Declare Function rasCreateRasterizer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Rows As IntPtr, ByVal Buffer As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal ScanlineLen As Integer, ByVal PixFmt As TPDFPixFormat) As IntPtr
        Private Declare Function rasCreateRasterizerEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DC As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal PixFmt As TPDFPixFormat) As IntPtr
        Private Declare Sub rasDeleteRasterizer Lib "dynapdf.dll" (ByRef RasPtr As IntPtr)
        Private Declare Function rasGetWidthHeight Lib "dynapdf.dll" (ByVal PagePtr As IntPtr, ByVal Flags As TRasterFlags, ByRef Width As Single, ByRef Height As Single, ByVal Rotate As Integer, ByRef BBox As IntPtr) As Integer
        Private Declare Sub rasRedraw Lib "dynapdf.dll" (ByVal RasPtr As IntPtr, ByVal DC As IntPtr, ByVal DestX As Integer, ByVal DestY As Integer)
        Private Declare Function rasResizeBitmap Lib "dynapdf.dll" (ByVal RasPtr As IntPtr, ByVal DC As IntPtr, ByVal Width As Integer, ByVal Height As Integer) As Integer

        ' ----------------------------------- Page cache --------------------------------------

        Public Declare Sub rasChangeBackColor Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Value As Integer)
        Public Declare Sub rasCloseFile Lib "dynapdf.dll" (ByVal CachePtr As IntPtr)
        Public Declare Function rasCreatePageCache Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PixFmt As TPDFPixFormat, ByVal HBorder As Integer, ByVal VBorder As Integer, ByVal BackColor As Integer) As Integer
        Public Declare Sub rasDeletePageCache Lib "dynapdf.dll" (ByRef CachePtr As IntPtr)
        Public Declare Function rasExecBookmark Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Handle As Integer, ByRef NewX As Integer, ByRef NewY As Integer, ByRef NewZoom As Single, ByRef NewPageScale As TPDFPageScale, ByVal Action As IntPtr) As TUpdBmkAction
        Public Declare Function rasGetCurrPage Lib "dynapdf.dll" (ByVal IPDF As Integer) As Integer
        Public Declare Function rasGetCurrZoom Lib "dynapdf.dll" (ByVal CachePtr As IntPtr) As Single
        Public Declare Function rasGetDefPageLayout Lib "dynapdf.dll" (ByVal CachePtr As IntPtr) As TPageLayout
        Public Declare Function rasGetPageAt Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal ScrollX As Integer, ByVal ScrollY As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
        Public Declare Function rasGetPageLayout Lib "dynapdf.dll" (ByVal CachePtr As IntPtr) As TPageLayout
        Public Declare Function rasGetPageMatrix Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal CursorX As Integer, ByVal CursorY As Integer, ByRef DestX As Integer, ByRef DestY As Integer, ByRef Width As Integer, ByRef Height As Integer, ByRef Matrix As TCTM) As Integer
        Public Declare Function rasGetPageScale Lib "dynapdf.dll" (ByVal CachePtr As IntPtr) As TPDFPageScale
        Public Declare Function rasGetRotate Lib "dynapdf.dll" (ByVal CachePtr As IntPtr) As Integer
        Public Declare Function rasGetScrollLineDelta Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Vertical As Integer) As Integer
        Public Declare Function rasGetScrollPos Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Vertical As Integer, ByVal PageNum As Integer) As Integer
        Public Declare Function rasGetScrollRange Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Vertical As Integer, ByRef Max As Integer, ByRef SmallChange As Integer, ByRef LargeChange As Integer) As Integer
        Public Declare Function rasInitBaseObjects Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal Flags As TInitCacheFlags) As Integer
        Public Declare Function rasInitColorManagement Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByRef Profiles As TPDFColorProfiles, ByVal Flags As TPDFInitCMFlags) As Integer
        Public Declare Function rasInitialize Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Priority As TPDFThreadPriority) As Integer
        Public Declare Function rasMouseDown Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal X As Integer, ByVal Y As Integer) As TPDFCursor
        Public Declare Function rasMouseMove Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal hWnd As Integer, ByVal LeftBtnDown As Integer, ByRef ScrollX As Integer, ByRef ScrollY As Integer, ByVal X As Integer, ByVal Y As Integer) As TUpdScrollbar
        Public Declare Function rasPaint Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal DC As Integer, ByRef ScrollX As Integer, ByRef ScrollY As Integer) As TUpdScrollbar
        Public Declare Sub rasResetMousePos Lib "dynapdf.dll" (ByVal CachePtr As IntPtr)
        Public Declare Function rasResize Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Width As Integer, ByVal Height As Integer) As Integer
        Public Declare Function rasProcessErrors Lib "dynapdf.dll" (ByVal CachePtr As IntPtr) As Integer
        Public Declare Function rasScroll Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Vertical As Integer, ByVal ScrollCode As Integer, ByRef ScrollX As Integer, ByRef ScrollY As Integer) As TUpdScrollbar
        Public Declare Ansi Function rasSetCMapDirA Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal APath As String, ByVal Flags As TLoadCMapFlags) As Integer
        Public Declare Unicode Function rasSetCMapDirW Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal APath As String, ByVal Flags As TLoadCMapFlags) As Integer
        Public Declare Sub rasSetDefPageLayout Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Value As TPageLayout)
        Public Declare Function rasSetMinLineWidth Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Value As Single) As Integer
        Public Declare Function rasSetOCGState Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Handle As Integer, ByVal Visible As Integer, ByVal SaveState As Integer) As Integer
        Public Declare Sub rasSetOnPaintCallback Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal UserData As IntPtr, ByVal Callback As TOnUpdateWindow)
        Public Declare Sub rasSetPageLayout Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Value As TPageLayout)
        Public Declare Sub rasSetPageScale Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Value As TPDFPageScale)
        Public Declare Sub rasSetRotate Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Value As Integer)
        Public Declare Function rasSetScrollLineDelta Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Vertical As Integer, ByVal Value As Integer) As Integer
        Public Declare Sub rasSetThreadPriority Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal UpdateThread As TPDFThreadPriority, ByVal RenderThread As TPDFThreadPriority)
        Public Declare Function rasZoom Lib "dynapdf.dll" (ByVal CachePtr As IntPtr, ByVal Value As Single, ByRef HorzPos As Integer, ByRef VertPos As Integer) As Integer

        ' -------------------------------------------------------------------------------------

        Private Declare Function pdfAddActionToObj Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As Integer, ByVal AEvent As Integer, ByVal ActHandle As Integer, ByVal ObjHandle As Integer) As Integer
        Private Declare Function pdfAddAnnotToPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal Handle As Integer) As Integer
        Private Declare Function pdfAddArticle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Ansi Function pdfAddBookmarkA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Unicode Function pdfAddBookmarkW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Ansi Function pdfAddBookmarkExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Unicode Function pdfAddBookmarkExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Ansi Function pdfAddBookmarkEx2AA Lib "dynapdf.dll" Alias "pdfAddBookmarkEx2A" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As String, ByVal Unicode As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Ansi Function pdfAddBookmarkEx2AW Lib "dynapdf.dll" Alias "pdfAddBookmarkEx2A" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal NamedDest As String, ByVal Unicode As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Unicode Function pdfAddBookmarkEx2WA Lib "dynapdf.dll" Alias "pdfAddBookmarkEx2W" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal NamedDest As String, ByVal Unicode As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Unicode Function pdfAddBookmarkEx2WW Lib "dynapdf.dll" Alias "pdfAddBookmarkEx2W" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As String, ByVal Unicode As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Auto Function pdfAddButtonImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal BtnHandle As Integer, ByVal State As Integer, ByVal Caption As String, ByVal ImgFile As String) As Integer
        Private Declare Function pdfAddButtonImageEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal BtnHandle As Integer, ByVal State As Integer, ByVal Caption As String, ByVal hBitmap As IntPtr) As Integer
        Private Declare Ansi Function pdfAddContinueTextA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String) As Integer
        Private Declare Unicode Function pdfAddContinueTextW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String) As Integer
        Private Declare Ansi Function pdfAddDeviceNProcessColorants Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DeviceNCS As Integer, ByVal Colorants() As String, ByVal NumColorants As Integer, ByVal ProcessCS As Integer, ByVal Handle As Integer) As Integer
        Private Declare Ansi Function pdfAddDeviceNSeparations Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DeviceNCS As Integer, ByVal Colorants() As String, ByVal SeparationCS() As Integer, ByVal NumColorants As Integer) As Integer
        Private Declare Function pdfAddFieldToFormAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Action As Integer, ByVal Field As Integer, ByVal DoInclude As Integer) As Integer
        Private Declare Function pdfAddFieldToHideAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal HideAct As Integer, ByVal Field As Integer) As Integer
        Private Declare Ansi Function pdfAddFileCommentA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String) As Integer
        Private Declare Unicode Function pdfAddFileCommentW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String) As Integer
        Private Declare Ansi Function pdfAddFontSearchPathA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal APath As String, ByVal Recursive As Integer) As Integer
        Private Declare Unicode Function pdfAddFontSearchPathW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal APath As String, ByVal Recursive As Integer) As Integer
        Private Declare Ansi Function pdfAddImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Filter As Integer, ByVal Flags As Integer, ByRef Image As TPDFImage) As Integer
        Private Declare Function pdfAddInkList Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal InkAnnot As Integer, ByVal Points() As TFltPoint, ByVal NumPoints As Integer) As Integer
        Private Declare Ansi Function pdfAddJavaScriptA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Script As String) As Integer
        Private Declare Unicode Function pdfAddJavaScriptW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, <MarshalAs(UnmanagedType.LPStr)> ByVal Name As String, ByVal Script As String) As Integer
        Private Declare Ansi Function pdfAddLayerToDisplTreeA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Parent As IntPtr, ByVal Layer As Integer, ByVal Title As String) As IntPtr
        Private Declare Unicode Function pdfAddLayerToDisplTreeW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Parent As IntPtr, ByVal Layer As Integer, ByVal Title As String) As IntPtr
        Private Declare Function pdfAddMaskImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal BaseImage As Integer, ByVal Buffer As IntPtr, ByVal BufSize As Integer, ByVal Stride As Integer, ByVal BitsPerPixel As Integer, ByVal Width As Integer, ByVal Height As Integer) As Integer
        Private Declare Function pdfAddMaskImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal BaseImage As Integer, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Stride As Integer, ByVal BitsPerPixel As Integer, ByVal Width As Integer, ByVal Height As Integer) As Integer
        Private Declare Function pdfAddObjectToLayer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OCG As Integer, ByVal ObjType As TOCObject, ByVal Handle As Integer) As Integer
        Private Declare Function pdfAddOCGToAppEvent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Events As TOCAppEvent, ByVal Categories As TOCGUsageCategory) As Integer
        Private Declare Ansi Function pdfAddOutputIntentA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ICCFile As String) As Integer
        Private Declare Unicode Function pdfAddOutputIntentW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ICCFile As String) As Integer
        Private Declare Function pdfAddOutputIntentEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer) As Integer
        Private Declare Ansi Function pdfAddPageLabelA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal StartRange As Integer, ByVal Format As TPageLabelFormat, ByVal Prefix As String, ByVal AddNum As Integer) As Integer
        Private Declare Unicode Function pdfAddPageLabelW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal StartRange As Integer, ByVal Format As TPageLabelFormat, ByVal Prefix As String, ByVal AddNum As Integer) As Integer
        Private Declare Ansi Function pdfAddRasImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal RasPtr As IntPtr, ByVal Filter As TCompressionFilter) As Integer
        Private Declare Ansi Function pdfAddRenderingIntentA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ICCFile As String) As Integer
        Private Declare Unicode Function pdfAddRenderingIntentW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ICCFile As String) As Integer
        Private Declare Function pdfAddRenderingIntentEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer) As Integer
        Private Declare Ansi Function pdfAddValToChoiceFieldA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal ExpValue As String, ByVal Value As String, ByVal Selected As Integer) As Integer
        Private Declare Unicode Function pdfAddValToChoiceFieldW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal ExpValue As String, ByVal Value As String, ByVal Selected As Integer) As Integer
        Private Declare Function pdfAppend Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfApplyAppEvent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AppEvent As TOCAppEvent, ByVal SaveResult As Integer) As Integer
        Private Declare Function pdfApplyPattern Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PattHandle As Integer, ByVal ColorMode As Integer, ByVal Color As Integer) As Integer
        Private Declare Function pdfApplyShading Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ShadHandle As Integer) As Integer
        Private Declare Function pdfAssociateEmbFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DestObject As TAFDestObject, ByVal DestHandle As Integer, ByVal Relationship As TAFRelationship, ByVal EmbFile As Integer) As Integer
        Private Declare Ansi Function pdfAttachFileA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FilePath As String, ByVal Description As String, ByVal Compress As Integer) As Integer
        Private Declare Unicode Function pdfAttachFileW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FilePath As String, ByVal Description As String, ByVal Compress As Integer) As Integer
        Private Declare Ansi Function pdfAttachFileExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal FileName As String, ByVal Description As String, ByVal Compress As Integer) As Integer
        Private Declare Unicode Function pdfAttachFileExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal FileName As String, ByVal Description As String, ByVal Compress As Integer) As Integer
        Private Declare Function pdfAutoTemplate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Templ As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfBeginClipPath Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfBeginContinueText Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double) As Integer
        Private Declare Function pdfBeginLayer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OCG As Integer) As Integer
        Private Declare Ansi Function pdfBeginPageTemplate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal UseAutoTemplates As Integer) As Integer
        Private Declare Function pdfBeginPattern Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PatternType As Integer, ByVal TilingType As Integer, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfBeginTemplate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfBeginTransparencyGroup Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal Isolated As Integer, ByVal Knockout As Integer, ByVal CS As TExtColorSpace, ByVal CSHandle As Integer) As Integer
        Private Declare Function pdfBezier_1_2_3 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Integer
        Private Declare Function pdfBezier_1_3 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x3 As Double, ByVal y3 As Double) As Integer
        Private Declare Function pdfBezier_2_3 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Integer
        Private Declare Function pdfCalcWidthHeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OrgWidth As Double, ByVal OrgHeight As Double, ByVal ScaledWidth As Double, ByVal ScaledHeight As Double) As Double
        Private Declare Ansi Function pdfCaretAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Unicode Function pdfCaretAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Ansi Function pdfChangeAnnotNameA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Name As String) As Integer
        Private Declare Unicode Function pdfChangeAnnotNameW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Name As String) As Integer
        Private Declare Function pdfChangeAnnotPos Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Ansi Function pdfChangeBookmarkA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ABmk As Integer, ByVal ATitle As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Unicode Function pdfChangeBookmarkW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ABmk As Integer, ByVal ATitle As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Function pdfChangeFont Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer) As Integer
        Private Declare Function pdfChangeFontSize Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Size As Double) As Integer
        Private Declare Function pdfChangeFontStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Style As Integer) As Integer
        Private Declare Function pdfChangeFontStyleEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Style As Integer) As Integer
        Private Declare Ansi Function pdfChangeJavaScriptA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByVal NewScript As String) As Integer
        Private Declare Unicode Function pdfChangeJavaScriptW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByVal NewScript As String) As Integer
        Private Declare Ansi Function pdfChangeJavaScriptActionA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByVal NewScript As String) As Integer
        Private Declare Unicode Function pdfChangeJavaScriptActionW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByVal NewScript As String) As Integer
        Private Declare Ansi Function pdfChangeJavaScriptNameA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByVal Name As String) As Integer
        Private Declare Unicode Function pdfChangeJavaScriptNameW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByVal Name As String) As Integer
        Private Declare Ansi Function pdfChangeLinkAnnot Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal URL As String) As Integer
        Private Declare Function pdfChangeSeparationColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal CSHandle As Integer, ByVal NewColor As Integer, ByVal Alternate As TExtColorSpace, ByVal AltHandle As Integer) As Integer
        Private Declare Function pdfCheckFieldNames Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfCheckCollection Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfCheckConformance Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ConfType As Integer, ByVal Options As Integer, ByVal UserData As IntPtr, ByVal OnFontNotFound As TOnFontNotFoundProc, ByVal OnReplaceICCProfile As TOnReplaceICCProfile) As Integer
        Private Declare Ansi Function pdfCircleAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Unicode Function pdfCircleAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Function pdfClearAutoTemplates Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Sub pdfClearErrorLog Lib "dynapdf.dll" (ByVal IPDF As IntPtr)
        Private Declare Function pdfClearHostFonts Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfClipPath Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ClipMode As Integer, ByVal FillMode As Integer) As Integer
        Private Declare Ansi Function pdfCloseAndSignFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal CertFile As String, ByVal Password As String, ByVal Reason As String, ByVal Location As String) As Integer
        Private Declare Ansi Function pdfCloseAndSignFileEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As Integer, ByVal Restrict As Integer, ByVal CertFile As String, ByVal Password As String, ByVal Reason As String, ByVal Location As String) As Integer
        Private Declare Ansi Function pdfCloseAndSignFileExt Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef SigParms As TPDFSigParms_I) As Integer
        Private Declare Function pdfCloseFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfCloseFileEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As Integer, ByVal Restrict As Integer) As Integer
        Private Declare Ansi Function pdfCloseImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfCloseImportFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfCloseImportFileEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Ansi Function pdfClosePath Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FillMode As Integer) As Integer
        Private Declare Ansi Function pdfCloseTag Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfComputeBBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef BBox As TPDFRect, ByVal Flags As TCompBBoxFlags) As Integer
        Private Declare Ansi Function pdfConvColor Lib "dynapdf.dll" (ByVal Color As IntPtr, ByVal NumComps As Integer, ByVal SourceCS As Integer, ByVal IColorSpace As IntPtr, ByVal DestCS As Integer) As Integer
        Private Declare Ansi Function pdfConvertColors Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer, ByVal Add() As Single) As Integer
        Private Declare Auto Function pdfConvertEMFSpool Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SpoolFile As String, ByVal LeftMargin As Double, ByVal TopMargin As Double, ByVal Flags As TSpoolConvFlags) As Integer
        Private Declare Ansi Function pdfConvToUnicode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AString As String, ByVal CP As Integer) As IntPtr
        Private Declare Ansi Function pdfCopyChoiceValues Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Source As Integer, ByVal Dest As Integer, ByVal Share As Integer) As Integer
        Private Declare Ansi Function pdfCreate3DAnnot Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Name As String, ByVal U3DFile As String, ByVal Image As String) As Integer
        Private Declare Ansi Function pdfCreate3DBackground Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal IView As IntPtr, ByVal BackColor As Integer) As Integer
        Private Declare Ansi Function pdfCreate3DGotoViewAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Base3DAnnot As Integer, ByVal IView As IntPtr, ByVal Named As Integer) As Integer
        Private Declare Ansi Function pdfCreate3DProjection Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal IView As IntPtr, ByVal ProjType As Integer, ByVal ScaleType As Integer, ByVal Diameter As Double, ByVal FOV As Double) As Integer
        Private Declare Ansi Function pdfCreate3DView Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Base3DAnnot As Integer, ByVal Name As String, ByVal SetAsDefault As Integer, ByVal Matrix() As Double, ByVal CamDistance As Double, ByVal RM As Integer, ByVal LS As Integer) As IntPtr
        Private Declare Ansi Function pdfCreateAnnotAP Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Annot As Integer) As Integer
        Private Declare Ansi Function pdfCreateArticleThreadA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ThreadName As String) As Integer
        Private Declare Unicode Function pdfCreateArticleThreadW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ThreadName As String) As Integer
        Private Declare Function pdfCreateAxialShading Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal sx As Double, ByVal sy As Double, ByVal eX As Double, ByVal eY As Double, ByVal SCenter As Double, ByVal SColor As Integer, ByVal EColor As Integer, ByVal Extend1 As Integer, ByVal Extend2 As Integer) As Integer
        Private Declare Function pdfCreateBarcodeField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByRef Barcode As TPDFBarcode_I) As Integer
        Private Declare Ansi Function pdfCreateButtonA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Caption As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Unicode Function pdfCreateButtonW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, <MarshalAs(UnmanagedType.LPStr)> ByVal Name As String, ByVal Caption As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Ansi Function pdfCreateCheckBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal ExpValue As String, ByVal Checked As Integer, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfCreateCIEColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Base As Integer, ByRef WhitePoint() As Single, ByRef BlackPoint() As Single, ByRef Gamma() As Single, ByRef Matrix() As Single) As Integer
        Private Declare Function pdfCreateCollection Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal View As Integer) As Integer
        Private Declare Ansi Function pdfCreateCollectionFieldA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ColType As Integer, ByVal Column As Integer, ByVal Name As String, ByVal Key As String, ByVal Visible As Integer, ByVal Editable As Integer) As Integer
        Private Declare Unicode Function pdfCreateCollectionFieldW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ColType As Integer, ByVal Column As Integer, ByVal Name As String, <MarshalAs(UnmanagedType.LPStr)> ByVal Key As String, ByVal Visible As Integer, ByVal Editable As Integer) As Integer
        Private Declare Function pdfCreateColItemDate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal EmbFile As Integer, ByVal Key As String, ByVal DateVal As Integer, ByVal Prefix As String) As Integer
        Private Declare Function pdfCreateColItemNumber Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal EmbFile As Integer, ByVal Key As String, ByVal Value As Double, ByVal Prefix As String) As Integer
        Private Declare Ansi Function pdfCreateColItemStringA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal EmbFile As Integer, ByVal Key As String, ByVal Value As String, ByVal Prefix As String) As Integer
        Private Declare Unicode Function pdfCreateColItemStringW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal EmbFile As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal Key As String, ByVal Value As String, ByVal Prefix As String) As Integer
        Private Declare Ansi Function pdfCreateComboBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Sort As Integer, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Ansi Function pdfCreateDeviceNColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Colorants() As String, ByVal NumColorants As Integer, ByVal PostScriptFunc As String, ByVal Alternate As Integer, ByVal Handle As Integer) As Integer
        Private Declare Function pdfCreateExtGState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef GS As TPDFExtGState) As Integer
        Private Declare Function pdfCreateGoToAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DestType As Integer, ByVal PageNum As Integer, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Integer
        Private Declare Function pdfCreateGoToActionEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal NamedDest As Integer) As Integer
        Private Declare Ansi Function pdfCreateGoToEActionA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Location As TEmbFileLocation, ByVal Source As String, ByVal SrcPage As Integer, ByVal Target As String, ByVal DestName As String, ByVal DestPage As Integer, ByVal NewWindow As Integer) As Integer
        Private Declare Unicode Function pdfCreateGoToEActionW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Location As TEmbFileLocation, ByVal Source As String, ByVal SrcPage As Integer, ByVal Target As String, ByVal DestName As String, ByVal DestPage As Integer, ByVal NewWindow As Integer) As Integer
        Private Declare Ansi Function pdfCreateGoToRActionA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PageNum As Integer) As Integer
        Private Declare Unicode Function pdfCreateGoToRActionW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PageNum As Integer) As Integer
        Private Declare Ansi Function pdfCreateGoToRActionExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Integer) As Integer
        Private Declare Unicode Function pdfCreateGoToRActionExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, <MarshalAs(UnmanagedType.LPStr)> ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Integer) As Integer
        Private Declare Ansi Function pdfCreateGoToRActionExUA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Integer) As Integer
        Private Declare Unicode Function pdfCreateGoToRActionExUW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal DestName As String, ByVal NewWindow As Integer) As Integer
        Private Declare Ansi Function pdfCreateGroupField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Parent As Integer) As Integer
        Private Declare Function pdfCreateHideAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Hide As Integer) As Integer
        Private Declare Ansi Function pdfCreateICCBasedColorSpaceA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ICCProfile As String) As Integer
        Private Declare Unicode Function pdfCreateICCBasedColorSpaceW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ICCProfile As String) As Integer
        Private Declare Ansi Function pdfCreateICCBasedColorSpaceEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer) As Integer
        Private Declare Ansi Function pdfCreateImageA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal Format As Integer) As Integer
        Private Declare Unicode Function pdfCreateImageW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal Format As Integer) As Integer
        Private Declare Ansi Function pdfCreateImpDataActionA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DataFile As String) As Integer
        Private Declare Unicode Function pdfCreateImpDataActionW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DataFile As String) As Integer
        Private Declare Function pdfCreateIndexedColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Base As Integer, ByVal Handle As Integer, ByRef ColorTable() As Byte, ByVal NumColors As Integer) As Integer
        Private Declare Ansi Function pdfCreateJSActionA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Script As String) As Integer
        Private Declare Unicode Function pdfCreateJSActionW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Script As String) As Integer
        Private Declare Ansi Function pdfCreateLaunchAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OP As Integer, ByVal FileName As String, ByVal DefDir As String, ByVal Param As String, ByVal NewWindow As Integer) As Integer
        Private Declare Ansi Function pdfCreateLaunchActionExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal NewWindow As Integer) As Integer
        Private Declare Unicode Function pdfCreateLaunchActionExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal NewWindow As Integer) As Integer
        Private Declare Ansi Function pdfCreateListBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Sort As Integer, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfCreateNamedAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Action As Integer) As Integer
        Private Declare Ansi Function pdfCreateNamedDestA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal DestPage As Integer, ByVal DestType As TDestType, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Integer
        Private Declare Unicode Function pdfCreateNamedDestW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal DestPage As Integer, ByVal DestType As TDestType, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Integer
        Private Declare Ansi Function pdfCreateNewPDFA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutPDF As String) As Integer
        Private Declare Unicode Function pdfCreateNewPDFW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutPDF As String) As Integer
        Private Declare Ansi Function pdfCreateOCGA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal DisplayInUI As Integer, ByVal Visible As Integer, ByVal Intent As TOCGIntent) As Integer
        Private Declare Unicode Function pdfCreateOCGW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal DisplayInUI As Integer, ByVal Visible As Integer, ByVal Intent As TOCGIntent) As Integer
        Private Declare Function pdfCreateOCMD Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Visibility As TOCVisibility, ByVal OCGs() As Integer, ByVal Count As Integer) As Integer
        Private Declare Function pdfCreateRadialShading Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal sx As Double, ByVal sy As Double, ByVal R1 As Double, ByVal eX As Double, ByVal eY As Double, ByVal R2 As Double, ByVal SCenter As Double, ByVal SColor As Integer, ByVal EColor As Integer, ByVal Extend1 As Integer, ByVal Extend2 As Integer) As Integer
        Private Declare Ansi Function pdfCreateRadioButton Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal ExpValue As String, ByVal Checked As Integer, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfCreateResetAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfCreateSeparationCS Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Colorant As String, ByVal Alternate As Integer, ByVal Handle As Integer, ByVal Color As Integer) As Integer
        Private Declare Function pdfCreateSetOCGStateAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal _On() As Integer, ByVal OnCount As Integer, ByVal Off() As Integer, ByVal OffCount As Integer, ByVal Toggle() As Integer, ByVal ToggleCount As Integer, ByVal PreserveRB As Integer) As Integer
        Private Declare Ansi Function pdfCreateSigField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Parent As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfCreateSigFieldAP Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SigField As Integer) As Integer
        Private Declare Function pdfCreateSoftMask Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TranspGroup As Integer, ByVal MaskType As TSoftMaskType, ByVal BackColor As Integer) As Integer
        Private Declare Function pdfCreateStdPattern Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Pattern As Integer, ByVal LineWidth As Double, ByVal Distance As Double, ByVal LineColor As Integer, ByVal BackColor As Integer) As Integer
        Private Declare Function pdfCreateStructureTree Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfCreateSubmitAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer, ByVal URL As String) As Integer
        Private Declare Ansi Function pdfCreateTextField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Parent As Integer, ByVal Multiline As Integer, ByVal MaxLen As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfCreateURIAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal URL As String) As Integer
        Private Declare Ansi Function pdfDecryptPDFA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PwdType As Integer, ByVal Password As String) As Integer
        Private Declare Unicode Function pdfDecryptPDFW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PwdType As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal Password As String) As Integer
        Private Declare Sub pdfDeleteAcroForm Lib "dynapdf.dll" (ByVal IPDF As IntPtr)
        Private Declare Function pdfDeleteActionFromObj Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As Integer, ByVal ActHandle As Integer, ByVal ObjHandle As Integer) As Integer
        Private Declare Function pdfDeleteActionFromObjEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As Integer, ByVal ObjHandle As Integer, ByVal ActIndex As Integer) As Integer
        Private Declare Function pdfDeleteAnnotation Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfDeleteAnnotationFromPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal Handle As Integer) As Integer
        Private Declare Function pdfDeleteAppEvents Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ApplyEvent As Integer, ByVal AppEvent As TOCAppEvent) As Integer
        Private Declare Function pdfDeleteBookmark Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ABmk As Integer) As Integer
        Private Declare Function pdfDeleteEmbeddedFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfDeleteField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Ansi Function pdfDeleteFieldEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String) As Integer
        Private Declare Sub pdfDeleteJavaScripts Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DelJavaScriptActions As Integer)
        Private Declare Function pdfDeleteOCGFromAppEvent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Events As TOCAppEvent, ByVal Categories As TOCGUsageCategory, ByVal DelCategoryOnly As Integer) As Integer
        Private Declare Function pdfDeleteOutputIntent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer) As Integer
        Private Declare Function pdfDeletePage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer) As Integer
        Private Declare Function pdfDeletePDF Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Sub pdfDeletePageLabels Lib "dynapdf.dll" (ByVal IPDF As IntPtr)
        Private Declare Function pdfDeleteSeparationInfo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AllPages As Integer) As Integer
        Private Declare Function pdfDeleteTemplate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfDeleteTemplateEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer) As Integer
        Private Declare Sub pdfDeleteXFAForm Lib "dynapdf.dll" (ByVal IPDF As IntPtr)
        Private Declare Function pdfDrawArc Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal StartAngle As Double, ByVal EndAngle As Double) As Integer
        Private Declare Function pdfDrawArcEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal StartAngle As Double, ByVal EndAngle As Double) As Integer
        Private Declare Function pdfDrawChord Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal StartAngle As Double, ByVal EndAngle As Double, ByVal FillMode As Integer) As Integer
        Private Declare Function pdfDrawCircle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal FillMode As Integer) As Integer
        Private Declare Function pdfDrawPie Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal StartAngle As Double, ByVal EndAngle As Double, ByVal FillMode As Integer) As Integer
        Private Declare Function pdfEditPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer) As Integer
        Private Declare Function pdfEditTemplate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer) As Integer
        Private Declare Function pdfEditTemplate2 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfEllipse Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal FillMode As Integer) As Integer
        Private Declare Ansi Function pdfEncryptPDFA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal OpenPwd As String, ByVal OwnerPwd As String, ByVal KeyLen As Integer, ByVal Restrict As Integer) As Integer
        Private Declare Unicode Function pdfEncryptPDFW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal OpenPwd As String, <MarshalAs(UnmanagedType.LPStr)> ByVal OwnerPwd As String, ByVal KeyLen As Integer, ByVal Restrict As Integer) As Integer
        Private Declare Function pdfEndContinueText Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfEndLayer Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfEndPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfEndPattern Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfEndTemplate Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfEnumDocFonts Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Data As IntPtr, ByVal EnumProc As TEnumFontsProc2) As Integer
        Private Declare Function pdfEnumHostFonts Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Data As IntPtr, ByVal EnumProc As TEnumFontsProc) As Integer
        Private Declare Function pdfEnumHostFontsEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Data As IntPtr, ByVal EnumProc As TEnumFontsProcEx) As Integer
        Private Declare Function pdfExchangeBookmarks Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Bmk1 As Integer, ByVal Bmk2 As Integer) As Integer
        Private Declare Function pdfExchangePages Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal First As Integer, ByVal Second As Integer) As Integer
        Private Declare Function pdfExtractText Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal Flags As TTextExtractionFlags, ByRef Area As TFltRect, ByRef Text As IntPtr, ByRef TextLen As Integer) As Integer
        Private Declare Ansi Function pdfFileAttachAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Icon As Integer, ByVal Author As String, ByVal Desc As String, ByVal AFile As String, ByVal Compress As Integer) As Integer
        Private Declare Unicode Function pdfFileAttachAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Icon As Integer, ByVal Author As String, ByVal Desc As String, ByVal AFile As String, ByVal Compress As Integer) As Integer
        Private Declare Auto Function pdfFileAttachAnnotEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Icon As TFileAttachIcon, ByVal FileName As String, ByVal Author As String, ByVal Desc As String, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Compress As Integer) As Integer
        Private Declare Ansi Function pdfFileLinkA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal AFilePath As String) As Integer
        Private Declare Unicode Function pdfFileLinkW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal AFilePath As String) As Integer
        Private Declare Ansi Function pdfFindBookmarkA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DestPage As Integer, ByVal Title As String) As Integer
        Private Declare Unicode Function pdfFindBookmarkW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DestPage As Integer, ByVal Title As String) As Integer
        Private Declare Auto Function pdfFindEmbeddedFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String) As Integer
        Private Declare Ansi Function pdfFindFieldA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String) As Integer
        Private Declare Unicode Function pdfFindFieldW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String) As Integer
        Private Declare Ansi Function pdfFindLinkAnnot Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal URL As String) As Integer
        Private Declare Function pdfFindNextBookmark Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfFinishSignature Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PKCS7Obj() As Byte, ByVal Length As Integer) As Integer
        Private Declare Function pdfFlattenAnnots Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As TAnnotFlattenFlags) As Integer
        Private Declare Function pdfFlattenForm Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfFlushPageContent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Stack As TPDFStack) As Integer
        Private Declare Function pdfFlushPages Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As TFlushPageFlags) As Integer
        Private Declare Sub pdfFreeImageBuffer Lib "dynapdf.dll" (ByVal IPDF As IntPtr)
        Private Declare Function pdfFreeImageObj Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfFreeImageObjEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ImagePtr As IntPtr) As Integer
        Private Declare Function pdfFreePDF Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfFreeTextAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal AText As String, ByVal Align As Integer) As Integer
        Private Declare Unicode Function pdfFreeTextAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal AText As String, ByVal Align As Integer) As Integer
        Private Declare Function pdfFreeUniBuf Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfGet3DAnnotStream Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Annot As Integer, ByRef Data As IntPtr, ByRef Size As Integer, ByRef Subtype As String) As Integer
        Private Declare Function pdfGetActionCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetActionHandle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As Integer, ByVal ObjHandle As Integer, ByVal ActIndex As Integer) As Integer
        Private Declare Function pdfGetActionType Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ActHandle As Integer) As Integer
        Private Declare Function pdfGetActionTypeEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As Integer, ByVal ObjHandle As Integer, ByVal ActIndex As Integer) As Integer
        Private Declare Function pdfGetActiveFont Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetAllocBy Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetAnnot Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Annot As TPDFAnnotation_I) As Integer
        Private Declare Function pdfGetAnnotBBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef BBox As TPDFRect) As Integer
        Private Declare Function pdfGetAnnotCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetAnnotEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Annot As TPDFAnnotationEx_I) As Integer
        Private Declare Function pdfGetAnnotFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetAnnotLink Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As IntPtr
        Private Declare Function pdfGetAnnotType Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfGetAscent Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetBarcodeDict Lib "dynapdf.dll" (ByVal IBarcode As IntPtr, ByRef Barcode As TPDFBarcode_I) As Integer
        Private Declare Function pdfGetBBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Boundary As Integer, ByRef BBox As TPDFRect) As Integer
        Private Declare Function pdfGetBidiMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetBookmark Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByRef Bmk As TBookmark_I) As Integer
        Private Declare Function pdfGetBookmarkCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetBorderStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetBuffer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef BufSize As Integer) As IntPtr
        Private Declare Function pdfGetCapHeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetCharacterSpacing Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetCheckBoxChar Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetCheckBoxCharEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetCheckBoxDefState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetCMap Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef CMap As TPDFCMap) As Integer
        Private Declare Function pdfGetCMapCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetColorSpaceCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetColorSpaceObj Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef cs As TPDFColorSpaceObj_I) As Integer
        Private Declare Function pdfGetColorSpaceObjEx Lib "dynapdf.dll" (ByVal IColorSpace As IntPtr, ByRef cs As TPDFColorSpaceObj_I) As Integer
        Private Declare Function pdfGetCompressionFilter Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetCompressionLevel Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetContent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Buffer As IntPtr) As Integer
        Private Declare Function pdfGetDefBitsPerPixel Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetDescent Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetDeviceNAttributes Lib "dynapdf.dll" (ByVal IAttributes As IntPtr, ByRef Attributes As TDeviceNAttributes_I) As Integer
        Private Declare Function pdfGetDocInfo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DInfo As Integer, ByRef Value As IntPtr) As Integer
        Private Declare Function pdfGetDocInfoCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetDocInfoEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef DInfo As Integer, ByRef Key As IntPtr, ByRef Value As IntPtr, ByRef IsUnicode As Integer) As Integer
        Private Declare Function pdfGetDocUsesTransparency Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer) As Integer
        Private Declare Function pdfGetDrawDirection Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetDynaPDFVersion Lib "dynapdf.dll" () As IntPtr
        Private Declare Function pdfGetEmbeddedFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef FileSpec As TPDFFileSpec_I, ByVal Decompress As Integer) As Integer
        Private Declare Function pdfGetEmbeddedFileCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetEmbeddedFileNode Lib "dynapdf.dll" (ByVal IEF As IntPtr, ByRef F As TPDFEmbFileNode_I, ByVal Decompress As Integer) As Integer
        Private Declare Function pdfGetEMFPatternDistance Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetErrLogMessage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Err As TPDFError_I) As Integer
        Private Declare Function pdfGetErrLogMessageCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetErrorMessage Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As IntPtr
        Private Declare Function pdfGetErrorMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByRef Field As TPDFField_I) As Integer
        Private Declare Function pdfGetFieldBackColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetFieldBorderColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetFieldBorderStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldBorderWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Double
        Private Declare Function pdfGetFieldChoiceValue Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal ValIndex As Integer, ByRef Value As TPDFChoiceValue_I) As Integer
        Private Declare Function pdfGetFieldColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal ColorType As Integer, ByRef ColorSpace As Integer, ByRef Color As Integer) As Integer
        Private Declare Function pdfGetFieldCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetFieldEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Field As TPDFFieldEx_I) As Integer
        Private Declare Function pdfGetFieldEx2 Lib "dynapdf.dll" (ByVal IField As IntPtr, ByRef Field As TPDFFieldEx_I) As Integer
        Private Declare Function pdfGetFieldExpValCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldExpValue Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByRef Value As IntPtr) As Integer
        Private Declare Function pdfGetFieldExpValueEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal ValIndex As Integer, ByRef Value As IntPtr, ByRef ExpValue As IntPtr, ByRef Selected As Integer) As Integer
        Private Declare Function pdfGetFieldFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldGroupType Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldHighlightMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldIndex Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldMapName Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByRef Value As IntPtr, ByRef bUnicode As Integer) As Integer
        Private Declare Function pdfGetFieldName Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByRef Name As IntPtr) As Integer
        Private Declare Function pdfGetFieldOrientation Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldTextAlign Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFieldTextColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetFieldToolTip Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByRef Value As IntPtr, ByRef bUnicode As Integer) As Integer
        Private Declare Function pdfGetFieldType Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer) As Integer
        Private Declare Function pdfGetFileSpec Lib "dynapdf.dll" (ByVal IFS As IntPtr, ByRef F As TPDFFileSpecEx_I) As Integer
        Private Declare Function pdfGetFillColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetFontCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetFontEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Font As TPDFFontObj_I) As Integer
        Private Declare Function pdfGetFontInfoEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Font As TPDFFontInfo_I) As Integer
        Private Declare Function pdfGetFontOrigin Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Sub pdfGetFontSearchOrder Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Order() As TFontBaseType)
        Private Declare Function pdfGetFontSelMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetFontWeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfGetFTextHeightA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Align As Integer, ByVal AText As String) As Double
        Private Declare Unicode Function pdfGetFTextHeightW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Align As Integer, ByVal AText As String) As Double
        Private Declare Ansi Function pdfGetFTextHeightExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Width As Double, ByVal Align As Integer, ByVal AText As String) As Double
        Private Declare Unicode Function pdfGetFTextHeightExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Width As Double, ByVal Align As Integer, ByVal AText As String) As Double
        Private Declare Function pdfGetGlyphIndex Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer) As Integer
        Private Declare Function pdfGetGlyphOutline Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByVal Outline As TPDFGlyphOutline_I) As Integer
        Private Declare Function pdfGetGoToAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFGoToAction_I) As Integer
        Private Declare Function pdfGetGoToRAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFGoToAction_I) As Integer
        Private Declare Function pdfGetGStateFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetHideAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFHideAction_I) As Integer
        Private Declare Function pdfGetIconColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetImageBuffer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef BufSize As Integer) As IntPtr
        Private Declare Auto Function pdfGetImageCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String) As Integer
        Private Declare Function pdfGetImageHeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer) As Integer
        Private Declare Function pdfGetImageObj Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Flags As TParseFlags, ByRef Image As TPDFImage) As Integer
        Private Declare Function pdfGetImageObjCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetImageObjEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ImagePtr As IntPtr, ByVal Flags As TParseFlags, ByRef Image As TPDFImage) As Integer
        Private Declare Function pdfGetImageWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer) As Integer
        Private Declare Function pdfGetImportDataAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFImportDataAction_I) As Integer
        Private Declare Function pdfGetImportFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetImportFlags2 Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInBBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal Boundary As Integer, ByRef BBox As TPDFRect) As Integer
        Private Declare Function pdfGetInDocInfo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DInfo As Integer, ByRef Value As IntPtr) As Integer
        Private Declare Function pdfGetInDocInfoCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInDocInfoEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef DInfo As Integer, ByRef Key As IntPtr, ByRef Value As IntPtr, ByRef IsUnicode As Integer) As Integer
        Private Declare Function pdfGetInEncryptionFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInFieldCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInIsCollection Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInIsEncrypted Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInIsSigned Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInIsTrapped Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInIsXFAForm Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInkList Lib "dynapdf.dll" (ByVal List As IntPtr, ByRef Points As IntPtr, ByRef Count As Integer) As Integer
        Private Declare Function pdfGetInMetadata Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByRef Buffer As IntPtr, ByRef BufSize As Integer) As Integer
        Private Declare Function pdfGetInOrientation Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer) As Integer
        Private Declare Function pdfGetInPrintSettings Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Settings As TPDFPrintSettings_I) As Integer
        Private Declare Function pdfGetInPageCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInPDFVersion Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetInRepairMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetIsFixedPitch Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetIsTaggingEnabled Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetItalicAngle Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetJavaScript Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByRef Len As Integer, ByRef bUnicode As Integer) As IntPtr
        Private Declare Function pdfGetJavaScriptAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AHandle As Integer, ByRef Len As Integer, ByRef bUnicode As Integer) As IntPtr
        Private Declare Function pdfGetJavaScriptAction2 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As Integer, ByVal ObjHandle As Integer, ByVal ActIndex As Integer, ByRef Len As Integer, ByRef bUnicode As Integer, ByRef ObjEvent As TObjEvent) As IntPtr
        Private Declare Function pdfGetJavaScriptActionEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFJavaScriptAction_I) As Integer
        Private Declare Function pdfGetJavaScriptCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetJavaScriptEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByRef Len As Integer, ByRef bUnicode As Integer) As IntPtr
        Private Declare Function pdfGetJavaScriptName Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef NameLen As Integer, ByRef IsUnicode As Integer) As IntPtr
        Private Declare Function pdfGetJPEGQuality Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetLanguage Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As IntPtr
        Private Declare Function pdfGetLastTextPosX Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetLastTextPosY Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetLayerConfig Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Config As TPDFOCLayerConfig_I) As Integer
        Private Declare Function pdfGetLayerConfigCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetLaunchAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFLaunchAction_I) As Integer
        Private Declare Function pdfGetLeading Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetLineCapStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetLineJoinStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetLineWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetLinkHighlightMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Auto Function pdfGetLogMetafileSize Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByRef R As TRectL) As Integer
        Private Declare Function pdfGetLogMetafileSizeEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByRef R As TRectL) As Integer
        Private Declare Function pdfGetMatrix Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Matrix As TCTM) As Integer
        Private Declare Function pdfGetMaxFieldLen Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TxtField As Integer) As Integer
        Private Declare Function pdfGetMeasureObj Lib "dynapdf.dll" (ByVal Measure As IntPtr, ByRef Value As TPDFMeasure_I) As Integer
        Private Declare Function pdfGetMetaConvFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetMetadata Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As TMetadataObj, ByVal Handle As Integer, ByRef Buffer As IntPtr, ByRef BufSize As Integer) As Integer
        Private Declare Function pdfGetMissingGlyphs Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Count As Integer) As IntPtr
        Private Declare Function pdfGetMiterLimit Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetMovieAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFMovieAction_I) As Integer
        Private Declare Function pdfGetNamedAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFNamedAction_I) As Integer
        Private Declare Function pdfGetNamedDest Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Dest As TPDFNamedDest_I) As Integer
        Private Declare Function pdfGetNamedDestCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetNeedAppearance Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetNumberFormatObj Lib "dynapdf.dll" (ByVal NumberFmt As IntPtr, ByRef Value As TPDFNumberFormat_I) As Integer
        Private Declare Function pdfGetObjActionCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As Integer, ByVal ObjHandle As Integer) As Integer
        Private Declare Function pdfGetObjActions Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As TObjType, ByVal ObjHandle As Integer, ByRef Actions As TPDFObjActions) As Integer
        Private Declare Function pdfGetObjEvent Lib "dynapdf.dll" (ByVal IEvent As IntPtr, ByRef ObjEvent As TPDFObjEvent) As Integer
        Private Declare Function pdfGetOCG Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Value As TPDFOCG) As Integer
        Private Declare Function pdfGetOCGContUsage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Value As TPDFOCGContUsage) As Integer
        Private Declare Function pdfGetOCGCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetOCGUsageUserName Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Index As Integer, ByRef NameA As IntPtr, ByRef NameW As IntPtr) As Integer
        Private Declare Function pdfGetOCHandle Lib "dynapdf.dll" (ByVal OC As IntPtr) As Integer
        Private Declare Function pdfGetOCUINode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Node As IntPtr, ByVal OutNode As TPDFOCUINode_I) As IntPtr
        Private Declare Function pdfGetOpacity Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetOrientation Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetOutputIntent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Intent As TPDFOutputIntent_I) As Integer
        Private Declare Function pdfGetOutputIntentCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageAnnot Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Annot As TPDFAnnotation_I) As Integer
        Private Declare Function pdfGetPageAnnotCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageAnnotEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Annot As TPDFAnnotationEx_I) As Integer
        Private Declare Function pdfGetPageBBox Lib "dynapdf.dll" (ByVal PagePtr As IntPtr, ByVal Boundary As TPageBoundary, ByRef BBox As TFltRect) As Integer
        Private Declare Function pdfGetPageCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Field As TPDFField_I) As Integer
        Private Declare Function pdfGetPageFieldCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageFieldEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Field As TPDFFieldEx_I) As Integer
        Private Declare Function pdfGetPageHeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetPageLayout Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageNum Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageObject Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer) As IntPtr
        Private Declare Function pdfGetPageOrientation Lib "dynapdf.dll" (ByVal PagePtr As IntPtr) As Integer
        Private Declare Function pdfGetPageText Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Stack As TPDFStack) As Integer
        Private Declare Function pdfGetPageWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetPDFVersion Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPrintSettings Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Settings As TPDFPrintSettings_I) As Integer
        Private Declare Function pdfGetPtDataArray Lib "dynapdf.dll" (ByVal PtData As IntPtr, ByVal Index As Integer, ByRef DataType As IntPtr, ByRef Values As IntPtr, ByRef ValCount As Integer) As Integer
        Private Declare Function pdfGetPtDataObj Lib "dynapdf.dll" (ByVal PtData As IntPtr, ByRef Subtype As IntPtr, ByRef NumArrays As Integer) As Integer
        Private Declare Function pdfGetRelFileNode Lib "dynapdf.dll" (ByVal IRF As IntPtr, ByRef F As TPDFRelFileNode_I, ByVal Decompress As Integer) As Integer
        Private Declare Function pdfGetResetAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Value As TPDFResetFormAction_I) As Integer
        Private Declare Function pdfGetResolution Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetSaveNewImageFormat Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetPageLabel Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByRef Label As TPDFPageLabel_I) As Integer
        Private Declare Function pdfGetPageLabelCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetSeparationInfo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Colorant As IntPtr, ByRef CS As TExtColorSpace) As Integer
        Private Declare Function pdfGetSigDict Lib "dynapdf.dll" (ByVal ISignature As IntPtr, ByRef SigDict As TPDFSigDict_I) As Integer
        Private Declare Function pdfGetStrokeColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetSubmitAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Value As TPDFSubmitFormAction_I) As Integer
        Private Declare Function pdfGetSysFontInfo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Value As TPDFSysFont_I) As Integer
        Private Declare Function pdfGetTabLen Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetTemplCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetTemplHandle Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetTemplHeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Double
        Private Declare Function pdfGetTemplWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Double
        Private Declare Function pdfGetTextDrawMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetTextFieldValue Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByRef Value As IntPtr, ByRef ValIsUnicode As Integer, ByRef DefValue As IntPtr, ByRef DefValIsUnicode As Integer) As Integer
        Private Declare Function pdfGetTextRect Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef PosX As Double, ByRef PosY As Double, ByRef Width As Double, ByRef Height As Double) As Integer
        Private Declare Function pdfGetTextRise Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetTextScaling Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Ansi Function pdfGetTextWidthA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String) As Double
        Private Declare Unicode Function pdfGetTextWidthW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String) As Double
        Private Declare Ansi Function pdfGetTextWidthExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String, ByVal Len As Integer) As Double
        Private Declare Unicode Function pdfGetTextWidthExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String, ByVal Len As Integer) As Double
        Private Declare Function pdfGetTransparentColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetTrapped Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetURIAction Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Action As TPDFURIAction_I) As Integer
        Private Declare Function pdfGetUseExactPwd Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetUseGlobalImpFiles Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetUserRights Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetUserUnit Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Single
        Private Declare Function pdfGetUseStdFonts Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetUseSystemFonts Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetUsesTransparency Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer) As Integer
        Private Declare Function pdfGetUseTransparency Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetUseVisibleCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetViewerPrefrences Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Preference As Integer, ByRef AddVal As Integer) As Integer
        Private Declare Function pdfGetViewport Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal Index As Integer, ByRef Value As TPDFViewport_I) As Integer
        Private Declare Function pdfGetViewportCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer) As Integer
        Private Declare Function pdfGetWMFDefExtent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Width As Integer, ByRef Height As Integer) As Integer
        Private Declare Function pdfGetWMFPixelPerInch Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfGetWordSpacing Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Double
        Private Declare Function pdfGetXFAStream Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer, ByVal Out As TPDFXFAStream_I) As Integer
        Private Declare Function pdfGetXFAStreamCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfHaveOpenDoc Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfHaveOpenPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfHighlightAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SubType As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Unicode Function pdfHighlightAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SubType As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Color As Integer, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Function pdfImportBookmarks Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfImportCatalogObjects Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfImportDocInfo Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfImportEncryptionSettings Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfImportOCProperties Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfImportPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer) As Integer
        Private Declare Function pdfImportPageEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal ScaleX As Double, ByVal ScaleY As Double) As Integer
        Private Declare Function pdfImportPDFFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DestPage As Integer, ByVal ScaleX As Double, ByVal ScaleY As Double) As Integer
        Private Declare Function pdfInitColorManagement Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Profiles As TPDFColorProfiles, ByVal DestSpace As TPDFColorSpace, ByVal Flags As TPDFInitCMFlags) As Integer
        Private Declare Function pdfInitColorManagementEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Profiles As TPDFColorProfilesEx, ByVal DestSpace As TPDFColorSpace, ByVal Flags As TPDFInitCMFlags) As Integer
        Private Declare Function pdfInitExtGState Lib "dynapdf.dll" (ByRef GS As TPDFExtGState) As Integer
        Private Declare Function pdfInitOCGContUsage Lib "dynapdf.dll" (ByRef Value As TPDFOCGContUsage) As Integer
        Private Declare Function pdfInitStack Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Stack As TPDFStack) As Integer
        Private Declare Ansi Function pdfInkAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Points() As TFltPoint, ByVal NumPoints As Integer, ByVal LineWidth As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Unicode Function pdfInkAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Points() As TFltPoint, ByVal NumPoints As Integer, ByVal LineWidth As Double, ByVal Color As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Function pdfInsertBMPFromBuffer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal Buffer() As Byte) As Integer
        Private Declare Function pdfInsertBMPFromHandle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal hBitmap As IntPtr) As Integer
        Private Declare Ansi Function pdfInsertBookmarkA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Integer, ByVal AddChildren As Integer) As Integer
        Private Declare Unicode Function pdfInsertBookmarkW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal DestPage As Integer, ByVal DoOpen As Integer, ByVal AddChildren As Integer) As Integer
        Private Declare Ansi Function pdfInsertBookmarkExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Integer, ByVal AddChildren As Integer) As Integer
        Private Declare Unicode Function pdfInsertBookmarkExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Title As String, ByVal Parent As Integer, ByVal NamedDest As Integer, ByVal DoOpen As Integer, ByVal AddChildren As Integer) As Integer
        Private Declare Function pdfInsertImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal AFile As String) As Integer
        Private Declare Auto Function pdfInsertImageEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal AFile As String, ByVal Index As Integer) As Integer
        Private Declare Function pdfInsertImageFromBuffer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Index As Integer) As Integer
        Private Declare Function pdfInsertImageFromBuffer2 Lib "dynapdf.dll" Alias "pdfInsertImageFromBuffer" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByVal Buffer As IntPtr, ByVal BufSize As Integer, ByVal Index As Integer) As Integer
        Private Declare Auto Function pdfInsertMetafile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfInsertMetafileEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Auto Function pdfInsertMetafileExt Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByRef View As TRectL, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfInsertMetafileExtEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByRef View As TRectL, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfInsertMetafileFromHandle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal hEnhMetafile As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfInsertMetafileFromHandleEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal hEnhMetafile As IntPtr, ByRef View As TRectL, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfInsertRawImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BitsPerPixel As Integer, ByVal ColorCount As Integer, ByVal ImgWidth As Integer, ByVal ImgHeight As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Integer
        Private Declare Function pdfInsertRawImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer As IntPtr, ByVal BitsPerPixel As Integer, ByVal ColorCount As Integer, ByVal ImgWidth As Integer, ByVal ImgHeight As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Integer
        Private Declare Function pdfInsertRawImageEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double, ByRef Image As TPDFRawImage) As Integer
        Private Declare Function pdfIsBidiText Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String) As Integer
        Private Declare Function pdfIsColorPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal GrayIsColor As Integer) As Integer
        Private Declare Function pdfIsEmptyPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfLineAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Unicode Function pdfLineAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Function pdfLineTo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double) As Integer
        Private Declare Ansi Function pdfLoadCMap Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal CMapName As String, ByVal Embed As Integer) As Integer
        Private Declare Ansi Function pdfLoadFDFDataA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal Password As String, ByVal Flags As Integer) As Integer
        Private Declare Unicode Function pdfLoadFDFDataW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, <MarshalAs(UnmanagedType.LPStr)> ByVal Password As String, ByVal Flags As Integer) As Integer
        Private Declare Ansi Function pdfLoadFDFDataEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Password As String, ByVal Flags As Integer) As Integer
        Private Declare Function pdfLoadFont Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Style As Integer, ByVal Size As Double, ByVal Embed As Integer, ByVal CP As Integer) As Integer
        Private Declare Auto Function pdfLoadFontEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FontFile As String, ByVal Index As Integer, ByVal Style As Integer, ByVal Size As Double, ByVal Embed As Integer, ByVal CP As Integer) As Integer
        Private Declare Function pdfLoadLayerConfig Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Index As Integer) As Integer
        Private Declare Function pdfLockLayer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Layer As Integer) As Integer
        Private Declare Function pdfMovePage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Source As Integer, ByVal Dest As Integer) As Integer
        Private Declare Function pdfMoveTo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double) As Integer
        Private Declare Function pdfNewPDF Lib "dynapdf.dll" () As IntPtr
        Private Declare Function pdfOpenImportBuffer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal PwdType As Integer, ByVal Password As String) As Integer
        Private Declare Ansi Function pdfOpenImportFileA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PwdType As Integer, ByVal Password As String) As Integer
        Private Declare Unicode Function pdfOpenImportFileW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PwdType As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal Password As String) As Integer
        Private Declare Ansi Function pdfOpenOutputFileA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutPDF As String) As Integer
        Private Declare Unicode Function pdfOpenOutputFileW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutPDF As String) As Integer
        Private Declare Auto Function pdfOpenOutputFileEncrypted Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutPDF As String, <MarshalAs(UnmanagedType.LPStr)> ByVal OpenPwd As String, <MarshalAs(UnmanagedType.LPStr)> ByVal OwnerPwd As String, ByVal KeyLen As TKeyLen, ByVal Restrict As TRestrictions) As Integer
        Private Declare Ansi Function pdfOpenTagA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Tag As Integer, ByVal Lang As String, ByVal AltText As String, ByVal Expansion As String) As Integer
        Private Declare Unicode Function pdfOpenTagW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Tag As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal Lang As String, ByVal AltText As String, ByVal Expansion As String) As Integer
        Private Declare Function pdfOptimize Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As TOptimizeFlags, ByVal Parms As TOptimizeParams) As Integer
        Private Declare Function pdfPageLink Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal DestPage As Integer) As Integer
        Private Declare Function pdfPageLink2 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal NamedDest As Integer) As Integer
        Private Declare Ansi Function pdfPageLink3A Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal NamedDest As String) As Integer
        Private Declare Unicode Function pdfPageLink3W Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal NamedDest As String) As Integer
        Private Declare Function pdfPageLinkEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal DestType As Integer, ByVal DestPage As Integer, ByVal a As Double, ByVal b As Double, ByVal C As Double, ByVal d As Double) As Integer
        Private Declare Function pdfParseContent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Data As IntPtr, ByRef Stack As TPDFParseInterface, ByVal Flags As TParseFlags) As Integer
        Private Declare Function pdfPlaceImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ImgHandle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Integer
        Private Declare Function pdfPlaceSigFieldValidateIcon Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SigField As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfPlaceTemplate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TmplHandle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Integer
        Private Declare Function pdfPlaceTemplateEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TmplHandle As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal ScaleWidth As Double, ByVal ScaleHeight As Double) As Integer
        Private Declare Ansi Function pdfPolygonAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Vertices() As TFltPoint, ByVal NumVertices As Integer, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Unicode Function pdfPolygonAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Vertices() As TFltPoint, ByVal NumVertices As Integer, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Ansi Function pdfPolyLineAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Vertices() As TFltPoint, ByVal NumVertices As Integer, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Unicode Function pdfPolyLineAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Vertices() As TFltPoint, ByVal NumVertices As Integer, ByVal LineWidth As Double, ByVal lStart As TLineEndStyle, ByVal lEnd As TLineEndStyle, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Content As String) As Integer
        Private Declare Auto Function pdfPrintPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal DocName As String, ByVal DC As IntPtr, ByVal Flags As TPDFPrintFlags, ByRef Margin As TRectL, ByVal Parms As TPDFPrintParams) As Integer
        Private Declare Auto Function pdfPrintPDFFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TmpDir As String, ByVal DocName As String, ByVal DC As IntPtr, ByVal Flags As TPDFPrintFlags, ByRef Margin As TRectL, ByVal Parms As TPDFPrintParams) As Integer
        Private Declare Auto Function pdfReadImageFormat Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByRef Width As Integer, ByRef Height As Integer, ByRef BitsPerPixel As Integer, ByRef UseZip As Integer) As Integer
        Private Declare Auto Function pdfReadImageFormat2 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal Index As Integer, ByRef Width As Integer, ByRef Height As Integer, ByRef BitsPerPixel As Integer, ByRef UseZip As Integer) As Integer
        Private Declare Function pdfReadImageFormatEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal hBitmap As IntPtr, ByRef Width As Integer, ByRef Height As Integer, ByRef BitsPerPixel As Integer, ByRef UseZip As Integer) As Integer
        Private Declare Function pdfReadImageFormatFromBuffer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Index As Integer, ByRef Width As Integer, ByRef Height As Integer, ByRef BitsPerPixel As Integer, ByRef UseZip As Integer) As Integer
        Private Declare Auto Function pdfReadImageResolution Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal Index As Integer, ByRef resX As Integer, ByRef resY As Integer) As Integer
        Private Declare Function pdfReadImageResolutionEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer As Byte(), ByVal BufSize As Integer, ByVal Index As Integer, ByRef resX As Integer, ByRef resY As Integer) As Integer
        Private Declare Function pdfRectangle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal FillMode As Integer) As Integer
        Private Declare Ansi Function pdfReEncryptPDFA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PwdType As Integer, ByVal InPwd As String, ByVal NewOpenPwd As String, ByVal NewOwnerPwd As String, ByVal NewKeyLen As Integer, ByVal Restrict As Integer) As Integer
        Private Declare Unicode Function pdfReEncryptPDFW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FileName As String, ByVal PwdType As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal InPwd As String, <MarshalAs(UnmanagedType.LPStr)> ByVal NewOpenPwd As String, <MarshalAs(UnmanagedType.LPStr)> ByVal NewOwnerPwd As String, ByVal NewKeyLen As Integer, ByVal Restrict As Integer) As Integer
        Private Declare Function pdfRenameSpotColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Colorant As String, ByVal NewName As String) As Integer
        Private Declare Ansi Function pdfRenderAnnotOrField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal IsAnnot As Integer, ByVal State As TButtonState, ByRef Matrix As TCTM, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByRef Out As TPDFBitmap) As Integer
        Private Declare Ansi Function pdfRenderPage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PagePtr As IntPtr, ByVal RasPtr As IntPtr, ByRef Img As TPDFRasterImage) As Integer
        Private Declare Function pdfRenderPageEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DC As IntPtr, ByRef DestX As Integer, ByRef DestY As Integer, ByVal PagePtr As IntPtr, ByVal RasPtr As IntPtr, ByRef Img As TPDFRasterImage) As Integer
        Private Declare Auto Function pdfRenderPageToImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageNum As Integer, ByVal OutFile As String, ByVal Resolution As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Integer
        Private Declare Auto Function pdfRenderPDFFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutFile As String, ByVal Resolution As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Integer
        Private Declare Ansi Function pdfRenderPDFFileA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutFile As String, ByVal Resolution As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Integer
        Private Declare Unicode Function pdfRenderPDFFileW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutFile As String, ByVal Resolution As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Integer
        Private Declare Auto Function pdfRenderPDFFileEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OutFile As String, ByVal Resolution As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Flags As TRasterFlags, ByVal PixFmt As TPDFPixFormat, ByVal Filter As TCompressionFilter, ByVal Format As TImageFormat) As Integer
        Private Declare Function pdfReOpenImportFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Ansi Function pdfReplaceFontA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PDFFont As IntPtr, ByVal Name As String, ByVal Style As TFStyle, ByVal NameIsFamilyName As Integer) As Integer
        Private Declare Unicode Function pdfReplaceFontW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PDFFont As IntPtr, ByVal Name As String, ByVal Style As TFStyle, ByVal NameIsFamilyName As Integer) As Integer
        Private Declare Ansi Function pdfReplaceFontExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PDFFont As IntPtr, ByVal FontFile As String, ByVal Embed As Integer) As Integer
        Private Declare Unicode Function pdfReplaceFontExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PDFFont As IntPtr, ByVal FontFile As String, ByVal Embed As Integer) As Integer
        Private Declare Ansi Function pdfReplaceICCProfileA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ColorSpace As Integer, ByVal ICCFile As String) As Integer
        Private Declare Unicode Function pdfReplaceICCProfileW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ColorSpace As Integer, ByVal ICCFile As String) As Integer
        Private Declare Function pdfReplaceICCProfileEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ColorSpace As Integer, ByVal Buffer As Byte(), ByVal BufSize As Integer) As Integer
        Private Declare Auto Function pdfReplaceImage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Source As IntPtr, ByVal Image As String, ByVal Index As Integer, ByVal CS As TExtColorSpace, ByVal CSHandle As Integer, ByVal Flags As TReplaceImageFlags) As Integer
        Private Declare Ansi Function pdfReplaceImageEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Source As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Index As Integer, ByVal CS As TExtColorSpace, ByVal CSHandle As Integer, ByVal Flags As TReplaceImageFlags) As Integer
        Private Declare Ansi Function pdfReplacePageTextA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal NewText As String, ByRef Stack As TPDFStack) As Integer
        Private Declare Ansi Function pdfReplacePageTextExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal NewText As String, ByRef Stack As TPDFStack) As Integer
        Private Declare Unicode Function pdfReplacePageTextExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal NewText As String, ByRef Stack As TPDFStack) As Integer
        Private Declare Function pdfResetEncryptionSettings Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfResetLineDashPattern Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfRestoreGraphicState Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfRotateCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal alpha As Double, ByVal OriginX As Double, ByVal OriginY As Double) As Integer
        Private Declare Function pdfRoundRect Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Radius As Double, ByVal FillMode As Integer) As Integer
        Private Declare Function pdfRoundRectEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal rWidth As Double, ByVal rHeight As Double, ByVal FillMode As Integer) As Integer
        Private Declare Function pdfSaveGraphicState Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfScaleCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal sx As Double, ByVal sy As Double) As Integer
        Private Declare Function pdfSelfTest Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfSet3DAnnotProps Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Annot As Integer, ByVal ActType As Integer, ByVal DeActType As Integer, ByVal InstType As Integer, ByVal DeInstType As Integer, ByVal DisplToolbar As Integer, ByVal DisplModelTree As Integer) As Integer
        Private Declare Ansi Function pdfSet3DAnnotScriptA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Annot As Integer, ByVal Value As String, ByVal Len As Integer) As Integer
        Private Declare Function pdfSetAllocBy Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetAnnotBorderStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Style As TBorderStyle) As Integer
        Private Declare Function pdfSetAnnotBorderWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal LineWidth As Double) As Integer
        Private Declare Function pdfSetAnnotColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal ColorType As TAnnotColor, ByVal CS As TPDFColorSpace, ByVal Color As Integer) As Integer
        Private Declare Function pdfSetAnnotFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer) As Integer
        Private Declare Function pdfSetAnnotFlagsEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Flags As Integer) As Integer
        Private Declare Function pdfSetAnnotHighlightMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Mode As THighlightMode) As Integer
        Private Declare Function pdfSetAnnotIcon Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Icon As TAnnotIcon) As Integer
        Private Declare Function pdfSetAnnotLineDashPattern Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Dash() As Single, ByVal NumValues As Integer) As Integer
        Private Declare Function pdfSetAnnotLineEndStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal StartStyle As TLineEndStyle, ByVal EndStyle As TLineEndStyle) As Integer
        Private Declare Auto Function pdfSetAnnotMigrationState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Annot As Integer, ByVal State As TAnnotState, ByVal User As String) As Integer
        Private Declare Function pdfSetAnnotOpacity Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Value As Double) As Integer
        Private Declare Function pdfSetAnnotOpenState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Function pdfSetAnnotOrFieldDate Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal IsField As Integer, ByVal DateType As TDateType, ByVal DateTime As Integer) As Integer
        Private Declare Function pdfSetAnnotQuadPoints Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Value() As TFltPoint, ByVal Count As Integer) As Integer
        Private Declare Ansi Function pdfSetAnnotStringA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal StringType As Integer, ByVal Value As String) As Integer
        Private Declare Unicode Function pdfSetAnnotStringW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal StringType As Integer, ByVal Value As String) As Integer
        Private Declare Ansi Function pdfSetAnnotSubjectA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Value As String) As Integer
        Private Declare Unicode Function pdfSetAnnotSubjectW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Value As String) As Integer
        Private Declare Function pdfSetBBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Boundary As Integer, ByVal LeftX As Double, ByVal LeftY As Double, ByVal RightX As Double, ByVal RightY As Double) As Integer
        Private Declare Function pdfSetBidiMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Mode As Integer) As Integer
        Private Declare Function pdfSetBookmarkDest Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ABmk As Integer, ByVal DestType As Integer, ByVal a As Double, ByVal b As Double, ByVal C As Double, ByVal d As Double) As Integer
        Private Declare Function pdfSetBookmarkStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ABmk As Integer, ByVal Style As Integer, ByVal RGBColor As Integer) As Integer
        Private Declare Function pdfSetAnnotBorderEffect Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Value As TBorderEffect) As Integer
        Private Declare Function pdfSetBorderStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Style As Integer) As Integer
        Private Declare Function pdfSetCharacterSpacing Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetCheckBoxChar Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal CheckBoxChar As Integer) As Integer
        Private Declare Function pdfSetCheckBoxDefState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Checked As Integer) As Integer
        Private Declare Function pdfSetCheckBoxState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Checked As Integer) As Integer
        Private Declare Ansi Function pdfSetCIDFontA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal CMapHandle As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Integer) As Integer
        Private Declare Unicode Function pdfSetCIDFontW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal CMapHandle As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Integer) As Integer
        Private Declare Ansi Function pdfSetCMapDirA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Path As String, ByVal Flags As TLoadCMapFlags) As Integer
        Private Declare Unicode Function pdfSetCMapDirW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Path As String, ByVal Flags As TLoadCMapFlags) As Integer
        Private Declare Function pdfSetColDefFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal EmbFile As Integer) As Integer
        Private Declare Function pdfSetColorMask Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ImageHandle As Integer, ByVal Mask() As Integer, ByVal Count As Integer) As Integer
        Private Declare Function pdfSetColors Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color As Integer) As Integer
        Private Declare Function pdfSetColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ColorSpace As Integer) As Integer
        Private Declare Function pdfSetColSortField Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ColField As Integer, ByVal AscendingOrder As Integer) As Integer
        Private Declare Function pdfSetCompressionFilter Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ComprFilter As Integer) As Integer
        Private Declare Function pdfSetCompressionLevel Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal CompressLevel As Integer) As Integer
        Private Declare Function pdfSetContent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Buffer() As Byte, ByVal BufSize As Integer) As Integer
        Private Declare Function pdfSetDateTimeFormat Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TxtField As Integer, ByVal Fmt As Integer) As Integer
        Private Declare Function pdfSetDefBitsPerPixel Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Ansi Function pdfSetDocInfoA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DInfo As Integer, ByVal Value As String) As Integer
        Private Declare Unicode Function pdfSetDocInfoW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DInfo As Integer, ByVal Value As String) As Integer
        Private Declare Ansi Function pdfSetDocInfoExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DInfo As Integer, ByVal Key As String, ByVal Value As String) As Integer
        Private Declare Unicode Function pdfSetDocInfoExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DInfo As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal Key As String, ByVal Value As String) As Integer
        Private Declare Function pdfSetDrawDirection Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Direction As Integer) As Integer
        Private Declare Function pdfSetEMFFrameDPI Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal DPIX As Integer, ByVal DPIY As Integer) As Integer
        Private Declare Function pdfSetEMFPatternDistance Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetErrorMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ErrMode As Integer) As Integer
        Private Declare Function pdfSetExtColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfSetExtFillColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfSetExtGState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfSetExtStrokeColorSpace Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfSetFieldBackColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AColor As Integer) As Integer
        Private Declare Function pdfSetFieldBBox Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByRef BBox As TPDFRect) As Integer
        Private Declare Function pdfSetFieldBorderColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AColor As Integer) As Integer
        Private Declare Function pdfSetFieldBorderStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Style As Integer) As Integer
        Private Declare Function pdfSetFieldBorderWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal LineWidth As Double) As Integer
        Private Declare Function pdfSetFieldColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal ColorType As Integer, ByVal CS As Integer, ByVal Color As Integer) As Integer
        Private Declare Ansi Function pdfSetFieldExpValueA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal ValIndex As Integer, ByVal Value As String, ByVal ExpValue As String, ByVal Selected As Integer) As Integer
        Private Declare Ansi Function pdfSetFieldExpValueW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal ValIndex As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal Value As String, ByVal ExpValue As String, ByVal Selected As Integer) As Integer
        Private Declare Function pdfSetFieldExpValueEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AField As Integer, ByVal ValIndex As Integer, ByVal Selected As Integer, ByVal DefSelected As Integer) As Integer
        Private Declare Function pdfSetFieldFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Flags As Integer, ByVal DoReset As Integer) As Integer
        Private Declare Auto Function pdfSetFieldFont Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Size As Double, ByVal Embed As Integer, ByVal CP As TCodepage) As Integer
        Private Declare Function pdfSetFieldFontEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Handle As Integer, ByVal FontSize As Double) As Integer
        Private Declare Function pdfSetFieldFontSize Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal FontSize As Double) As Integer
        Private Declare Function pdfSetFieldHighlightMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Mode As Integer) As Integer
        Private Declare Function pdfSetFieldIndex Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Index As Integer) As Integer
        Private Declare Ansi Function pdfSetFieldMapNameA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String) As Integer
        Private Declare Unicode Function pdfSetFieldMapNameW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String) As Integer
        Private Declare Auto Function pdfSetFieldName Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal NewName As String) As Integer
        Private Declare Function pdfSetFieldOrientation Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetFieldTextAlign Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Align As Integer) As Integer
        Private Declare Function pdfSetFieldTextColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color As Integer) As Integer
        Private Declare Ansi Function pdfSetFieldToolTipA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String) As Integer
        Private Declare Unicode Function pdfSetFieldToolTipW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String) As Integer
        Private Declare Function pdfSetFillColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color As Integer) As Integer
        Private Declare Function pdfSetFillColorEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color() As Byte, ByVal NumComponents As Integer) As Integer
        Private Declare Function pdfSetFillColorF Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color() As Single, ByVal NumComponents As Integer) As Integer
        Private Declare Function pdfSetFloatPrecision Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal NumTextDecDigits As Integer, ByVal NumVectDecDigits As Integer) As Integer
        Private Declare Ansi Function pdfSetFontA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Style As Integer, ByVal Size As Double, ByVal Embed As Integer, ByVal CP As Integer) As Integer
        Private Declare Unicode Function pdfSetFontW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Style As Integer, ByVal Size As Double, ByVal Embed As Integer, ByVal CP As Integer) As Integer
        Private Declare Ansi Function pdfSetFontExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Style As Integer, ByVal Size As Double, ByVal Embed As Integer, ByVal CP As Integer) As Integer
        Private Declare Unicode Function pdfSetFontExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Name As String, ByVal Style As Integer, ByVal Size As Double, ByVal Embed As Integer, ByVal CP As Integer) As Integer
        Private Declare Function pdfSetFontOrigin Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Origin As Integer) As Integer
        Private Declare Sub pdfSetFontSearchOrder Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Order() As TFontBaseType)
        Private Declare Sub pdfSetFontSearchOrderEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal S1 As TFontBaseType, ByVal S2 As TFontBaseType, ByVal S3 As TFontBaseType, ByVal S4 As TFontBaseType)
        Private Declare Function pdfSetFontSelMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Mode As Integer) As Integer
        Private Declare Function pdfSetFontWeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Weight As Integer) As Integer
        Private Declare Sub pdfSetGStateFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer, ByVal Reset As Integer)
        Private Declare Function pdfSetIconColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color As Integer) As Integer
        Private Declare Function pdfSetImportFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer) As Integer
        Private Declare Function pdfSetImportFlags2 Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer) As Integer
        Private Declare Function pdfSetItalicAngle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Angle As Double) As Integer
        Private Declare Function pdfSetJPEGQuality Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Ansi Function pdfSetLanguage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ISOTag As String) As Integer
        Private Declare Function pdfSetLeading Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Ansi Function pdfSetLicenseKey Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As String) As Integer
        Private Declare Function pdfSetLineAnnotParms Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal FontHandle As Integer, ByVal FontSize As Double, ByVal Parms As TLineAnnotParms) As Integer
        Private Declare Function pdfSetLineCapStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Style As Integer) As Integer
        Private Declare Ansi Function pdfSetLineDashPattern Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Dash As String, ByVal Phase As Integer) As Integer
        Private Declare Function pdfSetLineDashPatternEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Dash() As Double, ByVal NumValues As Integer, ByVal Phase As Integer) As Integer
        Private Declare Function pdfSetLineJoinStyle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Style As Integer) As Integer
        Private Declare Function pdfSetLineWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetLinkHighlightMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Mode As Integer) As Integer
        Private Declare Function pdfSetListFont Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfSetMatrix Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Matrix As TCTM) As Integer
        Private Declare Sub pdfSetMaxErrLogMsgCount Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer)
        Private Declare Function pdfSetMaxFieldLen Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TxtField As Integer, ByVal MaxLen As Integer) As Integer
        Private Declare Function pdfSetMetaConvFlags Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Flags As Integer) As Integer
        Private Declare Function pdfSetMetadata Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal ObjType As TMetadataObj, ByVal Handle As Integer, ByVal Buffer() As Byte, ByVal BufSize As Integer) As Integer
        Private Declare Function pdfSetMiterLimit Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetNeedAppearance Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Ansi Function pdfSetNumberFormat Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TxtField As Integer, ByVal Sep As Integer, ByVal DecPlaces As Integer, ByVal NegStyle As Integer, ByVal CurrStr As String, ByVal Prepend As Integer) As Integer
        Private Declare Function pdfSetOCGContUsage Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByRef Value As TPDFOCGContUsage) As Integer
        Private Declare Function pdfSetOCGState Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Visible As Integer, ByVal SaveState As Integer) As Integer
        Private Declare Function pdfSetOnErrorProc Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Data As IntPtr, ByVal ErrProc As TErrorProc) As Integer
        Private Declare Function pdfSetOnPageBreakProc Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Data As IntPtr, ByVal OnBreakProc As TOnPageBreakProc) As Integer
        Private Declare Function pdfSetOpacity Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetOrientation Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetOrientationEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetPageBBox Lib "dynapdf.dll" (ByVal PagePtr As IntPtr, ByVal Boundary As TPageBoundary, ByRef BBox As TFltRect) As Integer
        Private Declare Function pdfSetPageCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PageCoords As Integer) As Integer
        Private Declare Function pdfSetPageFormat Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetPageHeight Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetPageLayout Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Layout As Integer) As Integer
        Private Declare Function pdfSetPageMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Mode As Integer) As Integer
        Private Declare Function pdfSetPageWidth Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetPDFVersion Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Version As Integer) As Integer
        Private Declare Function pdfSetPrintSettings Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Mode As Integer, ByVal PickTrayByPDFSize As Integer, ByVal NumCopies As Integer, ByVal PrintScaling As Integer, ByVal PrintRanges() As Integer, ByVal NumRanges As Integer) As Integer
        Private Declare Function pdfSetProgressProc Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Data As IntPtr, ByVal InitProgress As TInitProgress, ByVal Progress As TProgress) As Integer
        Private Declare Function pdfSetResolution Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetSaveNewImageFormat Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetSeparationInfo Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer) As Integer
        Private Declare Function pdfSetStrokeColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color As Integer) As Integer
        Private Declare Function pdfSetStrokeColorEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color() As Byte, ByVal NumComponents As Integer) As Integer
        Private Declare Function pdfSetStrokeColorF Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Color() As Single, ByVal NumComponents As Integer) As Integer
        Private Declare Function pdfSetTabLen Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal TabLen As Integer) As Integer
        Private Declare Function pdfSetTextDrawMode Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Mode As Integer) As Integer
        Private Declare Ansi Function pdfSetTextFieldValueA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String, ByVal DefValue As String, ByVal Align As Integer) As Integer
        Private Declare Unicode Function pdfSetTextFieldValueW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String, ByVal DefValue As String, ByVal Align As Integer) As Integer
        Private Declare Ansi Function pdfSetTextFieldValueExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String) As Integer
        Private Declare Unicode Function pdfSetTextFieldValueExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Field As Integer, ByVal Value As String) As Integer
        Private Declare Function pdfSetTextRect Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Function pdfSetTextRise Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetTextScaling Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSetTransparentColor Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AColor As Integer) As Integer
        Private Declare Sub pdfSetTrapped Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer)
        Private Declare Function pdfSetUseExactPwd Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetUseGlobalImpFiles Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetUseImageInterpolation Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Handle As Integer, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetUseImageInterpolationEx Lib "dynapdf.dll" (ByVal Image As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetUserUnit Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Single) As Integer
        Private Declare Function pdfSetUseStdFonts Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetUseSwapFile Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SwapContents As Integer, ByVal SwapLimit As Integer) As Integer
        Private Declare Ansi Function pdfSetUseSwapFileEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SwapContents As Integer, ByVal SwapLimit As Integer, ByVal SwapDir As String) As Integer
        Private Declare Function pdfSetUseSystemFonts Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetUseTransparency Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetUseVisibleCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetViewerPreferences Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer, ByVal AddVal As Integer) As Integer
        Private Declare Function pdfSetWMFDefExtent Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Width As Integer, ByVal Height As Integer) As Integer
        Private Declare Function pdfSetWMFPixelPerInch Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Integer) As Integer
        Private Declare Function pdfSetWordSpacing Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Value As Double) As Integer
        Private Declare Function pdfSkewCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal alpha As Double, ByVal beta As Double, ByVal OriginX As Double, ByVal OriginY As Double) As Integer
        Private Declare Function pdfSortFieldsByIndex Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Function pdfSortFieldsByName Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Ansi Function pdfSquareAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Unicode Function pdfSquareAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal LineWidth As Double, ByVal FillColor As Integer, ByVal StrokeColor As Integer, ByVal CS As TPDFColorSpace, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Ansi Function pdfStampAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SubType As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Unicode Function pdfStampAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal SubType As Integer, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Subject As String, ByVal Comment As String) As Integer
        Private Declare Function pdfStrokePath Lib "dynapdf.dll" (ByVal IPDF As IntPtr) As Integer
        Private Declare Auto Function pdfTestGlyphsEx Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal FontHandle As Integer, ByVal Text As String, ByVal Length As Integer) As Integer
        Private Declare Ansi Function pdfTextAnnotA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Text As String, ByVal Icon As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Unicode Function pdfTextAnnotW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Author As String, ByVal Text As String, ByVal Icon As Integer, ByVal DoOpen As Integer) As Integer
        Private Declare Function pdfTranslateCoords Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal OriginX As Double, ByVal OriginY As Double) As Integer
        Private Declare Function pdfTriangle Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal FillMode As Integer) As Integer
        Private Declare Function pdfUnLockLayer Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Layer As Integer) As Integer
        Private Declare Function pdfWatermarkAnnot Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double) As Integer
        Private Declare Ansi Function pdfWebLinkA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal URL As String) As Integer
        Private Declare Unicode Function pdfWebLinkW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal URL As String) As Integer
        Private Declare Ansi Function pdfWriteAngleTextA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String, ByVal Angle As Double, ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal YOrigin As Double) As Integer
        Private Declare Unicode Function pdfWriteAngleTextW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AText As String, ByVal Angle As Double, ByVal PosX As Double, ByVal PosY As Double, ByVal Radius As Double, ByVal YOrigin As Double) As Integer
        Private Declare Ansi Function pdfWriteFTextA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Align As Integer, ByVal AText As String) As Integer
        Private Declare Unicode Function pdfWriteFTextW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal Align As Integer, ByVal AText As String) As Integer
        Private Declare Ansi Function pdfWriteFTextExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Align As Integer, ByVal AText As String) As Integer
        Private Declare Unicode Function pdfWriteFTextExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal Width As Double, ByVal Height As Double, ByVal Align As Integer, ByVal AText As String) As Integer
        Private Declare Ansi Function pdfWriteTextA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String) As Integer
        Private Declare Unicode Function pdfWriteTextW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String) As Integer
        Private Declare Ansi Function pdfWriteTextExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String, ByVal Len As Integer) As Integer
        Private Declare Unicode Function pdfWriteTextExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal PosX As Double, ByVal PosY As Double, ByVal AText As String, ByVal Len As Integer) As Integer
        Private Declare Ansi Function pdfWriteTextMatrixA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Matrix As TCTM, ByVal AText As String) As Integer
        Private Declare Unicode Function pdfWriteTextMatrixW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Matrix As TCTM, ByVal AText As String) As Integer
        Private Declare Ansi Function pdfWriteTextMatrixExA Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Matrix As TCTM, ByVal AText As String, ByVal Len As Integer) As Integer
        Private Declare Unicode Function pdfWriteTextMatrixExW Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByRef Matrix As TCTM, ByVal AText As String, ByVal Len As Integer) As Integer
        ' Helper function to copy a kerning record into the managed type TTextRecordW.
        Public Declare Function CopyKernRecord Lib "dynapdf.dll" Alias "pdfCopyMem" (ByVal Source As IntPtr, ByRef Record As TTextRecordW, ByVal Len As Integer) As Integer
        Private Declare Function pdfCopyMemUInt Lib "dynapdf.dll" Alias "pdfCopyMem" (ByVal Source As IntPtr, ByVal Dest() As System.UInt32, ByVal Len As Integer) As Integer
        Private Declare Function pdfCopyMemIntPtr Lib "dynapdf.dll" Alias "pdfCopyMem" (ByVal Source As IntPtr, ByVal Dest() As IntPtr, ByVal Len As Integer) As Integer
        Private Declare Function pdfStrLenA Lib "dynapdf.dll" (ByVal AStr As IntPtr) As Integer
    End Class

    Public Enum TCellContType
        cctText
        cctImage
        cctTable
        cctTemplate
    End Enum

    Public Enum TDeleteContent
        dcText = &H1              ' Text is always a foreground object
        dcImage = &H2
        dcTemplate = &H4          ' PDF or EMF objects
        dcTable = &H8
        dcAllCont = &H1F          ' Delete all content types
        dcForeGround = &H10000000
        dcBackGround = &H20000000
        dcBoth = &H30000000       ' Delete both foreground and background objects
    End Enum

    Public Enum TCellAlign
        coLeft = 0
        coTop = coLeft
        coRight = 1
        coBottom = coRight
        coCenter = 2
    End Enum

    Public Enum TColumnAdjust
        coaUniqueWidth ' Set the column widths uniquely to TableWidth / NumColumns
        coaAdjLeft     ' Decrease or increase the column widths starting from the left side
        coaAdjRight    ' Decrease or increase the column widths starting from the right side
    End Enum

    Public Enum TTableColor
        tcBackColor = 0     ' Table, Columns, Rows, Cells -> default none (transparent)
        tcBorderColor = 1   ' Table, Columns, Rows, Cells -> default black
        tcGridHorzColor = 2 ' Table                       -> default black
        tcGridVertColor = 3 ' Table                       -> default black
        tcImageColor = 4    ' Table, Columns, Rows, Cells -> default RGB black
        tcTextColor = 5     ' Table, Columns, Rows, Cells -> default black
    End Enum

    Public Enum TTableBoxProperty
        tbpBorderWidth = 0 ' Table, Columns, Rows, Cells -> default (0 0 0 0)
        tbpCellSpacing = 1 ' Table, Columns, Rows, Cells -> default (0 0 0 0)
        tbpCellPadding = 2 ' Table, Columns, Rows, Cells -> default (0 0 0 0)
    End Enum

    Public Enum TTableFlags
        tfDefault = 0
        tfStatic = 1      ' This flag marks a row column or cell as static to avoid the deletion of the content with ClearContent().
        tfHeaderRow = 2   ' Header rows are drawn first after a page break occurred
        tfNoLineBreak = 4 ' Prohibit line breaks in cells whith text -> Can be set to the entire table, columns, rows, and cells
        tfScaleToRect = 8 ' If set, the specified output width and height represents the maximum size of the image or template.
        ' The image or template is scaled into this rectangle without changing the aspect ratio.
        tfUseImageCS = 16 ' If set, images are inserted in the native image color space.
        tfAddFlags = 32   ' If set, the new flags are added to the current ones. If absent, the new flags override the previous value.
    End Enum

    Public Class CPDFTable
        Private m_PDF As Modul_PDF
        Private m_SubTables As ArrayList
        Private m_Table As IntPtr

        Public Sub New(ByVal PDFInst As Modul_PDF, ByVal AllocRows As Integer, ByVal NumCols As Integer, ByVal Width As Single, ByVal DefRowHeight As Single)
            MyBase.New()
            m_PDF = PDFInst ' This is just to make sure that the CPDF instance will not be deleted before the table class is released
            m_Table = tblCreateTable(m_PDF.GetPDFInstance(), AllocRows, NumCols, Width, DefRowHeight)
            If IntPtr.Zero.Equals(m_Table) Then Throw New System.Exception("Out of Memory")
        End Sub

        Protected Overrides Sub Finalize()
            m_SubTables = Nothing
            If Not IntPtr.Zero.Equals(m_Table) Then tblDeleteTable(m_Table)
            m_PDF = Nothing
            MyBase.Finalize()
        End Sub

        Public Function AddColumn(ByVal Left As Integer, ByVal Width As Single) As Integer
            Return tblAddColumn(m_Table, Left, Width)
        End Function

        Public Function AddRow(ByVal Height As Single) As Integer
            Return tblAddRow(m_Table, Height)
        End Function

        Public Function AddRows(ByVal Count As Integer, ByVal Height As Single) As Integer
            AddRows = tblAddRows(m_Table, Count, Height)
        End Function

        Public Sub ClearColumn(ByVal Col As Integer, ByVal Types As TDeleteContent)
            tblClearColumn(m_Table, Col, Types)
        End Sub

        Public Sub ClearContent(ByVal Types As TDeleteContent)
            tblClearContent(m_Table, Types)
        End Sub

        Public Sub ClearRow(ByVal Row As Integer, ByVal Types As TDeleteContent)
            tblClearRow(m_Table, Row, Types)
        End Sub

        Public Sub DeleteCol(ByVal Col As Integer)
            tblDeleteCol(m_Table, Col)
        End Sub

        Public Sub DeleteRow(ByVal Row As Integer)
            tblDeleteRow(m_Table, Row)
        End Sub

        Public Sub DeleteRows()
            tblDeleteRows(m_Table)
        End Sub

        Public Function DrawTable(ByVal x As Single, ByVal y As Single, ByVal MaxHeight As Single) As Single
            ' This is normally not required but we need to make sure that .Net doesn't delete the CPDF class before the table class is released
            tblSetPDFInstance(m_Table, m_PDF.GetPDFInstance())
            Return tblDrawTable(m_Table, x, y, MaxHeight)
        End Function

        Public Function GetFirstRow() As Integer
            Return tblGetFirstRow(m_Table)
        End Function

        Public Function GetFlags(ByVal Row As Integer, ByVal Col As Integer) As Integer
            Return tblGetFlags(m_Table, Row, Col)
        End Function

        Protected Function GetInstancePtr() As IntPtr
            Return m_Table
        End Function

        Public Function GetNextHeight(ByVal MaxHeight As Single, ByRef NextRow As Integer) As Single
            Return tblGetNextHeight(m_Table, MaxHeight, NextRow)
        End Function

        Public Function GetNextRow() As Integer
            Return tblGetNextRow(m_Table)
        End Function

        Public Function GetNumCols() As Integer
            Return tblGetNumCols(m_Table)
        End Function

        Public Function GetNumRows() As Integer
            Return tblGetNumRows(m_Table)
        End Function

        Public Function GetTableHeight() As Single
            Return tblGetTableHeight(m_Table)
        End Function

        Public Function GetTableWidth() As Single
            Return tblGetTableWidth(m_Table)
        End Function

        Public Function HaveMore() As Boolean
            Return CBool(tblHaveMore(m_Table))
        End Function

        Public Function SetBoxProperty(ByVal Row As Integer, ByVal Col As Integer, ByVal PropType As TTableBoxProperty, ByVal Left As Single, ByVal Right As Single, ByVal Top As Single, ByVal Bottom As Single) As Boolean
            Return CBool(tblSetBoxProperty(m_Table, Row, Col, PropType, Left, Right, Top, Bottom))
        End Function

        Public Function SetCellImage(ByVal Row As Integer, ByVal Col As Integer, ByVal ForeGround As Boolean, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal Width As Single, ByVal Height As Single, ByVal Image As String, ByVal Index As Integer) As Boolean
            Return CBool(tblSetCellImage(m_Table, Row, Col, Convert.ToInt32(ForeGround), HAlign, VAlign, Width, Height, Image, Index))
        End Function

        Public Function SetCellImageEx(ByVal Row As Integer, ByVal Col As Integer, ByVal ForeGround As Boolean, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal Width As Single, ByVal Height As Single, ByRef Buffer() As Byte, ByVal Index As Integer) As Boolean
            Return CBool(tblSetCellImageEx(m_Table, Row, Col, Convert.ToInt32(ForeGround), HAlign, VAlign, Width, Height, Buffer, Buffer.Length, Index))
        End Function

        Public Function SetCellOrientation(ByVal Row As Integer, ByVal Col As Integer, ByVal Orientation As Integer) As Boolean
            Return CBool(tblSetCellOrientation(m_Table, Row, Col, Orientation))
        End Function

        Public Function SetCellTable(ByVal Row As Integer, ByVal Col As Integer, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal SubTable As CPDFTable) As Boolean
            ' We store a reference of the sub table in an array to avoid the deletion. This ensures that the pointer of the sub table stays valid for the entire lifetime of the table that holds a reference of it.
            If m_SubTables Is Nothing Then
                m_SubTables = New ArrayList()
                m_SubTables.Add(SubTable)
            ElseIf m_SubTables.IndexOf(SubTable) < 0 Then
                m_SubTables.Add(SubTable)
            End If
            Return CBool(tblSetCellTable(m_Table, Row, Col, HAlign, VAlign, SubTable.GetInstancePtr()))
        End Function

        Public Function SetCellTemplate(ByVal Row As Integer, ByVal Col As Integer, ByVal ForeGround As Integer, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal TmplHandle As Integer, ByVal Width As Single, ByVal Height As Single) As Boolean
            Return CBool(tblSetCellTemplate(m_Table, Row, Col, ForeGround, HAlign, VAlign, TmplHandle, Width, Height))
        End Function

        Public Function SetCellText(ByVal Row As Integer, ByVal Col As Integer, ByVal HAlign As TTextAlign, ByVal VAlign As TCellAlign, ByVal Text As String) As Boolean
            Return CBool(tblSetCellTextW(m_Table, Row, Col, HAlign, VAlign, Text, Text.Length))
        End Function

        Public Function SetCellTextA(ByVal Row As Integer, ByVal Col As Integer, ByVal HAlign As TTextAlign, ByVal VAlign As TCellAlign, ByVal Text As String) As Boolean
            Return CBool(tblSetCellTextA(m_Table, Row, Col, HAlign, VAlign, Text, Text.Length))
        End Function

        Public Function SetCellTextW(ByVal Row As Integer, ByVal Col As Integer, ByVal HAlign As TTextAlign, ByVal VAlign As TCellAlign, ByVal Text As String) As Boolean
            Return CBool(tblSetCellTextW(m_Table, Row, Col, HAlign, VAlign, Text, Text.Length))
        End Function

        Public Function SetColor(ByVal Row As Integer, ByVal Col As Integer, ByVal ClrType As TTableColor, ByVal CS As TPDFColorSpace, ByVal Color As Integer) As Boolean
            Return CBool(tblSetColor(m_Table, Row, Col, ClrType, CS, Color))
        End Function

        Public Function SetColorEx(ByVal Row As Integer, ByVal Col As Integer, ByVal ClrType As TTableColor, ByVal Color() As Single, ByVal CS As TExtColorSpace, ByVal Handle As Integer) As Boolean
            Return CBool(tblSetColorEx(m_Table, Row, Col, ClrType, Color, Color.Length, CS, Handle))
        End Function

        Public Function SetColWidth(ByVal Col As Integer, ByVal Width As Single, ByVal ExtTable As Boolean) As Boolean
            Return CBool(tblSetColWidth(m_Table, Col, Width, Convert.ToInt32(ExtTable)))
        End Function

        Public Function SetFlags(ByVal Row As Integer, ByVal Col As Integer, ByVal Flags As TTableFlags) As Boolean
            Return CBool(tblSetFlags(m_Table, Row, Col, Flags))
        End Function

        Public Function SetFont(ByVal Row As Integer, ByVal Col As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Embed As Boolean, ByVal CP As TCodepage) As Boolean
            Return CBool(tblSetFont(m_Table, Row, Col, Name, Style, Convert.ToInt32(Embed), CP))
        End Function

        Public Function SetFontSelMode(ByVal Row As Integer, ByVal Col As Integer, ByVal Mode As TFontSelMode) As Boolean
            Return CBool(tblSetFontSelMode(m_Table, Row, Col, Mode))
        End Function

        Public Function SetFontSize(ByVal Row As Integer, ByVal Col As Integer, ByVal Value As Single) As Boolean
            Return CBool(tblSetFontSize(m_Table, Row, Col, Value))
        End Function

        Public Function SetGridWidth(ByVal Horz As Single, ByVal Vert As Single) As Boolean
            Return CBool(tblSetGridWidth(m_Table, Horz, Vert))
        End Function

        Public Function SetRowHeight(ByVal Row As Integer, ByVal Value As Single) As Boolean
            Return CBool(tblSetRowHeight(m_Table, Row, Value))
        End Function

        Public Sub SetTableWidth(ByVal Value As Single, ByVal AdjustType As TColumnAdjust, ByVal MinColWidth As Single)
            tblSetTableWidth(m_Table, Value, AdjustType, MinColWidth)
        End Sub

        Private Declare Function tblAddColumn Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Left As Integer, ByVal Width As Single) As Integer
        Private Declare Function tblAddRow Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Height As Single) As Integer
        Private Declare Function tblAddRows Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Count As Integer, ByVal Height As Single) As Integer
        Private Declare Sub tblClearColumn Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Col As Integer, ByVal Types As TDeleteContent)
        Private Declare Sub tblClearContent Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Types As TDeleteContent)
        Private Declare Sub tblClearRow Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Types As TDeleteContent)
        Private Declare Function tblCreateTable Lib "dynapdf.dll" (ByVal IPDF As IntPtr, ByVal AllocRows As Integer, ByVal NumCols As Integer, ByVal Width As Single, ByVal DefRowHeight As Single) As IntPtr
        Private Declare Sub tblDeleteCol Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Col As Integer)
        Private Declare Sub tblDeleteRow Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer)
        Private Declare Sub tblDeleteRows Lib "dynapdf.dll" (ByVal Table As IntPtr)
        Private Declare Sub tblDeleteTable Lib "dynapdf.dll" (ByRef Table As IntPtr)
        Private Declare Function tblDrawTable Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal x As Single, ByVal y As Single, ByVal MaxHeight As Single) As Single
        Private Declare Function tblGetFirstRow Lib "dynapdf.dll" (ByVal Table As IntPtr) As Integer
        Private Declare Function tblGetFlags Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer) As Integer
        Private Declare Function tblGetNextHeight Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal MaxHeight As Single, ByRef NextRow As Integer) As Single
        Private Declare Function tblGetNextRow Lib "dynapdf.dll" (ByVal Table As IntPtr) As Integer
        Private Declare Function tblGetNumCols Lib "dynapdf.dll" (ByVal Table As IntPtr) As Integer
        Private Declare Function tblGetNumRows Lib "dynapdf.dll" (ByVal Table As IntPtr) As Integer
        Private Declare Function tblGetTableHeight Lib "dynapdf.dll" (ByVal Table As IntPtr) As Single
        Private Declare Function tblGetTableWidth Lib "dynapdf.dll" (ByVal Table As IntPtr) As Single
        Private Declare Function tblHaveMore Lib "dynapdf.dll" (ByVal Table As IntPtr) As Integer
        Private Declare Function tblSetBoxProperty Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal PropType As TTableBoxProperty, ByVal Left As Single, ByVal Right As Single, ByVal Top As Single, ByVal Bottom As Single) As Integer
        Private Declare Auto Function tblSetCellImage Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal ForeGround As Integer, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal Width As Single, ByVal Height As Single, ByVal Image As String, ByVal Index As Integer) As Integer
        Private Declare Function tblSetCellImageEx Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal ForeGround As Integer, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal Width As Single, ByVal Height As Single, ByVal Buffer() As Byte, ByVal BufSize As Integer, ByVal Index As Integer) As Integer
        Private Declare Function tblSetCellOrientation Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal Orientation As Integer) As Integer
        Private Declare Function tblSetCellTable Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal SubTable As IntPtr) As Integer
        Private Declare Function tblSetCellTemplate Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal ForeGround As Integer, ByVal HAlign As TCellAlign, ByVal VAlign As TCellAlign, ByVal TmplHandle As Integer, ByVal Width As Single, ByVal Height As Single) As Integer
        Private Declare Ansi Function tblSetCellTextA Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal Align As TTextAlign, ByVal VAlign As TCellAlign, ByVal Text As String, ByVal sLen As Integer) As Integer
        Private Declare Unicode Function tblSetCellTextW Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal Align As TTextAlign, ByVal VAlign As TCellAlign, ByVal Text As String, ByVal sLen As Integer) As Integer
        Private Declare Function tblSetColor Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal ClrType As TTableColor, ByVal CS As TPDFColorSpace, ByVal Color As Integer) As Integer
        Private Declare Function tblSetColorEx Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal ClrType As TTableColor, ByVal Color() As Single, ByVal NumComps As Integer, ByVal CS As TExtColorSpace, ByVal Handle As Integer) As Integer
        Private Declare Function tblSetColWidth Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Col As Integer, ByVal Width As Single, ByVal ExtTable As Integer) As Integer
        Private Declare Function tblSetFlags Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal Flags As TTableFlags) As Integer
        Private Declare Auto Function tblSetFont Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal Name As String, ByVal Style As TFStyle, ByVal Embed As Integer, ByVal CP As TCodepage) As Integer
        Private Declare Function tblSetFontSelMode Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal Mode As TFontSelMode) As Integer
        Private Declare Function tblSetFontSize Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Col As Integer, ByVal Value As Single) As Integer
        Private Declare Function tblSetGridWidth Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Horz As Single, ByVal Vert As Single) As Integer
        Private Declare Sub tblSetPDFInstance Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal IPDF As IntPtr)
        Private Declare Function tblSetRowHeight Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Row As Integer, ByVal Value As Single) As Integer
        Private Declare Sub tblSetTableWidth Lib "dynapdf.dll" (ByVal Table As IntPtr, ByVal Value As Single, ByVal AdjustType As TColumnAdjust, ByVal MinColWidth As Single)
    End Class

End Namespace