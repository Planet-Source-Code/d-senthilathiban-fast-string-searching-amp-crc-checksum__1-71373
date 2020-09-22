Attribute VB_Name = "basListView"
Option Explicit

' Code by Randy Birch

Public Enum LedgerColours
  vbLedgerWhite = &HF9FEFF
  vbLedgerGreen = &HD0FFCC
  vbLedgerYellow = &HE1FAFF
  vbLedgerRed = &HE1E1FF
  vbLedgerGrey = &HE0E0E0
  vbLedgerbeige = &HD9F2F7
  vbLedgerSoftWhite = &HF7F7F7
  vbLedgerPureWhite = &HFFFFFF
  vbLedgerLightBlue = &HFBC0A2
End Enum

Public Enum ImageSizingTypes
   [sizeNone] = 0
   [sizeCheckBox]
   [sizeIcon]
End Enum
''''''this added extra D.Senthilathiban
Public StopSearch As Boolean
''''''''''''''''''''''''''''''''''''''
Public Sub SetListViewLedger(lv As ListView, _
                              Bar1Color As LedgerColours, _
                              Bar2Color As LedgerColours, _
                              nSizingType As ImageSizingTypes)

   Dim iBarHeight  As Long  '/* height of 1 line in the listview
   Dim lBarWidth   As Long  '/* width of listview
   Dim diff        As Long  '/* used in calculations of row height
   Dim twipsy      As Long  '/* variable holding Screen.TwipsPerPicture1elY
   
   iBarHeight = 0
   lBarWidth = 0
   diff = 0
   
   On Local Error GoTo SetListViewColor_Error
   
   twipsy = Screen.TwipsPerPixelY
   
   If lv.View = lvwReport Then
   
     '/* set up the listview properties
      With lv
        .Picture = Nothing  '/* clear picture
        .Refresh
        .Visible = 1
        .PictureAlignment = lvwTile
        lBarWidth = .Width
      End With  ' lv
        
     '/* set up the picture box properties
      With frmFileSearchCRC.Picture1
         .AutoRedraw = False       '/* clear/reset picture
         .Picture = Nothing
         .BackColor = vbWhite
         .Height = 1
         .AutoRedraw = True        '/* assure image draws
         .BorderStyle = vbBSNone   '/* other attributes
         .ScaleMode = vbTwips
         .Top = frmFileSearchCRC.Top - 10000  '/* move it way off screen
         .Width = Screen.Width
         .Visible = False
         .Font = lv.Font           '/* assure Picture1 font matched listview font
         
        '/* match picture box font properties
        '/* with those of listview
         With .Font
            .Bold = lv.Font.Bold
            .Charset = lv.Font.Charset
            .Italic = lv.Font.Italic
            .Name = lv.Font.Name
            .Strikethrough = lv.Font.Strikethrough
            .Underline = lv.Font.Underline
            .Weight = lv.Font.Weight
            .Size = lv.Font.Size
         End With  'Picture1.Font
         
        '/* here we calculate the height of each
        '/* bar in the listview. Several things
        '/*  can affect this height - the use
        '/* of item icons, the size of those icons,
        '/* the use of checkboxes and so on through
        '/* all the permutations.
        '/*
        '/* Shown here is code sufficient to calculate
        '/* this height based on three combinations of
        '/*  data, state icons, and imagelist icons:
        '/*
        '/* 1. text only
        '/* 2. text with checkboxes
        '/* 3. text with icons
        
       '/* used by all sizing routines
         iBarHeight = .TextHeight("W")

         Select Case nSizingType
            Case sizeNone:
              '/* 1. text only
               iBarHeight = iBarHeight + twipsy
               
            Case sizeCheckBox:
              '/* 2. text with checkboxes: add to textheight the
              '/*    difference between 18 Pixels and iBarHeight
              '/*    all calculated initially in Pixels,
              '/*    then converted to twips
               If (iBarHeight \ twipsy) > 18 Then
                  iBarHeight = iBarHeight + twipsy
               Else
                  diff = 18 - (iBarHeight \ twipsy)
                  iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
               End If
               
            Case sizeIcon:
              '/* 3. text with icons: add to textheight the
              '/*    difference between textheight and image
              '/*    height, all calculated initially in Pixels,
              '/*    then converted to twips. Handles 16x16 icons
               'diff = imagelist1.ImageHeight - (iBarHeight \ twipsy)
               'iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
               
         End Select
      
        '/* since we need two-tone bars, the
        '/* picturebox needs to be twice as high
         .Height = iBarHeight * 2
         .Width = lBarWidth
         
        '/* paint the two bars of color and refresh
        '/* Note: The line method does not support
        '/* With/End With blocks
         frmFileSearchCRC.Picture1.Line (0, 0)-(lBarWidth, iBarHeight), Bar1Color, BF
         frmFileSearchCRC.Picture1.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), Bar2Color, BF
      
         .AutoSize = True
         .Refresh
         
      End With  'Picture1
     
     '/* set the lv picture to the
     '/* Picture1 image
     
      lv.Refresh
      lv.Picture = frmFileSearchCRC.Picture1.Image
      
   Else
    
      lv.Picture = Nothing
        
   End If  'lv.View = lvwReport

SetListViewColor_Exit:
On Local Error GoTo 0
Exit Sub
    
SetListViewColor_Error:

  '/* clear the listview's picture and exit
   With lv
      .Picture = Nothing
      .Refresh
   End With
   
   Resume SetListViewColor_Exit
    
End Sub
