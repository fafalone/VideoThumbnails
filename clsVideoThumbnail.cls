VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideoThumbnail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6", "Yes"
Attribute VB_Ext_KEY = "Top_Level", "Yes"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 12/2020
' Modified for WinDevLib/twinBASIC x86/x64 by Jon Johnson, 2024 May 23

'Removed definitions taken over by x64 compatible ones in WinDevLib
'Removed procedures no longer required with interfaces.

Option Explicit

' ----==== Const ====----
 
Private Const API_True As Long = &H1&
 
 
Private Const GdiPlusVersion As Long = 1&

 
Private Const mimeJPEG As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
 
Private Const GUID_NULL As String = "{00000000-0000-0000-0000-000000000000}"
Private Const IPictureIID As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"

Private Const MF_VERSION As Long = &H1&
 



Private Type FRAMESIZE
    Hight As Long
    Width As Long
End Type

' ----==== Variablen ====----
Private m_boolInit As Boolean
Private m_boolLoaded As Boolean
Private m_boolSeekable As Boolean
Private m_pIMFSourceReader As IMFSourceReader
Private m_lngToken As LongPtr
Private m_lngVideoWidth As Long
Private m_lngVideoHeight As Long
Private m_curVideoDuration As LongLong

Private Sub Class_Initialize()
    
    Dim tGDIPlusStartupInput As GdiplusStartupInput
    Dim tGdiplusStartupOutput As GdiplusStartupOutput
    
    m_boolInit = False
    m_boolLoaded = False
    m_boolSeekable = False
    
    ' Media Foundation starten
    If MFStartup(MF_VERSION, MFSTARTUP_FULL) = S_OK Then
    
        tGDIPlusStartupInput.GdiplusVersion = GdiPlusVersion
    
        ' Gdi+ starten
        If GdiplusStartup(m_lngToken, tGDIPlusStartupInput, _
            VarPtr(tGdiplusStartupOutput)) = Ok Then
    
            ' alles initialisiert
            m_boolInit = True
            
        End If
    
    End If

End Sub

Private Sub Class_Terminate()

    ' wenn alles initialisiert ist
    If m_boolInit = True Then
    
        ' If m_pIMFSourceReader <> 0& Then
        
        '     ' IMFSourceReader freigeben
        '     Call Release(m_pIMFSourceReader)
            
        ' End If
        If (m_pIMFSourceReader Is Nothing) = False Then
            Set m_pIMFSourceReader = Nothing
        End If
        If m_lngToken <> 0& Then
        
            ' Gdi+ beenden
            Call GdiplusShutdown(m_lngToken)
            
        End If

        ' Media Foundation beenden
        Call MFShutdown
        
    End If

End Sub
                         
Public Property Let SetVideoFile(ByVal sVideoFile As String)
    On Error Resume Next
    'Dim pIMFAttributes As Long
    Dim pIMFAttributes As IMFAttributes
    ' wenn alles initialisiert ist
    If m_boolInit = True Then
    
        ' ist noch ein IMFSourceReader vorhanden
        If (m_pIMFSourceReader Is Nothing) = False Then
        
            ' IMFSourceReader freigeben
            Set m_pIMFSourceReader = Nothing
            
            m_boolLoaded = False
            
        End If
        
        ' IMFAttributes erstellen
        If MFCreateAttributes(pIMFAttributes, 1&) = S_OK Then
        
            ' Videobearbeitung aktivieren
            ' If Invoke(pIMFAttributes, SetUINT32, VarPtr(Str2Guid( _
            '     MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING)), _
            '     API_True) = S_OK Then
            pIMFAttributes.SetUINT32 MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, API_True
            If Err.LastHresult = S_OK Then
                ' IMFSourceReader von Datei erstellen
                If MFCreateSourceReaderFromURL(StrPtr(sVideoFile), _
                    pIMFAttributes, m_pIMFSourceReader) = S_OK Then
                    
                    ' VideoStream selektieren und auf das
                    ' Ausgabevideoformat RGB32 einstellen
                    If SelectStream = True Then
                    
                        ' Dimensionen des Videos ermitteln
                        If GetDimension(m_lngVideoWidth, m_lngVideoHeight) = True Then
                    
                            ' Spielzeit des Videos ermitteln
                            If GetDuration(m_curVideoDuration) = True Then
                            
                                ' kann ein Seek auf das Video angewendet werden
                                m_boolSeekable = GetSeekable
                        
                                ' Video geladen und erforderliche
                                ' Parameter ermittelt
                                m_boolLoaded = True
                                
                            End If
                        
                        End If
                    
                    End If
                    
                End If
                
            End If
            
            ' IMFAttributes freigeben
            Set pIMFAttributes = Nothing
            
        End If
        
    End If
    
End Property

' gibt die Spielzeit des Videos zurück
Public Property Get GetVideoDuration() As LongLong

    GetVideoDuration = m_curVideoDuration
    
End Property

' gibt die Höhe des Videos zurück
Public Property Get GetVideoHeight() As Long

    GetVideoHeight = m_lngVideoHeight
    
End Property

' gibt die Breite des Videos zurück
Public Property Get GetVideoWidth() As Long

    GetVideoWidth = m_lngVideoWidth
    
End Property

' gibt zurück ob alles vorbereitet ist
Public Property Get IsVideoLoaded() As Boolean

    IsVideoLoaded = m_boolLoaded
    
End Property

' gibt zurück ob man im Video zu einer anderen
' Position gehen kann
Public Property Get IsVideoSeekable() As Boolean

    IsVideoSeekable = m_boolSeekable
    
End Property

' neue Position im Video setzen
Public Property Let SetVideoPos(ByVal curPosition As LongLong)
    On Error Resume Next
    Dim varRet As Variant
    ' Dim tPROPVARIANT As PROPVARIANT
    
    ' ist alles vorbereitet
    If m_boolLoaded = True Then
    
        ' kann eine neue Position gesetzt werden
        If m_boolSeekable = True Then
        
            ' muss gleich oder größer 0 sein
            If curPosition >= 0 Then
                
                ' muss gleich oder kleiner der maximalen Spielzeit sein
                If curPosition <= m_curVideoDuration Then
    
                    ' neue Position setzen
                    ' If Invoke(m_pIMFSourceReader, SetCurrentPosition, _
                    '     VarPtr(Str2Guid(GUID_NULL)), _
                    '     VarPtr(Currency2PropVar(curPosition))) = S_OK Then
                    m_pIMFSourceReader.SetCurrentPosition UUID_NULL, CVar(curPosition)
                    If Err.LastHresult = S_OK Then
                    End If
                
                End If
            
            End If
            
        End If
        
    End If
    
End Property

' gibt die aktuelle Position im Video zurück
Public Property Get GetVideoPos() As LongLong
    On Error Resume Next
    Dim lngStreamIndex As Long
    Dim curTimeStamp As LongLong
    Dim curPos As LongLong
    Dim pIMFSample As IMFSample
    Dim eStreamFlags As MF_SOURCE_READER_FLAG
    
    ' ist alles vorbereitet
    If m_boolLoaded = True Then
    
        ' aktuelle Position ermitteln
        ' If Invoke(m_pIMFSourceReader, ReadSample, _
        '     MF_SOURCE_READER_FIRST_VIDEO_STREAM, 0&, _
        '     VarPtr(lngStreamIndex), VarPtr(eStreamFlags), _
        '     VarPtr(curTimeStamp), VarPtr(pIMFSample)) = S_OK Then
        Set pIMFSample = m_pIMFSourceReader.ReadSample(MF_SOURCE_READER_FIRST_VIDEO_STREAM, 0, lngStreamIndex, eStreamFlags, curTimeStamp)
        If Err.LastHresult = S_OK Then
            curPos = curTimeStamp' * 10000
            
            ' IMFSample freigeben
            Set pIMFSample = Nothing
            
        End If
        
    End If
    
    GetVideoPos = curPos
    
End Property

' gibt von der aktuelle Position im Video ein Bild zurück
Public Function GetVideoThumbnail(Optional ByVal lMaxWidth As Long = 0&) As StdPicture
    'Debug.Print "GetVideoThumbnail(" & lMaxWidth & ")"
    On Error Resume Next
    Dim bolOK As Boolean
    Dim hThumbnail As LongPtr
    Dim lngCurLength As Long
    Dim lngMaxLength As Long
    Dim lngStreamIndex As Long
    Dim curTimeStamp As LongLong
    Dim pIMFSample As IMFSample
    Dim pLockBuffer As LongPtr
    Dim pIMFMediaBuffer As IMFMediaBuffer
    Dim eStreamFlags As MF_SOURCE_READER_FLAG
    Dim oIPicture As IPicture
    Dim tPictDesc As PICTDESC

    bolOK = False

    ' ist alles vorbereitet
    If m_boolLoaded = True Then
    
        ' ein IMFSample auslesen
        ' If Invoke(m_pIMFSourceReader, ReadSample, _
        '     MF_SOURCE_READER_FIRST_VIDEO_STREAM, 0&, _
        '     VarPtr(lngStreamIndex), VarPtr(eStreamFlags), _
        '     VarPtr(curTimeStamp), VarPtr(pIMFSample)) = S_OK Then
        Set pIMFSample = m_pIMFSourceReader.ReadSample(MF_SOURCE_READER_FIRST_VIDEO_STREAM, 0, lngStreamIndex, eStreamFlags, curTimeStamp)
        If Err.LastHresult = S_OK Then
            ' nur wenn das Ende des VideoStreams nicht erreicht ist
            If CBool(eStreamFlags And MF_SOURCE_READERF_ENDOFSTREAM) = False Then
            
                ' konvertiert mehrere IMFMediaBuffer im IMFSample zu
                ' einem IMFMediaBuffer und gibt diesen zurück
                ' If Invoke(pIMFSample, ConvertToContiguousBuffer, _
                '     VarPtr(pIMFMediaBuffer)) = S_OK Then
                Set pIMFMediaBuffer = pIMFSample.ConvertToContiguousBuffer()
                If Err.LastHresult = S_OK Then
                    ' Pointer auf die Daten im IMFMediaBuffer holen
                    ' If Invoke(pIMFMediaBuffer, Lock_Buffer, VarPtr(pLockBuffer), _
                    '     VarPtr(lngMaxLength), VarPtr(lngCurLength)) = S_OK Then
                    pIMFMediaBuffer.Lock pLockBuffer, lngMaxLength, lngCurLength
                    If Err.LastHresult = S_OK Then
                        ' BitmapHandle aus den Daten im IMFMediaBuffer
                        hThumbnail = GetThumbnailHandle(pLockBuffer, lMaxWidth)
                    
                        ' ist ein Handle vorhanden
                        If hThumbnail <> 0& Then
                    
                            With tPictDesc
                                .cbSizeofstruct = LenB(tPictDesc)
                                .picType = vbPicTypeBitmap
                                .hImage = hThumbnail
                            End With
    
                            ' ein Standard-Picture aus dem Handle erstellen
                            If OleCreatePictureIndirect(tPictDesc, IID_IPictureDisp, _
                                API_True, oIPicture) = S_OK Then
                            
                                bolOK = True
        
                            End If
                    
                        End If
                    
                        ' Lock auf den Pointer freigeben
                        pIMFMediaBuffer.Unlock
                    
                    End If
                
                    ' IMFMediaBuffer freigeben
                    Set pIMFMediaBuffer = Nothing
                
                End If
            
            End If
            
            ' IMFSample freigeben
            Set pIMFSample = Nothing
            
        End If
        
    End If
    
    If bolOK = True Then
    
        ' Standard-Picture zurück geben
        Set GetVideoThumbnail = oIPicture
    
    End If
    
End Function
                         
Public Function SaveVideoThumbnailAsJPG(ByVal hPicture As LongPtr, _
    ByVal sJpgFile As String, Optional ByVal Quality As Long = 85) As Boolean

    Dim bolRet As Boolean
    Dim pBitmap As LongPtr
    Dim tEncoderParameters As EncoderParameters

    bolRet = False

    ' ist ein BitmapHandle und ein Dateiname vorhanden
    If hPicture <> 0& And Len(sJpgFile) <> 0& Then
        
        ' Parameter begrenzen
        If Quality > 100 Then Quality = 100
        If Quality < 0 Then Quality = 0

        ' vom BitmapHandle ein Gdi+ Bitmap erstellen
        If GdipCreateBitmapFromHBITMAP(hPicture, 0&, pBitmap) = Ok Then
            
            ' Dim cx As Long, cy As Long
            ' GdipGetImageWidth(pBitmap, cx)
            ' GdipGetImageHeight(pBitmap, cy)
            ' Paramter für JPG
            tEncoderParameters.Count = 1
        
            With tEncoderParameters.Parameter(0)
                .Guid = EncoderQuality
                .NumberOfValues = 1
                .Type = EncoderParameterValueTypeLong
                .Value = VarPtr(Quality)
            End With
        
            ' speichert das Gdi+ Bitmap als JPG
            Dim gs As GpStatus
            gs = GdipSaveImageToFile(pBitmap, StrPtr(sJpgFile), _
                        Str2Guid(mimeJPEG), tEncoderParameters)
            If gs = Ok Then
                    
                bolRet = True
            Else
                Debug.Print "GdipSaveImageToFile fail code " & gs
            End If
        
            ' Gdi+ Bitmap freigeben
            Call GdipDisposeImage(pBitmap)
            
        End If
    
    End If

    SaveVideoThumbnailAsJPG = bolRet
    
End Function
                         
' ----==== String -> GUID ====----
Private Function Str2Guid(ByVal str As String) As UUID

    Call CLSIDFromString(str, Str2Guid)
    
End Function

Private Function SelectStream() As Boolean
    On Error Resume Next
    Dim bolRet As Boolean
    Dim pIMFMediaType As IMFMediaType
    
    bolRet = False
    
    ' IMFMediaType erstellen
    If MFCreateMediaType(pIMFMediaType) = S_OK Then
    
        ' MediaType = Video
        ' If Invoke(pIMFMediaType, SetGUID, VarPtr(Str2Guid(MF_MT_MAJOR_TYPE)), _
        '     VarPtr(Str2Guid(MFMediaType_Video))) = S_OK Then
        pIMFMediaType.SetGuid MF_MT_MAJOR_TYPE, MFMediaType_Video
        If Err.LastHresult = S_OK Then
            ' MediaType = VideoFormat RGB32
            ' If Invoke(pIMFMediaType, SetGUID, VarPtr(Str2Guid(MF_MT_SUBTYPE)), _
            '     VarPtr(Str2Guid(MFVideoFormat_RGB32))) = S_OK Then
            pIMFMediaType.SetGuid MF_MT_SUBTYPE, MFVideoFormat_RGB32
            If Err.LastHresult = S_OK Then
                ' MediaType für den ersten VideoStream im IMFSourceReader zuweisen
                ' If Invoke(m_pIMFSourceReader, SetCurrentMediaType, _
                '     MF_SOURCE_READER_FIRST_VIDEO_STREAM, 0&, pIMFMediaType) = S_OK Then
                m_pIMFSourceReader.SetCurrentMediaType MF_SOURCE_READER_FIRST_VIDEO_STREAM, 0, pIMFMediaType
                If Err.LastHresult = S_OK Then
                    ' ersten VideoStream selektieren
                    ' If Invoke(m_pIMFSourceReader, SetStreamSelection, _
                    '     MF_SOURCE_READER_FIRST_VIDEO_STREAM, API_True) = S_OK Then
                    m_pIMFSourceReader.SetStreamSelection MF_SOURCE_READER_FIRST_VIDEO_STREAM, API_True
                    If Err.LastHresult = S_OK Then
                        bolRet = True
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
        Set pIMFMediaType = Nothing
        
    End If
    
    SelectStream = bolRet
    
End Function

Private Function GetDimension(ByRef lWidth As Long, ByRef lHeight As _
    Long) As Boolean
    On Error Resume Next
    Dim bolRet As Boolean
    Dim pIMFMediaType As IMFMediaType
    Dim tGUID As UUID
    Dim tFRAMESIZE As FRAMESIZE
    Dim llFrameSize As LongLong
    bolRet = False
    
    ' IMFMediaType erstellen
    ' If Invoke(m_pIMFSourceReader, GetCurrentMediaType, _
    '     MF_SOURCE_READER_FIRST_VIDEO_STREAM, _
    '     VarPtr(pIMFMediaType)) = S_OK Then
    Set pIMFMediaType = m_pIMFSourceReader.GetCurrentMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM)
    If Err.LastHresult = S_OK Then
        ' FrameSize ermitteln
        ' If Invoke(pIMFMediaType, GetUINT64, _
        '     VarPtr(Str2Guid(MF_MT_FRAME_SIZE)), _
        '     VarPtr(tFRAMESIZE)) = S_OK Then
        llFrameSize = pIMFMediaType.GetUInt64(MF_MT_FRAME_SIZE)
        If Err.LastHresult = S_OK Then
            CopyMemory tFRAMESIZE, llFrameSize, 8
            lWidth = tFRAMESIZE.Width
            lHeight = tFRAMESIZE.Hight
            
            bolRet = True
            
        End If
        
        ' IMFMediaType freigeben
        Set pIMFMediaType = Nothing
        
    End If
    
    GetDimension = bolRet
    
End Function

Private Function GetDuration(ByRef cDuration As LongLong) As Boolean

    Dim bRet As Boolean
    Dim pIMFSourceReader As Long
    'Dim tPROPVARIANT As PROPVARIANT
    Dim tPROPVARIANT As Variant
    bRet = False
    
    ' SourceDuration ermitteln
    ' If Invoke(m_pIMFSourceReader, GetPresentationAttribute, _
    '     MF_SOURCE_READER_MEDIASOURCE, VarPtr(Str2Guid(MF_PD_DURATION)), _
    '     VarPtr(tPROPVARIANT)) = S_OK Then
    tPROPVARIANT = m_pIMFSourceReader.GetPresentationAttribute(MF_SOURCE_READER_MEDIASOURCE, MF_PD_DURATION)
    If Err.LastHresult = S_OK Then
        ' PROPVARIANT zu Currency konvertieren
        If PropVariantToInt64(tPROPVARIANT, cDuration) = S_OK Then
        
            ' nur wenn cDuration > 0
            If cDuration > 0 Then
            
                'cDuration = cDuration * 10000
                
                bRet = True
                
            End If
            
        End If
        
        ' PROPVARIANT freigeben
        Call PropVariantClear(tPROPVARIANT)
        
    End If
    
    GetDuration = bRet
    
End Function

Private Function GetSeekable() As Boolean
    On Error Resume Next
    Dim bolRet As Boolean
    Dim varRet As Variant
    Dim tPROPVARIANT As Variant
    
    bolRet = False
    
    ' PresentationAttribute ermitteln
    ' If Invoke(m_pIMFSourceReader, GetPresentationAttribute, _
    '     MF_SOURCE_READER_MEDIASOURCE, VarPtr(Str2Guid( _
    '     MF_SOURCE_READER_MEDIASOURCE_CHARACTERISTICS)), _
    '     VarPtr(tPROPVARIANT)) = S_OK Then
    tPROPVARIANT = m_pIMFSourceReader.GetPresentationAttribute(MF_SOURCE_READER_MEDIASOURCE, MF_SOURCE_READER_MEDIASOURCE_CHARACTERISTICS)
    If Err.LastHresult = S_OK Then
        ' ist ermittelte der Wert ein VT_UI4
        If VarType(tPROPVARIANT) = VARENUM.VT_UI4 Then
        
            ' PROPVARIANT zu VT_UI4 konvertieren
            If PropVariantChangeType(varRet, tPROPVARIANT, _
                0&, VT_I4) = S_OK Then
                
                ' ist der Flag MFMEDIASOURCE_CAN_SEEK vorhanden
                If CBool(CLng(varRet) And _
                    MFMEDIASOURCE_CAN_SEEK) = True Then
                    
                    bolRet = True
                    
                End If
                
            End If
            
        End If
        
    End If
    
    GetSeekable = bolRet
    
End Function

Private Function GetThumbnailHandle(ByVal pData As LongPtr, _
    ByVal lMaxWidth As Long) As LongPtr
    
    Dim bolOK As Boolean
    Dim hBitmapNew As LongPtr
    Dim hThumbnail As LongPtr
    Dim pGraphics As LongPtr
    Dim pBitmapOrg As LongPtr
    Dim pBitmapNew As LongPtr
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim dblScale As Double
    Dim dblRatioX As Double
    Dim dblRatioY As Double
    
    lngWidth = m_lngVideoWidth
    lngHeight = m_lngVideoHeight
    
    bolOK = False
    
    ' Gdi+ 32bppRGB Bitmap in den Dimensionen des
    ' Videos und aus den Bilddaten erstellen
    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, lngWidth * 4, _
        PixelFormat32bppRGB, ByVal pData, pBitmapOrg) = Ok Then
        
        ' Skalierung berechnen
        If lMaxWidth = 0& Then
                
                dblRatioX = 1
                dblRatioY = 1
        Else
        
                dblRatioX = lMaxWidth / lngWidth
                dblRatioY = lMaxWidth / lngHeight
        End If
        
        If dblRatioX > dblRatioY Then dblScale = dblRatioY Else dblScale = dblRatioX

        lngWidth = CLng(lngWidth * dblScale)
        lngHeight = CLng(lngHeight * dblScale)
        
        ' neue Gdi+ 32bppRGB Bitmap in der sklaierten Größe erstellen
        If GdipCreateBitmapFromScan0(lngWidth, lngHeight, 0&, _
            PixelFormat32bppRGB, ByVal 0&, pBitmapNew) = Ok Then
            
            ' Gdi+ Graphics von neuer Gdi+ Bitmap erstellen
            If GdipGetImageGraphicsContext(pBitmapNew, pGraphics) = Ok Then
            
                ' Gdi+ Bitmap vom Video in das skaliert
                ' in das neue Gdi+ Bitmap zeichnen
                If GdipDrawImageRectI(pGraphics, pBitmapOrg, 0, 0, _
                    lngWidth, lngHeight) = Ok Then
                    
                    ' zeichnen war erfolgreich
                    bolOK = True
                    
                End If
                
                ' Gdi+ Graphics freigeben
                Call GdipDeleteGraphics(pGraphics)
                
            End If
            
            ' war das zeichnen erfolgreich
            If bolOK = True Then
            
                bolOK = False
                
                ' Gdi32 BitmapHandle vom neuen Gdi+ Bitmap ermitteln
                If GdipCreateHBITMAPFromBitmap(pBitmapNew, _
                    hBitmapNew, 0&) = Ok Then
                    
                    hThumbnail = hBitmapNew
                    
                End If
                
            End If
            
            ' neues Gdi+ Bitmap freigeben
            Call GdipDisposeImage(pBitmapNew)
            
        End If
        
        ' Gdi+ Bitmap vom Video freigeben
        Call GdipDisposeImage(pBitmapOrg)
        
    End If
    
    ' Gdi32 BitmapHandle zurück geben
    GetThumbnailHandle = hThumbnail
    
End Function

Public Function Time2String(ByVal curNanoSec As LongLong, Optional ByVal _
    strTimeFormat As String = "hh:mm:ss") As String

    On Error Resume Next

    Dim lngHour As Long
    Dim lngMinu As Long
    Dim dblSeco As Double

    dblSeco = CDbl(curNanoSec / 10000000)

    lngHour = dblSeco \ 3600
    dblSeco = dblSeco Mod 3600
    lngMinu = dblSeco \ 60
    dblSeco = dblSeco Mod 60

    Time2String = Format$(TimeSerial(lngHour, lngMinu, dblSeco), strTimeFormat)

End Function