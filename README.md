# VideoThumbnails
Using Media Foundation and GDI+ to create a thumbnail of any point of a video

![image](https://github.com/fafalone/VideoThumbnails/assets/7834493/4ac193c5-291c-4fd4-8139-640b5e2aa2a1)

This is a modified version of -Franky-'s [VBC_MF_GDI+_VideoThumbnail project](https://www.activevb.de/cgi-bin/upload/upload.pl). He's a big fan of using `DispCallFunc` over typelib-defined interfaces, and I'm the opposite. This project shows converting the `DispCallFunc` calls into use of WinDevLib interfaces, and letting the APIs take over to easily also add 64bit compatibility. I've also made the thumbnail size selectable, displayed the video as Form-sized rather than thumbnail sized, and translated the string (with Google Translate, so apologies if inaccurate). Also simplified the propvariant use and time calculations by using tB's `LongLong` instead of `Currency`. 

The project itself is great; it shows how to open a video with the Media Foundation interfaces, and use a horizontal scrollbar to advance it little by little, then grab a thumbnail (or full sized frame grab) of the current point in time, rather than just the first frame like normal thumbnails. GDI+ is then used to save it to a JPG.

Overall, a good intro to Media Foundation.

(Note: There's currently a bug in tB that prevents -Franky-'s original, unmodified version from working in tB. It does work in VB6).

Example:

```vba
            If Invoke(pIMFAttributes, SetUINT32, VarPtr(Str2Guid( _
                MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING)), _
                API_True) = S_OK Then
```

easily becomes `pIMFAttributes.SetUINT32 MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, API_True` which I think is far more readable.


Thanks to -Franky- for his excellent work on the original project and for letting me post this update!
