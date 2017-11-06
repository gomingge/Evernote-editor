#SingleInstance Force
#NoTrayIcon

{
	#NoEnv						;不检查空变量是否为环境变量
	SetBatchLines, -1			;行之间运行不留时间空隙,默认是有10ms的间隔
	SetKeyDelay, -1, -1			;发送按键不留时间空隙
	SetMouseDelay, -1			;每次鼠标移动或点击后自动的延时=0   
	SetDefaultMouseSpeed, 0		;设置在 Click 和 MouseMove/Click/Drag 中没有指定鼠标速度时使用的速度 = 瞬间移动.
	SetWinDelay, 0
	SetControlDelay, 0
	SendMode Input

	#InstallKeybdHook		;安装键盘和鼠标钩子 像Input和A_PriorKey，都需要钩子
	#InstallMouseHook
	SetTitleMatchMode Regex	;更改进程匹配模式为正则
	#SingleInstance force	;决定当脚本已经运行时是否允许它再次运行。
	#Persistent				;持续运行不退出
	#MaxThreadsPerHotkey 5

    ;Menu, tray, tip, 印象笔记-编辑增强小工具
	;TrayTip, 提示, 印象笔记-编辑增强小工具, , 1
	Sleep, 1000
	TrayTip
  
	;evernote编辑器增强函数
	evernoteEdit(eFoward, eEnd)
	{
		clipboard =
		Send ^c
		ClipWait, 1
		t := WinClip.GetHtml3()
		html = %eFoward%%t%%eEnd%
		WinClip.Clear()
		WinClip.SetHTML(html)
		Sleep, 300
		Send ^v
		Return
	}
	
	;evernote不保留原格式，增强函数
	evernoteEditText(eFoward, eEnd)
	{
		clipboard =
		Send ^c
		ClipWait, 1
		t := WinClip.GetText()
		html = %eFoward%%t%%eEnd%
		WinClip.Clear()
		WinClip.SetHTML(html)
		Sleep, 300
		Send ^v
		Return
	}
	
	;evernote无原文本的插入html增强函数
	evernoteInsertHTML(html)
	{
		clipboard =
		WinClip.SetHTML(html)
		Sleep, 300
		Send ^v
		Return
	}

	;自定义setLink函数
	setLink()
	{
		clipboard = %clipboard%
		homeStr = <a href="
		endStr = ">
		div = %homeStr%%clipboard%%endStr%
		Return div
	}
	WinClip.GetHtml2 := Func("GetHtml2")		; 也可以直接覆盖原来的函数 -> WinClip.GetHtml := Func("GetHtml2")
	WinClip.GetHtml3 := Func("GetHtml_DOM")
	
	;操作HTML DOM，比GetHTML函数更实用
	GetHtml_DOM(this, Encoding := "UTF-8")
	{
		html := this.GetHtml2(Encoding)
		static doc := ComObjCreate("htmlFile")
		doc.Write(html), doc.Close()
		return doc.all.tags("span")[0].InnerHtml
	}

	;WinClip中Get的UTF-8改写，支持中文
	GetHtml2(this, Encoding := "UTF-8")
	{
	  if !( clipSize := this._fromclipboard( clipData ) )
		return ""
	  if !( out_size := this._getFormatData( out_data, clipData, clipSize, "HTML Format" ) )
		return ""
	  return strget( &out_data, out_size, Encoding )
	}

}

#IfWinActive ahk_class (ENSingleNoteView|ENMainFrame)
{		
	;快捷键: 非编辑器部分
	{
		^Space::controlsend, , ^{Space}, A   	;简化格式
		$RButton::		;双击右键高亮
		{
			CountStp := ++CountStp
			SetTimer, TimerPrtSc, -500
			Return
			TimerPrtSc:
				if CountStp = 1 ;只按一次时执行
					SendInput, {RButton}
				if CountStp = 2 ;按两次时...
					SendInput, ^+h
				CountStp := 0 ;最后把记录的变量设置为0,于下次记录.
				Return
		}
	}
	
	;颜色 字体格式等
	{
		;方框环绕
		!f::evernoteEdit("<div style='margin-top: 5px; margin-bottom: 9px; word-wrap: break-word; padding: 8.5px; border-top-left-radius: 4px; border-top-right-radius: 4px; border-bottom-right-radius: 4px; border-bottom-left-radius: 4px; background-color: rgb(245, 245, 245); border: 1px solid rgba(0, 0, 0, 0.148438)'>", "</div></br>")
		;常规标题设置
		!1::evernoteEditText("<h1 style='font-family:黑体;margin: 10px 0 10px;padding: 0;font-weight: 700;cursor: text;position: relative;color:#2f2f2f;font-size:26px;color:#b45b3e;'>","</h1>")
		!2::evernoteEditText("<h2 style='font-family:黑体;margin: 6px 0 10px;padding: 0;font-weight: 700;cursor: text;position: relative;color:#2f2f2f;font-size:22px;color:#b45b3e;'>","</h1>")
		!3::evernoteEditText("<h3 style='font-family:黑体;margin: 6px 0 10px;padding: 0;font-weight: 700;cursor: text;position: relative;color:#2f2f2f;font-size:18px;color:#b45b3e;'>","</h1>")
		;引用
		!s::evernoteEdit("<div style='margin:0.8em 0px; line-height:1.5em; border-left-width:5px; border-left-style:solid; border-left-color:rgb(127, 192, 66); padding-left:1.5em; '>", "</div>")

		;为制定文本设置超级链接
		!g::evernoteEditText(setLink(),"</a>")
		;超级标题
		!a::evernoteEdit("<div style='margin:1px 0px; color:rgb(255, 255, 255); background-color:#8BAAD0; border-top-left-radius:5px; border-top-right-radius:5px; border-bottom-right-radius:5px; border-bottom-left-radius:5px; text-align:center;'><b>", "</b></div></br>")
		;虚线
		!d::evernoteEdit("<blockquote style='width: 100%;color: #8b8b8b;margin: 15px auto;padding: 10px;clear: both;border: 2px dashed #ddd;background: #f9f9f9;'>", "</blockquote>")
		;红框框
		!q::evernoteEdit("<div style='color: #c66;background: #ffecea -1px -1px no-repeat;border: 1px solid #ebb1b1;overflow: hidden;margin: 10px 0;padding: 15px 15px 15px 15px;'>", "</div>")
		;黄框框
		!w::evernoteEdit("<div style='color: #ad9948;background: #fff4b9 -1px -1px no-repeat;border: 1px solid #eac946;overflow: hidden;margin: 10px 0;padding: 15px 15px 15px 15px;'>", "</div>" )
		;灰框框
		!e::evernoteEdit("<div style='color: #777;background: #eaeaea -1px -1px no-repeat;border: 1px solid #ccc;overflow: hidden;margin: 10px 0;padding: 15px 15px 15px 15px;'>", "</div>" )
		;绿框框
		!r::evernoteEdit("<div style='color: #7da33c;background: #ecf2d6 -1px -1px no-repeat;border: 1px solid #aac66d;overflow: hidden;margin: 10px 0;padding: 15px 15px 15px 15px;'>", "</div>" )
		
		;v6下用evernoteEditText()回帖，前面都会多一个空格，无解。但删除一下也不麻烦，聊胜于无吧
		;背景色黄色
		^!1::evernoteEditText("<span style='background: #FFFAA5;'>", "</span>")
		;背景色蓝色
		^!2::evernoteEditText("<span style='background: #ADD8E6;'>", "</span>")
		;背景色灰色
		^!3::evernoteEditText("<span style='background: #D3D3D3;'>", "</span>")
		;背景色绿色
		^!4::evernoteEditText("<span style='background: #90EE90;'>", "</span>")
		
		;字体红色
		!+1::evernoteEditText("<span style='color: #F02E37;'><b>", "</b></span>")
		;字体蓝色
		!+2::evernoteEditText("<span style='color: #3740E6;'><b>", "</b></span>")
		;字体灰色
		!+3::evernoteEditText("<span style='color: #D6D6D6;'>", "</span>")
		;字体绿色
		!+4::evernoteEditText("<span style='color: #0F820F;'><b>", "</b></span>")
		
		;v6版本，鼠标点击方式，实现修改文字颜色
		evernoteMouseChangeColor(r, g, b) {
			CoordMode, Mouse, Screen	;鼠标坐标全屏幕模式，方便鼠标回归原位
			MouseGetPos, xpos, ypos 
			文字:=""
			文字.="|<>52.0000300000000200000000401zzU007k03tw000V0073U002400840008k0000000R0008"
			if 查找文字(929,181,150000,150000,文字,"*147",X,Y,OCR,0.2,0.2)
			{
				CoordMode, Mouse
				Click, %X%, %Y%		;点击颜色按钮
				Y1 := Y + 180
				Click, %X%, %Y1%	;点击更多颜色
			}
			else
			{
				MsgBox, 没有找到颜色选择框,找字模块失败!
			}
			;SendL("M")			;进入更多颜色		
			WinWait, 颜色
			WinMove, 10, 10
			CoordMode, Mouse, Client	;鼠标坐标Client模式
			Click, 116, 333		;进入自定义颜色
			SendInput, {Tab}{Tab}{Tab}
			SendInput %r%{Tab}%g%{Tab}%b%{Tab}{Space}
			Click, 21, 259		;点击设定好自定义颜色
			SendInput, {Tab}{Space}
			CoordMode, Mouse, Screen	;鼠标坐标全屏幕模式
			MouseMove, %xpos%, %ypos%, 0
			return
		}
		
		;字体红色
		^!F1::
			evernoteMouseChangeColor(240, 46, 55)
			SendInput, ^b
			return
		;字体蓝色
		^!F2::
			evernoteMouseChangeColor(55, 64, 230)
			SendInput, ^b
			return
		;字体灰色
		^!F3::
			evernoteMouseChangeColor(214, 214, 214)
			return
		;字体绿色
		^!F4::
			evernoteMouseChangeColor(15, 130, 15)
			SendInput, ^b
			return
		;字体白色
		^!F5::
			evernoteMouseChangeColor(255, 255, 255)
			return
	}
}

#IfWinActive

;~ 下面是库函数
{
class WinClip_base
{
  __Call( aTarget, aParams* ) {
    if ObjHasKey( WinClip_base, aTarget )
      return WinClip_base[ aTarget ].( this, aParams* )
    throw Exception( "Unknown function '" aTarget "' requested from object '" this.__Class "'", -1 )
  }
  
  Err( msg ) {
    throw Exception( this.__Class " : " msg ( A_LastError != 0 ? "`n" this.ErrorFormat( A_LastError ) : "" ), -2 )
  }
  
  ErrorFormat( error_id ) {
    VarSetCapacity(msg,1000,0)
    if !len := DllCall("FormatMessageW"
          ,"UInt",FORMAT_MESSAGE_FROM_SYSTEM := 0x00001000 | FORMAT_MESSAGE_IGNORE_INSERTS := 0x00000200		;dwflags
          ,"Ptr",0		;lpSource
          ,"UInt",error_id	;dwMessageId
          ,"UInt",0			;dwLanguageId
          ,"Ptr",&msg			;lpBuffer
          ,"UInt",500)			;nSize
      return
    return 	strget(&msg,len)
  }
}

class WinClipAPI_base extends WinClip_base
{
  __Get( name ) {
    if !ObjHasKey( this, initialized )
      this.Init()
    else
      throw Exception( "Unknown field '" name "' requested from object '" this.__Class "'", -1 )
  }
}

class WinClipAPI extends WinClip_base
{
  memcopy( dest, src, size ) {
    return DllCall( "msvcrt\memcpy", "ptr", dest, "ptr", src, "uint", size )
  }
  GlobalSize( hObj ) {
    return DllCall( "GlobalSize", "Ptr", hObj )
  }
  GlobalLock( hMem ) {
    return DllCall( "GlobalLock", "Ptr", hMem )
  }
  GlobalUnlock( hMem ) {
    return DllCall( "GlobalUnlock", "Ptr", hMem )
  }
  GlobalAlloc( flags, size ) {
    return DllCall( "GlobalAlloc", "Uint", flags, "Uint", size )
  }
  OpenClipboard() {
    return DllCall( "OpenClipboard", "Ptr", 0 )
  }
  CloseClipboard() {
    return DllCall( "CloseClipboard" )
  }
  SetClipboardData( format, hMem ) {
    return DllCall( "SetClipboardData", "Uint", format, "Ptr", hMem )
  }
  GetClipboardData( format ) {
    return DllCall( "GetClipboardData", "Uint", format ) 
  }
  EmptyClipboard() {
    return DllCall( "EmptyClipboard" )
  }
  EnumClipboardFormats( format ) {
    return DllCall( "EnumClipboardFormats", "UInt", format )
  }
  CountClipboardFormats() {
    return DllCall( "CountClipboardFormats" )
  }
  GetClipboardFormatName( iFormat ) {
    size := VarSetCapacity( bufName, 255*( A_IsUnicode ? 2 : 1 ), 0 )
    DllCall( "GetClipboardFormatName", "Uint", iFormat, "str", bufName, "Uint", size )
    return bufName
  }
  GetEnhMetaFileBits( hemf, ByRef buf ) {
    if !( bufSize := DllCall( "GetEnhMetaFileBits", "Ptr", hemf, "Uint", 0, "Ptr", 0 ) )
      return 0
    VarSetCapacity( buf, bufSize, 0 )
    if !( bytesCopied := DllCall( "GetEnhMetaFileBits", "Ptr", hemf, "Uint", bufSize, "Ptr", &buf ) )
      return 0
    return bytesCopied
  }
  SetEnhMetaFileBits( pBuf, bufSize ) {
    return DllCall( "SetEnhMetaFileBits", "Uint", bufSize, "Ptr", pBuf )
  }
  DeleteEnhMetaFile( hemf ) {
    return DllCall( "DeleteEnhMetaFile", "Ptr", hemf )
  }
  ErrorFormat(error_id) {
    VarSetCapacity(msg,1000,0)
    if !len := DllCall("FormatMessageW"
          ,"UInt",FORMAT_MESSAGE_FROM_SYSTEM := 0x00001000 | FORMAT_MESSAGE_IGNORE_INSERTS := 0x00000200		;dwflags
          ,"Ptr",0		;lpSource
          ,"UInt",error_id	;dwMessageId
          ,"UInt",0			;dwLanguageId
          ,"Ptr",&msg			;lpBuffer
          ,"UInt",500)			;nSize
      return
    return 	strget(&msg,len)
  }
  IsInteger( var ) {
    if var is integer
      return True
    else 
      return False
  }
  LoadDllFunction( file, function ) {
      if !hModule := DllCall( "GetModuleHandleW", "Wstr", file, "UPtr" )
          hModule := DllCall( "LoadLibraryW", "Wstr", file, "UPtr" )
      
      ret := DllCall("GetProcAddress", "Ptr", hModule, "AStr", function, "UPtr")
      return ret
  }
  SendMessage( hWnd, Msg, wParam, lParam ) {
     static SendMessageW

     If not SendMessageW
        SendMessageW := this.LoadDllFunction( "user32.dll", "SendMessageW" )

     ret := DllCall( SendMessageW, "UPtr", hWnd, "UInt", Msg, "UPtr", wParam, "UPtr", lParam )
     return ret
  }
  GetWindowThreadProcessId( hwnd ) {
    return DllCall( "GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0 )
  }
  WinGetFocus( hwnd ) {
    GUITHREADINFO_cbsize := 24 + A_PtrSize*6
    VarSetCapacity( GuiThreadInfo, GUITHREADINFO_cbsize, 0 )	;GuiThreadInfoSize = 48
    NumPut(GUITHREADINFO_cbsize, GuiThreadInfo, 0, "UInt")
    threadWnd := this.GetWindowThreadProcessId( hwnd )
    if not DllCall( "GetGUIThreadInfo", "uint", threadWnd, "UPtr", &GuiThreadInfo )
        return 0
    return NumGet( GuiThreadInfo, 8+A_PtrSize,"UPtr")  ; Retrieve the hwndFocus field from the struct.
  }
  GetPixelInfo( ByRef DIB ) {
    ;~ typedef struct tagBITMAPINFOHEADER {
    ;~ DWORD biSize;              0
    ;~ LONG  biWidth;             4
    ;~ LONG  biHeight;            8
    ;~ WORD  biPlanes;            12
    ;~ WORD  biBitCount;          14
    ;~ DWORD biCompression;       16
    ;~ DWORD biSizeImage;         20
    ;~ LONG  biXPelsPerMeter;     24
    ;~ LONG  biYPelsPerMeter;     28
    ;~ DWORD biClrUsed;           32
    ;~ DWORD biClrImportant;      36
    
    bmi := &DIB  ;BITMAPINFOHEADER  pointer from DIB
    biSize := numget( bmi+0, 0, "UInt" )
    ;~ return bmi + biSize
    biSizeImage := numget( bmi+0, 20, "UInt" )
    biBitCount := numget( bmi+0, 14, "UShort" )
    if ( biSizeImage == 0 )
    {
      biWidth := numget( bmi+0, 4, "UInt" )
      biHeight := numget( bmi+0, 8, "UInt" )
      biSizeImage := (((( biWidth * biBitCount + 31 ) & ~31 ) >> 3 ) * biHeight )
      numput( biSizeImage, bmi+0, 20, "UInt" )
    }
    p := numget( bmi+0, 32, "UInt" )  ;biClrUsed
    if ( p == 0 && biBitCount <= 8 )
      p := 1 << biBitCount
    p := p * 4 + biSize + bmi
    return p
  }
  Gdip_Startup() {
    if !DllCall( "GetModuleHandleW", "Wstr", "gdiplus", "UPtr" )
          DllCall( "LoadLibraryW", "Wstr", "gdiplus", "UPtr" )
    
    VarSetCapacity(GdiplusStartupInput , 3*A_PtrSize, 0), NumPut(1,GdiplusStartupInput ,0,"UInt") ; GdiplusVersion = 1
    DllCall("gdiplus\GdiplusStartup", "Ptr*", pToken, "Ptr", &GdiplusStartupInput, "Ptr", 0)
    return pToken
  }
  Gdip_Shutdown(pToken) {
    DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)
    if hModule := DllCall( "GetModuleHandleW", "Wstr", "gdiplus", "UPtr" )
      DllCall("FreeLibrary", "Ptr", hModule)
    return 0
  }
  StrSplit(str,delim,omit = "") {
    if (strlen(delim) > 1)
    {
      StringReplace,str,str,% delim,ƒ,1 		;■¶╬
      delim = ƒ
    }
    ra := Array()
    loop, parse,str,% delim,% omit
      if (A_LoopField != "")
        ra.Insert(A_LoopField)
    return ra
  }
  RemoveDubls( objArray ) {
    while True
    {
      nodubls := 1
      tempArr := Object()
      for i,val in objArray
      {
        if tempArr.haskey( val )
        {
          nodubls := 0
          objArray.Remove( i )
          break
        }
        tempArr[ val ] := 1
      }
      if nodubls
        break
    }
    return objArray
  }
  RegisterClipboardFormat( fmtName ) {
    return DllCall( "RegisterClipboardFormat", "ptr", &fmtName )
  }
  GetOpenClipboardWindow() {
    return DllCall( "GetOpenClipboardWindow" )
  }
  IsClipboardFormatAvailable( iFmt ) {
    return DllCall( "IsClipboardFormatAvailable", "UInt", iFmt )
  }
  GetImageEncodersSize( ByRef numEncoders, ByRef size ) {
    return DllCall( "gdiplus\GdipGetImageEncodersSize", "Uint*", numEncoders, "UInt*", size )
  }
  GetImageEncoders( numEncoders, size, pImageCodecInfo ) {
    return DllCall( "gdiplus\GdipGetImageEncoders", "Uint", numEncoders, "UInt", size, "Ptr", pImageCodecInfo )
  }
  GetEncoderClsid( format, ByRef CLSID ) {
    ;format should be the following
    ;~ bmp
    ;~ jpeg
    ;~ gif
    ;~ tiff
    ;~ png
    if !format
      return 0
    format := "image/" format
    this.GetImageEncodersSize( num, size )
    if ( size = 0 )
      return 0
    VarSetCapacity( ImageCodecInfo, size, 0 )
    this.GetImageEncoders( num, size, &ImageCodecInfo )
    loop,% num
    {
      pici := &ImageCodecInfo + ( 48+7*A_PtrSize )*(A_Index-1)
      pMime := NumGet( pici+0, 32+4*A_PtrSize, "UPtr" )
      MimeType := StrGet( pMime, "UTF-16")
      if ( MimeType = format )
      {
        VarSetCapacity( CLSID, 16, 0 )
        this.memcopy( &CLSID, pici, 16 )
        return 1
      }
    }
    return 0
  }
}

class WinClip extends WinClip_base
{
  __New()
  {
    this.isinstance := 1
    this.allData := ""
  }
 
  _toclipboard( ByRef data, size )
  {
    if !WinClipAPI.OpenClipboard()
      return 0
    offset := 0
    lastPartOffset := 0
    WinClipAPI.EmptyClipboard()
    while ( offset < size )
    {
      if !( fmt := NumGet( data, offset, "UInt" ) )
        break
      offset += 4
      if !( dataSize := NumGet( data, offset, "UInt" ) )
        break
      offset += 4
      if ( ( offset + dataSize ) > size )
        break
      if !( pData := WinClipAPI.GlobalLock( WinClipAPI.GlobalAlloc( 0x0042, dataSize ) ) )
      {
        offset += dataSize
        continue
      }
      WinClipAPI.memcopy( pData, &data + offset, dataSize )
      if ( fmt == this.ClipboardFormats.CF_ENHMETAFILE )
        pClipData := WinClipAPI.SetEnhMetaFileBits( pData, dataSize )
      else
        pClipData := pData
      if !pClipData
        continue
      WinClipAPI.SetClipboardData( fmt, pClipData )
      if ( fmt == this.ClipboardFormats.CF_ENHMETAFILE )
        WinClipAPI.DeleteEnhMetaFile( pClipData )
      WinClipAPI.GlobalUnlock( pData )
      offset += dataSize
      lastPartOffset := offset
    }
    WinClipAPI.CloseClipboard()
    return lastPartOffset
  }
  
  _fromclipboard( ByRef clipData )
  {
    if !WinClipAPI.OpenClipboard()
      return 0
    nextformat := 0
    objFormats := object()
    clipSize := 0
    formatsNum := 0
    while ( nextformat := WinClipAPI.EnumClipboardFormats( nextformat ) )
    {
      if this.skipFormats.hasKey( nextformat )
        continue
      if ( dataHandle := WinClipAPI.GetClipboardData( nextformat ) )
      {
        pObjPtr := 0, nObjSize := 0
        if ( nextFormat == this.ClipboardFormats.CF_ENHMETAFILE )
        {
          if ( bufSize := WinClipAPI.GetEnhMetaFileBits( dataHandle, hemfBuf ) )
            pObjPtr := &hemfBuf, nObjSize := bufSize
        }
        else if ( nSize := WinClipAPI.GlobalSize( WinClipAPI.GlobalLock( dataHandle ) ) )
          pObjPtr := dataHandle, nObjSize := nSize
        else
          continue
        if !( pObjPtr && nObjSize )
          continue
        objFormats[ nextformat ] := { handle : pObjPtr, size : nObjSize }
        clipSize += nObjSize
        formatsNum++
      }
    }
    structSize := formatsNum*( 4 + 4 ) + clipSize  ;allocating 4 bytes for format ID and 4 for data size
    if !structSize
      return 0
    VarSetCapacity( clipData, structSize, 0 )
    ; array in form of:
    ; format   UInt
    ; dataSize UInt
    ; data     Byte[]
    offset := 0
    for fmt, params in objFormats
    {
      NumPut( fmt, &clipData, offset, "UInt" )
      offset += 4
      NumPut( params.size, &clipData, offset, "UInt" )
      offset += 4
      WinClipAPI.memcopy( &clipData + offset, params.handle, params.size )
      offset += params.size
      WinClipAPI.GlobalUnlock( params.handle )
    }
    WinClipAPI.CloseClipboard()
    return structSize
  }
  
  _IsInstance( funcName )
  {
    if !this.isinstance
    {
      throw Exception( "Error in '" funcName "':`nInstantiate the object first to use this method!", -1 )
      return 0
    }
    return 1
  }
  
  _loadFile( filePath, ByRef Data )
  {
    f := FileOpen( filePath, "r","CP0" )
    if !IsObject( f )
      return 0
    f.Pos := 0
    dataSize := f.RawRead( Data, f.Length )
    f.close()
    return dataSize
  }

  _saveFile( filepath, byRef data, size )
  {
    f := FileOpen( filepath, "w","CP0" )
    bytes := f.RawWrite( &data, size )
    f.close()
    return bytes
  }

  _setClipData( ByRef data, size )
  {
    if !size
      return 0
    if !ObjSetCapacity( this, "allData", size )
      return 0
    if !( pData := ObjGetAddress( this, "allData" ) )
      return 0
    WinClipAPI.memcopy( pData, &data, size )
    return size
  }
  
  _getClipData( ByRef data )
  {
    if !( clipSize := ObjGetCapacity( this, "allData" ) )
      return 0
    if !( pData := ObjGetAddress( this, "allData" ) )
      return 0
    VarSetCapacity( data, clipSize, 0 )
    WinClipAPI.memcopy( &data, pData, clipSize )
    return clipSize
  }
  
  __Delete()
  {
    ObjSetCapacity( this, "allData", 0 )
    return
  }
  
  _parseClipboardData( ByRef data, size )
  {
    offset := 0
    formats := object()
    while ( offset < size )
    {
      if !( fmt := NumGet( data, offset, "UInt" ) )
        break
      offset += 4
      if !( dataSize := NumGet( data, offset, "UInt" ) )
        break
      offset += 4
      if ( ( offset + dataSize ) > size )
        break
      params := { name : this._getFormatName( fmt ), size : dataSize }
      ObjSetCapacity( params, "buffer", dataSize )
      pBuf := ObjGetAddress( params, "buffer" )
      WinClipAPI.memcopy( pBuf, &data + offset, dataSize )
      formats[ fmt ] := params
      offset += dataSize
    }
    return formats
  }
  
  _compileClipData( ByRef out_data, objClip )
  {
    if !IsObject( objClip )
      return 0
    ;calculating required data size
    clipSize := 0
    for fmt, params in objClip
      clipSize += 8 + params.size
    VarSetCapacity( out_data, clipSize, 0 )
    offset := 0
    for fmt, params in objClip
    {
      NumPut( fmt, out_data, offset, "UInt" )
      offset += 4
      NumPut( params.size, out_data, offset, "UInt" )
      offset += 4
      WinClipAPI.memcopy( &out_data + offset, ObjGetAddress( params, "buffer" ), params.size )
      offset += params.size
    }
    return clipSize
  }
  
  GetFormats()
  {
    if !( clipSize := this._fromclipboard( clipData ) )
      return 0
    return this._parseClipboardData( clipData, clipSize )
  }
  
  iGetFormats()
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return 0
    return this._parseClipboardData( clipData, clipSize )
  }
  
  Snap( ByRef data )
  {
    return this._fromclipboard( data )
  }
  
  iSnap()
  {
    this._IsInstance( A_ThisFunc )
    if !( dataSize := this._fromclipboard( clipData ) )
      return 0
    return this._setClipData( clipData, dataSize )
  }

  Restore( ByRef clipData )
  {
    clipSize := VarSetCapacity( clipData )
    return this._toclipboard( clipData, clipSize )
  }

  iRestore()
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return 0
    return this._toclipboard( clipData, clipSize )
  }

  Save( filePath )
  {
    if !( size := this._fromclipboard( data ) )
      return 0
    return this._saveFile( filePath, data, size )
  }
  
  iSave( filePath )
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
          return 0
    return this._saveFile( filePath, clipData, clipSize )
  }

  Load( filePath )
  {
    if !( dataSize := this._loadFile( filePath, dataBuf ) )
      return 0
    return this._toclipboard( dataBuf, dataSize )
  }

  iLoad( filePath )
  {
    this._IsInstance( A_ThisFunc )
    if !( dataSize := this._loadFile( filePath, dataBuf ) )
      return 0
    return this._setClipData( dataBuf, dataSize )
  }

  Clear()
  {
    if !WinClipAPI.OpenClipboard()
      return 0
    WinClipAPI.EmptyClipboard()
    WinClipAPI.CloseClipboard()
    return 1
  }
  
  iClear()
  {
    this._IsInstance( A_ThisFunc )
    ObjSetCapacity( this, "allData", 0 )
  }
  
  Copy( timeout = 1, method = 1 )
  {
    this.Snap( data )
    this.Clear()    ;clearing the clipboard
    if( method = 1 )
      SendInput, ^{Ins}
    else
      SendInput, ^{vk43sc02E} ;ctrl+c
    ClipWait,% timeout, 1
    if ( ret := this._isClipEmpty() )
      this.Restore( data )
    return !ret
  }
  
  iCopy( timeout = 1, method = 1 )
  {
    this._IsInstance( A_ThisFunc )
    this.Snap( data )
    this.Clear()    ;clearing the clipboard
    if( method = 1 )
      SendInput, ^{Ins}
    else
      SendInput, ^{vk43sc02E} ;ctrl+c
    ClipWait,% timeout, 1
    bytesCopied := 0
    if !this._isClipEmpty()
    {
      this.iClear()   ;clearing the variable containing the clipboard data
      bytesCopied := this.iSnap()
    }
    this.Restore( data )
    return bytesCopied
  }
  
  Paste( plainText = "", method = 1 )
  {
    ret := 0
    if ( plainText != "" )
    {
      this.Snap( data )
      this.Clear()
      ret := this.SetText( plainText )
    }
    if( method = 1 )
      SendInput, +{Ins}
    else
      SendInput, ^{vk56sc02F} ;ctrl+v
    this._waitClipReady( 3000 )
    if ( plainText != "" )
    {
      this.Restore( data )
    }
    else
      ret := !this._isClipEmpty()
    return ret
  }
  
  iPaste( method = 1 )
  {
    this._IsInstance( A_ThisFunc )
    this.Snap( data )
    if !( bytesRestored := this.iRestore() )
      return 0
    if( method = 1 )
      SendInput, +{Ins}
    else
      SendInput, ^{vk56sc02F} ;ctrl+v
    this._waitClipReady( 3000 )
    this.Restore( data )
    return bytesRestored
  }
  
  IsEmpty()
  {
    return this._isClipEmpty()
  }
  
  iIsEmpty()
  {
    return !this.iGetSize()
  }
  
  _isClipEmpty()
  {
    return !WinClipAPI.CountClipboardFormats()
  }
  
  _waitClipReady( timeout = 10000 )
  {
    start_time := A_TickCount
    sleep 100
    while ( WinClipAPI.GetOpenClipboardWindow() && ( A_TickCount - start_time < timeout ) )
      sleep 100
  }

  iSetText( textData )
  {
    if ( textData = "" )
      return 0
    this._IsInstance( A_ThisFunc )
    clipSize := this._getClipData( clipData )
    if !( clipSize := this._appendText( clipData, clipSize, textData, 1 ) )
      return 0
    return this._setClipData( clipData, clipSize )
  }
  
  SetText( textData )
  {
    if ( textData = "" )
      return 0
    clipSize :=  this._fromclipboard( clipData )
    if !( clipSize := this._appendText( clipData, clipSize, textData, 1 ) )
      return 0
    return this._toclipboard( clipData, clipSize )
  }

  GetRTF()
  {
    if !( clipSize := this._fromclipboard( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, "Rich Text Format" ) )
      return ""
    return strget( &out_data, out_size, "CP0" )
  }
  
  iGetRTF()
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, "Rich Text Format" ) )
      return ""
    return strget( &out_data, out_size, "CP0" )
  }
  
  SetRTF( textData )
  {
    if ( textData = "" )
      return 0
    clipSize :=  this._fromclipboard( clipData )
    if !( clipSize := this._setRTF( clipData, clipSize, textData ) )
          return 0
    return this._toclipboard( clipData, clipSize )
  }
  
  iSetRTF( textData )
  {
    if ( textData = "" )
      return 0
    this._IsInstance( A_ThisFunc )
    clipSize :=  this._getClipData( clipData )
    if !( clipSize := this._setRTF( clipData, clipSize, textData ) )
          return 0
    return this._setClipData( clipData, clipSize )
  }

  _setRTF( ByRef clipData, clipSize, textData )
  {
    objFormats := this._parseClipboardData( clipData, clipSize )
    uFmt := WinClipAPI.RegisterClipboardFormat( "Rich Text Format" )
    objFormats[ uFmt ] := object()
    sLen := StrLen( textData )
    ObjSetCapacity( objFormats[ uFmt ], "buffer", sLen )
    StrPut( textData, ObjGetAddress( objFormats[ uFmt ], "buffer" ), sLen, "CP0" )
    objFormats[ uFmt ].size := sLen
    return this._compileClipData( clipData, objFormats )
  }

  iAppendText( textData )
  {
    if ( textData = "" )
      return 0
    this._IsInstance( A_ThisFunc )
    clipSize := this._getClipData( clipData )
    if !( clipSize := this._appendText( clipData, clipSize, textData ) )
      return 0
    return this._setClipData( clipData, clipSize )
  }
  
  AppendText( textData )
  {
    if ( textData = "" )
      return 0
    clipSize :=  this._fromclipboard( clipData )
    if !( clipSize := this._appendText( clipData, clipSize, textData ) )
      return 0
    return this._toclipboard( clipData, clipSize )
  }

  SetHTML( html, source = "" )
  {
    if ( html = "" )
      return 0
    clipSize :=  this._fromclipboard( clipData )
    if !( clipSize := this._setHTML( clipData, clipSize, html, source ) )
      return 0
    return this._toclipboard( clipData, clipSize )
  }

  iSetHTML( html, source = "" )
  {
    if ( html = "" )
      return 0
    this._IsInstance( A_ThisFunc )
    clipSize := this._getClipData( clipData )
    if !( clipSize := this._setHTML( clipData, clipSize, html, source ) )
      return 0
    return this._setClipData( clipData, clipSize )
  }

  _calcHTMLLen( num )
  {
    while ( StrLen( num ) < 10 )
      num := "0" . num
    return num
  }

  _setHTML( ByRef clipData, clipSize, htmlData, source )
  {
    objFormats := this._parseClipboardData( clipData, clipSize )
    uFmt := WinClipAPI.RegisterClipboardFormat( "HTML Format" )
    objFormats[ uFmt ] := object()
    encoding := "UTF-8"
    htmlLen := StrPut( htmlData, encoding ) - 1   ;substract null
    srcLen := 2 + 10 + StrPut( source, encoding ) - 1      ;substract null
    StartHTML := this._calcHTMLLen( 105 + srcLen )
    EndHTML := this._calcHTMLLen( StartHTML + htmlLen + 76 )
    StartFragment := this._calcHTMLLen( StartHTML + 38 )
    EndFragment := this._calcHTMLLen( StartFragment + htmlLen )
    html =
    ( Join`r`n
Version:0.9
StartHTML:%StartHTML%
EndHTML:%EndHTML%
StartFragment:%StartFragment%
EndFragment:%EndFragment%
SourceURL:%source%
<html>
<body>
<!--StartFragment-->
%htmlData%
<!--EndFragment-->
</body>
</html>
    )
    sLen := StrPut( html, encoding )
    ObjSetCapacity( objFormats[ uFmt ], "buffer", sLen )
    StrPut( html, ObjGetAddress( objFormats[ uFmt ], "buffer" ), sLen, encoding )
    objFormats[ uFmt ].size := sLen
    return this._compileClipData( clipData, objFormats )
  }
  
  _appendText( ByRef clipData, clipSize, textData, IsSet = 0 )
  {
    objFormats := this._parseClipboardData( clipData, clipSize )
    uFmt := this.ClipboardFormats.CF_UNICODETEXT
    str := ""
    if ( objFormats.haskey( uFmt ) && !IsSet )
      str := strget( ObjGetAddress( objFormats[ uFmt ],  "buffer" ), "UTF-16" )
    else
      objFormats[ uFmt ] := object()
    str .= textData
    sLen := ( StrLen( str ) + 1 ) * 2
    ObjSetCapacity( objFormats[ uFmt ], "buffer", sLen )
    StrPut( str, ObjGetAddress( objFormats[ uFmt ], "buffer" ), sLen, "UTF-16" )
    objFormats[ uFmt ].size := sLen
    return this._compileClipData( clipData, objFormats )
  }
  
  _getFiles( pDROPFILES )
  {
    fWide := numget( pDROPFILES + 0, 16, "uchar" ) ;getting fWide value from DROPFILES struct
    pFiles := numget( pDROPFILES + 0, 0, "UInt" ) + pDROPFILES  ;getting address of files list
    list := ""
    while numget( pFiles + 0, 0, fWide ? "UShort" : "UChar" )
    {
      lastPath := strget( pFiles+0, fWide ? "UTF-16" : "CP0" )
      list .= ( list ? "`n" : "" ) lastPath
      pFiles += ( StrLen( lastPath ) + 1 ) * ( fWide ? 2 : 1 )
    }
    return list
  }
  
  _setFiles( ByRef clipData, clipSize, files, append = 0, isCut = 0 )
  {
    objFormats := this._parseClipboardData( clipData, clipSize )
    uFmt := this.ClipboardFormats.CF_HDROP
    if ( append && objFormats.haskey( uFmt ) )
      prevList := this._getFiles( ObjGetAddress( objFormats[ uFmt ], "buffer" ) ) "`n"
    objFiles := WinClipAPI.StrSplit( prevList . files, "`n", A_Space A_Tab )
    objFiles := WinClipAPI.RemoveDubls( objFiles )
    if !objFiles.MaxIndex()
      return 0
    objFormats[ uFmt ] := object()
    DROP_size := 20 + 2
    for i,str in objFiles
      DROP_size += ( StrLen( str ) + 1 ) * 2
    VarSetCapacity( DROPFILES, DROP_size, 0 )
    NumPut( 20, DROPFILES, 0, "UInt" )  ;offset
    NumPut( 1, DROPFILES, 16, "uchar" ) ;NumPut( 20, DROPFILES, 0, "UInt" )
    offset := &DROPFILES + 20
    for i,str in objFiles
    {
      StrPut( str, offset, "UTF-16" )
      offset += ( StrLen( str ) + 1 ) * 2
    }
    ObjSetCapacity( objFormats[ uFmt ], "buffer", DROP_size )
    WinClipAPI.memcopy( ObjGetAddress( objFormats[ uFmt ], "buffer" ), &DROPFILES, DROP_size )
    objFormats[ uFmt ].size := DROP_size
    prefFmt := WinClipAPI.RegisterClipboardFormat( "Preferred DropEffect" )
    objFormats[ prefFmt ] := { size : 4 }
    ObjSetCapacity( objFormats[ prefFmt ], "buffer", 4 )
    NumPut( isCut ? 2 : 5, ObjGetAddress( objFormats[ prefFmt ], "buffer" ), 0 "UInt" )
    return this._compileClipData( clipData, objFormats )
  }
  
  SetFiles( files, isCut = 0 )
  {
    if ( files = "" )
      return 0
    clipSize := this._fromclipboard( clipData )
    if !( clipSize := this._setFiles( clipData, clipSize, files, 0, isCut ) )
      return 0
    return this._toclipboard( clipData, clipSize )
  }
  
  iSetFiles( files, isCut = 0 )
  {
    this._IsInstance( A_ThisFunc )
    if ( files = "" )
      return 0
    clipSize := this._getClipData( clipData )
    if !( clipSize := this._setFiles( clipData, clipSize, files, 0, isCut ) )
      return 0
    return this._setClipData( clipData, clipSize )
  }
  
  AppendFiles( files, isCut = 0 )
  {
    if ( files = "" )
      return 0
    clipSize := this._fromclipboard( clipData )
    if !( clipSize := this._setFiles( clipData, clipSize, files, 1, isCut ) )
      return 0
    return this._toclipboard( clipData, clipSize )
  }
  
  iAppendFiles( files, isCut = 0 )
  {
    this._IsInstance( A_ThisFunc )
    if ( files = "" )
      return 0
    clipSize := this._getClipData( clipData )
    if !( clipSize := this._setFiles( clipData, clipSize, files, 1, isCut ) )
      return 0
    return this._setClipData( clipData, clipSize )
  }
  
  GetFiles()
  {
    if !( clipSize := this._fromclipboard( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_HDROP ) )
      return ""
    return this._getFiles( &out_data )
  }
  
  iGetFiles()
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_HDROP ) )
      return ""
    return this._getFiles( &out_data )
  }
  
  _getFormatData( ByRef out_data, ByRef data, size, needleFormat )
  {
    needleFormat := WinClipAPI.IsInteger( needleFormat ) ? needleFormat : WinClipAPI.RegisterClipboardFormat( needleFormat )
    if !needleFormat
      return 0
    offset := 0
    while ( offset < size )
    {
      if !( fmt := NumGet( data, offset, "UInt" ) )
        break
      offset += 4
      if !( dataSize := NumGet( data, offset, "UInt" ) )
        break
      offset += 4
      if ( fmt == needleFormat )
      {
        VarSetCapacity( out_data, dataSize, 0 )
        WinClipAPI.memcopy( &out_data, &data + offset, dataSize )
        return dataSize
      }
      offset += dataSize
    }
    return 0
  }
  
  _DIBtoHBITMAP( ByRef dibData )
  {
    ;http://ebersys.blogspot.com/2009/06/how-to-convert-dib-to-bitmap.html
    pPix := WinClipAPI.GetPixelInfo( dibData )
    gdip_token := WinClipAPI.Gdip_Startup()
    DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", &dibData, "Ptr", pPix, "Ptr*", pBitmap )
    DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", hBitmap, "int", 0xffffffff )
    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
    WinClipAPI.Gdip_Shutdown( gdip_token )
    return hBitmap
  }
  
  GetBitmap()
  {
    if !( clipSize := this._fromclipboard( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
      return ""
    return this._DIBtoHBITMAP( out_data )
  }
  
  iGetBitmap()
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
      return ""
    return this._DIBtoHBITMAP( out_data )
  }
  
  _BITMAPtoDIB( bitmap, ByRef DIB )
  {
    if !bitmap
      return 0
    if !WinClipAPI.IsInteger( bitmap )
    {
      gdip_token := WinClipAPI.Gdip_Startup()
      DllCall("gdiplus\GdipCreateBitmapFromFileICM", "wstr", bitmap, "Ptr*", pBitmap )
      DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", hBitmap, "int", 0xffffffff )
      DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
      WinClipAPI.Gdip_Shutdown( gdip_token )
      bmMade := 1
    }
    else
      hBitmap := bitmap, bmMade := 0
    if !hBitmap
        return 0
    ;http://www.codeguru.com/Cpp/G-M/bitmap/article.php/c1765
    if !( hdc := DllCall( "GetDC", "Ptr", 0 ) )
      goto, _BITMAPtoDIB_cleanup
    hPal := DllCall( "GetStockObject", "UInt", 15 ) ;DEFAULT_PALLETE
    hPal := DllCall( "SelectPalette", "ptr", hdc, "ptr", hPal, "Uint", 0 )
    DllCall( "RealizePalette", "ptr", hdc )
    size := DllCall( "GetObject", "Ptr", hBitmap, "Uint", 0, "ptr", 0 )
    VarSetCapacity( bm, size, 0 )
    DllCall( "GetObject", "Ptr", hBitmap, "Uint", size, "ptr", &bm )
    biBitCount := NumGet( bm, 16, "UShort" )*NumGet( bm, 18, "UShort" )
    nColors := (1 << biBitCount)
	if ( nColors > 256 ) 
		nColors := 0
	bmiLen  := 40 + nColors * 4
    VarSetCapacity( bmi, bmiLen, 0 )
    ;BITMAPINFOHEADER initialization
    NumPut( 40, bmi, 0, "Uint" )
    NumPut( NumGet( bm, 4, "Uint" ), bmi, 4, "Uint" )   ;width
    NumPut( biHeight := NumGet( bm, 8, "Uint" ), bmi, 8, "Uint" ) ;height
    NumPut( 1, bmi, 12, "UShort" )
    NumPut( biBitCount, bmi, 14, "UShort" )
    NumPut( 0, bmi, 16, "UInt" ) ;compression must be BI_RGB

    ; Get BITMAPINFO. 
    if !DllCall("GetDIBits"
              ,"ptr",hdc
              ,"ptr",hBitmap
              ,"uint",0 
              ,"uint",biHeight
              ,"ptr",0      ;lpvBits 
              ,"ptr",&bmi  ;lpbi 
              ,"uint",0)    ;DIB_RGB_COLORS
      goto, _BITMAPtoDIB_cleanup
    biSizeImage := NumGet( &bmi, 20, "UInt" )
    if ( biSizeImage = 0 )
    {
      biBitCount := numget( &bmi, 14, "UShort" )
      biWidth := numget( &bmi, 4, "UInt" )
      biHeight := numget( &bmi, 8, "UInt" )
      biSizeImage := (((( biWidth * biBitCount + 31 ) & ~31 ) >> 3 ) * biHeight )
      ;~ dwCompression := numget( bmi, 16, "UInt" )
      ;~ if ( dwCompression != 0 ) ;BI_RGB
        ;~ biSizeImage := ( biSizeImage * 3 ) / 2
      numput( biSizeImage, &bmi, 20, "UInt" )
    }
    DIBLen := bmiLen + biSizeImage
    VarSetCapacity( DIB, DIBLen, 0 )
    WinClipAPI.memcopy( &DIB, &bmi, bmiLen )
    if !DllCall("GetDIBits"
              ,"ptr",hdc
              ,"ptr",hBitmap
              ,"uint",0 
              ,"uint",biHeight
              ,"ptr",&DIB + bmiLen     ;lpvBits 
              ,"ptr",&DIB  ;lpbi 
              ,"uint",0)    ;DIB_RGB_COLORS
      goto, _BITMAPtoDIB_cleanup
_BITMAPtoDIB_cleanup:
    if bmMade
      DllCall( "DeleteObject", "ptr", hBitmap )
    DllCall( "SelectPalette", "ptr", hdc, "ptr", hPal, "Uint", 0 )
    DllCall( "RealizePalette", "ptr", hdc )
    DllCall("ReleaseDC","ptr",hdc)
    if ( A_ThisLabel = "_BITMAPtoDIB_cleanup" )
      return 0
    return DIBLen
  }
  
  _setBitmap( ByRef DIB, DIBSize, ByRef clipData, clipSize )
  {
    objFormats := this._parseClipboardData( clipData, clipSize )
    uFmt := this.ClipboardFormats.CF_DIB
    objFormats[ uFmt ] := { size : DIBSize }
    ObjSetCapacity( objFormats[ uFmt ], "buffer", DIBSize )
    WinClipAPI.memcopy( ObjGetAddress( objFormats[ uFmt ], "buffer" ), &DIB, DIBSize )
    return this._compileClipData( clipData, objFormats )
  }
  
  SetBitmap( bitmap )
  {
    if ( DIBSize := this._BITMAPtoDIB( bitmap, DIB ) )
    {
      clipSize := this._fromclipboard( clipData )
      if ( clipSize := this._setBitmap( DIB, DIBSize, clipData, clipSize ) )
        return this._toclipboard( clipData, clipSize )
    }
    return 0
  }
  
  iSetBitmap( bitmap )
  {
    this._IsInstance( A_ThisFunc )
    if ( DIBSize := this._BITMAPtoDIB( bitmap, DIB ) )
    {
      clipSize := this._getClipData( clipData )
      if ( clipSize := this._setBitmap( DIB, DIBSize, clipData, clipSize ) )
        return this._setClipData( clipData, clipSize )
    }
    return 0
  }
  
  GetText()
  {
    if !( clipSize := this._fromclipboard( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_UNICODETEXT ) )
      return ""
    return strget( &out_data, out_size, "UTF-16" )
  }

  iGetText()
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_UNICODETEXT ) )
      return ""
    return strget( &out_data, out_size, "UTF-16" )
  }
  
  GetHtml()
  {
    if !( clipSize := this._fromclipboard( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, "HTML Format" ) )
      return ""
    return strget( &out_data, out_size, "CP0" )
  }
  
  iGetHtml()
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return ""
    if !( out_size := this._getFormatData( out_data, clipData, clipSize, "HTML Format" ) )
      return ""
    return strget( &out_data, out_size, "CP0" )
  }
  
  _getFormatName( iformat )
  {
    if this.formatByValue.HasKey( iformat )
      return this.formatByValue[ iformat ]
    else
      return WinClipAPI.GetClipboardFormatName( iformat )
  }
  
  iGetData( ByRef Data )
  {
    this._IsInstance( A_ThisFunc )
    return this._getClipData( Data )
  }
  
  iSetData( ByRef data )
  {
    this._IsInstance( A_ThisFunc )
    return this._setClipData( data, VarSetCapacity( data ) )
  }
  
  iGetSize()
  {
    this._IsInstance( A_ThisFunc )
    return ObjGetCapacity( this, "alldata" )
  }
  
  HasFormat( fmt )
  {
    if !fmt
      return 0
    return WinClipAPI.IsClipboardFormatAvailable( WinClipAPI.IsInteger( fmt ) ? fmt 
                                                                                  : WinClipAPI.RegisterClipboardFormat( fmt )  )
  }
  
  iHasFormat( fmt )
  {
    this._IsInstance( A_ThisFunc )
    if !( clipSize := this._getClipData( clipData ) )
      return 0
    return this._hasFormat( clipData, clipSize, fmt )
  }

  _hasFormat( ByRef data, size, needleFormat )
  {
    needleFormat := WinClipAPI.IsInteger( needleFormat ) ? needleFormat 
                                                                  : WinClipAPI.RegisterClipboardFormat( needleFormat )
    if !needleFormat
      return 0
    offset := 0
    while ( offset < size )
    {
      if !( fmt := NumGet( data, offset, "UInt" ) )
        break
      if ( fmt == needleFormat )
        return 1
      offset += 4
      if !( dataSize := NumGet( data, offset, "UInt" ) )
        break
      offset += 4 + dataSize
    }
    return 0
  }
  
  iSaveBitmap( filePath, format )
  {
    this._IsInstance( A_ThisFunc )
    if ( filePath = "" || format = "" )
      return 0
    if !( clipSize := this._getClipData( clipData ) )
      return 0
    if !( DIBsize := this._getFormatData( DIB, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
      return 0
    gdip_token := WinClipAPI.Gdip_Startup()
    if !WinClipAPI.GetEncoderClsid( format, CLSID )
      return 0
    DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", &DIB, "Ptr", WinClipAPI.GetPixelInfo( DIB ), "Ptr*", pBitmap )
    DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "wstr", filePath, "Ptr", &CLSID, "Ptr", 0 )
    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
    WinClipAPI.Gdip_Shutdown( gdip_token )
    return 1
  }
  
  SaveBitmap( filePath, format )
  {
    if ( filePath = "" || format = "" )
      return 0
    if !( clipSize := this._fromclipboard( clipData ) )
      return 0
    if !( DIBsize := this._getFormatData( DIB, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
      return 0
    gdip_token := WinClipAPI.Gdip_Startup()
    if !WinClipAPI.GetEncoderClsid( format, CLSID )
      return 0
    DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", &DIB, "Ptr", WinClipAPI.GetPixelInfo( DIB ), "Ptr*", pBitmap )
    DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "wstr", filePath, "Ptr", &CLSID, "Ptr", 0 )
    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
    WinClipAPI.Gdip_Shutdown( gdip_token )
    return 1
  }
  
  static ClipboardFormats := { CF_BITMAP : 2 ;A handle to a bitmap (HBITMAP).
                              ,CF_DIB : 8  ;A memory object containing a BITMAPINFO structure followed by the bitmap bits.
                              ,CF_DIBV5 : 17 ;A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits.
                              ,CF_DIF : 5 ;Software Arts' Data Interchange Format.
                              ,CF_DSPBITMAP : 0x0082 ;Bitmap display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data.
                              ,CF_DSPENHMETAFILE : 0x008E ;Enhanced metafile display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data.
                              ,CF_DSPMETAFILEPICT : 0x0083 ;Metafile-picture display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data.
                              ,CF_DSPTEXT : 0x0081 ;Text display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data.
                              ,CF_ENHMETAFILE : 14 ;A handle to an enhanced metafile (HENHMETAFILE).
                              ,CF_GDIOBJFIRST : 0x0300 ;Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST.Handles associated with clipboard formats in this range are not automatically deleted using the GlobalFree function when the clipboard is emptied. Also, when using values in this range, the hMem parameter is not a handle to a GDI object, but is a handle allocated by the GlobalAlloc function with the GMEM_MOVEABLE flag.
                              ,CF_GDIOBJLAST : 0x03FF ;See CF_GDIOBJFIRST.
                              ,CF_HDROP : 15 ;A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by passing the handle to the DragQueryFile function.
                              ,CF_LOCALE : 16 ;The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text. An application that pastes text from the clipboard can retrieve this format to determine which character set was used to generate the text. Note that the clipboard does not support plain text in multiple character sets. To achieve this, use a formatted text data type such as RTF instead. The system uses the code page associated with CF_LOCALE to implicitly convert from CF_TEXT to CF_UNICODETEXT. Therefore, the correct code page table is used for the conversion.
                              ,CF_METAFILEPICT : 3 ;Handle to a metafile picture format as defined by the METAFILEPICT structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle.
                              ,CF_OEMTEXT : 7 ;Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
                              ,CF_OWNERDISPLAY : 0x0080 ;Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The hMem parameter must be NULL.
                              ,CF_PALETTE : 9 ;Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.If the clipboard contains data in the CF_PALETTE (logical color palette) format, the application should use the SelectPalette and RealizePalette functions to realize (compare) any other data in the clipboard against that logical palette.When displaying clipboard data, the clipboard always uses as its current palette any object on the clipboard that is in the CF_PALETTE format.
                              ,CF_PENDATA : 10 ;Data for the pen extensions to the Microsoft Windows for Pen Computing.
                              ,CF_PRIVATEFIRST : 0x0200 ;Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message.
                              ,CF_PRIVATELAST : 0x02FF ;See CF_PRIVATEFIRST.
                              ,CF_RIFF : 11 ;Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
                              ,CF_SYLK : 4 ;Microsoft Symbolic Link (SYLK) format.
                              ,CF_TEXT : 1 ;Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
                              ,CF_TIFF : 6 ;Tagged-image file format.
                              ,CF_UNICODETEXT : 13 ;Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
                              ,CF_WAVE : 12 } ;Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
  
  static WM_COPY := 0x301
        ,WM_CLEAR := 0x0303
        ,WM_CUT := 0x0300
        ,WM_PASTE := 0x0302
  
  static skipFormats := {   2      : 0 ;"CF_BITMAP"
                              ,17     : 0 ;"CF_DIBV5"
                              ,0x0082 : 0 ;"CF_DSPBITMAP"
                              ,0x008E : 0 ;"CF_DSPENHMETAFILE"
                              ,0x0083 : 0 ;"CF_DSPMETAFILEPICT"
                              ,0x0081 : 0 ;"CF_DSPTEXT"
                              ,0x0080 : 0 ;"CF_OWNERDISPLAY"
                              ,3      : 0 ;"CF_METAFILEPICT"
                              ,7      : 0 ;"CF_OEMTEXT"
                              ,1      : 0 } ;"CF_TEXT"
                              
  static formatByValue := { 2 : "CF_BITMAP"
                              ,8 : "CF_DIB"
                              ,17 : "CF_DIBV5"
                              ,5 : "CF_DIF"
                              ,0x0082 : "CF_DSPBITMAP"
                              ,0x008E : "CF_DSPENHMETAFILE"
                              ,0x0083 : "CF_DSPMETAFILEPICT"
                              ,0x0081 : "CF_DSPTEXT"
                              ,14 : "CF_ENHMETAFILE"
                              ,0x0300 : "CF_GDIOBJFIRST"
                              ,0x03FF : "CF_GDIOBJLAST"
                              ,15 : "CF_HDROP"
                              ,16 : "CF_LOCALE"
                              ,3 : "CF_METAFILEPICT"
                              ,7 : "CF_OEMTEXT"
                              ,0x0080 : "CF_OWNERDISPLAY"
                              ,9 : "CF_PALETTE"
                              ,10 : "CF_PENDATA"
                              ,0x0200 : "CF_PRIVATEFIRST"
                              ,0x02FF : "CF_PRIVATELAST"
                              ,11 : "CF_RIFF"
                              ,4 : "CF_SYLK"
                              ,1 : "CF_TEXT"
                              ,6 : "CF_TIFF"
                              ,13 : "CF_UNICODETEXT"
                              ,12 : "CF_WAVE" }
}


;---- 将后面的函数附加到自己的脚本中 ----


;-----------------------------------------
; 查找屏幕文字/图像字库及OCR识别
; 注意：参数中的x、y为中心点坐标，w、h为左右上下偏移
; cha1、cha0分别为0、_字符的容许误差百分比
;-----------------------------------------
查找文字(x,y,w,h,wz,c,ByRef rx="",ByRef ry="",ByRef ocr=""
  , cha1=0, cha0=0)
{
  xywh2xywh(x-w,y-h,2*w+1,2*h+1,x,y,w,h)
  if (w<1 or h<1)
    Return, 0
  bch:=A_BatchLines
  SetBatchLines, -1
  ;--------------------------------------
  GetBitsFromScreen(x,y,w,h,Scan0,Stride,bits)
  ;--------------------------------------
  ; 设定图内查找范围，注意不要越界
  sx:=0, sy:=0, sw:=w, sh:=h
  if PicOCR(Scan0,Stride,sx,sy,sw,sh,wz,c
    ,rx,ry,ocr,cha1,cha0)
  {
    rx+=x, ry+=y
    SetBatchLines, %bch%
    Return, 1
  }
  ; 容差为0的若失败则使用 5% 的容差再找一次
  if (cha1=0 and cha0=0)
    and PicOCR(Scan0,Stride,sx,sy,sw,sh,wz,c
      ,rx,ry,ocr,0.05,0.05)
  {
    rx+=x, ry+=y
    SetBatchLines, %bch%
    Return, 1
  }
  SetBatchLines, %bch%
  Return, 0
}

;-- 规范输入范围在屏幕范围内
xywh2xywh(x1,y1,w1,h1,ByRef x,ByRef y,ByRef w,ByRef h)
{
  ; 获取包含所有显示器的虚拟屏幕范围
  SysGet, zx, 76
  SysGet, zy, 77
  SysGet, zw, 78
  SysGet, zh, 79
  left:=x1, right:=x1+w1-1, up:=y1, down:=y1+h1-1
  left:=left<zx ? zx:left, right:=right>zx+zw-1 ? zx+zw-1:right
  up:=up<zy ? zy:up, down:=down>zy+zh-1 ? zy+zh-1:down
  x:=left, y:=up, w:=right-left+1, h:=down-up+1
}

;-- 获取屏幕图像的内存数据，图像包括透明窗口
GetBitsFromScreen(x,y,w,h,ByRef Scan0,ByRef Stride,ByRef bits)
{
  VarSetCapacity(bits, w*h*4, 0)
  Ptr:=A_PtrSize ? "Ptr" : "UInt"
  ; 桌面窗口对应包含所有显示器的虚拟屏幕
  win:=DllCall("GetDesktopWindow", Ptr)
  hDC:=DllCall("GetWindowDC", Ptr,win, Ptr)
  mDC:=DllCall("CreateCompatibleDC", Ptr,hDC, Ptr)
  hBM:=DllCall("CreateCompatibleBitmap", Ptr,hDC
    , "int",w, "int",h, Ptr)
  oBM:=DllCall("SelectObject", Ptr,mDC, Ptr,hBM, Ptr)
  DllCall("BitBlt", Ptr,mDC, "int",0, "int",0, "int",w, "int",h
    , Ptr,hDC, "int",x, "int",y, "uint",0x00CC0020|0x40000000)
  ;--------------------------
  VarSetCapacity(bi, 40, 0)
  NumPut(40, bi, 0, "int"), NumPut(w, bi, 4, "int")
  NumPut(-h, bi, 8, "int"), NumPut(1, bi, 12, "short")
  NumPut(bpp:=32, bi, 14, "short"), NumPut(0, bi, 16, "int")
  ;--------------------------
  DllCall("GetDIBits", Ptr,mDC, Ptr,hBM
    , "int",0, "int",h, Ptr,&bits, Ptr,&bi, "int",0)
  DllCall("SelectObject", Ptr,mDC, Ptr,oBM)
  DllCall("DeleteObject", Ptr,hBM)
  DllCall("DeleteDC", Ptr,mDC)
  DllCall("ReleaseDC", Ptr,win, Ptr,hDC)
  Scan0:=&bits, Stride:=((w*bpp+31)//32)*4
}

;-----------------------------------------
; 图像内查找文字/图像字符串及OCR函数
;-----------------------------------------
PicOCR(Scan0, Stride, sx, sy, sw, sh, wenzi, c
  , ByRef rx, ByRef ry, ByRef ocr, cha1, cha0)
{
  static MyFunc
  if !MyFunc
  {
    x32:="5589E55383C4808B45240FAF451C8B5520C1E20201D0894"
    . "5F08B5528B80000000029D0C1E00289C28B451C01D08945ECC"
    . "745E800000000C745D400000000C745D0000000008B4528894"
    . "5CC8B452C8945C8C745C400000000837D08000F85660100008"
    . "B450CC1E81025FF0000008945C08B450CC1E80825FF0000008"
    . "945BC8B450C25FF0000008945B88B4510C1E81025FF0000008"
    . "945B48B4510C1E80825FF0000008945B08B451025FF0000008"
    . "945AC8B45C02B45B48945A88B45BC2B45B08945A48B45B82B4"
    . "5AC8945A08B55C08B45B401D089459C8B55BC8B45B001D0894"
    . "5988B55B88B45AC01D0894594C745F400000000E9BF000000C"
    . "745F800000000E99D0000008B45F083C00289C28B451801D00"
    . "FB6000FB6C03B45A87C798B45F083C00289C28B451801D00FB"
    . "6000FB6C03B459C7F618B45F083C00189C28B451801D00FB60"
    . "00FB6C03B45A47C498B45F083C00189C28B451801D00FB6000"
    . "FB6C03B45987F318B55F08B451801D00FB6000FB6C03B45A07"
    . "C1E8B55F08B451801D00FB6000FB6C03B45947F0B8B55E88B4"
    . "53401D0C600318345F8018345F0048345E8018B45F83B45280"
    . "F8C57FFFFFF8345F4018B45EC0145F08B45F43B452C0F8C35F"
    . "FFFFFE917020000837D08010F85A30000008B450C83C001C1E"
    . "00789450CC745F400000000EB7DC745F800000000EB628B45F"
    . "083C00289C28B451801D00FB6000FB6C06BD0268B45F083C00"
    . "189C18B451801C80FB6000FB6C06BC04B8D0C028B55F08B451"
    . "801D00FB6000FB6D089D0C1E00429D001C83B450C730B8B55E"
    . "88B453401D0C600318345F8018345F0048345E8018B45F83B4"
    . "5287C968345F4018B45EC0145F08B45F43B452C0F8C77FFFFF"
    . "FE96A010000C745F400000000EB7BC745F800000000EB608B5"
    . "5E88B45308D0C028B45F083C00289C28B451801D00FB6000FB"
    . "6C06BD0268B45F083C00189C38B451801D80FB6000FB6C06BC"
    . "04B8D1C028B55F08B451801D00FB6000FB6D089D0C1E00429D"
    . "001D8C1F80788018345F8018345F0048345E8018B45F83B452"
    . "87C988345F4018B45EC0145F08B45F43B452C0F8C79FFFFFF8"
    . "B452883E8018945908B452C83E80189458CC745F401000000E"
    . "9B0000000C745F801000000E9940000008B45F40FAF452889C"
    . "28B45F801D08945E88B55E88B453001D00FB6000FB6D08B450"
    . "C01D08945EC8B45E88D50FF8B453001D00FB6000FB6C03B45E"
    . "C7F488B45E88D50018B453001D00FB6000FB6C03B45EC7F328"
    . "B45E82B452889C28B453001D00FB6000FB6C03B45EC7F1A8B5"
    . "5E88B452801D089C28B453001D00FB6000FB6C03B45EC7E0B8"
    . "B55E88B453401D0C600318345F8018B45F83B45900F8C60FFF"
    . "FFF8345F4018B45F43B458C0F8C44FFFFFFC745E800000000E"
    . "9E30000008B45E88D1485000000008B454401D08B008945E08"
    . "B45E08945E48B45E48945F08B45E883C0018D1485000000008"
    . "B454401D08B008945908B45E883C0028D1485000000008B454"
    . "401D08B0089458CC745F400000000EB7CC745F800000000EB6"
    . "78B45F08D50018955F089C28B453801D00FB6003C3175278B4"
    . "5E48D50018955E48D1485000000008B453C01C28B45F40FAF4"
    . "52889C18B45F801C88902EB258B45E08D50018955E08D14850"
    . "00000008B454001C28B45F40FAF452889C18B45F801C889028"
    . "345F8018B45F83B45907C918345F4018B45F43B458C0F8C78F"
    . "FFFFF8345E8078B45E83B45480F8C11FFFFFF8B45D00FAF452"
    . "889C28B45D401D08945E4C745F800000000E909030000C745F"
    . "400000000E9ED0200008B45F40FAF452889C28B45F801C28B4"
    . "5E401D08945F0C745E800000000E9BB0200008B45E883C0018"
    . "D1485000000008B454401D08B008945908B45E883C0028D148"
    . "5000000008B454401D08B0089458C8B55F88B459001D03B45C"
    . "C0F8F770200008B55F48B458C01D03B45C80F8F660200008B4"
    . "5E88D1485000000008B454401D08B008945E08B45E883C0038"
    . "D1485000000008B454401D08B008945888B45E883C0048D148"
    . "5000000008B454401D08B008945848B45E883C0058D1485000"
    . "000008B454401D08B008945DC8B45E883C0068D14850000000"
    . "08B454401D08B008945D88B45883945840F4D4584894580C74"
    . "5EC00000000E9820000008B45EC3B45887D378B55E08B45EC0"
    . "1D08D1485000000008B453C01D08B108B45F001D089C28B453"
    . "401D00FB6003C31740E836DDC01837DDC000F88980100008B4"
    . "5EC3B45847D378B55E08B45EC01D08D1485000000008B45400"
    . "1D08B108B45F001D089C28B453401D00FB6003C30740E836DD"
    . "801837DD8000F885C0100008345EC018B45EC3B45800F8C72F"
    . "FFFFF837DC4000F85840000008B55208B45F801C28B454C891"
    . "08B454C83C0048B4D248B55F401CA89108B454C8D50088B459"
    . "089028B454C8D500C8B458C8902C745C4040000008B45F42B4"
    . "58C8945D08B558C89D001C001D08945C88B558C89D0C1E0020"
    . "1D001C083C0648945CC837DD0007907C745D0000000008B452"
    . "C2B45D03B45C87D338B452C2B45D08945C8EB288B55088B451"
    . "401D03B45F87D1B8B45C48D50018955C48D1485000000008B4"
    . "54C01D0C700FFFFFFFF8B459083E8018945088B45C48D50018"
    . "955C48D1485000000008B454C01D08B55E883C2078910817DC"
    . "4FD0300000F8FA4000000C745EC00000000EB298B55E08B45E"
    . "C01D08D1485000000008B453C01D08B108B45F001D089C28B4"
    . "53401D0C600308345EC018B45EC3B45887CCF8B45F883C0010"
    . "145D48B45282B45D43B45CC0F8D13FDFFFF8B45282B45D4894"
    . "5CCE905FDFFFF90EB0490EB01908345E8078B45E83B45480F8"
    . "C39FDFFFF8345F4018B45F43B45C80F8C07FDFFFF8345F8018"
    . "B45F83B45CC0F8CEBFCFFFF837DC4007508B800000000EB1B9"
    . "08B45C48D1485000000008B454C01D0C70000000000B801000"
    . "00083EC805B5DC24800"
    x64:="554889E54883C480894D108955184489452044894D288B4"
    . "5480FAF45388B5540C1E20201D08945F48B5550B8000000002"
    . "9D0C1E00289C28B453801D08945F0C745EC00000000C745D80"
    . "0000000C745D4000000008B45508945D08B45588945CCC745C"
    . "800000000837D10000F85850100008B4518C1E81025FF00000"
    . "08945C48B4518C1E80825FF0000008945C08B451825FF00000"
    . "08945BC8B4520C1E81025FF0000008945B88B4520C1E80825F"
    . "F0000008945B48B452025FF0000008945B08B45C42B45B8894"
    . "5AC8B45C02B45B48945A88B45BC2B45B08945A48B55C48B45B"
    . "801D08945A08B55C08B45B401D089459C8B55BC8B45B001D08"
    . "94598C745F800000000E9DE000000C745FC00000000E9BC000"
    . "0008B45F483C0024863D0488B45304801D00FB6000FB6C03B4"
    . "5AC0F8C910000008B45F483C0024863D0488B45304801D00FB"
    . "6000FB6C03B45A07F768B45F483C0014863D0488B45304801D"
    . "00FB6000FB6C03B45A87C5B8B45F483C0014863D0488B45304"
    . "801D00FB6000FB6C03B459C7F408B45F44863D0488B4530480"
    . "1D00FB6000FB6C03B45A47C288B45F44863D0488B45304801D"
    . "00FB6000FB6C03B45987F108B45EC4863D0488B45684801D0C"
    . "600318345FC018345F4048345EC018B45FC3B45500F8C38FFF"
    . "FFF8345F8018B45F00145F48B45F83B45580F8C16FFFFFFE95"
    . "9020000837D10010F85B60000008B451883C001C1E00789451"
    . "8C745F800000000E98D000000C745FC00000000EB728B45F48"
    . "3C0024863D0488B45304801D00FB6000FB6C06BD0268B45F48"
    . "3C0014863C8488B45304801C80FB6000FB6C06BC04B8D0C028"
    . "B45F44863D0488B45304801D00FB6000FB6D089D0C1E00429D"
    . "001C83B451873108B45EC4863D0488B45684801D0C60031834"
    . "5FC018345F4048345EC018B45FC3B45507C868345F8018B45F"
    . "00145F48B45F83B45580F8C67FFFFFFE999010000C745F8000"
    . "00000E98D000000C745FC00000000EB728B45EC4863D0488B4"
    . "560488D0C028B45F483C0024863D0488B45304801D00FB6000"
    . "FB6C06BD0268B45F483C0014C63C0488B45304C01C00FB6000"
    . "FB6C06BC04B448D04028B45F44863D0488B45304801D00FB60"
    . "00FB6D089D0C1E00429D04401C0C1F80788018345FC018345F"
    . "4048345EC018B45FC3B45507C868345F8018B45F00145F48B4"
    . "5F83B45580F8C67FFFFFF8B455083E8018945948B455883E80"
    . "1894590C745F801000000E9CA000000C745FC01000000E9AE0"
    . "000008B45F80FAF455089C28B45FC01D08945EC8B45EC4863D"
    . "0488B45604801D00FB6000FB6D08B451801D08945F08B45EC4"
    . "898488D50FF488B45604801D00FB6000FB6C03B45F07F538B4"
    . "5EC4898488D5001488B45604801D00FB6000FB6C03B45F07F3"
    . "88B45EC2B45504863D0488B45604801D00FB6000FB6C03B45F"
    . "07F1D8B55EC8B455001D04863D0488B45604801D00FB6000FB"
    . "6C03B45F07E108B45EC4863D0488B45684801D0C600318345F"
    . "C018B45FC3B45940F8C46FFFFFF8345F8018B45F83B45900F8"
    . "C2AFFFFFFC745EC00000000E9100100008B45EC4898488D148"
    . "500000000488B85880000004801D08B008945E48B45E48945E"
    . "88B45E88945F48B45EC48984883C001488D148500000000488"
    . "B85880000004801D08B008945948B45EC48984883C002488D1"
    . "48500000000488B85880000004801D08B00894590C745F8000"
    . "00000E98C000000C745FC00000000EB778B45F48D50018955F"
    . "44863D0488B45704801D00FB6003C31752C8B45E88D5001895"
    . "5E84898488D148500000000488B45784801C28B45F80FAF455"
    . "089C18B45FC01C88902EB2D8B45E48D50018955E44898488D1"
    . "48500000000488B85800000004801C28B45F80FAF455089C18"
    . "B45FC01C889028345FC018B45FC3B45947C818345F8018B45F"
    . "83B45900F8C68FFFFFF8345EC078B45EC3B85900000000F8CE"
    . "1FEFFFF8B45D40FAF455089C28B45D801D08945E8C745FC000"
    . "00000E988030000C745F800000000E96C0300008B45F80FAF4"
    . "55089C28B45FC01C28B45E801D08945F4C745EC00000000E93"
    . "70300008B45EC48984883C001488D148500000000488B85880"
    . "000004801D08B008945948B45EC48984883C002488D1485000"
    . "00000488B85880000004801D08B008945908B55FC8B459401D"
    . "03B45D00F8FE10200008B55F88B459001D03B45CC0F8FD0020"
    . "0008B45EC4898488D148500000000488B85880000004801D08"
    . "B008945E48B45EC48984883C003488D148500000000488B858"
    . "80000004801D08B0089458C8B45EC48984883C004488D14850"
    . "0000000488B85880000004801D08B008945888B45EC4898488"
    . "3C005488D148500000000488B85880000004801D08B008945E"
    . "08B45EC48984883C006488D148500000000488B85880000004"
    . "801D08B008945DC8B458C3945880F4D4588894584C745F0000"
    . "00000E9950000008B45F03B458C7D3F8B55E48B45F001D0489"
    . "8488D148500000000488B45784801D08B108B45F401D04863D"
    . "0488B45684801D00FB6003C31740E836DE001837DE0000F88C"
    . "E0100008B45F03B45887D428B55E48B45F001D04898488D148"
    . "500000000488B85800000004801D08B108B45F401D04863D04"
    . "88B45684801D00FB6003C30740E836DDC01837DDC000F88870"
    . "100008345F0018B45F03B45840F8C5FFFFFFF837DC8000F859"
    . "70000008B55408B45FC01C2488B85980000008910488B85980"
    . "000004883C0048B4D488B55F801CA8910488B8598000000488"
    . "D50088B45948902488B8598000000488D500C8B45908902C74"
    . "5C8040000008B45F82B45908945D48B559089D001C001D0894"
    . "5CC8B559089D0C1E00201D001C083C0648945D0837DD400790"
    . "7C745D4000000008B45582B45D43B45CC7D3B8B45582B45D48"
    . "945CCEB308B55108B452801D03B45FC7D238B45C88D5001895"
    . "5C84898488D148500000000488B85980000004801D0C700FFF"
    . "FFFFF8B459483E8018945108B45C88D50018955C84898488D1"
    . "48500000000488B85980000004801D08B55EC83C2078910817"
    . "DC8FD0300000F8FAF000000C745F000000000EB318B55E48B4"
    . "5F001D04898488D148500000000488B45784801D08B108B45F"
    . "401D04863D0488B45684801D0C600308345F0018B45F03B458"
    . "C7CC78B45FC83C0010145D88B45502B45D83B45D00F8D97FCF"
    . "FFF8B45502B45D88945D0E989FCFFFF90EB0490EB01908345E"
    . "C078B45EC3B85900000000F8CBAFCFFFF8345F8018B45F83B4"
    . "5CC0F8C88FCFFFF8345FC018B45FC3B45D00F8C6CFCFFFF837"
    . "DC8007508B800000000EB23908B45C84898488D14850000000"
    . "0488B85980000004801D0C70000000000B8010000004883EC8"
    . "05DC3909090909090909090909090909090"
    MCode(MyFunc, A_PtrSize=8 ? x64:x32)
  }
  ;--------------------------------------
  ; 统计字库文字的个数和宽高，将解释文字存入数组并删除<>
  ;--------------------------------------
  wenzitab:=[], num:=0, wz:="", j:=""
  Loop, Parse, wenzi, |
  {
    v:=A_LoopField, txt:="", e1:=cha1, e0:=cha0
    ; 用角括号输入每个字库字符串的识别结果文字
    if RegExMatch(v,"<([^>]*)>",r)
      v:=StrReplace(v,r), txt:=r1
    ; 可以用中括号输入每个文字的两个容差，以逗号分隔
    if RegExMatch(v,"\[([^\]]*)]",r)
    {
      v:=StrReplace(v,r), r2:=""
      StringSplit, r, r1, `,
      e1:=r1, e0:=r2
    }
    ; 记录每个文字的起始位置、宽、高、10字符的数量和容差
    StringSplit, r, v, .
    w:=r1, v:=base64tobit(r2), h:=StrLen(v)//w
    if (r0<2 or w>sw or h>sh or StrLen(v)!=w*h)
      Continue
    len1:=StrLen(StrReplace(v,"0"))
    len0:=StrLen(StrReplace(v,"1"))
    e1:=Round(len1*e1), e0:=Round(len0*e0)
    j.=StrLen(wz) "|" w "|" h
      . "|" len1 "|" len0 "|" e1 "|" e0 "|"
    wz.=v, wenzitab[++num]:=Trim(txt)
  }
  IfEqual, wz,, Return, 0
  ;--------------------------------------
  ; wz 使用Astr参数类型可以自动转为ANSI版字符串
  ; in 输入各文字的起始位置等信息，out 返回结果
  ; ss 等为临时内存，jiange 超过间隔就会加入*号
  ;--------------------------------------
  mode:=InStr(c,"**") ? 2 : InStr(c,"*") ? 1 : 0
  c:=StrReplace(c,"*"), jiange:=5, num*=7
  if mode=0
  {
    c:=StrReplace(c,"0x") . "-0"
    StringSplit, r, c, -
    c:=Round("0x" r1), dc:=Round("0x" r2)
  }
  VarSetCapacity(in,num*4,0), i:=-4
  Loop, Parse, j, |
    if (A_Index<=num)
      NumPut(A_LoopField, in, i+=4, "int")
  VarSetCapacity(gs, sw*sh)
  VarSetCapacity(ss, sw*sh, Asc("0"))
  k:=StrLen(wz)*4
  VarSetCapacity(s1, k, 0), VarSetCapacity(s0, k, 0)
  VarSetCapacity(out, 1024*4, 0)
  if DllCall(&MyFunc, "int",mode, "uint",c, "uint",dc
    , "int",jiange, "ptr",Scan0, "int",Stride
    , "int",sx, "int",sy, "int",sw, "int",sh
    , "ptr",&gs, "ptr",&ss
    , "Astr",wz, "ptr",&s1, "ptr",&s0
    , "ptr",&in, "int",num, "ptr",&out)
  {
    ocr:="", i:=-4  ; 返回第一个文字的中心位置
    x:=NumGet(out,i+=4,"int"), y:=NumGet(out,i+=4,"int")
    w:=NumGet(out,i+=4,"int"), h:=NumGet(out,i+=4,"int")
    rx:=x+w//2, ry:=y+h//2
    While (k:=NumGet(out,i+=4,"int"))
      v:=wenzitab[k//7], ocr.=v="" ? "*" : v
    Return, 1
  }
  Return, 0
}

MCode(ByRef code, hex)
{
  ListLines, Off
  bch:=A_BatchLines
  SetBatchLines, -1
  VarSetCapacity(code, StrLen(hex)//2)
  Loop, % StrLen(hex)//2
    NumPut("0x" . SubStr(hex,2*A_Index-1,2)
      , code, A_Index-1, "char")
  Ptr:=A_PtrSize ? "Ptr" : "UInt"
  DllCall("VirtualProtect", Ptr,&code, Ptr
    ,VarSetCapacity(code), "uint",0x40, Ptr . "*",0)
  SetBatchLines, %bch%
  ListLines, On
}

base64tobit(s) {
  ListLines, Off
  s:=RegExReplace(s,"\s+")
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  StringCaseSense, On
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:=(i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,A_LoopField,v)
  }
  StringCaseSense, Off
  s:=SubStr(s,1,InStr(s,"1",0,0)-1)
  ListLines, On
  Return, s
}

bit2base64(s) {
  ListLines, Off
  s:=RegExReplace(s,"\s+")
  s.=SubStr("100000",1,6-Mod(StrLen(s),6))
  s:=RegExReplace(s,".{6}","|$0")
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:="|" . (i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,v,A_LoopField)
  }
  ListLines, On
  Return, s
}


/************  机器码的C源码 ************

int __attribute__((__stdcall__)) OCR( int mode
  , unsigned int c, unsigned int dc
  , int jiange, unsigned char * Bmp, int Stride
  , int sx, int sy, int sw, int sh
  , unsigned char * gs, char * ss
  , char * wz, int * s1, int * s0
  , int * in, int num, int * out )
{
  int x, y, o=sy*Stride+sx*4, j=Stride-4*sw, i=0;
  int o1, o2, w, h, max, len1, len0, e1, e0;
  int sx1=0, sy1=0, sw1=sw, sh1=sh, Ptr=0;

  //准备工作一：先将图像各点在ss中转化为01字符
  if (mode==0)    //颜色模式
  {
    int R=(c>>16)&0xFF, G=(c>>8)&0xFF, B=c&0xFF;
    int dR=(dc>>16)&0xFF, dG=(dc>>8)&0xFF, dB=dc&0xFF;
    int R1=R-dR, G1=G-dG, B1=B-dB;
    int R2=R+dR, G2=G+dG, B2=B+dB;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
      {
        if ( Bmp[2+o]>=R1 && Bmp[2+o]<=R2
          && Bmp[1+o]>=G1 && Bmp[1+o]<=G2
          && Bmp[o]  >=B1 && Bmp[o]  <=B2 )
            ss[i]='1';
      }
  }
  else if (mode==1)    //灰度阀值模式
  {
    c=(c+1)*128;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
        if (Bmp[2+o]*38+Bmp[1+o]*75+Bmp[o]*15<c)
          ss[i]='1';
  }
  else    //mode==2，边缘灰差模式
  {
    for (y=0; y<sh; y++, o+=j)
    {
      for (x=0; x<sw; x++, o+=4, i++)
        gs[i]=(Bmp[2+o]*38+Bmp[1+o]*75+Bmp[o]*15)>>7;
    }
    w=sw-1; h=sh-1;
    for (y=1; y<h; y++)
    {
      for (x=1; x<w; x++)
      {
        i=y*sw+x; j=gs[i]+c;
        if (gs[i-1]>j || gs[i+1]>j
          || gs[i-sw]>j || gs[i+sw]>j)
            ss[i]='1';
      }
    }
  }

  //准备工作二：生成s1、s0查表数组
  for (i=0; i<num; i+=7)
  {
    o=o1=o2=in[i]; w=in[i+1]; h=in[i+2];
    for (y=0; y<h; y++)
    {
      for (x=0; x<w; x++)
      {
        if (wz[o++]=='1')
          s1[o1++]=y*sw+x;
        else
          s0[o2++]=y*sw+x;
      }
    }
  }

  //正式工作：ss中每一点都进行一次全字库匹配
  NextWenzi:
  o1=sy1*sw+sx1;
  for (x=0; x<sw1; x++)
  {
    for (y=0; y<sh1; y++)
    {
      o=y*sw+x+o1;
      for (i=0; i<num; i+=7)
      {
        w=in[i+1]; h=in[i+2];
        if (x+w>sw1 || y+h>sh1)
          continue;
        o2=in[i]; len1=in[i+3]; len0=in[i+4];
        e1=in[i+5]; e0=in[i+6];
        max=len1>len0 ? len1 : len0;
        for (j=0; j<max; j++)
        {
          if (j<len1 && ss[o+s1[o2+j]]!='1' && (--e1)<0)
            goto NoMatch;
          if (j<len0 && ss[o+s0[o2+j]]!='0' && (--e0)<0)
            goto NoMatch;
        }
        //成功找到文字或图像
        if (Ptr==0)
        {
          out[0]=sx+x; out[1]=sy+y;
          out[2]=w; out[3]=h; Ptr=4;
          //找到第一个字就确定后续查找的上下范围和右边范围
          sy1=y-h; sh1=h*3; sw1=h*10+100;
          if (sy1<0)
            sy1=0;
          if (sh1>sh-sy1)
            sh1=sh-sy1;
        }
        else if (x>mode+jiange)  //与前一字间隔较远就添加*号
          out[Ptr++]=-1;
        mode=w-1; out[Ptr++]=i+7;
        if (Ptr>1021)    //返回的int数组中元素个数不超过1024
          goto ReturnOK;
        //清除找到文字，后续查找的左边范围为找到位置的X坐标+1
        for (j=0; j<len1; j++)
          ss[o+s1[o2+j]]='0';
        sx1+=x+1;
        if (sw1>sw-sx1)
          sw1=sw-sx1;
        goto NextWenzi;
        //------------
        NoMatch:
        continue;
      }
    }
  }
  if (Ptr==0)
    return 0;
  ReturnOK:
  out[Ptr]=0;
  return 1;
}

*/


;============ 脚本结束 =================

;
}
