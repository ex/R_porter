
AtlTagps.dll: dlldata.obj AtlTag_p.obj AtlTag_i.obj
	link /dll /out:AtlTagps.dll /def:AtlTagps.def /entry:DllMain dlldata.obj AtlTag_p.obj AtlTag_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del AtlTagps.dll
	@del AtlTagps.lib
	@del AtlTagps.exp
	@del dlldata.obj
	@del AtlTag_p.obj
	@del AtlTag_i.obj
