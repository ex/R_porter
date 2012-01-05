// exAtlTag.h : Declaration of the CexAtlTag

#ifndef __EXATLTAG_H_
#define __EXATLTAG_H_

#include "resource.h"       // main symbols
#include <atlctl.h>


#include "ExTag.h"


/////////////////////////////////////////////////////////////////////////////
// CexAtlTag
class ATL_NO_VTABLE CexAtlTag : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<IexAtlTag2, &IID_IexAtlTag2, &LIBID_ATLTAGLib>,
	public CComControl<CexAtlTag>,
	public IPersistStreamInitImpl<CexAtlTag>,
	public IOleControlImpl<CexAtlTag>,
	public IOleObjectImpl<CexAtlTag>,
	public IOleInPlaceActiveObjectImpl<CexAtlTag>,
	public IViewObjectExImpl<CexAtlTag>,
	public IOleInPlaceObjectWindowlessImpl<CexAtlTag>,
	public IPersistStorageImpl<CexAtlTag>,
	public ISpecifyPropertyPagesImpl<CexAtlTag>,
	public IQuickActivateImpl<CexAtlTag>,
	public IDataObjectImpl<CexAtlTag>,
	public IProvideClassInfo2Impl<&CLSID_exAtlTag, NULL, &LIBID_ATLTAGLib>,
	public CComCoClass<CexAtlTag, &CLSID_exAtlTag>
{
public:
	CexAtlTag()
	{
		mn_Bitrate = 0;
		mn_ErrorNumber = 0;
		mn_Bitrate = 0;
		mn_Mpeg = 0;
		mn_Layer = 0;
		mn_SampleRate = 0;
		mn_Mode = 0;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_EXATLTAG)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CexAtlTag)
	COM_INTERFACE_ENTRY(IexAtlTag)
	COM_INTERFACE_ENTRY(IexAtlTag2)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IViewObjectEx)
	COM_INTERFACE_ENTRY(IViewObject2)
	COM_INTERFACE_ENTRY(IViewObject)
	COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceObject)
	COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
	COM_INTERFACE_ENTRY(IOleControl)
	COM_INTERFACE_ENTRY(IOleObject)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
	COM_INTERFACE_ENTRY(IQuickActivate)
	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY(IDataObject)
	COM_INTERFACE_ENTRY(IProvideClassInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)
END_COM_MAP()

BEGIN_PROP_MAP(CexAtlTag)
	PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
	PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
	// Example entries
	// PROP_ENTRY("Property Description", dispid, clsid)
	// PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()

BEGIN_MSG_MAP(CexAtlTag)
	CHAIN_MSG_MAP(CComControl<CexAtlTag>)
	DEFAULT_REFLECTION_HANDLER()
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);



// IViewObjectEx
	DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

// IexAtlTag
private:

	INT		mn_ErrorNumber;
	INT		mn_Bitrate;
	short	mn_Mpeg;
	short	mn_Layer;
	LONG	mn_SampleRate;
	short	mn_Mode;
	char	ms_Path[MAX_PATH + 1];
	CExTag	clxTag;

public:
	STDMETHOD(FreeBuffer)();
	STDMETHOD(SetPathFile2)(/*[in]*/ BSTR pathFile, /*[in]*/ BYTE bVerifyBR, /*[in]*/ LONG cod);
	STDMETHOD(SetBufferLength)(INT BufferLen);
	STDMETHOD(get_ErrorNumber)(/*[out, retval]*/ short *pVal);
	STDMETHOD(get_Mode)(/*[out, retval]*/ short *pVal);
	STDMETHOD(get_SampleRate)(/*[out, retval]*/ INT *pVal);
	STDMETHOD(get_Layer)(/*[out, retval]*/ short *pVal);
	STDMETHOD(get_Mpeg)(/*[out, retval]*/ short *pVal);
	STDMETHOD(get_Bitrate)(/*[out, retval]*/ INT *pVal);
	STDMETHOD(SetPathFile)(/*[in]*/ BSTR pathFile, /*[in]*/ LONG BufferLen, /*[in]*/ BYTE bVerifyBR, /*[in]*/ LONG cod);

	HRESULT OnDraw(ATL_DRAWINFO& di)
	{
		RECT& rc = *(RECT*)di.prcBounds;
		Rectangle(di.hdcDraw, rc.left, rc.top, rc.right, rc.bottom);

		SetTextAlign(di.hdcDraw, TA_CENTER|TA_BASELINE);
		LPCTSTR pszText = _T("ex");
		TextOut(di.hdcDraw, (rc.left + rc.right) / 2, (rc.top + rc.bottom) / 2, pszText, lstrlen(pszText));

		return S_OK;
	}
};

#endif //__EXATLTAG_H_
