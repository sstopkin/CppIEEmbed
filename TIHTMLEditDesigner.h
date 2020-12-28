#include <windows.h>
#include <mshtmhst.h>
#include <mshtmdid.h>

class TIHTMLEditDesigner : public IHTMLEditDesigner
{
public:
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void __RPC_FAR *__RPC_FAR *ppvObject) {
		HRESULT hrRet = S_OK;
		*ppvObject = NULL;
		if (IsEqualIID(riid, IID_IUnknown))
			*ppvObject = (IUnknown *)this;
		else if (IsEqualIID(riid, IID_IHTMLEditDesigner))
			*ppvObject = (IHTMLEditDesigner *)this;
		else
			hrRet = E_NOINTERFACE;
		if (S_OK == hrRet)
			((IUnknown *)*ppvObject)->AddRef();
		return hrRet;
	}
	virtual ULONG   STDMETHODCALLTYPE AddRef(void) {
		return ++m_uRefCount;
	};
	virtual ULONG   STDMETHODCALLTYPE Release(void) {
		--m_uRefCount;
		if (m_uRefCount == 0) delete this;
		return m_uRefCount;
	};

	virtual HRESULT STDMETHODCALLTYPE PreHandleEvent(DISPID inEvtDispId, IHTMLEventObj *pIEventObj) {
		if (inEvtDispId == DISPID_KEYPRESS) {
			IHTMLElement* ele = NULL;
			if (FAILED(pIEventObj->get_srcElement(&ele))) return S_FALSE;
			VARIANT_BOOL bIsEdit = 0;
			if (FAILED(ele->get_isTextEdit(&bIsEdit))) return S_FALSE;
			ele->Release();
			VARIANT_BOOL bControl = 0;
			if (FAILED(pIEventObj->get_ctrlKey(&bControl))) return S_FALSE;
			if (bControl) {
				long keycode = 0;
				if (FAILED(pIEventObj->get_keyCode(&keycode))) return S_FALSE;
				if (keycode == 3)
					m_webform->ExecWB(OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER, NULL, NULL);
				if (keycode == 22 && bIsEdit)
					m_webform->ExecWB(OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER, NULL, NULL);
				if (keycode == 24 && bIsEdit)
					m_webform->ExecWB(OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER, NULL, NULL);
				if (keycode == 1 && bIsEdit)
					m_webform->ExecWB(OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER, NULL, NULL);
			}
		}
		if (inEvtDispId == DISPID_MOUSEDOWN) {
			IHTMLElement* ele = NULL;
			if (FAILED(pIEventObj->get_srcElement(&ele))) return S_FALSE;
		}
		return S_FALSE;
	};
	virtual HRESULT STDMETHODCALLTYPE PostHandleEvent(DISPID inEvtDispId, IHTMLEventObj *pIEventObj) {
		return S_FALSE;
	};
	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(DISPID inEvtDispId, IHTMLEventObj *pIEventObj) {
		return S_FALSE;
	};
	virtual HRESULT STDMETHODCALLTYPE PostEditorEventNotify(DISPID inEvtDispId, IHTMLEventObj *pIEventObj) {
		return S_FALSE;
	};

	TIHTMLEditDesigner(WebForm* wf) : m_webform(wf) {};
	~TIHTMLEditDesigner() {};

private:
	ULONG m_uRefCount = 0;
	WebForm* m_webform;
};
