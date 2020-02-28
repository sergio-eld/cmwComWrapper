#include "com_wrapper.h"

#include <exception>


using namespace cmw;

COMContext::COMContext(bool multithreaded)
{
    tagCOINIT threads = multithreaded ? 
        COINIT_MULTITHREADED :
        COINIT_APARTMENTTHREADED;

    HRESULT res = CoInitializeEx(0, threads);

    assert(SUCCEEDED(res) && "Failed to initialize COM");

    if (!SUCCEEDED(res))
        throw std::exception("Failed to initialize COM");
}

COMContext::~COMContext()
{
    CoUninitialize();
}


std::variant<ComPtr<IConnectionPoint>, HRESULT> FindConnectionPoint<void>::Find(IConnectionPointContainer & cpContainer, REFIID riid)
{
    IConnectionPoint *pCp = nullptr;
    HRESULT hr = cpContainer.FindConnectionPoint(riid, &pCp);
    if (!SUCCEEDED(hr))
        return hr;

    return ComPtr<IConnectionPoint>(pCp);
}



std::unique_ptr<Listener> cmw::Listener::Create(REFIID connectionIID)
{
    return std::unique_ptr<Listener>(new Listener(connectionIID));
}

REFIID cmw::Listener::Interface(size_t n) const
{
    assert(!n && "Only one interface connectible!");
    if (n)
        throw std::out_of_range("Inerface index is out of bounds!");

    return connectionIID_;
}

size_t cmw::Listener::NumInterfaces() const
{
    return 1;
}

void cmw::Listener::SetCallback(DISPID dispiid, std::function<disp_inv_t>&& callback, REFIID)
{
    std::unique_lock<std::shared_mutex> uLock;
    callbackMap_.emplace(dispiid, std::move(callback));
}

size_t cmw::Listener::NumConnections() const
{
    return connections_.NumConnections();
}

void cmw::Listener::RegConnection(DWORD cookie, ComPtr<IConnectionPoint>& cpoint)
{
    connections_.RegConnection(cookie, cpoint);
}

std::variant<HRESULT, bool> cmw::Listener::Disconnect(DWORD cookie)
{
    return connections_.Disconnect(cookie);
}

HRESULT cmw::Listener::DisconnectAll()
{
    return connections_.DisconnectAll();
}

ULONG __stdcall cmw::Listener::AddRef(void)
{
    return refCounter_.AddRef();
}

ULONG __stdcall cmw::Listener::Release(void)
{
    ULONG refs = refCounter_.Release();
    return refs;
}

// default implementation

HRESULT __stdcall cmw::Listener::QueryInterface(REFIID riid, void ** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (riid == IID_IUnknown)
    {
        *ppvObject = static_cast<IUnknown*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == IID_IDispatch ||
        riid == connectionIID_)
    {
        *ppvObject = static_cast<IDispatch*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

HRESULT __stdcall cmw::Listener::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult, EXCEPINFO * pExcepInfo, UINT * puArgErr)
{
    std::shared_lock<std::shared_mutex> sharedLock;

    decltype(callbackMap_)::iterator found = callbackMap_.find(dispIdMember);

    if (found == callbackMap_.cend())
        return DISP_E_MEMBERNOTFOUND;
    /*
    return found->second->Invoke(dispIdMember, riid,
        lcid, wFlags,
        pDispParams,
        pVarResult, pExcepInfo, puArgErr);*/

    return found->second(dispIdMember, riid,
        lcid, wFlags,
        pDispParams,
        pVarResult, pExcepInfo, puArgErr);
}

HRESULT __stdcall cmw::Listener::GetTypeInfoCount(UINT * pctinfo)
{
    if (!pctinfo)
        return E_INVALIDARG;

    return E_NOTIMPL;
}

HRESULT __stdcall cmw::Listener::GetIDsOfNames(REFIID riid, LPOLESTR * rgszNames, UINT cNames, LCID lcid, DISPID * rgDispId)
{
    return E_NOTIMPL;
}

HRESULT __stdcall cmw::Listener::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo ** ppTInfo)
{
    return E_NOTIMPL;
}
