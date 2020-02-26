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
