
#include "com_wrapper.h"

#include <iostream>



struct Bar
{
    int i = 5;

    static void foo(int, float)
    {

    }
    void bar(double, char)
    {

    }

    static HRESULT printDispID(DISPID num)
    {
        std::cout << "Bar static method called: " << num << std::endl;
        return S_OK;
    }

    HRESULT printNum()
    {
        std::cout << "Bar: " << i << std::endl;
        return S_OK;
    }

    HRESULT receiveParams(DISPPARAMS *pparams)
    {
        std::cout << "Bar id_" << i << "has ";
        if (!pparams)
            std::cout << "not ";
        std::cout << "received params" << std::endl;
        return S_OK;
    }

};



int main(int argc, const char **argv)
{

    std::function<HRESULT(DISPID, DISPPARAMS*)> dummy{ [](DISPID, DISPPARAMS*) 
    {
        std::cout << "Invoking reduced function" << std::endl;

        return S_OK;
    } };

    std::function<HRESULT(DISPPARAMS*, DISPID)> dummy_reversed{ [](DISPPARAMS*, DISPID)
    {
        std::cout << "Invoking reduced function with reversed args' order" << std::endl;

        return S_OK;
    } };

    Bar bar;
    std::unique_ptr<cmw::Listener> listener = cmw::Listener::Create(IID());

    cmw::RegisterCallback(*listener, 0, std::move(dummy));
    cmw::RegisterCallback(*listener, 1, std::move(dummy_reversed));
    cmw::RegisterCallback(*listener, 2, &Bar::printDispID);
    cmw::RegisterCallback(*listener, 3, &bar, &Bar::printNum);
    cmw::RegisterCallback(*listener, 4, &bar, &Bar::receiveParams);

    HRESULT hr = S_OK;

    hr = listener->Invoke(0, IID(), LCID(), WORD(), nullptr,
        nullptr, nullptr, nullptr);
    if (!SUCCEEDED(hr))
        return -1;

    hr = listener->Invoke(1, IID(), LCID(), WORD(), nullptr,
        nullptr, nullptr, nullptr);
    if (!SUCCEEDED(hr))
        return -1;

    hr = listener->Invoke(2, IID(), LCID(), WORD(), nullptr,
        nullptr, nullptr, nullptr);
    if (!SUCCEEDED(hr))
        return -1;

    hr = listener->Invoke(3, IID(), LCID(), WORD(), nullptr,
        nullptr, nullptr, nullptr);
    if (!SUCCEEDED(hr))
        return -1;

    hr = listener->Invoke(4, IID(), LCID(), WORD(), nullptr,
        nullptr, nullptr, nullptr);
    if (!SUCCEEDED(hr))
        return -1;

   return 0;
}