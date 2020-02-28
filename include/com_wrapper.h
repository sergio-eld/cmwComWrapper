#pragma once

#include <map>
#include <set>
#include <unordered_map>

#include <shared_mutex>

#include <type_traits>
#include <variant>
#include <cassert>

#include <atomic>
#include <functional>

#include <combaseapi.h>
#include <comdef.h>

#undef interface
#undef max

namespace cmw
{
    struct COMContext
    {
        COMContext(bool multithreaded = true);
        ~COMContext();
    };

    template <typename T, 
        class = std::enable_if_t<std::is_base_of_v<IUnknown, T>>>
    class ComPtr
    {
        // dummy class, workaround. 
        // Otherwise error due to not being able to inherit from final class 
        template <class C, bool = std::is_final_v<C>>
        struct hide_refs {};

        // hide IUnknown methods if T is not final
        template <class C>
        struct hide_refs<C, false> : public C
        {

        protected:
            using IUnknown::AddRef;
            using IUnknown::Release;
            using IUnknown::QueryInterface;

            friend class ComPtr<T>;
        };

        using ptr_type = std::conditional_t<std::is_final_v<T>, T, hide_refs<T>>;

        ptr_type *rawPtr_;

    public:

        ComPtr()
            : rawPtr_(nullptr)
        {}

        ComPtr(const ComPtr& other)
            : rawPtr_(other.rawPtr_)
        {
            if (IsValid())
                rawPtr_->AddRef();
        }

        ComPtr(ComPtr&& other)
            : rawPtr_(other.rawPtr_)
        {
            other.rawPtr_ = nullptr;
        }

        ComPtr& operator=(const ComPtr& other)
        {
            release_if_valid(rawPtr_);
            rawPtr_ = other.rawPtr_;
            if (IsValid())
                rawPtr_->AddRef();
            return *this;
        }

        ComPtr& operator=(ComPtr&& other)
        {
            release_if_valid(rawPtr_);
            rawPtr_ = other.rawPtr_;
            other.rawPtr_ = nullptr;

            return *this;
        }

        // ownership must be transfered to this object
        void Init(T* rawPtr)
        {
            rawPtr_ = static_cast<ptr_type*>(rawPtr);
            assert(assert_ptr(rawPtr) && "Invalid raw pointer!");
        }

        // ownership must be transfered to this object
        ComPtr(T* rawPtr)
            : rawPtr_(static_cast<ptr_type*>(rawPtr))
        {
            assert(assert_ptr(rawPtr) && "Invalid raw pointer!");
        }

        bool IsValid() const
        {
            return (bool)rawPtr_;
        }

        operator bool() const
        {
            return IsValid();
        }

        template <typename Q, class = std::enable_if_t<std::is_base_of_v<IUnknown, Q>>>
        ComPtr(Q *raw_parent)
            : rawPtr_(nullptr)
        {
            HRESULT hr = raw_parent->QueryInterface((T**)&rawPtr_);
            assert(SUCCEEDED(hr) && "Failed to query interface!");

            if (!SUCCEEDED(hr))
                throw _com_error(hr);
        }

        // TODO: return hresut code, do not throw
        template <typename Q, class = std::enable_if_t<std::is_base_of_v<IUnknown, Q>>>
        std::variant<ComPtr<Q>, HRESULT> QueryInterface()
        {
            Q *rawOut = nullptr;
            HRESULT hr = rawPtr_->QueryInterface(__uuidof(Q), (void**)&rawOut);

            if (!SUCCEEDED(hr))
                return hr;
                //throw _com_error(hr);

            return ComPtr<Q>(rawOut);
        }

        template <typename Q, class = std::enable_if_t<std::is_base_of_v<IUnknown, Q>>>
        operator ComPtr<Q>()
        {
            std::variant<ComPtr<Q>, HRESULT> ret = QueryInterface<Q>();
            assert(std::holds_alternative<ComPtr<Q>>(ret) && "Failed to query interface!");
            if (std::holds_alternative<HRESULT>(ret))
                throw _com_error(std::get<1>(ret));

            return std::move(std::get<0>(ret));
        }

        ptr_type* operator->()
        {
            assert(IsValid() && "Accessing nullptr!");
            return rawPtr_;
        }

        const ptr_type* operator->() const
        {
            assert(IsValid() && "Accessing nullptr!");
            return rawPtr_;
        }

        ptr_type& operator*()
        {
            assert(IsValid() && "Dereferencing nullptr!");
            return *rawPtr_;
        }

        const ptr_type& operator*() const
        {
            assert(IsValid() && "Dereferencing nullptr!");
            return *rawPtr_;
        }

        T* GetRaw()
        {
            return rawPtr_;
        }

        const T* GetRaw() const
        {
            return rawPtr_;
        }

        size_t RefsCount() const
        {
            if (!IsValid())
                return 0;

            ULONG refs = rawPtr_->AddRef();
            refs = rawPtr_->Release();

            return (size_t)refs;
        }

        ~ComPtr()
        {
            if (rawPtr_)
                rawPtr_->Release();
        }

    private:
        
        static bool assert_ptr(T *ptr)
        {
            if (!ptr)
                return false;
            return true;
        }

        static void release_if_valid(T *ptr)
        {
            if (ptr)
                ptr->Release();
        }
    };

    template <class Interface>
    struct transfer_com_ptr
    {

        using v_interface = std::variant<ComPtr<Interface>, HRESULT>;
        using ptr_com = ComPtr<Interface>;

        v_interface temp_;

        constexpr transfer_com_ptr(v_interface&& temp)
            : temp_(std::move(temp))
        {}

        operator v_interface()
        {
            return std::move(temp_);
        }

        operator ptr_com()
        {
            assert(!std::holds_alternative<HRESULT>(temp_) &&
                "Invalid Connection point received!");
            if (std::holds_alternative<HRESULT>(temp_))
                throw _com_error(std::get<HRESULT>(temp_));

            return std::move(std::get<ptr_com>(temp_));
        }
    };

    template <class Interface, class CoClass>
    class CreateInstance : protected transfer_com_ptr<Interface>
    {
        using transfer = transfer_com_ptr<Interface>;

    public:
        static std::variant<ComPtr<Interface>, HRESULT> Create(tagCLSCTX clsContext,
            IUnknown * pAggregate = nullptr)
        {
            Interface *pRes = nullptr;
            HRESULT hr = CoCreateInstance(__uuidof(CoClass), pAggregate, clsContext,
                __uuidof(Interface), (void**)&pRes);
            if (!SUCCEEDED(hr))
                return hr;

            assert(pRes && "Interface is nullptr!");
            return ComPtr<Interface>(pRes);
        }

        CreateInstance(tagCLSCTX clsContext,
            IUnknown * pAggregate = nullptr) noexcept
            : transfer(Create(clsContext, pAggregate))
        {}

        using transfer::operator transfer::v_interface;
        using transfer::operator transfer::ptr_com;

    };

    template <class Interface, class CoClass, class Dispatch>
    class ComObj
    {
    protected:

        ComPtr<Interface> pInterface_;
        ComPtr<Dispatch> pDispInterface_;

    public:

        using coclass = CoClass;
        using interface = Interface;
        using dispinterface = Dispatch;

        ComObj() = default;

        ComObj(const ComPtr<Interface>& pInterface)
            : pInterface_(pInerface),
            pDispInterface_(pInterface_)
        {}

        operator const ComPtr<Interface>&()
        {
            return pInterface_;
        }
        operator const ComPtr<Dispatch>&()
        {
            return pDispInterface_;
        }

        HRESULT CreateInstance(tagCLSCTX clsContext, IUnknown *pAggregate = nullptr)
        {
            std::variant<ComPtr<Interface>, HRESULT> vInterface =
                cmw::CreateInstance<Interface, CoClass>(clsContext, pAggregate);

            if (std::holds_alternative<HRESULT>(vInterface))
                return std::get<HRESULT>(vInterface);

            pInterface_ = std::move(std::get<0>(vInterface));
            return S_OK;
        }

        ComObj(tagCLSCTX clsContext,
            IUnknown * pAggregate = nullptr)
        {
            try
            {
                pInterface_ = CreateInstance<Interface, CoClass>(clsContext, pAggregate);
                pDispInterface_ = pInterface_;
            }
            catch (const _com_error& error)
                throw error;
        }

    private:

        static void check_dispatch()
        {
            if constexpr (std::is_same_v<void, Dispatch>)
                return;
            else
                static_assert(std::is_base_of_v<IDispatch, Dispatch>,
                    "Dispinterface must inherit from IDispatch!");
        }
    };


    template <class Interface = void, bool = std::is_void_v<Interface>>
    class FindConnectionPoint : protected transfer_com_ptr<IConnectionPoint>
    {
        using transfer = transfer_com_ptr<IConnectionPoint>;

    public:

        static std::variant<ComPtr<IConnectionPoint>, HRESULT>
            Find(IConnectionPointContainer& cpContainer, REFIID riid);

        FindConnectionPoint(IConnectionPointContainer& cpContainer, REFIID riid) noexcept
            : transfer(Find(cpContainer, riid))
        {}

        using transfer::operator transfer::v_interface;
        using transfer::operator transfer::ptr_com;

    };

    template <class Interface>
    class FindConnectionPoint<Interface, false> : public FindConnectionPoint<void>
    {
        static_assert(std::is_base_of_v<IUnknown, Interface>,
            "Invalid interface! Must inherit from IUnknown!");

        static_assert(!(std::is_same_v<IUnknown, Interface> ||
            std::is_same_v<IDispatch, Interface> ||
            std::is_same_v<IConnectionPoint, Interface>), "Class must have a valid unique REFIID");
   
        using base = FindConnectionPoint<void>;
    public:

        static std::variant<ComPtr<IConnectionPoint>, HRESULT>
            Find(IConnectionPointContainer& cpContainer)
        {
            // TODO: assert __uuidof is not null
            return FindConnectionPoint(cpContainer, __uuidof(Interface));
        }

        FindConnectionPoint(IConnectionPointContainer& cpContainer) noexcept
            : base(cpContainer, __uuidof(Interface))
        {}

    };
    
    // helper class. Implements thread safe reference counting
    class reference_counter
    {
        std::atomic<ULONG> refs_ = 1;

    public:
        ULONG AddRef()
        {
            ++refs_;
            return refs_;
        }

        ULONG Release()
        {
            --refs_;
            assert((refs_ != std::numeric_limits<ULONG>::max()) && "Invalid value!");
            return refs_;
        }

        ULONG RefsCount() const
        {
            return refs_;
        }

        constexpr reference_counter() = default;
        ~reference_counter() = default;

    };

    // this class implements multiple connections for a com object. 
    // Connectible object may inherit from it to provide RegConnection 
    class com_connections
    {
        // pointers to IConnectionPoints must be 'Alive' by the moment Unadvise is called 
        std::map<DWORD, ComPtr<IConnectionPoint>> connections_;

    public:

        size_t NumConnections() const
        {
            return connections_.size();
        }

        void RegConnection(DWORD cookie, ComPtr<IConnectionPoint>& cpoint)
        {
            connections_.emplace(cookie, cpoint);
        }

        std::variant<HRESULT, bool> Disconnect(DWORD cookie)
        {
            auto found = connections_.find(cookie);
            if (found == connections_.end())
                return false;

            ComPtr<IConnectionPoint>& cp = found->second;
            HRESULT hr = cp->Unadvise(found->first);

            connections_.erase(found);

            return hr;
        }

        // 
        HRESULT DisconnectAll()
        {
            HRESULT res = S_OK;

            using iter = decltype(connections_)::iterator;

            iter it = connections_.begin();
            while (it != connections_.end())
            {
                ComPtr<IConnectionPoint>& cp = it->second;
                HRESULT hr = cp->Unadvise(it->first);
                if (!SUCCEEDED(hr))
                    res = hr;
                ++it;
            }

            connections_.clear();
            return res;
        }

        com_connections() = default;
        ~com_connections()
        {
            HRESULT hr = DisconnectAll();
            assert(SUCCEEDED(hr));
        }

    };



    template <class Dispatch>
    class listener_traits
    {
        using t_reg_cpoint = void(Dispatch::*)(DWORD, IConnectionPoint&);
        using t_disconnect = std::variant<HRESULT, bool>(Dispatch::*)(DWORD);

        template <class T, class = std::void_t<>>
        struct has_reg_cpoint : std::false_type {};

        template<class T>
        struct has_reg_cpoint<T,
            std::void_t<decltype(t_reg_cpoint(&T::RegConnection))>> :
            std::true_type {};

        template <class T, class = std::void_t<>>
        struct has_disconnect : std::false_type {};

        template<class T>
        struct has_disconnect<T,
            std::void_t<decltype(t_reg_cpoint(&T::Disconnect))>> :
            std::true_type {};

    public:

        using type = Dispatch;
        constexpr static bool inherit_idispatch = std::is_base_of_v<IDispatch, Dispatch>;
        constexpr static bool has_reg_cpoint_v = has_reg_cpoint<Dispatch>();
        constexpr static bool has_disconnect_v = has_disconnect<Dispatch>();

        constexpr static bool is_connectible = //inherit_idispatch &&
            has_reg_cpoint_v &&
            has_disconnect_v;

    };

    template <class T>
    struct tag_iid {};

    template <class Connectible,
        class = std::enable_if_t<listener_traits<Connectible>::is_connectible>>
        class ConnectListener
    {
        HRESULT hr_ = S_OK;

    public:

        // does not register dword cookie in Connectible's connections map
        static std::variant<DWORD, HRESULT> Connect(ComPtr<IUnknown>& pSink,
            IConnectionPoint& cpoint)
        {
            DWORD cookie = 0;
            HRESULT hr = cpoint.Advise(pSink.GetRaw(), &cookie);
            if (!SUCCEEDED(hr))
                return hr;
            return cookie;
        }

        static HRESULT Connect(ComPtr<Connectible>& connectible,
            ComPtr<IConnectionPoint>& cpoint)
        {
            ComPtr<IUnknown> pSink = connectible;
            std::variant<DWORD, HRESULT> vCookie = Connect(pSink, *cpoint);
            if (std::holds_alternative<HRESULT>(vCookie))
                return std::get<HRESULT>(vCookie);

            connectible->RegConnection(std::get<DWORD>(vCookie), cpoint);

            return S_OK;
        }

        static std::variant<HRESULT, bool> Disconnect(ComPtr<Connectible>& connectible, 
            DWORD cookie)
        {
            return connectible->Disconnect(cookie);
        }

        ConnectListener(ComPtr<Connectible>& connectible,
            ComPtr<IConnectionPoint>& cpoint)
            : hr_(Connect(connectible, cpoint))
        {}

        // something is wrong. If this method is used, UnAdvise returns "Object is not connected to server"
        template <class Interface, class Provider>
        ConnectListener(ComPtr<Connectible>& connectible, ComPtr<Provider>& cpProvider, 
            tag_iid<Interface>)
        {
            
            try
            {
                ComPtr<IConnectionPointContainer> cpContainter = cpProvider;
                ComPtr<IConnectionPoint> cPoint =
                    FindConnectionPoint<Interface>(*cpContainter);
                hr_ = Connect(connectible, cPoint);
            }
            catch (const _com_error& error)
            {
                hr_ = error.Error();
            }
           
        }

        operator HRESULT() const
        {
            return hr_;
        }

    };


    enum class disp_inv_args : size_t
    {
        dispIDMember = 0,
        riid,
        locale,
        wFlags,
        pDispParams,
        pVarResult,
        pExcepInfo,
        puArgErr
    };

    using disp_inv_t = HRESULT(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);

    template <typename T>
    struct disp_arg_indx : std::integral_constant<size_t, std::numeric_limits<size_t>::max()> {};

    template <> 
    struct disp_arg_indx<DISPID> : std::integral_constant<size_t, (size_t)disp_inv_args::dispIDMember> {};
    template <>
    struct disp_arg_indx<REFIID> : std::integral_constant<size_t, (size_t)disp_inv_args::riid> {};
    template <>
    struct disp_arg_indx<LCID> : std::integral_constant<size_t, (size_t)disp_inv_args::locale> {};
    template <>
    struct disp_arg_indx<WORD> : std::integral_constant<size_t, (size_t)disp_inv_args::wFlags> {};
    template <>
    struct disp_arg_indx<DISPPARAMS*> : std::integral_constant<size_t, (size_t)disp_inv_args::pDispParams> {};
    template <>
    struct disp_arg_indx<VARIANT*> : std::integral_constant<size_t, (size_t)disp_inv_args::pVarResult> {};
    template <>
    struct disp_arg_indx<EXCEPINFO*> : std::integral_constant<size_t, (size_t)disp_inv_args::pExcepInfo> {};
    template <>
    struct disp_arg_indx<UINT*> : std::integral_constant<size_t, (size_t)disp_inv_args::puArgErr> {};

    template <typename T>
    constexpr inline size_t disp_arg_indx_v = disp_arg_indx<T>();

    template <typename ... A>
    class reduce_disp_inv_args
    {
        std::function<disp_inv_t> reduced_;

        template <size_t ... arg_i>
        constexpr static std::function<disp_inv_t> reduce_args(std::function<HRESULT(A...)>&& f,
            std::index_sequence<arg_i...>)
        {
            static_assert(((arg_i != disp_arg_indx<void>()) && ...),
                "Callback function contains invalid arguement types!");

            return std::function<disp_inv_t>(
                [f = std::move(f)](DISPID dispIDMember,
                    REFIID riid, LCID lcid, WORD wFlags,
                    DISPPARAMS *pDispParams, VARIANT *pVarResult,
                    EXCEPINFO *pExcepInfo, UINT *puArgErr)
            {
                // tuple of default args collection
                auto fwd = std::tie(dispIDMember, riid, lcid,
                    wFlags, pDispParams,
                    pVarResult, pExcepInfo, puArgErr);

                return f(std::get<arg_i>(fwd)...);
            });
        }

        reduce_disp_inv_args() = delete;
        reduce_disp_inv_args(const reduce_disp_inv_args&) = delete;
        reduce_disp_inv_args(reduce_disp_inv_args&&) = delete;

    public:
        constexpr reduce_disp_inv_args(std::function<HRESULT(A...)>&& callback)
            : reduced_(reduce_args(std::move(callback),
                std::index_sequence<disp_arg_indx_v<A>...>()))
        {}

        constexpr operator std::function<disp_inv_t>()
        {
            return std::move(reduced_);
        }
    };

    template <auto ptr, typename = decltype(ptr)>
    struct function_traits;
    
    template <auto ptr, typename R, typename ... Args>
    struct function_traits<ptr, R(*)(Args...)>
    {
        constexpr static size_t args_count = sizeof...(Args);
        using type = R(Args...);
    };

    template <auto ptr, typename R, typename C, typename ... Args>
    struct function_traits<ptr, R(C::*)(Args...)>
    {
        constexpr static size_t args_count = sizeof...(Args);
        using type = R(Args...);
    };

    template <auto ptr>
    constexpr inline size_t fn_args_count_v = function_traits<ptr>::args_count;

    template <auto ptr>
    using fn_invoke_t = typename function_traits<ptr>::type;

    template <size_t i>
    struct pholder{};

    template <size_t i>
    struct std::is_placeholder<pholder<i>> : public std::integral_constant<int, i + 1>{};
   
    template <typename R, typename ... Args>
    class bind_function
    {
        std::function<R(Args...)> bound_;

        template <size_t ... arg_i>
        constexpr static std::function<R(Args...)> Bind(R(*ptr)(Args...), std::index_sequence<arg_i...>)
        {
            return std::bind(ptr, pholder<arg_i>()...);
        }

        template <class T, size_t ... arg_i>
        constexpr static std::function<R(Args...)> Bind(T* obj, R(T::*ptr)(Args...), std::index_sequence<arg_i...>)
        {
            return std::bind(ptr, obj, pholder<arg_i>()...);
        }

    public:
        constexpr bind_function(R(*ptr)(Args...))
            : bound_(Bind(ptr, std::make_index_sequence<sizeof...(Args)>()))
        {}

        template <class T>
        constexpr bind_function(T* obj, R(T::*ptr)(Args...))
            : bound_(Bind(obj, ptr, std::make_index_sequence<sizeof...(Args)>()))
        {}

        constexpr operator std::function<R(Args...)>()
        {
            return std::move(bound_);
        }
    };

    // default implementation has one-to-one interface connection 
    class Listener : public IDispatch
    {
        // destroy last to keep track of references till the end
        reference_counter refCounter_;
        IID connectionIID_;

        // must be destroyed after connections
        //std::unordered_map<DISPID, std::unique_ptr<idisp_callback>> callbackMap_;
        std::unordered_map<DISPID, std::function<disp_inv_t>> callbackMap_;
        com_connections connections_;

        std::shared_mutex mutexMap_;

    public:

        // object address must be unique
        Listener(const Listener&) = delete;
        Listener(Listener&&) = delete;

        // RAII. Terminate connections on destruction
        static std::unique_ptr<Listener> Create(REFIID connectionIID);

        // IConnectible

        virtual REFIID Interface(size_t n = 0) const;
        virtual size_t NumInterfaces() const;

        virtual void SetCallback(DISPID dispiid, std::function<disp_inv_t>&& callback,
                REFIID = IID());

        size_t NumConnections() const;
        void RegConnection(DWORD cookie, ComPtr<IConnectionPoint>& cpoint);
        std::variant<HRESULT, bool> Disconnect(DWORD cookie);
        HRESULT DisconnectAll();

        // IUnknown

        ULONG __stdcall AddRef(void) override;
        ULONG __stdcall Release(void) override;

        // default implementation
        virtual HRESULT __stdcall QueryInterface(REFIID riid, void ** ppvObject) override;

        // IDispatch

        virtual HRESULT __stdcall Invoke(DISPID dispIdMember, 
            REFIID riid, LCID lcid, WORD wFlags, 
            DISPPARAMS * pDispParams, 
            VARIANT * pVarResult, EXCEPINFO * pExcepInfo, 
            UINT * puArgErr) override;

        virtual ~Listener() = default;

    protected:

        Listener(REFIID connectionIID)
            : connectionIID_(connectionIID)
        {}

        // these methods are not implemented
        virtual HRESULT __stdcall GetTypeInfoCount(UINT * pctinfo) override;
        virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR * rgszNames,
            UINT cNames, LCID lcid, DISPID * rgDispId) override;
        virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo ** ppTInfo) override;


    };


    // TODO: make Listener a template parameter?
    class RegisterCallback
    {
        static void Register(Listener& listener, DISPID dispIDMember,
            std::function<disp_inv_t>&& callback)
        {
            listener.SetCallback(dispIDMember, std::move(callback));
        }

    public:

        RegisterCallback(Listener& listener, DISPID dispIDMember,
            std::function<disp_inv_t>&& callback)
        {
            Register(listener, dispIDMember, std::move(callback));
        }

        template <typename ... A>
        RegisterCallback(Listener& listener, DISPID dispIDMember,
            std::function<HRESULT(A...)>&& callback)
            : RegisterCallback(listener, dispIDMember,
                reduce_disp_inv_args(std::move(callback)))
        {
        }

        template <typename ... A>
        RegisterCallback(Listener& listener, DISPID dispIDMember,
            HRESULT(*pCallback)(A...))
            : RegisterCallback(listener, dispIDMember, 
            (std::function<HRESULT(A...)>)bind_function(pCallback))
            // why doesn't this work?
            //: RegisterCallback(listener, dispIDMember, 
            //    reduce_disp_inv_args(bind_function(pCallback)))
        {
        }

        template <class T, typename ... A>
        RegisterCallback(Listener& listener, DISPID dispIDMember, T *pObj,
            HRESULT(T::*pCallback)(A...))
            : RegisterCallback(listener, dispIDMember,
            (std::function<HRESULT(A...)>)bind_function(pObj, pCallback))
           // : RegisterCallback(listener, dispIDMember, 
           //     reduce_disp_inv_args(bind_function(pObj, pCallback)))
        {
        }

    };


    /*
    // TODO: implement
    class ListenerMultiple : public Listener
    {
    };*/

    /*
    template <class COM, class Interface, class Disp>
    struct com_traits
    {
        static_assert(std::is_base_of_v<IDispatch, Disp>, "Invalid Dispathc Interface");
        static_assert(std::is_base_of_v<IUnknown, Interface>, "Type must inherit from IUnknown");

        using coclass = COM;
        using interface = Interface;
        using dispinterface = Disp;

    };
    */
    
}