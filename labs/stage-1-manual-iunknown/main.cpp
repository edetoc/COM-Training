#include "Calculator.h"
#include <cstdio>
#include <cassert>

static void ProveIdentityRule()
{
    printf("\n=== Rule 1: reflexive / identity ===\n");
    ICalculator* pCalc = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&pCalc, "identity");

    IAdvancedCalculator* pAdv = nullptr;
    pCalc->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);

    printf("pCalc = %p, pAdv = %p  -> DIFFERENT addresses, same object!\n", pCalc, pAdv);

    IUnknown* u1 = nullptr; IUnknown* u2 = nullptr;
    pCalc->QueryInterface(IID_IUnknown, (void**)&u1);
    pAdv ->QueryInterface(IID_IUnknown, (void**)&u2);
    printf("u1 = %p, u2 = %p  -> MUST be equal: %s\n", u1, u2, (u1 == u2) ? "YES" : "NO!!!");
    assert(u1 == u2);

    u2->Release(); u1->Release(); pAdv->Release(); pCalc->Release();
}

static void ProveSymmetryAndTransitivity()
{
    printf("\n=== Rules 2 & 3: symmetric, transitive ===\n");
    ICalculator* a = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&a, "sym");

    IAdvancedCalculator* b = nullptr;
    assert(SUCCEEDED(a->QueryInterface(IID_IAdvancedCalculator, (void**)&b)));   // A -> B

    ICalculator* backToA = nullptr;
    assert(SUCCEEDED(b->QueryInterface(IID_ICalculator, (void**)&backToA)));     // B -> A
    printf("symmetry OK\n");

    backToA->Release(); b->Release(); a->Release();
}

static void ProveQIAddRefs()
{
    printf("\n=== Rule 5: QI always AddRefs, even for the SAME iid ===\n");
    ICalculator* p1 = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&p1, "qi-addref");   // count 1

    ICalculator* p2 = nullptr;
    p1->QueryInterface(IID_ICalculator, (void**)&p2);              // count 2  <-- watch the trace
    printf("p1 == p2 : %s, but the count went to 2.\n", (p1 == p2) ? "yes" : "no");

    p2->Release();   // 1
    p1->Release();   // 0 -> DESTROY
}

static void ShowFailureBehaviour()
{
    printf("\n=== E_NOINTERFACE and [out] param hygiene ===\n");
    ICalculator* p = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&p, "fail");

    // 1. A failed QI must NULL the [out] pointer, even if it arrived holding garbage.
    IClassFactory* pCF = reinterpret_cast<IClassFactory*>(static_cast<UINT_PTR>(0xDEADBEEF));
    HRESULT hr = p->QueryInterface(IID_IClassFactory, (void**)&pCF);
    printf("QI(IClassFactory) hr = 0x%08X (E_NOINTERFACE), pCF = %p (MUST be null)\n", hr, pCF);
    assert(hr == E_NOINTERFACE && pCF == nullptr);

    // 2. A FAILING METHOD must null its [out] params too - not just QueryInterface.
    IAdvancedCalculator* pAdv = nullptr;
    hr = p->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);
    assert(SUCCEEDED(hr));

    long r = 999;                        // garbage the callee is obliged to overwrite
    hr = pAdv->Divide(10, 0, &r);
    printf("Div hr = 0x%08X (E_INVALIDARG), r = %ld (MUST be 0)\n", hr, r);
    assert(hr == E_INVALIDARG && r == 0);

    pAdv->Release();
    p->Release();
}

int main()
{
    ProveIdentityRule();
    ProveSymmetryAndTransitivity();
    ProveQIAddRefs();
    ShowFailureBehaviour();
    printf("\nDone. Every CREATE above must have a matching DESTROY.\n");
    return 0;
}
