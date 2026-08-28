#include "Calculator.h"
#include <cstdio>
#include <cassert>

// Rule 1 - IDENTITY. QI(IID_IUnknown) returns the same pointer whichever
// interface you start from. That is the ONLY legal way to ask "are these two
// pointers the same object?"; comparing interface pointers directly is not,
// as the printout below shows.
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

// Rules 2 and 3 - SYMMETRIC and TRANSITIVE. If you can get from A to B you can
// always get back, and anything reachable from one interface is reachable from
// all of them. Together they mean the interface set is fixed for the object's
// lifetime: QI can never start succeeding, or stop.
static void ProveSymmetryAndTransitivity()
{
    printf("\n=== Rules 2 & 3: symmetric, transitive ===\n");

    ICalculator* a = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&a, "sym");
    printf("  start with  a = ICalculator* = %p\n\n", a);

    // Rule 2 - SYMMETRY: if A can reach B, then B can reach A.
    printf("  Rule 2, symmetry\n");
    IAdvancedCalculator* b = nullptr;
    HRESULT hr = a->QueryInterface(IID_IAdvancedCalculator, (void**)&b);
    printf("    a -> IAdvancedCalculator   hr=0x%08X  b       = %p\n", hr, b);
    assert(SUCCEEDED(hr));

    ICalculator* backToA = nullptr;
    hr = b->QueryInterface(IID_ICalculator, (void**)&backToA);
    printf("    b -> ICalculator           hr=0x%08X  backToA = %p  <- must succeed\n",
           hr, backToA);
    assert(SUCCEEDED(hr));
    printf("    backToA == a ? %s\n\n", (backToA == a) ? "yes" : "no");

    // Rule 3 - TRANSITIVITY: a reaches b, b reaches IUnknown, so a must reach
    // IUnknown too. And since it is the same object, both answers must match.
    printf("  Rule 3, transitivity\n");
    IUnknown* unkFromB = nullptr;
    hr = b->QueryInterface(IID_IUnknown, (void**)&unkFromB);
    printf("    b -> IUnknown              hr=0x%08X  unkFromB = %p\n", hr, unkFromB);
    assert(SUCCEEDED(hr));

    IUnknown* unkFromA = nullptr;
    hr = a->QueryInterface(IID_IUnknown, (void**)&unkFromA);
    printf("    a -> IUnknown              hr=0x%08X  unkFromA = %p  <- must succeed\n",
           hr, unkFromA);
    assert(SUCCEEDED(hr));

    printf("    unkFromA == unkFromB ? %s  <- one object, one identity\n\n",
           (unkFromA == unkFromB) ? "YES" : "NO!!!");
    assert(unkFromA == unkFromB);

    printf("  So the interface set is FIXED: reachable from one pointer means\n"
           "  reachable from every pointer, for as long as the object lives.\n");

    // Five successful QIs and one create = five references. All must go back.
    unkFromA->Release(); unkFromB->Release();
    backToA->Release(); b->Release(); a->Release();
}

// Rule 5 - QI ALWAYS AddRefs on success, even when it hands back the very
// pointer you called it on. Two variables means two references, so this
// function owes two Releases. Watch the count in the trace: 1 -> 2 -> 1 -> 0.
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

// [out] PARAMETER HYGIENE, in both places it matters: a failed QueryInterface
// must null its [out] pointer even when it arrived holding garbage, and any
// failing method must do the same for its own [out] params. Callers depend on
// this to avoid reading stale values after an error.
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
