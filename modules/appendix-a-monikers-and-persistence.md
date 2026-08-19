# Appendix A — Monikers, persistence, and structured storage

**Reference material. Read A.1 before Module 7 (§7.5 uses a moniker); read A.2 when you meet an `IPersist*` interface.**

Two families of interfaces that the main modules use but don't teach:

- **Monikers** answer *"how do I name an object that doesn't exist yet?"*
- **`IPersist*` / structured storage** answer *"how do I save an object and get it back?"*

---

# A.1 Monikers

## A.1.1 The idea

`CoCreateInstance` says *"make me a new, empty object of class X."* That's often not what you want. You want:

- "the object represented by `C:\Reports\Q3.xlsx`"
- "the WMI namespace `root\cimv2` on server `FS01`"
- "the Excel instance that is **already running**"
- "an instance of class X, but **elevated**"

None of those are expressible as a CLSID. They need a *name* — and something that knows how to turn that name into an object.

> **A moniker is an object that knows how to name another object, and how to produce it on demand.**

The key word is *object*. A moniker isn't a string; it's a COM object implementing `IMoniker`. It has a string form (its **display name**), but the moniker itself carries the logic to resolve that name. "A name that knows how to resolve itself."

The act of turning a moniker into the object it names is called **binding**.

## A.1.2 `IMoniker`

```cpp
struct IMoniker : public IPersistStream        // note: monikers are themselves persistable
{
    HRESULT BindToObject(IBindCtx* pbc, IMoniker* pmkToLeft, REFIID riid, void** ppv);
    HRESULT BindToStorage(IBindCtx* pbc, IMoniker* pmkToLeft, REFIID riid, void** ppv);
    HRESULT Reduce(IBindCtx* pbc, DWORD dwReduceHowFar, IMoniker** ppmkToLeft, IMoniker** ppmkReduced);
    HRESULT ComposeWith(IMoniker* pmkRight, BOOL fOnlyIfNotGeneric, IMoniker** ppmkComposite);
    HRESULT Enum(BOOL fForward, IEnumMoniker** ppenumMoniker);
    HRESULT IsEqual(IMoniker* pmkOtherMoniker);
    HRESULT Hash(DWORD* pdwHash);
    HRESULT IsRunning(IBindCtx* pbc, IMoniker* pmkToLeft, IMoniker* pmkNewlyRunning);
    HRESULT GetTimeOfLastChange(IBindCtx* pbc, IMoniker* pmkToLeft, FILETIME* pFileTime);
    HRESULT Inverse(IMoniker** ppmk);
    HRESULT CommonPrefixWith(IMoniker* pmkOther, IMoniker** ppmkPrefix);
    HRESULT RelativePathTo(IMoniker* pmkOther, IMoniker** ppmkRelPath);
    HRESULT GetDisplayName(IBindCtx* pbc, IMoniker* pmkToLeft, LPOLESTR* ppszDisplayName);
    HRESULT ParseDisplayName(IBindCtx* pbc, IMoniker* pmkToLeft, LPOLESTR pszDisplayName,
                             ULONG* pchEaten, IMoniker** ppmkOut);
    HRESULT IsSystemMoniker(DWORD* pdwMksys);
};
```

You will almost never implement this. You will frequently *use* `BindToObject` — usually indirectly.

`IBindCtx` is a scratchpad for one binding operation: a timeout, options, and a table of objects bound so far (so a composite moniker doesn't bind the same thing twice). `CreateBindCtx(0, &pbc)` gives you a default one.

## A.1.3 The standard moniker types

| Moniker | Display name form | Names |
|---|---|---|
| **File** | `C:\Reports\Q3.xlsx` | An object stored in a file |
| **Item** | `!Sheet1!A1:B7` | A piece *inside* another object |
| **Class** (OBJREF) | `clsid:0002DF01-0000-0000-C000-000000000046:` | A new instance of a CLSID |
| **Pointer** | *(none)* | Wraps an existing pointer so it can be used where a moniker is required |
| **Anti** | *(none)* | The "`..`" of monikers — cancels the moniker to its left |
| **Composite** | `C:\Book.xlsx!Sheet1!A1` | Two or more monikers joined left-to-right |
| **URL** | `https://server/doc` | A resource by URL |
| **Queue** | `queue:/new:MyApp.Component` | A COM+ queued component (Appendix B) |
| **Session** | `session:2!clsid:...` | An object in a specific terminal-services session |
| **Elevation** | `Elevation:Administrator!new:{CLSID}` | An elevated instance (Module 7 §7.5) |

**Composition** is what makes monikers more than a lookup table. `C:\Book.xlsx!Sheet1!A1:B7` is a file moniker composed with two item monikers. Binding walks left to right: bind the file, ask *it* to resolve `Sheet1`, ask *that* to resolve `A1:B7`. Each moniker only understands its own step.

## A.1.4 Using monikers — the three APIs

### `CoGetObject` — the one you'll actually use

```cpp
IUnknown* pUnk = nullptr;
HRESULT hr = CoGetObject(L"C:\\Reports\\Q3.xlsx", nullptr, IID_IUnknown, (void**)&pUnk);
```

It parses the display name into a moniker, binds it, and releases the moniker — the whole sequence in one call. Module 7's elevation moniker uses exactly this, with a `BIND_OPTS3` to carry the parent window and class context.

### `MkParseDisplayName` — when you need the moniker itself

```cpp
IBindCtx* pbc = nullptr;
CreateBindCtx(0, &pbc);

ULONG chEaten = 0;
IMoniker* pmk = nullptr;
HRESULT hr = MkParseDisplayName(pbc, L"C:\\Reports\\Q3.xlsx", &chEaten, &pmk);
if (SUCCEEDED(hr))
{
    IUnknown* pUnk = nullptr;
    hr = pmk->BindToObject(pbc, nullptr, IID_IUnknown, (void**)&pUnk);
    // ...
    if (pUnk) pUnk->Release();
    pmk->Release();
}
pbc->Release();
```

> **`chEaten` is a diagnostic gift.** On failure it tells you **how many characters were successfully parsed** before the parser gave up. That's the exact offset of the syntax error in the display name — invaluable when a customer's connection string or moniker fails with a generic error.

### Scripting

```vbscript
' GetObject binds a moniker. CreateObject calls CoCreateInstance. Different verbs!
Set xl   = GetObject("C:\Reports\Q3.xlsx")           ' file moniker
Set wmi  = GetObject("winmgmts:\\FS01\root\cimv2")   ' WMI moniker
Set run  = GetObject(, "Excel.Application")          ' ROT lookup - already-running instance
```

```powershell
$wmi = [WMI]"\\.\root\cimv2:Win32_Service.Name='RpcSs'"     # moniker under the hood
$svc = [WMIClass]"\\.\root\cimv2:Win32_Service"
```

**Every WMI connection string is a moniker.** `winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2` is a WMI moniker with binding options. When a WMI ticket says "invalid syntax" or "invalid namespace," you're debugging moniker parsing.

## A.1.5 The Running Object Table (ROT)

The ROT is a process-independent, machine-wide table mapping monikers → **already-running** objects. It's how `GetObject(, "Excel.Application")` finds the Excel you already have open instead of starting a new one.

```cpp
// --- Publish yourself ---
IRunningObjectTable* pROT = nullptr;
GetRunningObjectTable(0, &pROT);

IMoniker* pmk = nullptr;
CreateFileMoniker(L"C:\\Reports\\Q3.xlsx", &pmk);

DWORD dwRegister = 0;
pROT->Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, pMyObject, pmk, &dwRegister);

// --- Later, MANDATORY ---
pROT->Revoke(dwRegister);
```

```cpp
// --- Find something already running ---
IUnknown* pUnk = nullptr;
HRESULT hr = pROT->GetObject(pmk, &pUnk);       // MK_E_UNAVAILABLE if not present
```

Enumerate everything currently registered — a genuinely useful diagnostic:

```cpp
IEnumMoniker* pEnum = nullptr;
pROT->EnumRunning(&pEnum);
IMoniker* pmk = nullptr;
while (pEnum->Next(1, &pmk, nullptr) == S_OK)
{
    LPOLESTR name = nullptr;
    pmk->GetDisplayName(pbc, nullptr, &name);
    wprintf(L"%s\n", name);
    CoTaskMemFree(name);
    pmk->Release();
}
```

### ROT support issues

| Symptom | Cause |
|---|---|
| A dead process still appears in the ROT | `Revoke` never called (crash, or a missing cleanup path). Entries are cleaned lazily. |
| `GetObject(, "Excel.Application")` returns a hung instance | Same — a zombie registration |
| Object never released; process won't exit | `ROTFLAGS_REGISTRATIONKEEPSALIVE` holds a **strong** reference. No `Revoke` = permanent leak, exactly like a missing `Unadvise` (Module 5) |
| Can't see another user's registration | The ROT is per-session and honours security; `ROTFLAGS_ALLOWANYCLIENT` changes this and requires a matching AppID/RunAs setup |

> **Rule, same shape as Module 5's:** every `IRunningObjectTable::Register` needs a `Revoke`. Wrap it in RAII.

## A.1.6 Moniker error codes

| HRESULT | Symbol | Meaning |
|---|---|---|
| `0x800401E4` | `MK_E_SYNTAX` | The display name couldn't be parsed — check `chEaten` for where |
| `0x800401E3` | `MK_E_NOOBJECT` | The named object doesn't exist |
| `0x800401E5` | `MK_E_NOOBJECT`/`MK_E_UNAVAILABLE` | Not in the ROT / not running |
| `0x800401E6` | `MK_E_INVALIDEXTENSION` | No handler for that file type |
| `0x800401EA` | `MK_E_NOTBINDABLE` | Binding not supported for this moniker |
| `0x800401F3` | `CO_E_CLASSSTRING` | The ProgID/CLSID string in the moniker is invalid |
| `0x80040154` | `REGDB_E_CLASSNOTREG` | Binding resolved to a CLSID that isn't registered — **back to Module 2** |

Note the last row: a moniker failure very often bottoms out in an ordinary activation failure. Parse the display name mentally, find the CLSID it implies, and run the Module 2 flow on it.

---

# A.2 Persistence — the `IPersist` family

## A.2.1 The problem

A COM object's state is private. The client can't serialize it — it doesn't know the layout (Module 0). So the object must serialize *itself*, through a standard interface, into a medium the client supplies.

```cpp
struct IPersist : public IUnknown
{
    HRESULT GetClassID(CLSID* pClassID);     // "what class am I?" - so it can be recreated
};
```

That single method is the root of the family, and it's the key to the whole design: the persisted data records the **CLSID**, so a loader can `CoCreateInstance` the right class and hand it its own data back.

## A.2.2 The family

| Interface | Medium | Typical use |
|---|---|---|
| `IPersistStream` | `IStream` (a flat byte stream) | Simple objects; monikers themselves |
| `IPersistStreamInit` | `IStream` + `InitNew()` | Same, but distinguishes "new" from "loaded" — preferred for controls |
| `IPersistFile` | A file path | Shell links (`.lnk`), documents |
| `IPersistStorage` | `IStorage` (a hierarchical store) | Compound documents, embedded objects |
| `IPersistPropertyBag` | Named properties | Controls in HTML/designers; text-based persistence |
| `IPersistMemory` | A fixed memory block | Objects with a known-size state |
| `IPersistFolder` / `IPersistIDList` | Shell item IDs | Shell extensions |

```cpp
struct IPersistStream : public IPersist
{
    HRESULT IsDirty();                                  // S_OK = dirty, S_FALSE = clean
    HRESULT Load(IStream* pStm);
    HRESULT Save(IStream* pStm, BOOL fClearDirty);
    HRESULT GetSizeMax(ULARGE_INTEGER* pcbSize);        // upper bound, for pre-allocation
};
```

> **`IsDirty` returns `S_FALSE` for "clean."** Both values are success. This is the Module 1 §1.5 trap in its natural habitat — `if (hr == S_OK)` here silently means "always dirty."

You already used this family in Module 0's hello-world:

```cpp
IPersistFile* pFile = nullptr;
pLink->QueryInterface(IID_IPersistFile, (void**)&pFile);
pFile->Save(L"C:\\Temp\\test.lnk", TRUE);
```

`CShellLink` implements `IShellLink` (its behaviour) and `IPersistFile` (its storage). Two orthogonal contracts on one object — exactly what `QueryInterface` is for.

## A.2.3 `IStream`

COM's byte-stream abstraction. Module 3 used one as a marshaling transport; here it's genuine storage.

```cpp
struct IStream : public ISequentialStream        // ISequentialStream: Read, Write
{
    HRESULT Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition);
    HRESULT SetSize(ULARGE_INTEGER libNewSize);
    HRESULT CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten);
    HRESULT Commit(DWORD grfCommitFlags);
    HRESULT Revert();
    HRESULT LockRegion(...); HRESULT UnlockRegion(...);
    HRESULT Stat(STATSTG* pstatstg, DWORD grfStatFlag);
    HRESULT Clone(IStream** ppstm);
};
```

Handy implementations you get for free:

```cpp
IStream* pStm = nullptr;
CreateStreamOnHGlobal(nullptr, TRUE, &pStm);         // memory-backed, auto-freeing
SHCreateStreamOnFileEx(path, STGM_READ, 0, FALSE, nullptr, &pStm);   // file-backed
SHCreateMemStream(pInit, cbInit);                    // returns IStream* directly
```

`CreateStreamOnHGlobal` is the standard way to serialize an object into memory:

```cpp
IStream* pStm = nullptr;
CreateStreamOnHGlobal(nullptr, TRUE, &pStm);

IPersistStream* pPS = nullptr;
pObj->QueryInterface(IID_IPersistStream, (void**)&pPS);
pPS->Save(pStm, TRUE);

HGLOBAL hg = nullptr;
GetHGlobalFromStream(pStm, &hg);       // now you have the bytes
```

## A.2.4 Structured storage — `IStorage`

**A file system inside a single file.** An `IStorage` is a directory; it contains named sub-storages and named `IStream`s.

```
  Q3.xlsx  (a compound file)
   ├── [storage] Workbook
   │      ├── [stream] Book
   │      └── [stream] Styles
   ├── [stream]  SummaryInformation
   └── [storage] Macros
          └── [stream] VBA
```

```cpp
IStorage* pStg = nullptr;
HRESULT hr = StgCreateStorageEx(L"C:\\Temp\\doc.stg",
                                STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
                                STGFMT_STORAGE, 0, nullptr, nullptr,
                                IID_IStorage, (void**)&pStg);

IStream* pStm = nullptr;
pStg->CreateStream(L"Contents", STGM_CREATE | STGM_WRITE | STGM_SHARE_EXCLUSIVE,
                   0, 0, &pStm);
pStm->Write("hello", 5, nullptr);
pStm->Commit(STGC_DEFAULT);
pStm->Release();

pStg->Commit(STGC_DEFAULT);      // transacted mode: nothing is durable until this
pStg->Release();
```

Open an existing one:

```cpp
StgOpenStorageEx(path, STGM_READ | STGM_SHARE_DENY_WRITE, STGFMT_ANY,
                 0, nullptr, nullptr, IID_IStorage, (void**)&pStg);

// Is this file a compound file at all?
HRESULT hr = StgIsStorageFile(path);      // S_OK = yes, S_FALSE = no
```

### Where you still meet it

- `.msi` installer databases
- Legacy Office binary formats (`.doc`, `.xls`, `.ppt`) — modern `.docx`/`.xlsx` are ZIP, not structured storage
- Thumbnail caches, `.msg` Outlook messages
- OLE embedded objects in any compound document
- `IPropertySetStorage` — the classic "Summary Information" property set

### Transacted mode

`STGM_TRANSACTED` makes changes invisible until `Commit`, and discardable via `Revert`. Powerful, but it's the source of the classic complaint that **compound files only grow**: freed space is retained internally. `StgCreateDocfile` + copy-out is the "compact" operation.

### Storage error codes

| HRESULT | Symbol | Meaning |
|---|---|---|
| `0x80030002` | `STG_E_FILENOTFOUND` | Path or element missing |
| `0x80030005` | `STG_E_ACCESSDENIED` | ACL, or a sharing mode conflict |
| `0x80030020` | `STG_E_SHAREVIOLATION` | Another handle holds an incompatible `STGM_SHARE_*` |
| `0x80030050` | `STG_E_FILEALREADYEXISTS` | `STGM_CREATE` vs `STGM_FAILIFTHERE` |
| `0x80030109` | `STG_E_INVALIDHEADER` | **Not a compound file**, or corrupt |
| `0x8003001D` | `STG_E_WRITEFAULT` | Disk/IO error |
| `0x80030103` | `STG_E_INVALIDFLAG` | Contradictory `STGM_*` flags |

> **`STG_E_SHAREVIOLATION` is the most common one in tickets**, and it's nearly always mismatched `STGM_SHARE_*` flags between two openers — not a file lock in the usual sense. Ask what mode *each* side opens with.

---

## A.3 Checkpoint

1. What is a moniker, and why isn't it just a string?
2. `CreateObject` vs `GetObject` in VBScript — what does each actually call?
3. `MkParseDisplayName` fails. What does `chEaten` give you, and why is that unusually helpful?
4. Why do dead processes sometimes linger in the ROT, and what's the leak analogous to?
5. `IPersistStream::IsDirty` returns `S_FALSE`. Is the object dirty? What bug does the obvious reading cause?
6. Why does `IPersist::GetClassID` exist — what would break without it?
7. A customer reports `STG_E_SHAREVIOLATION` opening a compound file that "nothing else has open." What do you ask?
8. A WMI script fails with `0x800401E4`. Where's the bug, and which module do you go to if it were `0x80040154` instead?

<details>
<summary>Answers</summary>

1. A moniker is a **COM object** that names another object *and knows how to resolve that name into the object* (binding). It isn't just a string because the resolution logic is type-specific — a file moniker launches a document handler, a WMI moniker connects to a namespace, an item moniker asks its left-hand neighbour to resolve a fragment. The string is only the moniker's *display name*; the behaviour is the point.

2. `CreateObject("X")` → `CLSIDFromProgID` + `CoCreateInstance` — makes a **new** object. `GetObject("name")` → moniker parse + bind (`CoGetObject`) — resolves an **existing** named thing. `GetObject(, "X")` with an empty first argument looks up the **Running Object Table** for an already-running instance.

3. `chEaten` returns the number of characters successfully parsed before the failure, i.e. the exact offset of the syntax error in the display name. It's unusually helpful because moniker display names (especially WMI connection strings) are long and composite, and the HRESULT alone (`MK_E_SYNTAX`) says nothing about *where*.

4. Because `Revoke` was never called — the process crashed or lacked a cleanup path. Cleanup is lazy, so stale entries persist and clients bind to zombie objects. With `ROTFLAGS_REGISTRATIONKEEPSALIVE` the registration holds a **strong** reference, so it's exactly analogous to Module 5's missing `Unadvise`: a permanent leak that keeps both objects, and possibly the whole process, alive.

5. **Not dirty.** `S_FALSE` means clean; `S_OK` means dirty. Both are successes. Writing `if (SUCCEEDED(hr))` or `if (hr == S_OK)`-style tests wrongly gives "always dirty," so the app re-saves unchanged documents — or, inverted, never saves changed ones. It's Module 1 §1.5's rule in the wild.

6. So that persisted data records **which class** wrote it. Without it, a loader holding a stream of bytes would have no way to know which `CoCreateInstance` to call before handing the bytes back via `Load`. It's what makes persistence round-trip across processes and machines.

7. Which `STGM_SHARE_*` mode **each** opener uses. `STG_E_SHAREVIOLATION` is usually two openers with incompatible sharing modes rather than an OS file lock — e.g. one opened `STGM_SHARE_EXCLUSIVE` and another wants `STGM_READ | STGM_SHARE_DENY_WRITE`. Also ask whether their own process opened it twice (a leaked earlier handle), and check with Process Explorer's handle search.

8. `0x800401E4` is `MK_E_SYNTAX` — the bug is in the **connection-string text** (a malformed `winmgmts:` moniker: bad namespace path, mistyped option braces, missing `!`). If it were `0x80040154` the moniker parsed fine and binding reached an unregistered CLSID — that's an ordinary activation failure, so go to **Module 2** (bitness, hive, ProcMon).

</details>

---

## A.4 Rules

1. A moniker is an object, not a string; the string is its display name.
2. `CreateObject` makes new; `GetObject` binds existing. Different verbs, different APIs.
3. `CoGetObject` is the one-call route; use `MkParseDisplayName` only when you need the moniker itself.
4. On parse failure, read `chEaten` — it points at the character that broke.
5. Every ROT `Register` needs a `Revoke`. It's a strong reference.
6. Moniker failures often bottom out as Module 2 activation failures. Find the implied CLSID.
7. `IsDirty` returns `S_FALSE` for clean. Never test `== S_OK`.
8. Persisted data carries the CLSID — that's what `IPersist::GetClassID` is for.
9. `STG_E_SHAREVIOLATION` means mismatched `STGM_SHARE_*`, not necessarily an OS lock.
10. `StgIsStorageFile` returns `S_FALSE` for "not a compound file" — success, not failure.

---

**Back to:** [Module 2 — Activation](02-activation-and-registry.md) · [Module 7 — DCOM & security](07-dcom-and-security.md) · [Curriculum](../README.md)
