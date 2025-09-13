# Gsync — Outlook Message Change Tracking (VSTO Add‑in, .NET Framework 4.8.1)

> **Branch:** `feature/message-change-tracker`  •  **Commit:** `b39cbdd3b7ec`  •  **Solution:** `Gsync.sln`

Gsync is a Visual Studio Tools for Office (VSTO) **Outlook add‑in** that lays the foundation for **message change tracking**. It provides a clean abstraction layer over Outlook COM objects, safe “detached” snapshots for diffing, robust comparers, and testable utilities (threading, serialization, logging) so the change‑tracker feature can be built in small, verifiable increments.

---

## At a Glance

- **Tech:** VSTO add‑in, .NET Framework **4.8.1**, C#
- **Targets:** Microsoft Outlook (Desktop)
- **Key pillars:**
  - Outlook interop **wrappers** (`IItem`, `IMailItem`, `OutlookItemWrapper`, `MailItemWrapper`, `Detached*`)
  - **Detached snapshots** for safe comparison & persistence (`DetachedOutlookItem`, `DetachedMailItem`)
  - **Comparers** for equality and similarity (`MailItemEqualityComparer`, `MailItemSimilarityComparer`, `IItemEqualityComparer`, `IItemSimilarityComparer`)
  - **Store discovery** (`StoresWrapper`, `StoreWrapper`, `GmailStoreLocator`)
  - **Threading & UI helpers** (`UiWrapper`, `UiTask`, timers) to keep the UI responsive
  - **SmartSerializable** framework for **config/persistence** with DI‑friendly file system & dialogs
  - **Ribbon** plumbing with **Development** actions to exercise the stack
  - **SVGControl** WinForms library for crisp, scalable UI assets
- **Logging:** log4net; logs emitted under `%APPDATA%\Gsync\Logs` (e.g., `debug yyyy-MM-dd-HH-mm.log`, `trace yyyy-MM-dd-HH-mm.log`).

---

## Solution Structure

```
Gsync.sln
├─ Gsync               # Outlook add‑in (VSTO) – core code, wrappers, ribbon, logging
├─ Gsync.Test          # MSTest-based unit tests (uses Moq)
└─ SVGControl          # WinForms SVG rendering & controls
```

### Notable Namespaces & Components

- **OutlookInter​op**
  - `Interfaces/Items`: **`IItem`**, **`IMailItem`**
  - `Item/_OutlookItem`: `OutlookItemWrapper`, `OutlookItemWrapperLoose`, `DetachedOutlookItem`
  - `Item/MailItem`: `MailItemWrapper`, `DetachedMailItem`, `MailItemEqualityComparer`, `MailItemSimilarityComparer`
  - `Store`: `StoresWrapper`, `StoreWrapper`
  - `Folder`: `FolderMinimalWrapper`
  - `Extensions`: WinForms interop & UI helpers
- **GmailLink**: `GmailStoreLocator` (simple heuristics to find Gmail stores)
- **Utilities**
  - **Threading:** `UiWrapper`, `UiTask`, `ThreadSafeSingleShotGuard`, timers (`TimerWrapper`, `Timed*`)
  - **SmartSerializable:** generic serialization/persistence helpers (+ Config Viewer)
  - **Collections:** `ConcurrentObservableDictionary<TKey,TValue>`
  - **Newtonsoft JSON converters** & path helpers
- **Ribbon**: Ribbon XML and callbacks (`_Ribbon/RibbonGsync.xml`, `RibbonGsync`, `DevelopmentMethods`)
- **AppGlobals**: App‑wide DI and bootstrapping (`OutlookApplication`, `FS`, `UI`, `StoresWrapper`)

---

## Requirements

- **Windows** with **Microsoft Outlook (Desktop)** installed
- **Visual Studio 2022** (17.x) with **Office/SharePoint development** workload (VSTO)
- **.NET Framework 4.8.1** developer pack
- **NuGet** package restore enabled
- (Optional, for tests) **MSTest.TestFramework** / **MSTest.TestAdapter**, **Moq** (restored via NuGet)

> The add‑in uses COM interop assemblies (e.g., `Microsoft.Office.Interop.Outlook`) and log4net. NuGet packages are restored on build where applicable (e.g., `SVGControl/packages.config`).

---

## Getting Started (Developer)

1. **Clone** the repository and open `Gsync.sln` in Visual Studio 2022.
2. Ensure **NuGet restore** is enabled (`Tools → NuGet Package Manager → Package Manager Settings → Restore`).
3. Set **Gsync** as the **startup project**.
4. **Build** the solution (Debug).
5. **Run (F5)**. Visual Studio will start Outlook with the add‑in loaded for debugging.
6. In Outlook, open the **Gsync** ribbon tab → **Development** group to run helper actions (e.g., **LoopInbox**).

> If Outlook doesn’t load the add‑in: check **File → Options → Add‑ins → COM Add‑ins**; verify the Gsync add‑in isn’t disabled. You may need VSTO runtime and to trust the solution on first run.

---

## Logging

- Logs are written under: **`%APPDATA%\Gsync\Logs`**
- Files include a timestamped **debug** log and **trace** log (method calls), configured via `log4net.config`.
- On startup, `ThisAddIn` ensures the log directory exists and points appender file paths to the current timestamp.

---

## Testing

- **Framework:** MSTest (attributes like `[TestClass]`, `[TestMethod]`), with **Moq** for mocking Outlook wrappers.
- **Scope:** wrappers (property forwarding, event wiring), comparers, detached snapshots, store discovery, threading utilities, SmartSerializable, JSON converters, collections.
- **Run:** open **Test Explorer** in Visual Studio and run all tests in the **Gsync.Test** project.

---

## Development Notes

- **Detached Pattern:** `Detached*` items decouple comparisons from live COM objects—safe to serialize, diff, and persist.
- **Comparers:** Use equality for strict identity; **similarity** comparers enable fuzzy matching for near‑duplicates.
- **Stores:** `StoresWrapper` hydrates all session stores; `StoreWrapper` provides account, inbox, and folder accessors.
- **Threading:** Use `UiTask`/`UiWrapper` to marshal work to the UI thread. Avoid long‑running operations on the Outlook UI thread.
- **Persistence:** `SmartSerializable<T>` centralizes JSON I/O with DI‑friendly file system and dialog interfaces (testable).

---

## Roadmap (Feature Increments)

1. **Snapshot & Diff Engine** for message change tracking (built on `Detached*` + comparers)
2. **Change Surface UI** (per‑item history, accept/revert flows)
3. **Background Scanning & Throttling** (timer‑backed with backoff)
4. **Retention & Persistence Policy** for snapshots/configs
5. **Diagnostics** (optional telemetry hooks; trace improvements)
6. **CI** integration for test coverage

---

## Repository Hygiene

- Generated/IDE files are excluded via `.gitignore`.
- Binary resources (e.g., SVGs) live under `SVGControl` and add‑in resource folders.
- Keep new code **layered**: Outlook COM types must **not** leak above wrapper abstractions.

---

## License

## License

**MIT License.** See [`LICENSE`](LICENSE) for the full text.
- Third-party license texts: see [`THIRD-PARTY-NOTICES.md`](THIRD-PARTY-NOTICES.md) and the `licenses/` folder.
- Suggested source-file header (C#):

  ```csharp
  // Copyright (c) 2025 Dan Moisan
  // Licensed under the MIT License. See LICENSE in the repo root for license information.
  ```


## Acknowledgments

- Built with Microsoft Office Interop, MSTest, Moq, Newtonsoft.Json, and custom SVG WinForms controls.
