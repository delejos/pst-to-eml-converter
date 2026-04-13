# PST to EML Converter

A **free, open-source PST to EML converter for Windows**, built with .NET 10 and WPF.  
Convert Outlook PST files into standard EML files while preserving folder structure — **no paid licenses, no trials, no limitations**.

> This is a fork of [kgounaris/pst-to-eml-converter](https://github.com/kgounaris/pst-to-eml-converter) with significant improvements — see [What's changed in this fork](#whats-changed-in-this-fork) below.

---

## ✨ Features

- ✅ Convert **PST → EML** (emails with attachments)
- 📇 Export **Contacts → vCard (.vcf)**
- 📅 Export **Calendar items → iCal (.ics)**
- ✅ Export **Tasks → iCal (.ics)**
- 📂 Supports **single PST file** or **folder with multiple PSTs**
- 🗂️ Preserves Outlook folder structure
- 📊 Real-time per-item progress reporting
- ⛔ Cancel conversion at any time
- 🧾 Detailed conversion log written to output directory
- 🆓 **100% free & open source**
- 🪟 Windows desktop application (WPF)
- 🔒 Runs completely **offline**

---

## ✅ No Outlook Required

> **Microsoft Outlook does NOT need to be installed.**

This fork replaces the original Outlook COM dependency with [XstReader.Api](https://github.com/iluvadev/XstReader) — a pure .NET PST/OST reader with no dependency on any Microsoft Office component.

---

## 📥 Building from Source

**Requirements:** .NET 10 SDK ([download](https://aka.ms/dotnet/download))

```bash
git clone https://github.com/delejos/pst-to-eml-converter
cd pst-to-eml-converter
dotnet build PstToEmlConverter/PstToEmlConverter.csproj
```

The output will be in `PstToEmlConverter/bin/Debug/net10.0-windows/`.

---

## 🛠️ How to Use

1. Choose **Source**
   - Single PST file **or**
   - Folder containing multiple PST files
2. Choose **Destination folder**
3. Select **Options**:
   - Preserve folder structure
   - Skip existing files
   - Toggle which item types to export (Emails, Contacts, Calendar, Tasks)
4. Click **Start**
5. Watch live progress in the status bar and log

---

## 📁 What Gets Converted?

| Item Type       | Output format | Toggleable |
|-----------------|---------------|------------|
| Emails          | `.eml`        | Always on  |
| Attachments     | Embedded in `.eml` | —     |
| Contacts        | `.vcf` (vCard 3.0) | ✅ Yes |
| Calendar items  | `.ics` (iCal VEVENT) | ✅ Yes |
| Tasks           | `.ics` (iCal VTODO) | ✅ Yes |

---

## 🧾 Logging

- Live progress shown in the app: current folder, item counts, percentage
- A `conversion_log.txt` is written to the output directory
- Per-item errors are logged but do not stop the overall conversion

---

## What's changed in this fork

### 1. Removed Outlook dependency
The original used Microsoft Outlook's COM interface (`Microsoft.Office.Interop.Outlook`), requiring Outlook to be installed on the same machine. This fork replaces it entirely with **[XstReader.Api](https://www.nuget.org/packages/XstReader.Api)** — a pure .NET PST/OST reader written in C# with no Office dependency.

### 2. Contacts, Calendar, and Tasks export
The original only converted emails and skipped everything else. This fork adds:
- **Contacts** (`IPM.Contact`) → exported as **vCard 3.0** (`.vcf`), including email addresses extracted from MAPI named properties
- **Calendar appointments** (`IPM.Appointment`) → exported as **iCal VEVENT** (`.ics`) with subject, description, start time, and location
- **Tasks** (`IPM.Task`) → exported as **iCal VTODO** (`.ics`)

Each item type can be toggled independently in the UI.

### 3. Better progress reporting
The original progress bar only advanced once per PST file. This fork reports progress **per item**:
- Accurate percentage based on total item count (pre-counted before conversion starts)
- Live status bar showing current PST and folder being processed
- Running item counts: emails saved / contacts saved / calendar saved / tasks saved / failed

### 4. UI improvements
- New two-column Options section with item-type checkboxes
- Status bar showing current folder and item counts while running
- Monospace log font for better readability
- Styled buttons and cleaner layout

### 5. RFC-compliant EML output
The original hand-rolled its own EML serialisation. This fork uses **[MimeKit](https://www.nuget.org/packages/MimeKit)** — the industry-standard .NET MIME library — producing fully RFC 5322-compliant `.eml` files with proper headers, multipart bodies, and base64-encoded attachments.

### 6. Proper cancellation
Cancellation is checked at each item boundary via `CancellationToken`, so the app stops cleanly mid-conversion without leaving partially-written files.

---

## 📦 Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| [XstReader.Api](https://www.nuget.org/packages/XstReader.Api) | 1.0.6 | PST/OST file reading (no Outlook required) |
| [MimeKit](https://www.nuget.org/packages/MimeKit) | 4.15.1 | RFC-compliant EML writing |
| [Ical.Net](https://www.nuget.org/packages/Ical.Net) | 5.2.0 | iCal (.ics) writing for calendar/tasks |

---

## 💡 Why This Exists

Most PST converters are:
- Extremely expensive
- Locked behind trials
- Require Outlook to be installed
- Only export emails, not contacts or calendar

This fork aims to be a **simple, dependency-free, complete alternative** for one-time PST migrations.

---

## 📜 License

This project is licensed under the **MIT License**.  
You are free to use, modify, and distribute it.

See [`LICENSE`](LICENSE) for details.

---

## 🧑‍💻 Contributing

Contributions are welcome!

- Bug reports
- Feature requests
- Code improvements

Just open an issue or pull request.

---

**⭐ If you find this useful, please star the repository — it helps others find it!**
