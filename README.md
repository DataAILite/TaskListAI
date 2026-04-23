# TaskListAI

TaskListAI is an open-source AI-powered project management and data analysis system built with ASP.NET and VB.NET.

---

## 🚀 Features

- AI-driven data analysis
- Task and workflow management
- Interactive reports
- OpenAI-powered insights

---

## ⚖️ License

TaskListAI is licensed under the **GNU General Public License v3.0 (GPL v3)**.

You are free to:
- Use the software
- Modify the source code
- Distribute the software

Requirements:
- Any distributed modifications MUST also be open source under GPL v3
- No warranty is provided

Full license: https://www.gnu.org/licenses/gpl-3.0.html

---

## 🛠 Installation

### Option 1 – Installer (Recommended)

1. Download `InstallTaskListAI.zip` from http://DataAI.link
2. Extract the archive
3. Right-click `InstallTaskListAI.exe`
4. Select **Run as Administrator**
5. Follow the installer instructions

---

### Option 2 – Manual Installation

1. Download `TaskListAI.zip`
2. Extract to:

wwwroot\TaskListAI\

3. Open IIS Manager
4. Convert the folder into an Application
5. Set Application Pool Identity to Anonymous Authentication
6. Update web.config with your settings

---

## ⚙️ Configuration

Update `web.config` with the following:

### Operational Database (for tickets information)

- Server
- Database / Namespace
- User ID
- Password
- Port (if required)

> MySQL is recommended for the operational database

---
System Database (Oracle / InterSystems IRIS only)

- System database password

### Web Settings

- Application URL: https://[your-domain]/TaskListAI/
- Website title
- Upload folder path
- Support email

---
### Email (SMTP)

- Email address
- Email password

### 🤖 OpenAI Settings

Configure:

- API Key
- Organization ID
- Base URL
- Model: gpt-4o / gpt-4o-mini / o3 / o3-mini
- Max Tokens:
  - 128000 → gpt-4o / gpt-4o-mini
  - 200000 → o3 / o3-mini

---

## 🌐 Running the Application

Open in browser:

https://[your-domain]/TaskListAI/Default.aspx

---

## 🔧 IIS Configuration

Ensure:
- Application created in IIS
- Correct Application Pool
- Anonymous Authentication enabled

---

## 📚 Documentation

http://DataAI.link

---

## 🆘 Support

https://oureports.net/OUReports/ContactUs.aspx

---

## ⚠️ Database Licensing Notice

If you are using:

- Oracle Database
- InterSystems IRIS (or other InterSystems products)

You MUST have a valid license from the respective vendor.

TaskListAI does NOT provide database licenses and does NOT include Oracle or InterSystems software.

It is the user’s responsibility to:
- Obtain proper licenses
- Comply with vendor terms
- Ensure legal usage of those database systems

---

## 🔒 Disclaimer

This software is provided "AS IS", without warranty of any kind.

---

## 🤝 Contributing

By contributing, you agree your contributions are licensed under GPL v3.
