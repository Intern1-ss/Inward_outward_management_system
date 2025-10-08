# 📋 Document Management System

A comprehensive web-based Document Management System built with Google Apps Script for tracking and managing inward and outward correspondence in organizations.

## 🌟 Overview

This Document Management System is designed to streamline the process of tracking incoming and outgoing documents in educational institutions, offices, and organizations. Built on Google Apps Script and Google Sheets, it provides a modern, user-friendly interface for document tracking, linking related entries, and generating reports.

## ✨ Key Features

### 📥 Document Entry Management
- **Dual Entry Types**: Manage both Inward and Outward documents
- **Auto-generated Reference Numbers**: Automatic INW/YYYY/XXX and OTW/YYYY/XXX numbering
- **Rich Metadata**: Track sender/recipient, subject, means of communication, file references, and postal tariffs
- **Date/Time Tracking**: Automatic timestamp recording for all entries

### 🔗 Advanced Linking System
- **Bidirectional Linking**: Connect related documents across Inward/Outward entries
- **UUID-based Tracking**: Unique identifiers for link groups
- **Link Visualization**: See all connected documents at a glance
- **Smart Search**: Filter by linked/unlinked entries

### ✅ Workflow Management
- **Status Tracking**: Monitor document lifecycle (Incomplete → Ready → Confirmed)
- **Action Required Alerts**: Flag entries needing attention
- **Completion Confirmation**: Mark physical work as complete with notes
- **Pending Work Notifications**: Weekly email reports for pending tasks

### 🔍 Powerful Search & Filtering
- **Hover Dropdown**: Preview all entries on hover
- **Advanced Filters**: Search by type, status, links, UUID
- **Real-time Search**: Instant results as you type
- **Linked Entries View**: Display all interconnected documents

### 📊 Reporting & Analytics
- **Financial Reports**: Track postal expenditure with cross-referencing
- **Multiple Export Formats**: CSV, JSON, and text reports
- **Custom Date Ranges**: Generate reports for specific periods
- **Statistics Dashboard**: Visual overview of document status

### 🎨 Modern UI/UX
- **Dual View Modes**: Toggle between Card and Table views
- **Responsive Design**: Works on desktop and mobile devices
- **Clean Interface**: Intuitive design with minimal learning curve
- **Real-time Updates**: Instant dashboard refresh

### ⚡ Performance Optimization
- **Smart Caching**: 5-minute cache for frequently accessed data
- **Optimized Queries**: Column-based reading for faster operations
- **Batch Processing**: Efficient bulk operations

## 🛠️ Tech Stack

- **Backend**: Google Apps Script (JavaScript)
- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Database**: Google Sheets
- **Caching**: Google Apps Script Cache Service
- **Email**: Gmail API integration

## 📋 Prerequisites

- Google Account
- Access to Google Sheets
- Google Apps Script enabled
- Basic understanding of Google Workspace

## 🚀 Installation & Setup

### 1. Create Google Spreadsheet
```bash
1. Go to Google Sheets (sheets.google.com)
2. Create a new spreadsheet
3. Name it "Document Management System"
```

### 2. Set Up Google Apps Script

1. In your spreadsheet, click **Extensions** → **Apps Script**
2. Delete any existing code in `Code.gs`
3. Copy the entire content from `Code.gs` (first document) and paste it
4. Click **File** → **New** → **HTML file**
5. Name it `Index`
6. Copy the entire content from `Index.html` (second document) and paste it
7. Click **Save** (💾 icon)

### 3. Configure Settings

In `Code.gs`, update the configuration:

```javascript
const CONFIG = {
  BOSS_EMAIL: 'your-supervisor@example.com', // Update this
  NOTIFICATION_SUBJECT: "Inward/Outward Pending Report",
  // ... other settings
};
```

### 4. Deploy as Web App

1. Click **Deploy** → **New deployment**
2. Click the gear icon ⚙️ → Select **Web app**
3. Fill in the details:
   - **Description**: Document Management System v1.0
   - **Execute as**: Me
   - **Who has access**: Anyone with Google account (or customize as needed)
4. Click **Deploy**
5. **Authorize** the app when prompted
6. Copy the **Web app URL**

### 5. Set Up Weekly Email (Optional)

Run this function once to enable weekly pending reports:

1. In Apps Script, select `setupWeeklyEmailTrigger` from the function dropdown
2. Click **Run** (▶️)
3. Authorize if prompted

## 📖 Usage Guide

### Creating New Entries

1. **Inward Entry**:
   - Click "NEW INWARD"
   - Fill in required fields (Date/Time, Means, From Whom, Subject, Taken By)
   - Optionally add Action Taken, File Reference, Postal Tariff
   - Click "Create Entry"

2. **Outward Entry**:
   - Click "NEW OUTWARD"
   - Fill in required fields (Date/Time, Means, To Whom, Subject, Sent By)
   - Optionally add Due Date, File Reference, Postal Tariff
   - Click "Create Entry"

### Linking Entries

1. Click the **🔗 Link** button on any entry
2. Search for related entries
3. Select one or more entries to link
4. Click "Link Entries"
5. A unique UUID is generated for the link group

### Marking Work Complete

1. Ensure entry has all required data filled
2. For Inward: Add "Action Taken" before marking complete
3. Click **✅ Mark Complete**
4. Add optional completion notes
5. Confirm to mark as physically processed

### Searching Entries

- **Quick Search**: Hover over search bar to see all entries
- **Advanced Search**: 
  - Use search types (All/Inward/Outward/Linked/UUID)
  - Filter by status, links, date range
- **View Linked**: Click "All Linked" to see entries with connections

### Generating Reports

1. Click **💰 FINANCIAL REPORT**
2. Select  date range
3. Choose export format (CSV/JSON/Text)
4. Download or view online

## 📁 Project Structure

```
Document-Management-System/
├── Code.gs                 # Backend logic (Google Apps Script)
│   ├── Configuration
│   ├── Caching System
│   ├── Sheet Setup
│   ├── Entry Management (CRUD)
│   ├── Linking System
│   ├── Search Functions
│   ├── Reporting
│   └── Email Notifications
│
├── Index.html              # Frontend UI
│   ├── HTML Structure
│   ├── CSS Styling
│   └── JavaScript Logic
│
└── Google Sheets (Auto-created)
    ├── Inward Sheet        # Inward entries
    ├── Outward Sheet       # Outward entries
    ├── Confirmations       # Work completion logs
    └── Entry_Links         # Entry relationships
```

## ⚙️ Configuration Options

### Admin Users
Add admin emails in `Code.gs` to grant elevated permissions (if implementing admin features):

```javascript
CONFIG.ADMIN_USERS = ['admin@example.com', 'supervisor@example.com'];
```

### Cache Duration
Adjust cache timeout (in seconds):

```javascript
const CACHE_DURATION = 300; // 5 minutes
```

### Email Schedule
Modify trigger timing in `setupWeeklyEmailTrigger()`:

```javascript
ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTION_NAME)
  .timeBased()
  .everyWeeks(1)
  .onWeekDay(ScriptApp.WeekDay.SATURDAY) // Change day
  .atHour(11) // Change hour (24-hour format)
  .create();
```

## 🎯 Workflow States

```
Entry Lifecycle:
┌─────────────┐
│ INCOMPLETE  │ → Missing required fields
└──────┬──────┘
       ↓
┌─────────────┐
│ READY       │ → All fields filled (Inward needs Action Taken)
└──────┬──────┘
       ↓
┌─────────────┐
│ CONFIRMED   │ → Physical work marked complete
└─────────────┘
```

## 📊 Database Schema

### Inward Sheet
| Column | Description |
|--------|-------------|
| Sl. No | Serial number |
| Means | Communication method |
| Inward No | Auto-generated code |
| From Whom | Sender details |
| Subject | Document subject |
| Taken By | Receiver name |
| Date & Time | Entry timestamp |
| Action Taken | Action description |
| File Reference | File location |
| Postal Tariff | Postal charges |

### Outward Sheet
| Column | Description |
|--------|-------------|
| Sl. No | Serial number |
| Means | Communication method |
| Outward No | Auto-generated code |
| To Whom | Recipient details |
| Subject | Document subject |
| Sent By | Sender name |
| Date & Time | Entry timestamp |
| Case Closed | Status (Yes/No) |
| File Reference | File location |
| Postal Tariff | Postal charges |
| Due Date | Response deadline |

## 🔒 Security Considerations

- Data is stored in Google Sheets (inherits Google's security)
- User authentication via Google accounts
- Restrict deployment access based on organization needs
- Email addresses are logged for audit trails
- No sensitive data should be stored in plain text

## 🐛 Troubleshooting

### Common Issues

**1. "Google Apps Script not available" error**
- Ensure you've deployed as a Web App
- Check authorization permissions
- Clear browser cache and reload

**2. Entries not loading**
- Check Sheet names match CONFIG values exactly
- Verify columns are in correct order
- Check Apps Script execution logs (View → Logs)

**3. Search dropdown not showing**
- Wait a few seconds for initial load
- Check browser console for JavaScript errors
- Try refreshing the page

**4. Email notifications not working**
- Verify `BOSS_EMAIL` is set correctly
- Check trigger is installed: `Edit → Current project's triggers`
- Ensure Gmail API permissions are granted

## 🤝 Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Guidelines
- Follow existing code style and structure
- Comment complex logic
- Test thoroughly before submitting
- Update documentation for new features

## 📝 Future Enhancements

- [ ] Multi-user real-time collaboration
- [ ] Document file attachments (Google Drive integration)
- [ ] Advanced analytics dashboard
- [ ] Mobile app version
- [ ] Role-based access control (RBAC)
- [ ] Notification system (in-app alerts)
- [ ] Barcode/QR code generation for entries
- [ ] Integration with other Google Workspace apps

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Your Name**
- GitHub: [@yourusername](https://github.com/yourusername)
- Email: your.email@example.com

## 🙏 Acknowledgments

- Built for Controller of Examination section (Sri Sathya Sai Institute of Higher Learning) document management needs
- Inspired by traditional register-based tracking systems
- Google Apps Script documentation and community

## 📞 Support

For support, please:
1. Check the [Troubleshooting](#-troubleshooting) section
2. Open an issue on GitHub
3. Contact the maintainer at saisathyajain@sssihl.edu.in

---

**Made with ❤️ using Google Apps Script**
