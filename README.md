# Provident Packaging - Smart Email Classifier

Automatically organize your team's Outlook emails into categories: **Purchase Orders**, **Enquiries**, **Invoices**, and **Shipping** — with zero manual effort.

## What This Does

Every 15 minutes, this system:
1. Checks all team mailboxes for new unread emails
2. Reads each email and understands what it's about
3. Automatically applies a color-coded category in Outlook
4. Tracks everything for reporting and improvement

Your team sees organized emails without clicking anything.

## Quick Start

### 1. Prerequisites

- Python 3.11+
- Microsoft 365 admin account (for app registration)
- OpenAI API key

### 2. Install

```bash
git clone <your-repo>
cd provident-email-classifier
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configure

```bash
cp .env.example .env
```

Edit `.env` with your credentials:

```
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
OPENAI_API_KEY=sk-your-openai-key
```

### 4. Microsoft 365 Setup

1. Go to [Azure Portal](https://portal.azure.com)
2. **Azure Active Directory** → **App registrations** → **New registration**
3. Name: `ProvidentEmailClassifier`
4. Go to **API permissions** → **Add permission** → **Microsoft Graph** → **Application permissions**
5. Add: `Mail.Read`, `Mail.ReadWrite`, `MailboxSettings.ReadWrite`
6. Click **Grant admin consent for [your org]**
7. Go to **Certificates & secrets** → **New client secret** → copy the value

### 5. Run

```bash
python src/scheduler.py
```

The system will:
- Run immediately on startup
- Check every 15 minutes thereafter
- Log all activity to the console

### 6. Deploy (Production)

**Recommended: Render**
1. Push code to GitHub
2. Create new **Background Worker** on [Render](https://render.com)
3. Connect your repo
4. Set environment variables from `.env`
5. Deploy

**Alternative: Local Server**
```bash
# Run with PM2 (keeps it alive)
pm install -g pm2
pm2 start src/scheduler.py --name email-classifier
pm2 save
pm2 startup
```

## Architecture

```
Microsoft 365 Outlook ←→ Graph API ←→ Your Service (Python) ←→ OpenAI API
     ↑                                                              ↓
     └────────────── Color-coded emails ←── Classifications ───────┘
```

## File Structure

```
provident-email-classifier/
├── src/
│   ├── config.py          # Settings & environment variables
│   ├── classifier.py      # AI email classification (OpenAI)
│   ├── graph_client.py    # Microsoft 365 connection
│   ├── scheduler.py       # Main automation loop
│   └── database.py        # SQLite tracking & stats
├── data/
│   └── processed_emails.db # Auto-created SQLite database
├── .env                   # Your secrets (never commit!)
├── requirements.txt       # Python dependencies
└── README.md              # This file
```

## How It Works

### Classification Categories

| Category | Color | Trigger Words |
|----------|-------|---------------|
| **Purchase Order** | 🟢 Green | PO, purchase order, order confirmation, quantity |
| **Enquiry** | 🔵 Blue | Quote, RFQ, pricing, how much, samples |
| **Invoice** | 🔴 Red | Invoice, payment due, billing, amount due |
| **Shipping** | 🟡 Yellow | Tracking, delivery, shipment, freight, logistics |
| **General** | ⚪ Gray | Everything else |

### The AI

Uses **OpenAI GPT-4o-mini** to understand email context. Not just keyword matching — it understands intent.

Example:
- *"Can you send me pricing for 500 boxes?"* → **Enquiry** (even without the word "quote")
- *"Please find attached PO #44592"* → **Purchase Order** (understands "attached PO")

### Fallback

If the AI service is down, the system falls back to keyword matching so emails never go unclassified.

## Monitoring

### View Statistics

The SQLite database tracks everything:

```python
from database import Database
db = Database()

# Daily stats
print(db.get_daily_stats(days=7))

# Accuracy
print(db.get_accuracy())

# Misclassified examples (for improvement)
print(db.get_misclassified_examples(10))
```

### Logs

```bash
# View logs
tail -f /var/log/provident-classifier.log

# Or if using PM2
pm2 logs email-classifier
```

## Costs

| Component | Monthly Cost |
|-----------|-------------|
| OpenAI API (~200 emails/day) | ~$30 |
| Hosting (Render free tier) | $0 |
| **Total** | **~$30/month** |

## Troubleshooting

### "Authentication failed"
- Check `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` in `.env`
- Ensure admin consent was granted in Azure Portal

### "No emails found"
- Check that users have unread emails in the lookback window (default 30 min)
- Verify the app has `Mail.Read` permission

### "Categories not applied"
- Categories are created per-user in Outlook
- The system auto-creates them on first run
- Check Outlook web to see if categories exist

### "AI classification failed"
- Check `OPENAI_API_KEY` is valid
- System falls back to keywords automatically
- Check OpenAI API status: https://status.openai.com

## Security

- All secrets in `.env` (never committed to git)
- Uses Microsoft Graph API (official, secure)
- Client credentials flow (no user passwords stored)
- Email content processed in memory, not stored long-term
- Database only stores metadata (subject, sender, category) — not full email bodies

## Support

For issues or questions:
1. Check the logs
2. Verify environment variables
3. Review Microsoft Graph API documentation
4. Contact your developer

## License

Private — for Provident Packaging internal use.
