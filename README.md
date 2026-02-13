# 🧠 Microsoft 365 Copilot Chat API Demo

This is a lightweight React app that integrates with the Microsoft 365 Copilot Chat API to enable conversational AI experiences grounded in your Microsoft 365 data (emails, meetings, files, etc.).

🚀 Built in ~3 hours as a proof-of-concept.

## ✨ Features

- 🔐 MSAL authentication with delegated Microsoft Graph permissions
- 💬 Multi-turn chat interface with Copilot
- 📡 Axios-based integration with the Microsoft Graph `/copilot/conversations` and `/chat` endpoints
- 🧠 Copilot responses grounded in your M365 context
- 🛠️ Error handling, retry logic, and debug logging

## 📸 Demo

> [Insert screen recording link here]

## 🧰 Tech Stack

- React (Vite)
- MSAL.js (Microsoft Authentication Library)
- Axios
- Microsoft Graph API (beta)
- Microsoft 365 Copilot Chat API (Preview)

## 🧪 Prerequisites

- Microsoft 365 tenant with Copilot licenses assigned
- Azure AD App Registration with the following delegated permissions:
  - User.Read
  - Sites.Read.All
  - Mail.Read
  - People.Read.All
  - OnlineMeetingTranscript.Read.All
  - Chat.Read
  - ChannelMessage.Read.All
  - ExternalItem.Read.All

## 🔧 Setup

1. Clone the repo:

```bash
git clone https://github.com/your-org/copilot-chat-demo.git
cd copilot-chat-demo
