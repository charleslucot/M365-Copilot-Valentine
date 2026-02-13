🧠 Microsoft 365 Copilot Chat API Demo
This is a lightweight React app that integrates with the Microsoft 365 Copilot Chat API to enable conversational AI experiences grounded in your Microsoft 365 data (emails, meetings, files, etc.).
🚀 Built in ~3 hours as a proof-of-concept.
✨ Features

🔐 MSAL authentication with delegated Microsoft Graph permissions
💬 Multi-turn chat interface with Copilot
📡 Axios-based integration with the Microsoft Graph /copilot/conversations and /chat endpoints
🧠 Copilot responses grounded in your M365 context
🛠️ Error handling, retry logic, and debug logging

📸 Demo
🎥 [Insert screen recording link here]
🧰 Tech Stack

React (Vite)
MSAL.js (Microsoft Authentication Library)
Axios
Microsoft Graph API (beta)
Microsoft 365 Copilot Chat API (Preview)

🧪 Prerequisites

Microsoft 365 tenant with Copilot licenses assigned
Azure AD App Registration with the following delegated permissions:

User.Read
Sites.Read.All
Mail.Read
People.Read.All
OnlineMeetingTranscript.Read.All
Chat.Read
ChannelMessage.Read.All
ExternalItem.Read.All



🔧 Setup

Clone the repo:

git clone https://github.com/your-org/copilot-chat-demo.git
cd copilot-chat-demo

Install dependencies:

npm install

Create a .env file in the root directory:

VITE_CLIENT_ID=your-azure-app-client-id
VITE_TENANT_ID=your-tenant-id
VITE_REDIRECT_URI=http://localhost:3000
VITE_DEBUG=true

Start the dev server:

npm run dev

Open http://localhost:3000 in your browser and sign in with a licensed Microsoft 365 account.

🧠 How It Works

On login, the app uses MSAL to acquire an access token.
It creates a new Copilot conversation via POST to:
https://graph.microsoft.com/beta/copilot/conversations
Messages are sent to:
https://graph.microsoft.com/beta/copilot/conversations/{conversationId}/chat
Responses are parsed from the messages array and displayed in the chat UI.

🐞 Debugging

Enable debug mode by setting VITE_DEBUG=true in your .env file.
All API requests and responses will be logged to the browser console.
If you see “Sorry, I could not generate a response,” check the console for the full API response and ensure your prompt is supported.

📚 Resources

Microsoft 365 Copilot Chat API Overview (Preview)
Microsoft Graph API Docs
MSAL.js Documentation


Built with ❤️ and Copilot.
