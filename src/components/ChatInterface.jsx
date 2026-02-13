// src/components/ChatInterface.jsx

import React, { useState, useEffect, useRef } from 'react';
import { useMsal, useAccount } from '@azure/msal-react';
import { loginRequest } from '../services/authConfig';
import copilotService from '../services/copilotService';
import LoginButton from './LoginButton';

function ChatInterface() {
  const { instance, accounts } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [messages, setMessages] = useState([]);
  const [inputMessage, setInputMessage] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const messagesEndRef = useRef(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const getAccessToken = async () => {
    try {
      const response = await instance.acquireTokenSilent({
        ...loginRequest,
        account: account,
      });
      return response.accessToken;
    } catch (error) {
      console.error('Token acquisition failed:', error);
      try {
        const response = await instance.acquireTokenPopup(loginRequest);
        return response.accessToken;
      } catch (popupError) {
        throw new Error('Failed to acquire access token');
      }
    }
  };

  /**
   * Extract Copilot's reply from the conversation response
   */
  const extractCopilotReply = (conversationData) => {
    console.log('🔍 Full API Response:', conversationData);

    const messages = conversationData.messages || [];
    console.log('📬 Messages array:', messages);

    // Find the last message from Copilot (responseMessage type)
    const copilotMsg = messages
      .filter(msg => msg['@odata.type']?.includes('copilotConversationResponseMessage'))
      .pop();

    console.log('🤖 Copilot message object:', copilotMsg);

    const replyText = copilotMsg?.text || null;
    console.log('💬 Extracted reply text:', replyText);

    return replyText;
  };

  const handleSendMessage = async (e) => {
    e.preventDefault();
    if (!inputMessage.trim() || isLoading) return;

    const userMessage = inputMessage.trim();
    setInputMessage('');
    setError(null);

    // Add user message to chat
    const newUserMessage = {
      id: Date.now(),
      role: 'user',
      content: userMessage,
      timestamp: new Date(),
    };
    setMessages((prev) => [...prev, newUserMessage]);
    setIsLoading(true);

    try {
      console.log('📤 Sending message to Copilot:', userMessage);

      const accessToken = await getAccessToken();
      const response = await copilotService.sendMessage(accessToken, userMessage);

      console.log('📥 Raw Copilot response:', response);

      // Extract Copilot's reply from the response
      const replyText = extractCopilotReply(response);

      if (!replyText) {
        console.warn('⚠️ No reply text found in response. Full response:', response);
        throw new Error('Copilot did not provide a response. This could be due to:\n- Insufficient data to answer your question\n- Missing permissions\n- Service timeout\n\nTry rephrasing your question or asking about specific M365 data (emails, meetings, files).');
      }

      // Add Copilot response to chat
      const copilotMessage = {
        id: Date.now() + 1,
        role: 'assistant',
        content: replyText,
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, copilotMessage]);

      console.log('✅ Successfully added Copilot reply to chat');

    } catch (err) {
      console.error('❌ Error in handleSendMessage:', err);
      setError(err.message || 'Failed to send message. Please try again.');

      // Add error message to chat
      const errorMessage = {
        id: Date.now() + 1,
        role: 'error',
        content: err.message || 'Something went wrong. Please try again.',
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  };

  const handleNewConversation = () => {
    copilotService.resetConversation();
    setMessages([]);
    setError(null);
    console.log('🔄 Conversation reset');
  };

  const formatTime = (date) => {
    return date.toLocaleTimeString('en-US', {
      hour: 'numeric',
      minute: '2-digit',
      hour12: true,
    });
  };

  if (!account) {
    return (
      <div className="chat-interface">
        <div className="welcome-section">
          <h2 className="welcome-title">Welcome to M365 Copilot 💝</h2>
          <p className="welcome-text">
            Sign in to start chatting with your AI companion
          </p>
          <LoginButton />
        </div>
      </div>
    );
  }

  return (
    <div className="chat-interface">
      <div className="chat-header">
        <LoginButton />
        {messages.length > 0 && (
          <button onClick={handleNewConversation} className="valentine-button new-chat-button">
            <span className="button-icon">✨</span>
            New Chat
          </button>
        )}
      </div>

      <div className="chat-messages">
        {messages.length === 0 ? (
          <div className="empty-state">
            <div className="empty-icon">💬</div>
            <h3>Start a conversation</h3>
            <p>Ask me anything about your M365 data!</p>
            <div className="suggestion-chips">
              <button
                onClick={() => setInputMessage('What meetings do I have today?')}
                className="suggestion-chip"
              >
                📅 My meetings
              </button>
              <button
                onClick={() => setInputMessage('Summarize my recent emails')}
                className="suggestion-chip"
              >
                📧 Recent emails
              </button>
              <button
                onClick={() => setInputMessage('What files did I work on this week?')}
                className="suggestion-chip"
              >
                📁 My files
              </button>
            </div>
          </div>
        ) : (
          <>
            {messages.map((message) => (
              <div key={message.id} className={`message message-${message.role}`}>
                <div className="message-avatar">
                  {message.role === 'user' ? '👤' : message.role === 'error' ? '⚠️' : '🤖'}
                </div>
                <div className="message-content">
                  <div className="message-text">{message.content}</div>
                  <div className="message-time">{formatTime(message.timestamp)}</div>
                </div>
              </div>
            ))}
            {isLoading && (
              <div className="message message-assistant">
                <div className="message-avatar">🤖</div>
                <div className="message-content">
                  <div className="typing-indicator">
                    <span></span>
                    <span></span>
                    <span></span>
                  </div>
                </div>
              </div>
            )}
            <div ref={messagesEndRef} />
          </>
        )}
      </div>

      {error && (
        <div className="error-banner">
          ⚠️ {error}
        </div>
      )}

      <form onSubmit={handleSendMessage} className="chat-input-form">
        <div className="input-wrapper">
          <input
            type="text"
            value={inputMessage}
            onChange={(e) => setInputMessage(e.target.value)}
            placeholder="Ask me anything... 💭"
            className="chat-input"
            disabled={isLoading}
          />
          <button
            type="submit"
            className="send-button"
            disabled={!inputMessage.trim() || isLoading}
          >
            <span className="send-icon">💌</span>
          </button>
        </div>
      </form>
    </div>
  );
}

export default ChatInterface;