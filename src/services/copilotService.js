// src/services/copilotService.js

import axios from 'axios';
import { graphConfig } from './authConfig';

/**
 * Microsoft 365 Copilot Chat API Service
 * 
 * This service handles communication with the M365 Copilot Chat API
 * API Reference: https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/chat/copilotroot-post-conversations
 */

class CopilotService {
  constructor() {
    this.conversationId = null;
  }

  /**
   * Create a new conversation with M365 Copilot
   * @param {string} accessToken - Bearer token from MSAL
   * @returns {Promise<string>} - Conversation ID
   */
  async createConversation(accessToken) {
    try {
      const response = await axios.post(
        graphConfig.copilotConversationsEndpoint,
        {},
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        }
      );

      this.conversationId = response.data.id;
      return this.conversationId;
    } catch (error) {
      console.error('Error creating conversation:', error.response?.data || error.message);
      throw new Error(`Failed to create conversation: ${error.response?.data?.error?.message || error.message}`);
    }
  }

  /**
   * Send a message to M365 Copilot
   * @param {string} accessToken - Bearer token
   * @param {string} message - User message
   * @param {boolean} isRetry - Internal flag to prevent infinite retry
   * @returns {Promise<object>} - Copilot response
   */
  async sendMessage(accessToken, message, isRetry = false) {
    try {
      if (!this.conversationId) {
        await this.createConversation(accessToken);
      }

      const response = await axios.post(
        `${graphConfig.copilotConversationsEndpoint}/${this.conversationId}/chat`,
        {
          message: { text: message },
          locationHint: {
            timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC',
          },
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        }
      );

      return response.data;
    } catch (error) {
      const status = error.response?.status;
      const errorMsg = error.response?.data?.error?.message || error.message;

      console.error('Error sending message:', errorMsg);

      // If conversation expired or invalid, reset and retry once
      if (status === 404 && !isRetry) {
        console.warn('Conversation ID may be invalid. Resetting and retrying...');
        this.conversationId = null;
        return this.sendMessage(accessToken, message, true);
      }

      throw new Error(`Failed to send message: ${errorMsg}`);
    }
  }

  /**
   * Reset the current conversation
   */
  resetConversation() {
    this.conversationId = null;
  }

  /**
   * Get conversation history (if needed)
   * @param {string} accessToken - Bearer token
   * @returns {Promise<object>} - Conversation history
   */
  async getConversation(accessToken) {
    if (!this.conversationId) {
      throw new Error('No active conversation');
    }

    try {
      const response = await axios.get(
        `${graphConfig.copilotConversationsEndpoint}/${this.conversationId}`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      return response.data;
    } catch (error) {
      console.error('Error getting conversation:', error.response?.data || error.message);
      throw new Error(`Failed to get conversation: ${error.response?.data?.error?.message || error.message}`);
    }
  }
}

export default new CopilotService();