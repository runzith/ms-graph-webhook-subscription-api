const express = require('express');
const router = express.Router();
const axios = require('axios');
const { v4: uuidv4 } = require('uuid');

// Store subscriptions in memory (for testing purposes)
const subscriptions = [];

/**
 * Helper function to get Microsoft Graph access token
 */
async function getAccessToken() {
  // const { GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID } = process.env;
  
  // if (!GRAPH_CLIENT_ID || !GRAPH_CLIENT_SECRET || !GRAPH_TENANT_ID) {
  //   throw new Error('Missing required environment variables for Graph API authentication');
  // }

  try {
    // const tokenEndpoint = `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`;
    
    // const params = new URLSearchParams();
    // params.append('client_id', GRAPH_CLIENT_ID);
    // params.append('client_secret', GRAPH_CLIENT_SECRET);
    // params.append('scope', 'https://graph.microsoft.com/.default');
    // params.append('grant_type', 'client_credentials');

    // const response = await axios.post(tokenEndpoint, params, {
    //   headers: {
    //     'Content-Type': 'application/x-www-form-urlencoded'
    //   }
    // });

    return ''; // get from KFA / hardcode
  } catch (error) {
    console.error('Failed to get access token:', error.response?.data || error.message);
    throw new Error('Failed to authenticate with Microsoft Graph');
  }
}

/**
 * API 2: Subscription Management
 * POST /api/subscription/create
 * 
 * Creates a new webhook subscription with Microsoft Graph.
 * This allows you to subscribe to changes in various Microsoft 365 resources.
 */
router.post('/create', async (req, res) => {
  try {
        const accessToken = await getAccessToken();
        
        if (!accessToken) {
            throw new Error('Failed to obtain access token');
        }
        const tunnelUrl = process.env.WEBHOOK_TUNNEL_URL || `https://joan-priced-feel-agencies.trycloudflare.com`;
        
        const clientState = uuidv4();
        const webhookUrl = `${tunnelUrl}/api/webhook/notifications`;
        const expirationDateTime = new Date();
        expirationDateTime.setDate(expirationDateTime.getDate() + 1);
        const subscriptionRequest = {
            changeType: "updated, created",
            notificationUrl: webhookUrl,
            resource: `/users/${userEmail}/events`,
            expirationDateTime: expirationDateTime.toISOString(),
            clientState: clientState
        };

        const response = await axios.post(
            `https://graph.microsoft.com/v1.0/subscriptions`,
            subscriptionRequest,
            {
                headers: {
                    "Authorization": `Bearer ${accessToken}`,
                    "Content-Type": "application/json"
                }
            }
        );

        const subscription = response.data;

        return subscription;

    } catch (error) {
        
        throw {
            code: error.response?.status || 400,
            message: `Error creating webhook subscription: ${error.message}`,
            detailError: error.response?.data,
            error
        };
    }
});

/**
 * GET /api/subscription/list
 * 
 * List all active subscriptions from Microsoft Graph
 */
router.get('/list', async (req, res) => {
  try {
    const accessToken = await getAccessToken();

    const response = await axios.get(
      'https://graph.microsoft.com/v1.0/subscriptions',
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        }
      }
    );

    const activeSubscriptions = response.data.value;

    res.json({
      message: 'Subscriptions retrieved successfully',
      count: activeSubscriptions.length,
      subscriptions: activeSubscriptions,
      localCache: subscriptions
    });

  } catch (error) {
    console.error('Error listing subscriptions:', error.response?.data || error.message);
    res.status(error.response?.status || 500).json({
      error: 'Failed to list subscriptions',
      message: error.response?.data?.error?.message || error.message
    });
  }
});

/**
 * DELETE /api/subscription/delete/:id
 * 
 * Delete a specific subscription by ID
 */
router.delete('/delete/:id', async (req, res) => {
  try {
    const { id } = req.params;
    

    if (!id) {
      return res.status(400).json({
        error: 'Subscription ID is required'
      });
    }

    const accessToken = await getAccessToken();

    await axios.delete(
      `https://graph.microsoft.com/v1.0/subscriptions/${id}`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        }
      }
    );

    // Remove from local cache
    const index = subscriptions.findIndex(sub => sub.id === id);
    if (index !== -1) {
      subscriptions.splice(index, 1);
    }

    console.log('Subscription deleted:', id);

    res.json({
      message: 'Subscription deleted successfully',
      subscriptionId: id,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('Error deleting subscription:', error.response?.data || error.message);
    res.status(error.response?.status || 500).json({
      error: 'Failed to delete subscription',
      message: error.response?.data?.error?.message || error.message
    });
  }
});

/**
 * PATCH /api/subscription/renew/:id
 * 
 * Renew/extend a subscription's expiration time
 */
router.patch('/renew/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { expirationDateTime } = req.body;

    if (!id) {
      return res.status(400).json({
        error: 'Subscription ID is required'
      });
    }

    const accessToken = await getAccessToken();

    // Calculate new expiration (default: 3 days from now)
    const newExpiration = expirationDateTime || 
      new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();

    const response = await axios.patch(
      `https://graph.microsoft.com/v1.0/subscriptions/${id}`,
      {
        expirationDateTime: newExpiration
      },
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );

    console.log('Subscription renewed:', id);
    console.log('   New expiration:', newExpiration);

    res.json({
      message: 'Subscription renewed successfully',
      subscription: response.data
    });

  } catch (error) {
    console.error('Error renewing subscription:', error.response?.data || error.message);
    res.status(error.response?.status || 500).json({
      error: 'Failed to renew subscription',
      message: error.response?.data?.error?.message || error.message
    });
  }
});

module.exports = router;


