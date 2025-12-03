require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const webhookRoutes = require('./routes/webhook');
const subscriptionRoutes = require('./routes/subscription');

const app = express();
const PORT = process.env.PORT || 3004;

// Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Routes
app.use('/api/webhook/', webhookRoutes);
app.use('/api/subscription', subscriptionRoutes);

// Health check endpoint
app.get('/health', (req, res) => {
  res.status(200).json({ 
    status: 'OK', 
    message: 'Graph Webhook Test API is running',
    timestamp: new Date().toISOString()
  });
});

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    message: 'Microsoft Graph Webhook Test API',
    description: 'Two APIs for Microsoft Graph webhook testing',
    apis: {
      'API 1 - Validation': {
        endpoint: '/api/webhook/validate',
        method: 'POST',
        description: 'Called by Graph to validate webhook URL (returns 200 with token)',
        usage: 'Microsoft Graph calls this during subscription creation'
      },
      'API 2 - Notifications': {
        endpoint: '/api/webhook/notifications',
        method: 'POST',
        description: 'Receives change notifications from Graph (returns 202 Accepted)',
        usage: 'Microsoft Graph calls this when changes occur'
      }
    },
    additionalEndpoints: {
      health: '/health',
      getNotifications: 'GET /api/webhook/notifications',
      createSubscription: 'POST /api/subscription/create',
      listSubscriptions: 'GET /api/subscription/list',
      deleteSubscription: 'DELETE /api/subscription/delete/:id'
    }
  });
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Error:', err);
  res.status(err.status || 500).json({
    error: err.message || 'Internal Server Error',
    timestamp: new Date().toISOString()
  });
});

// Start server
app.listen(PORT, () => {
  console.log(`Graph Webhook Test API running on port ${PORT}`);
});

module.exports = app;

