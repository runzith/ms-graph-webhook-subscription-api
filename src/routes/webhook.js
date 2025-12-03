const express = require('express');
const router = express.Router();
const { v4: uuidv4 } = require('uuid');

// Handle validation requests (POST with validationToken and notifications process) 
router.post('/notifications', (req, res) => {
  try {
    const validationToken = req?.query?.validationToken;
    const notifications = req.body?.value;

    
    if(notifications){
      console.log('Post notifications api triggered for Notification',JSON.stringify(notifications));
      if (!notifications || !Array.isArray(notifications)) {
          return {
              statusCode: 400,
              body: JSON.stringify({ error: 'Invalid notification payload' })
          };
      }else{
        return {
              statusCode: 200,
              body: JSON.stringify(notifications)
          };
      }

    }
    if(validationToken){
      console.log('Post notifications api triggered for validationToken ',validationToken);
      return res.status(200).type('text/plain').send(validationToken);
    }
    
    return false

  } catch (error) {
    console.error('Error during validation:', error);
    res.status(500).json({ 
      error: 'Validation failed',
      message: error.message
    });
  }
});


/**
 * GET /api/webhook/list
 * 
 * Retrieve all received notifications (for testing/debugging)
 */
router.get('/list', (req, res) => {
  try {
    const limit = parseInt(req.query.limit) || 50;
    const offset = parseInt(req.query.offset) || 0;
    
    const paginatedNotifications = notifications
      .slice()
      .reverse() // Most recent first
      .slice(offset, offset + limit);
    
    res.json({
      total: notifications.length,
      limit,
      offset,
      notifications: paginatedNotifications
    });
  } catch (error) {
    console.error('❌ Error retrieving notifications:', error);
    res.status(500).json({ 
      error: 'Failed to retrieve notifications',
      message: error.message
    });
  }
});

/**
 * DELETE /api/webhook/clear
 * 
 * Clear all stored notifications
 */
router.delete('/clear', (req, res) => {
  try {
    const count = notifications.length;
    notifications.length = 0; // Clear array
    
    res.json({
      message: 'All notifications cleared',
      cleared: count,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    console.error(' Error clearing notifications:', error);
    res.status(500).json({ 
      error: 'Failed to clear notifications',
      message: error.message
    });
  }
});

module.exports = router;

