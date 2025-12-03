# Microsoft Graph Webhook APIs - Complete Guide

## Overview

This project provides **TWO main APIs** that Microsoft Graph calls:

1. **Validation API** - Graph calls this to validate your webhook URL (returns **200 OK**)
2. **Notification API** - Graph calls this to send change notifications (returns **202 Accepted**)

---

## API 1: Webhook Validation Endpoint

### Purpose
Microsoft Graph calls this endpoint to validate that your webhook URL is valid and accessible **before creating a subscription**.

### Endpoint
```
POST /api/webhook/validate
```

### How Microsoft Graph Calls It
When you create a subscription, Microsoft Graph will:
1. Make a POST request to your `notificationUrl` 
2. Include a `validationToken` query parameter
3. Expect you to return the token with HTTP 200 status

### Request from Microsoft Graph
```http
POST https://your-domain.com/api/webhook/validate?validationToken=abc123xyz...
Content-Type: application/json
```

### Your Response (Required)
```
HTTP/1.1 200 OK
Content-Type: text/plain

abc123xyz...
```

**Important Requirements:**
- ✅ Must return **HTTP 200 OK** status code
- ✅ Must return **plain text** (not JSON)
- ✅ Response body must contain the exact validation token
- ✅ Must respond within 10 seconds

### Example Test
```bash
# Test the validation endpoint
curl -X POST "http://localhost:3000/api/webhook/validate?validationToken=test-token-123"

# Expected response (plain text):
test-token-123
```

### What Happens in the Code
```javascript
router.post('/validate', (req, res) => {
  const validationToken = req.query.validationToken;
  
  // Return 200 OK with plain text
  res.status(200)
     .type('text/plain')
     .send(validationToken);
});
```

---

## API 2: Webhook Notification Receiver

### Purpose
Microsoft Graph calls this endpoint to send you notifications when changes occur in the resources you subscribed to.

### Endpoint
```
POST /api/webhook/notifications
```

### How Microsoft Graph Calls It
When a change occurs (email received, calendar event updated, etc.):
1. Microsoft Graph makes a POST request to your `notificationUrl`
2. Sends an array of notifications in the request body
3. Expects you to return HTTP 202 Accepted within 3 seconds

### Request from Microsoft Graph
```http
POST https://your-domain.com/api/webhook/notifications
Content-Type: application/json

{
  "value": [
    {
      "subscriptionId": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
      "changeType": "created",
      "resource": "me/mailFolders('Inbox')/messages/AAMkAGI2...",
      "resourceData": {
        "@odata.type": "#Microsoft.Graph.Message",
        "@odata.id": "Users/user@example.com/Messages/AAMkAGI2...",
        "id": "AAMkAGI2..."
      },
      "clientState": "your-secret-client-state",
      "subscriptionExpirationDateTime": "2024-01-20T18:23:45.9356913Z",
      "tenantId": "84bd8158-6d4d-4958-8b9f-9d6445542f95"
    }
  ]
}
```

### Your Response (Required)
```http
HTTP/1.1 202 Accepted
Content-Type: application/json

{
  "message": "Notifications received and processed successfully",
  "count": 1,
  "timestamp": "2024-01-15T10:30:00.000Z"
}
```

**Important Requirements:**
- ✅ Must return **HTTP 202 Accepted** status code
- ✅ Must respond within **3 seconds**
- ✅ Response body is optional (can be empty)
- ✅ Process notifications asynchronously if needed

### Notification Properties

| Property | Description |
|----------|-------------|
| `subscriptionId` | The ID of the subscription that triggered this notification |
| `changeType` | Type of change: `created`, `updated`, or `deleted` |
| `resource` | The resource that changed (e.g., message, event, user) |
| `resourceData` | Contains the resource type and ID |
| `clientState` | The secret value you provided when creating the subscription |
| `subscriptionExpirationDateTime` | When the subscription expires |
| `tenantId` | The Azure AD tenant ID |

### Example Test
```bash
# Test the notification endpoint
curl -X POST http://localhost:3000/api/webhook/notifications \
  -H "Content-Type: application/json" \
  -d '{
    "value": [
      {
        "subscriptionId": "test-sub-123",
        "changeType": "created",
        "resource": "me/messages/msg-001",
        "resourceData": {
          "@odata.type": "#Microsoft.Graph.Message",
          "id": "msg-001"
        },
        "clientState": "my-secret"
      }
    ]
  }'

# Expected response:
# HTTP 202 Accepted
# {"message":"Notifications received and processed successfully","count":1,...}
```

### What Happens in the Code
```javascript
router.post('/notifications', (req, res) => {
  const { value } = req.body;
  
  // Process each notification
  value.forEach((notification) => {
    // Store notification
    notifications.push({
      id: uuidv4(),
      receivedAt: new Date().toISOString(),
      ...notification
    });
    
    console.log('Received:', notification.changeType, notification.resource);
  });
  
  // Return 202 Accepted
  res.status(202).json({ 
    message: 'Notifications received and processed successfully',
    count: value.length
  });
});
```

---

## Complete Flow: How It All Works Together

### Step 1: Create a Subscription
You call Microsoft Graph to create a subscription:

```bash
POST https://graph.microsoft.com/v1.0/subscriptions
Authorization: Bearer {access-token}
Content-Type: application/json

{
  "changeType": "created,updated",
  "notificationUrl": "https://your-domain.com/api/webhook/validate",
  "resource": "me/mailFolders('Inbox')/messages",
  "expirationDateTime": "2024-01-20T18:23:45.9356913Z",
  "clientState": "my-secret-state"
}
```

### Step 2: Microsoft Graph Validates Your Endpoint
Microsoft Graph immediately calls your validation endpoint:

```
POST https://your-domain.com/api/webhook/validate?validationToken=abc123...
```

**Your API 1 responds:**
```
HTTP 200 OK
Content-Type: text/plain

abc123...
```

### Step 3: Subscription is Created
If validation succeeds, Microsoft Graph creates the subscription and returns:

```json
{
  "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
  "resource": "me/mailFolders('Inbox')/messages",
  "changeType": "created,updated",
  "notificationUrl": "https://your-domain.com/api/webhook/notifications",
  ...
}
```

**Note:** The `notificationUrl` should now point to your notifications endpoint!

### Step 4: Changes Occur
When an email arrives or is updated, Microsoft Graph calls your notification endpoint:

```
POST https://your-domain.com/api/webhook/notifications
```

**Your API 2 responds:**
```
HTTP 202 Accepted
```

---

## Important Notes

### URL Configuration

When creating a subscription, you have two options:

**Option 1: Use the same URL for both (recommended for simplicity)**
```json
{
  "notificationUrl": "https://your-domain.com/api/webhook/notifications"
}
```
- During validation, Graph adds `?validationToken=...` to this URL
- Your code checks for the query parameter and handles validation
- After validation, Graph uses the same URL for notifications

**Option 2: Use separate URLs (clearer separation)**
```json
{
  "notificationUrl": "https://your-domain.com/api/webhook/validate"
}
```
- Use `/validate` endpoint during subscription creation
- Then update the subscription to use `/notifications` endpoint
- This project supports both approaches!

### Response Time Requirements

| Endpoint | Max Response Time |
|----------|------------------|
| Validation | 10 seconds |
| Notifications | 3 seconds |

### Status Codes

| Endpoint | Success Code | Purpose |
|----------|-------------|---------|
| Validation | **200 OK** | Confirms URL is valid |
| Notifications | **202 Accepted** | Acknowledges receipt |

### Security Best Practices

1. **Validate clientState**: Check that the `clientState` in notifications matches what you set
2. **Use HTTPS**: Microsoft Graph requires HTTPS in production
3. **Verify tenant**: Check the `tenantId` if you serve multiple tenants
4. **Process async**: Return 202 quickly, process notifications in background

---

## Testing the APIs

### Test Validation API
```bash
curl -X POST "http://localhost:3000/api/webhook/validate?validationToken=test-123"
```

Expected output (plain text):
```
test-123
```

### Test Notification API
```bash
curl -X POST http://localhost:3000/api/webhook/notifications \
  -H "Content-Type: application/json" \
  -d '{
    "value": [{
      "subscriptionId": "test",
      "changeType": "created",
      "resource": "test/resource",
      "resourceData": {"id": "123"}
    }]
  }'
```

Expected output (JSON):
```json
{
  "message": "Notifications received and processed successfully",
  "count": 1,
  "timestamp": "2024-01-15T10:30:00.000Z"
}
```

### View Received Notifications
```bash
curl http://localhost:3000/api/webhook/notifications
```

---

## Common Issues and Solutions

### Issue: Validation fails
**Cause**: Not returning 200 status or plain text
**Solution**: Ensure you return `res.status(200).type('text/plain').send(token)`

### Issue: Notifications not received
**Cause**: Subscription expired or URL not accessible
**Solution**: 
- Check subscription is active: `GET /api/subscription/list`
- Ensure your URL is publicly accessible (use ngrok for local testing)
- Verify you're returning 202 within 3 seconds

### Issue: "Invalid notificationUrl"
**Cause**: URL not publicly accessible or HTTPS required
**Solution**: Use ngrok for local testing: `ngrok http 3000`

---

## Summary

✅ **API 1 (Validation)**: `/api/webhook/validate`
- Called by Graph during subscription creation
- Returns **200 OK** with validation token as plain text
- Must respond within 10 seconds

✅ **API 2 (Notifications)**: `/api/webhook/notifications`
- Called by Graph when changes occur
- Returns **202 Accepted**
- Must respond within 3 seconds

Both APIs are ready to use and fully compliant with Microsoft Graph requirements!


