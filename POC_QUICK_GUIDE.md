# Microsoft Graph Webhook POC - Quick Implementation Guide
##  Table of Contents

1. [What This POC Does](#what-this-poc-does)
2. [Complete Implementation Flow](#complete-implementation-flow)
3. [Step-by-Step Implementation](#step-by-step-implementation)
4. [Testing the Implementation](#testing-the-implementation)

---

## What This POC Does

Track **meeting responses in real-time** using Microsoft Graph webhooks. Get instant notifications when attendees accept, decline, or respond "maybe" to meeting invitations.

**Flow**: `Create Meeting → Setup Webhook → Attendee Responds → Instant Notification → Process & Store`

---

## Complete Implementation Flow

### 📌 Phase 1: Event Creation & Subscription Setup

```
1. CREATE EVENT IN YOUR APP
   ↓
2. CHECK: Does webhook subscription exist for this user?
   - Query database: SELECT * FROM Subscriptions WHERE userEmail = ? AND isActive = true
   ↓
3. IF NOT EXISTS → CREATE WEBHOOK SUBSCRIPTION
   ↓
4. OBTAIN ACCESS TOKEN
   - Organization App Token (client_credentials)
   - OR User Delegated Token
   - OR Stored Token from KF Auth Service
   ↓
5. PREPARE WEBHOOK CALLBACK API
   - POST /api/webhook/notifications must be running
   - Endpoint must be publicly accessible (HTTPS)
   ↓
6. EXPOSE LOCAL URL (For Testing)
   - Run: cloudflared tunnel --url http://localhost:3004
   - Copy public URL: https://your-tunnel.trycloudflare.com
   ↓
7. VERIFY PUBLICLY ACCESSIBLE
   - Test: https://your-tunnel.com/health
   ↓
8. PREPARE SUBSCRIPTION PAYLOAD
   const subscriptionRequest = {
     changeType: "updated,created",
     notificationUrl: "https://your-tunnel.com/api/webhook/notifications",
     resource: `/users/${userEmail}/events`,
     expirationDateTime: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString(),
     clientState: uuidv4()
   };
   ↓
9. TRIGGER MICROSOFT GRAPH API
   POST https://graph.microsoft.com/v1.0/subscriptions
   Headers:
   - Authorization: Bearer {accessToken}
   - Content-Type: application/json
   ↓
10. MICROSOFT VALIDATES WEBHOOK
    - Graph sends POST with ?validationToken=xxx
    - Your API MUST return the token as plain text with 200 status
    ↓
11. RECEIVE SUBSCRIPTION RESPONSE
    {
      "id": "31cdd544-75ce-433d-9e15-d586458f77ec",
      "resource": "/users/email@domain.com/events",
      "changeType": "created,updated",
      "clientState": "42b6562f-112c-4606-8a8f-40ec503a4d3f",
      "notificationUrl": "https://your-url.com/api/webhook/notifications",
      "expirationDateTime": "2025-12-10T12:05:02.263Z"
    }
    ↓
12. STORE SUBSCRIPTION IN DATABASE
    - subscriptionId
    - userEmail
    - expirationDateTime
    - clientState
    - isActive = true
    
    WHY: Avoid repeated Graph API calls to check subscription status
    ↓
✅ WEBHOOK SUBSCRIPTION NOW ACTIVE
```

---

### 📌 Phase 2: Real-Time Notification Processing

```
13. USER ACCEPTS/DECLINES MEETING IN OUTLOOK
    - User clicks Accept/Decline/Tentative
    - OR Organizer updates meeting details
    ↓
14. MICROSOFT GRAPH DETECTS CHANGE
    - Change matches subscription filter: /users/{userId}/events
    - Change type: "updated" or "created"
    ↓
15. GRAPH SENDS WEBHOOK NOTIFICATION
    POST /api/webhook/notifications
    [{
      "subscriptionId": "31cdd544-75ce-433d-9e15-d586458f77ec",
      "changeType": "updated",
      "resource": "Users/7b98a0ea-.../Events/AAMkADdm...",
      "resourceData": {
        "@odata.type": "#Microsoft.Graph.Event",
        "id": "AAMkADdmZmEwYjA0LTNjZGUtNGM4Ny05NTYw..."
      },
      "clientState": "42b6562f-112c-4606-8a8f-40ec503a4d3f",
      "tenantId": "e9d21387-43f1-4e06-a253-f9ed9096dc48"
    }]
    ↓
16. YOUR API RECEIVES NOTIFICATION
    - Returns 202 Accepted IMMEDIATELY (< 30 seconds required)
    - Process asynchronously
    ↓
17. VALIDATE CLIENT STATE (Security)
    - Compare notification.clientState with stored value in DB
    - IF mismatch → REJECT (possible security threat)
    ↓
18. EXTRACT EVENT ID
    - eventId = notification.resourceData.id
    ↓
19. FETCH FULL EVENT DETAILS FROM GRAPH API
    GET https://graph.microsoft.com/v1.0/me/events/{eventId}
    
    Response includes attendees with status:
    {
      "attendees": [{
        "emailAddress": { "address": "john@company.com" },
        "status": {
          "response": "accepted",  // accepted, declined, tentativelyAccepted
          "time": "2025-12-03T10:30:00Z"
        }
      }]
    }
    ↓
20. RETRIEVE STORED EVENT DATA
    - Query by eventId from your database
    - Get previous attendee response states
    ↓
21. COMPARE & DETECT CHANGES
    - Compare old response vs new response for each attendee
    - Build change log:
      * John Doe: none → accepted
      * Jane Smith: tentativelyAccepted → declined
    ↓
22. UPDATE DATABASE
    - Insert/Update MeetingResponses table
    - Log response changes
    - Mark timestamp
    ↓
23. SEND NOTIFICATIONS TO STAKEHOLDERS
    - Email to event organizer
    - Notify schedulers/managers
    - Update dashboard (WebSocket/real-time)
    ↓
24. LOG NOTIFICATION SUCCESS
    - Store in NotificationLogs table
    - Mark notificationSent = true
    ↓
✅ RESPONSE TRACKED & NOTIFICATIONS SENT
```

---

### 📌 Phase 3: Subscription Lifecycle (Auto-Renewal)

```
25. SUBSCRIPTION EXPIRATION MONITORING
    - Background job runs daily
    - Query subscriptions expiring in next 2 days
    ↓
26. AUTOMATIC SUBSCRIPTION RENEWAL
    PATCH https://graph.microsoft.com/v1.0/subscriptions/{id}
    { "expirationDateTime": "2025-12-11T12:05:02.263Z" }
    
    Update database with new expiration date
    
    NOTE: Calendar subscriptions max lifetime = 7 days (4230 minutes)
```


### 2. Environment Configuration

```env
# Microsoft Graph API
# Webhook
WEBHOOK_TUNNEL_URL=https://your-tunnel.trycloudflare.com
```

### 4. Expose Local Server (for Testing Only)

```bash
# Terminal 1: Start your app
npm start

# Terminal 2: Expose to public
npm install -g cloudflared
cloudflared tunnel --url http://localhost:3004
```

Copy the public URL and set it as `WEBHOOK_TUNNEL_URL` in `.env`

---

## Implementation steps

### Step 1: Get Access Token

```javascript
async function getAccessToken() {}
```

---

### Step 2: Webhook Notification Handler

```javascript
// Handle validation and notifications
router.post('/notifications', async (req, res) => {
  try {
    const validationToken = req.query.validationToken;
    const notifications = req.body?.value;

    // STEP 1: Handle Microsoft Graph validation
    if (validationToken) {
      console.log('Validation request received:', validationToken);
      return res.status(200).type('text/plain').send(validationToken);
    }

    // STEP 2: Handle actual change notifications
    if (notifications && Array.isArray(notifications)) {
      console.log('Notifications received:', notifications.length);
      
      // Return 202 immediately (Microsoft requires response < 30 seconds)
      res.status(202).json({ 
        message: 'Notifications accepted',
        count: notifications.length 
      });

      // Process asynchronously
      for (const notification of notifications) {
        processNotification(notification).catch(err => {
          console.error('Error processing notification:', err);
        });
      }
      
      return;
    }

    return res.status(400).json({ error: 'Invalid request' });

  } catch (error) {
    console.error('Error in webhook handler:', error);
    res.status(500).json({ error: 'Processing failed' });
  }
});

```

---

### Step 3: Create Webhook Subscription

```javascript
    async function createSubscription(userEmail) {
    try {
        const accessToken = await getAccessToken();
        const tunnelUrl = process.env.WEBHOOK_TUNNEL_URL; // it can be KF api url
        const clientState = uuidv4();
        
        // Set expiration to 7 days (maximum for calendar events)
        const expirationDateTime = new Date();
        expirationDateTime.setDate(expirationDateTime.getDate() + 7);
        
        const subscriptionRequest = {
        changeType: "updated,created",
        notificationUrl: `${tunnelUrl}/api/webhook/notifications`,
        resource: `/users/${userEmail}/events`,
        expirationDateTime: expirationDateTime.toISOString(),
        clientState: clientState
        };

        const response = await axios.post(
        'https://graph.microsoft.com/v1.0/subscriptions',
        subscriptionRequest,
        {
            headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
            }
        }
        );

        const subscription = response.data;
        
        // Store subscription in database

        console.log('Subscription created:', subscription.id);
        return subscription;

    } catch (error) {
        console.error('Error creating subscription:', error.response?.data || error.message);
        throw error;
    }
    }
```

## Testing the Implementation in locally

### Test 1: Verify Setup

```bash
# Test health endpoint
curl http://localhost:3004/health

# Test public URL
curl https://your-tunnel.trycloudflare.com/health
```

Expected: `{"status":"OK",...}`

---

### Test 2: Test Webhook Validation

```bash
curl -X POST "https://your-tunnel.trycloudflare.com/api/webhook/notifications?validationToken=test123"
```

Expected: Returns `test123` as plain text

---

### Test 3: Create Subscription

```javascript
// Create subscription for a user
createSubscription('user@company.com')
  .then(sub => console.log('Subscription created:', sub.id))
  .catch(err => console.error('Failed:', err));
```

Expected Response:
```json
{
  "id": "31cdd544-75ce-433d-9e15-d586458f77ec",
  "resource": "/users/user@company.com/events",
  "changeType": "created,updated",
  "expirationDateTime": "2025-12-10T12:05:02.263Z"
}
```

---

### Test 4: Create/Update Event in Outlook

1. Open **Outlook Calendar**
2. Create new meeting with attendees
3. Click **Send**

**Watch your console** - you should see:
```
Notifications received: 1
Processing created notification for subscription 31cdd544...
Event ID: AAMkADdmZmEwYjA0...
```

---

### Test 5: Attendee Responds

1. Attendee opens meeting invitation
2. Clicks **Accept**

**Watch your console** - you should see:
```
Notifications received: 1
Processing updated notification for subscription 31cdd544...
Attendee changed response: none → accepted
Sending notifications for 1 response changes
```

---

## What to Do After Receiving Notification

### Process Flow After Notification Arrives

```javascript
// 1. NOTIFICATION RECEIVED
const notification = {
  subscriptionId: "31cdd544-75ce-433d-9e15-d586458f77ec",
  changeType: "updated",
  resourceData: { id: "AAMkADdm..." },
  clientState: "42b6562f-112c-4606-8a8f-40ec503a4d3f"
};

// 2. VALIDATE (Security Check)
validateClientState(notification.subscriptionId, notification.clientState);

// 3. GET EVENT ID
const eventId = notification.resourceData.id;

// 4. FETCH FULL EVENT FROM GRAPH API
const eventDetails = await fetchFromGraph(eventId);
// Returns: { subject, organizer, attendees: [{ emailAddress, status: { response, time }}] }

// 5. PROCESS THE DATA
// - Store in your database
// - Compare with previous state
// - Detect response changes
// - Send notifications to organizer
```

### Graph API Actions You Need

#### Action 1: Fetch Event Details

```http
GET https://graph.microsoft.com/v1.0/me/events/{eventId}
Authorization: Bearer {accessToken}
```

**Response Structure:**
```json
{
  "id": "AAMkADdm...",
  "subject": "Team Meeting",
  "organizer": {
    "emailAddress": {
      "name": "John Organizer",
      "address": "organizer@company.com"
    }
  },
  "attendees": [
    {
      "emailAddress": {
        "name": "Jane Attendee",
        "address": "jane@company.com"
      },
      "status": {
        "response": "accepted",
        "time": "2025-12-03T10:30:00Z"
      },
      "type": "required"
    }
  ],
  "start": { "dateTime": "2025-12-05T14:00:00", "timeZone": "UTC" },
  "end": { "dateTime": "2025-12-05T15:00:00", "timeZone": "UTC" }
}
```

#### Action 2: Get Subscription List (Optional - Check Status)

```http
GET https://graph.microsoft.com/v1.0/subscriptions
Authorization: Bearer {accessToken}
```

#### Action 3: Renew Subscription Before Expiry

```http
PATCH https://graph.microsoft.com/v1.0/subscriptions/{subscriptionId}
Authorization: Bearer {accessToken}
Content-Type: application/json

{
  "expirationDateTime": "2025-12-11T12:00:00.000Z"
}
```

---

## Key Implementation Notes

### Subscription Lifecycle

- **Maximum Duration**: 7 days for calendar events
- **Renewal Strategy**: Auto-renew 2 days before expiration
- **Renewal Frequency**: Run daily job to check expiring subscriptions


### Response Types

Microsoft Graph response values:
- `accepted` - Attendee accepted
- `declined` - Attendee declined
- `tentativelyAccepted` - Attendee responded maybe
- `notResponded` - No response yet
- `organizer` - Event organizer

---

## Quick Reference: API Endpoints

### Microsoft Graph Endpoints

```http
# Create Subscription
POST https://graph.microsoft.com/v1.0/subscriptions

# List Subscriptions
GET https://graph.microsoft.com/v1.0/subscriptions

# Renew Subscription
PATCH https://graph.microsoft.com/v1.0/subscriptions/{id}

# Delete Subscription
DELETE https://graph.microsoft.com/v1.0/subscriptions/{id}

# Get Event Details
GET https://graph.microsoft.com/v1.0/me/events/{eventId}
```

### Your Webhook Endpoints

```
# Webhook Notifications (Called by Microsoft Graph)
POST /api/webhook/notifications
```

---

## Troubleshooting

### Issue: Webhook Not Receiving Notifications

**Check:**
1. Is your webhook URL publicly accessible?
2. Does it return validation token correctly?
3. Is the subscription still active (not expired)?

**Test:**
```bash
curl https://your-url.com/api/webhook/notifications?validationToken=test
# Should return: test
```

### Issue: Subscription Creation Fails

**Check:**
1. Access token valid?
2. Required permissions granted?
3. Webhook URL uses HTTPS?

### Issue: Missing Attendee Responses

**Solution:**
Add 2-second delay before fetching event details (Graph API replication delay)

---

## Official Documentation Links

- [Microsoft Graph Webhooks](https://docs.microsoft.com/en-us/graph/webhooks)
- [Subscription Resource](https://docs.microsoft.com/en-us/graph/api/resources/subscription)
- [Create Subscription](https://docs.microsoft.com/en-us/graph/api/subscription-post-subscriptions)
- [Event Resource](https://docs.microsoft.com/en-us/graph/api/resources/event)
- [Calendar Permissions](https://docs.microsoft.com/en-us/graph/permissions-reference#calendar-permissions)

---

## Summary: Complete POC Flow

### From Start to Notification Received

```
START
  ↓
[1] Setup Azure AD App → Get Client ID, Secret, Tenant ID
  ↓
[2] Configure Webhook Endpoint → POST /api/webhook/notifications
  ↓
[3] Make Endpoint Public → Use cloudflared tunnel (testing) or HTTPS domain (prod)
  ↓
[4] Get Access Token → POST to login.microsoftonline.com
  ↓
[5] Create Subscription → POST to graph.microsoft.com/v1.0/subscriptions
  ↓
[6] Graph Validates → Sends validationToken, you return it as plain text
  ↓
[7] Subscription Active → Store subscriptionId, clientState, expirationDate
  ↓
[8] User Updates Event → In Outlook (Accept/Decline meeting)
  ↓
[9] Graph Sends Notification → POST to your webhook with event data
  ↓
[10] Validate & Process → Check clientState, extract eventId
  ↓
[11] Fetch Event Details → GET from graph.microsoft.com/v1.0/me/events/{eventId}
  ↓
[12] Process Response → Compare attendee status, detect changes
  ↓
END: Notification Received & Processed 
```

### Success Criteria

 **POC is Successful When:**

**Day 1:**
- Webhook subscription created successfully
- Validation token returned correctly
- Notification received when event is updated
- Event details fetched from Graph API
