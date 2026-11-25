# CID — C-nta Individual Dashboard

A React prototype for viewing residential service metrics and editing individual attributes that are not currently tracked in the existing system.

## Run locally

```bash
npm install
npm run dev
```

The local app runs at `http://localhost:3000`. Use that exact address rather than `127.0.0.1` or another port because Microsoft sign-in must return to the same registered origin.

## Configure Microsoft sign-in

1. Copy `.env.example` to `.env`.
2. Add the Application (client) ID and Directory (tenant) ID from the Azure app registration.
3. Add `http://localhost:3000/redirect` as a Single-page application redirect URI in Azure.
4. Restart the development server.

```env
VITE_ENTRA_CLIENT_ID=your-application-client-id
VITE_ENTRA_TENANT_ID=your-directory-tenant-id
VITE_REDIRECT_URI=http://localhost:3000/redirect
CNTAT_API_KEY=your-existing-cntat-api-key
```

The API key is read only by the local Vite proxy and is not included in browser JavaScript. Azure deployment will need an equivalent server-side proxy or function.

## Current prototype scope

- Live dashboard populated from existing client and house records
- Individual search and record selection
- Existing service and schedule attributes save to the C-NTA API
- New CID-only profile attributes persist in the API's shared `ClientProfiles` collection
- Responsive layout for desktop, tablet, and mobile
- MSAL authentication using the same Entra pattern as the Finance Dashboard

The local Vite proxy keeps the C-NTA API key out of browser JavaScript. Azure deployment should provide the same server-side proxy behavior.
