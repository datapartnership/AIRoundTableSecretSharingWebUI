# AI Round Table Secret Sharing — Web UI

React 18 + Vite partner/admin interface for the Privacy-Preserving Secure Aggregation system. Partners sign in with their organisational Azure AD account and the browser automatically runs the full four-phase ML-KEM-768 protocol — no manual coordination, no client secrets.

## Prerequisites

- Node.js 20+
- The .NET API running locally (`AIRoundTableSecretSharingAPI`, default port 5149)
- An Azure Entra ID account in the configured tenant that belongs to either the Partner or Admin security group

## Getting Started

```bash
npm install   # first time only
npm run dev
```

The UI runs on `http://localhost:3000`. All `/api/*` requests are proxied to `http://localhost:5149` by Vite.

## Pages

### Home
Protocol overview and sign-in prompt.

### Protocol Flow (`/protocol`)
Automated four-step wizard for registered partners:

| Step | What happens automatically |
|------|---------------------------|
| **1 — Key Generation** | Generates an ML-KEM-768 key pair in the browser; posts the encapsulation key (1184 bytes) to the API; waits for all epoch partners to register |
| **2 — Encapsulation** | For each partner with a smaller OID, encapsulates and posts a ciphertext (1088 bytes) to the API |
| **3 — Decapsulation** | Polls for incoming ciphertexts; decapsulates each using the local private key to derive shared secrets |
| **4 — Submit Data** | User enters actual MAU values per country / month; the UI applies HMAC-SHA256 noise and submits masked values |

Key material (key pair, shared secrets, epoch ID) is stored in `localStorage` keyed by the user's OID. On each page load the stored epoch ID is compared to the current epoch; stale state is cleared automatically when a new epoch is detected.

### Admin Panel (`/admin`)
Available to Admin group members only:

- View all registered partners
- Select a subset (≥ 2) to participate in a new epoch
- Clear all data

### Results (`/results`)
Query the aggregate for any country / month combination in the current epoch. Available once all epoch partners have submitted.

### Offline submission (computer without internet)
Some partners keep their data on a computer with no internet access. When the key exchange is done and the Quorum Check is complete, the epoch page shows **No internet on the computer with your data?**:

1. **Export key file** (`airt-epoch-<id>-key.json`): for each cell the partner owes, it holds the pad, which is the net noise their pairwise secrets add (`calculateMaskedValue(0, …)`). It also holds the epoch's months, mask width and quorum rules. Everything is encrypted with a passphrase (PBKDF2-SHA256, 600k iterations → AES-256-GCM). The shared secrets and the ML-KEM private key never leave the browser.
2. **Download offline tool** (`offline-masker.html`): one self-contained HTML file. Its CSP blocks every network request. On the offline computer the partner opens it from disk, opens the key file, uploads the CSV (same checks as online) and saves `airt-epoch-<id>-masked.json`. Each masked value is `(value + pad) mod 2^bits`, identical to online masking.
3. **Submit masked file**: in the same browser that exported the key file. The page checks the epoch, user, mask width and key fingerprint (SHA-256 of the registered public key), and that every expected cell appears exactly once with a value in range. It then submits through `submit-batch` with signature `web-ui-offline`.

The key file can unmask the partner's values, so partners are told to keep the passphrase separate and to delete the key file after submitting. The offline tool is built by `npm run build:offline` (`vite.offline.config.js` + `vite-plugin-singlefile`) into `public/offline-masker.html`, which is git-ignored. `npm run dev` and `npm run build` run it automatically. The library is `src/utils/offlineBundle.js` and the tool is `src/offline/`.

## Authentication

MSAL v5 (`@azure/msal-browser` + `@azure/msal-react`) with interactive redirect flow. The access token is injected into every API request automatically. The `groups` claim in the token determines Admin vs Partner access.

The web UI reads its Entra ID and group IDs from Vite environment variables:

| Variable | Purpose |
| --- | --- |
| `VITE_MSAL_CLIENT_ID` | SPA app registration client ID |
| `VITE_MSAL_AUTHORITY` | Entra authority URL |
| `VITE_MSAL_API_SCOPE` | API scope requested for access tokens |
| `VITE_ADMIN_GROUP_ID` | Admin security group object ID used for Admin navigation gating |
| `VITE_PARTNER_GROUP_ID` | Partner security group object ID, kept in sync with backend configuration |

## Project Structure

```
web-ui/
├── src/
│   ├── pages/
│   │   ├── Home.jsx          # Landing page (signed out) / epoch dashboard (signed in)
│   │   ├── AdminPanel.jsx    # Epoch management (admin only)
│   │   └── Results.jsx       # Aggregate results
│   ├── components/
│   │   ├── EpochDashboard.jsx # Open epochs + per-epoch key status; hosts the protocol
│   │   └── EpochProtocol.jsx  # Automated 4-step protocol for one epoch
│   ├── utils/
│   │   ├── localCrypto.js    # localStorage key/secret persistence + key status
│   │   ├── api.js            # All API calls; Entra token injected via acquireApiToken
│   │   ├── crypto.js         # ML-KEM-768 via `mlkem` npm package
│   │   └── noise.js          # HMAC-SHA256 noise; matches C# & Python exactly
│   ├── App.jsx               # Routing; admin flag from access token groups claim
│   ├── main.jsx              # MSAL provider setup
│   └── index.css             # Global styles
├── index.html
├── package.json
└── vite.config.js            # Proxy /api → localhost:5149
```

## Noise Compatibility

The JavaScript noise calculation in `noise.js` is mathematically identical to `SecureNoiseGenerator.cs` (C#):

```
noise  = HMAC-SHA256(sharedSecret, "{country}|{month}|{indicator}|{segment}")
       → first bits/8 bytes as LE signed integer (bits = epoch.maskBits: 128, or 64 for older epochs)
masked = (actual + Σ noise × sign) mod 2^bits     // sign +1 if myId < partnerId, else −1
```

The CSV check accepts values up to 2^(bits−1) − 1. Quorum Check answers always use 64 bits.

Both runtimes produce the same value for any given shared secret and cell, which is required for the noise to cancel in the aggregate.
