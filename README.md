# Excel Processor
 
Upload an Excel file in the browser → Express processes it → output is stored in
Backblaze B2 → the frontend receives a file ID and download URL.

```
saint_final/
├─ backend/    Express + multer + xlsx-js-style + Backblaze B2
└─ frontend/   Next.js (App Router) + Tailwind + shadcn button
```

## Run

Two terminals.

**Backend** (port 4000):

```bash
cd backend
cp .env.example .env   # fill in your B2 credentials
npm install
npm start
```

**Frontend** (port 3000):

```bash
cd frontend
npm install
npm run dev
```

Open http://localhost:3000, choose an `.xlsx` file, click **Upload & Process**,
then download the output or generate an import file.

## Backblaze B2 setup

Create a B2 bucket and an application key with read/write access, then set these
in `backend/.env`:

| Variable | Description |
|----------|-------------|
| `B2_ENDPOINT` | S3-compatible endpoint (e.g. `https://s3.us-east-005.backblazeb2.com`) |
| `B2_REGION` | Bucket region (e.g. `us-east-005`) |
| `B2_BUCKET` | Bucket name |
| `B2_KEY_ID` | Application key ID |
| `B2_APP_KEY` | Application key secret |
| `B2_URL_EXPIRY` | Presigned download URL lifetime in seconds (default `600`) |

## How it works

- `POST /process` accepts a multipart form field named `file`.
- The upload is processed in memory, then stored in B2.
- Returns JSON: `{ fileId, fileName, downloadUrl }`.
- `POST /import` accepts JSON `{ fileId }` referencing a stored output file.
- Fetches the output from B2, builds the import workbook, stores it in B2, and
  returns `{ fileId, fileName, downloadUrl }`.
- `GET /files/:fileId/download?type=output|import` returns a fresh presigned URL.

Optional form fields `date` (default `06/26`) and `currency` (default `AUD`)
override the defaults on `/process`.

The frontend reads the backend URL from `NEXT_PUBLIC_API_URL`
(defaults to `http://localhost:4000`).
