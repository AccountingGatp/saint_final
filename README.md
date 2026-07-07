# Excel Processor

Upload an Excel file in the browser → Express processes it in memory → the
generated workbook downloads back. **Nothing is stored on the server.**

```
saint_final/
├─ backend/    Express + multer (in-memory) + xlsx-js-style
└─ frontend/   Next.js (App Router) + Tailwind + shadcn button
```

## Run

Two terminals.

**Backend** (port 4000):

```bash
cd backend
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
and the output workbook downloads automatically.

## How it works

- `POST /process` accepts a multipart form field named `file`.
- The upload is held in memory only (`multer.memoryStorage`) — never written to disk.
- The buffer is transformed into a 3-sheet workbook (`Input`, `Sheet1`, `Sheet2`)
  with per-day USD→AUD conversion (live rates from the keyless Frankfurter API),
  then streamed straight back as the download.

Optional form fields `date` (default `06/26`) and `currency` (default `AUD`)
override the defaults if you add them to the request.

The frontend reads the backend URL from `NEXT_PUBLIC_API_URL`
(see `frontend/.env.local`, defaults to `http://localhost:4000`).
