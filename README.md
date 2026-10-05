# SIS Export Backend + Standalone Export UI

This project includes:

- A Node.js + Express backend for export operations
- PostgreSQL-ready schema and seed scripts
- A standalone React dashboard that works with local mock data when the backend is unavailable

## Backend

```bash
cd sis-export-backend
npm install
npm run db:init
npm run db:seed
npm start
```

Base URL: http://localhost:3002/api

### Authentication

Default login:

- username: admin
- password: password

### Endpoints

- GET /api/health
- POST /api/auth/login
- GET /api/bootstrap
- GET /api/coffee
- POST /api/coffee
- PUT /api/coffee/:id
- DELETE /api/coffee/:id
- GET /api/sesame
- POST /api/sesame
- PUT /api/sesame/:id
- DELETE /api/sesame/:id

## Frontend

```bash
cd sis-export-backend/frontend
npm install
npm run dev
```

The app tries to call the backend on load. If it fails, it falls back to local mock data automatically and keeps the UI running.
