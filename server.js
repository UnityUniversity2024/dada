import dotenv from 'dotenv';
import express from 'express';
import cors from 'cors';

import router from './routes/index.js';

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 3002);

app.use(cors({ origin: true }));
app.use(express.json());
app.use('/api', router);

app.listen(PORT, '0.0.0.0', () => {
  console.log(`SIS export backend running on http://0.0.0.0:${PORT}/api`);
});
