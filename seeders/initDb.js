import fs from 'fs/promises';
import pg from 'pg';
import dotenv from 'dotenv';

dotenv.config();

const { Client } = pg;

const config = {
  host: process.env.DB_HOST || 'localhost',
  port: Number(process.env.DB_PORT || 5432),
  user: process.env.DB_USER || 'postgres',
  password: process.env.DB_PASSWORD || '',
  database: 'postgres',
};

const dbName = process.env.DB_NAME || 'sis_export_db';

const createDatabase = async () => {
  const client = new Client(config);
  await client.connect();

  const exists = await client.query('SELECT 1 FROM pg_database WHERE datname = $1', [dbName]);
  if (!exists.rowCount) {
    await client.query(`CREATE DATABASE "${dbName}"`);
    console.log(`Created database: ${dbName}`);
  } else {
    console.log(`Database already exists: ${dbName}`);
  }

  await client.end();
};

const applySchema = async () => {
  const schema = await fs.readFile(new URL('../sql/schema.sql', import.meta.url), 'utf8');
  const client = new Client({
    ...config,
    database: dbName,
  });

  await client.connect();
  await client.query(schema);
  console.log('Applied schema successfully.');
  await client.end();
};

try {
  await createDatabase();
  await applySchema();
} catch (error) {
  console.error('Database init failed:', error.message);
  process.exit(1);
}
