import pg from 'pg';
import dotenv from 'dotenv';
import { coffeeSeed, sesameSeed } from '../services/exportService.js';

dotenv.config();

const { Client } = pg;

const client = new Client({
  host: process.env.DB_HOST || 'localhost',
  port: Number(process.env.DB_PORT || 5432),
  user: process.env.DB_USER || 'postgres',
  password: process.env.DB_PASSWORD || '',
  database: process.env.DB_NAME || 'sis_export_db',
});

const upsertUser = async () => {
  await client.query(`
    INSERT INTO app_users (username, password_hash, role)
    VALUES ($1, $2, $3)
    ON CONFLICT (username) DO UPDATE SET password_hash = EXCLUDED.password_hash, role = EXCLUDED.role;
  `, ['admin', 'password', 'admin']);
};

const upsertCoffee = async () => {
  for (const row of coffeeSeed) {
    await client.query(`
      INSERT INTO coffee_exports (
        week, shipment_no, buyer, destination, quantity, unit, containers, price_per_kg,
        gross_usd, paid_usd, payment_status, cleaning, product, vessel, contract,
        status, export_date, notes
      ) VALUES (
        $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18
      )
      ON CONFLICT (shipment_no) DO UPDATE SET
        buyer = EXCLUDED.buyer,
        destination = EXCLUDED.destination,
        quantity = EXCLUDED.quantity,
        unit = EXCLUDED.unit,
        containers = EXCLUDED.containers,
        price_per_kg = EXCLUDED.price_per_kg,
        gross_usd = EXCLUDED.gross_usd,
        paid_usd = EXCLUDED.paid_usd,
        payment_status = EXCLUDED.payment_status,
        cleaning = EXCLUDED.cleaning,
        product = EXCLUDED.product,
        vessel = EXCLUDED.vessel,
        contract = EXCLUDED.contract,
        status = EXCLUDED.status,
        export_date = EXCLUDED.export_date,
        notes = EXCLUDED.notes;
    `, [
      row.week,
      row.shipmentNo,
      row.buyer,
      row.destination,
      row.quantity,
      row.unit,
      row.containers,
      row.pricePerKg,
      row.grossUsd,
      row.paidUsd,
      row.paymentStatus,
      row.cleaning,
      row.product,
      row.vessel,
      row.contract,
      row.status,
      row.date,
      row.notes,
    ]);
  }
};

const upsertSesame = async () => {
  for (const row of sesameSeed) {
    await client.query(`
      INSERT INTO sesame_exports (
        week, shipment_no, buyer, destination, quantity, unit, containers, price_per_kg,
        gross_usd, paid_usd, payment_status, cleaning, product, vessel, contract,
        status, export_date, notes
      ) VALUES (
        $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18
      )
      ON CONFLICT (shipment_no) DO UPDATE SET
        buyer = EXCLUDED.buyer,
        destination = EXCLUDED.destination,
        quantity = EXCLUDED.quantity,
        unit = EXCLUDED.unit,
        containers = EXCLUDED.containers,
        price_per_kg = EXCLUDED.price_per_kg,
        gross_usd = EXCLUDED.gross_usd,
        paid_usd = EXCLUDED.paid_usd,
        payment_status = EXCLUDED.payment_status,
        cleaning = EXCLUDED.cleaning,
        product = EXCLUDED.product,
        vessel = EXCLUDED.vessel,
        contract = EXCLUDED.contract,
        status = EXCLUDED.status,
        export_date = EXCLUDED.export_date,
        notes = EXCLUDED.notes;
    `, [
      row.week,
      row.shipmentNo,
      row.buyer,
      row.destination,
      row.quantity,
      row.unit,
      row.containers,
      row.pricePerKg,
      row.grossUsd,
      row.paidUsd,
      row.paymentStatus,
      row.cleaning,
      row.product,
      row.vessel,
      row.contract,
      row.status,
      row.date,
      row.notes,
    ]);
  }
};

try {
  await client.connect();
  await upsertUser();
  await upsertCoffee();
  await upsertSesame();
  console.log('Seed data loaded successfully.');
} catch (error) {
  console.error('Seeding failed:', error.message);
  process.exit(1);
} finally {
  await client.end();
}
