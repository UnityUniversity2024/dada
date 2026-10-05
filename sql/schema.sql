CREATE TABLE IF NOT EXISTS app_users (
  id SERIAL PRIMARY KEY,
  username VARCHAR(100) UNIQUE NOT NULL,
  password_hash VARCHAR(255) NOT NULL,
  role VARCHAR(50) DEFAULT 'admin',
  created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS coffee_exports (
  id SERIAL PRIMARY KEY,
  week VARCHAR(20),
  shipment_no VARCHAR(100) UNIQUE NOT NULL,
  buyer VARCHAR(150),
  destination VARCHAR(150),
  quantity NUMERIC(12,2),
  unit VARCHAR(20),
  containers INTEGER,
  price_per_kg NUMERIC(12,4),
  gross_usd NUMERIC(14,2),
  paid_usd NUMERIC(14,2),
  payment_status VARCHAR(50),
  cleaning VARCHAR(50),
  product VARCHAR(100),
  vessel VARCHAR(150),
  contract VARCHAR(100),
  status VARCHAR(50),
  export_date DATE,
  notes TEXT,
  created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS sesame_exports (
  id SERIAL PRIMARY KEY,
  week VARCHAR(20),
  shipment_no VARCHAR(100) UNIQUE NOT NULL,
  buyer VARCHAR(150),
  destination VARCHAR(150),
  quantity NUMERIC(12,2),
  unit VARCHAR(20),
  containers INTEGER,
  price_per_kg NUMERIC(12,4),
  gross_usd NUMERIC(14,2),
  paid_usd NUMERIC(14,2),
  payment_status VARCHAR(50),
  cleaning VARCHAR(50),
  product VARCHAR(100),
  vessel VARCHAR(150),
  contract VARCHAR(100),
  status VARCHAR(50),
  export_date DATE,
  notes TEXT,
  created_at TIMESTAMP DEFAULT NOW()
);
