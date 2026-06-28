// db.js - SQLite database initialization using better-sqlite3
const Database = require('better-sqlite3');
// Database file will be created in the project root
const db = new Database('smartroute.db');

// Create tables if they don't exist
const createDepotTable = `
CREATE TABLE IF NOT EXISTS depots (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  address TEXT,
  lat REAL NOT NULL,
  lng REAL NOT NULL,
  isDepot INTEGER NOT NULL DEFAULT 1
);`;

const createCustomerTable = `
CREATE TABLE IF NOT EXISTS customers (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  address TEXT,
  lat REAL NOT NULL,
  lng REAL NOT NULL,
  isDepot INTEGER NOT NULL DEFAULT 0,
  route TEXT
);`;

const createHistoryTable = `
CREATE TABLE IF NOT EXISTS routing_history (
  timestamp INTEGER PRIMARY KEY,
  result TEXT NOT NULL
);`;

db.exec(createDepotTable);
db.exec(createCustomerTable);
db.exec(createHistoryTable);

module.exports = db;
