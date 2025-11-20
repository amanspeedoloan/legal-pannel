// DIRECT MYSQL CONNECTION (NO OTHER FILE)
const mysql = require("mysql2");

// === DB CONFIG ===
const db = mysql.createPool({
  host: "agrimuat.com",
  user: "admin",
  password: "!@#QET_14#14vin%13",
  database: "agrimdb",
  port: 3306,
  waitForConnections: true,
  connectionLimit: 10,
  ssl: false   // aapne bola DB_SSL=false
});

// === TEST CONNECTION ===
db.getConnection((err, connection) => {
  if (err) {
    console.log("❌ DB CONNECTION FAILED:", err.message);
  } else {
    console.log("✅ DB CONNECTED SUCCESSFULLY");
    connection.release();
  }
});
