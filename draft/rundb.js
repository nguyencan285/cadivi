const sqlite3 = require('sqlite3').verbose();

const db = new sqlite3.Database('./database.sqlite');

  db.run(`ALTER TABLE tickets ADD COLUMN photo_after TEXT;
  
  )`);

