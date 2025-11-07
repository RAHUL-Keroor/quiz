// app.js
const express = require('express');
const connectDB = require('./db');  // your db.js file
const app = express();

// Middleware (optional for JSON)
app.use(express.json());

// Connect to MongoDB
connectDB().then((db) => {
  console.log('✅ MongoDB connected successfully');

  // Simple route
  app.get('/', (req, res) => {
    res.send('Hello from MongoDB + Express server!');
  });

  // Example route to insert a user
  app.post('/add-user', async (req, res) => {
    const users = db.collection('users');
    const result = await users.insertOne({ name: 'Alice', age: 25 });
    res.json({ message: 'User added!', result });
  });

  // Example route to view all users
  app.get('/users', async (req, res) => {
    const users = db.collection('users');
    const allUsers = await users.find().toArray();
    res.json(allUsers);
  });

  // Start the server (this is the line that caused your error)
  app.listen(3000, '0.0.0.0', () => {
    console.log('🚀 Server running at http://localhost:3000');
  });
});
