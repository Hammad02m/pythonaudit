import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import { spawn } from "child_process";

const app = express();

// --- ✅ CORS Configuration ---
const whitelist = [
  "http://localhost:3000", // for local testing
  "https://helpful-treacle-7f8ee1.netlify.app" // your deployed frontend
];

const corsOptions = {
  origin: (origin, callback) => {
    // Allow requests with no origin (like curl or Postman)
    if (!origin || whitelist.includes(origin)) {
      callback(null, true);
    } else {
      callback(new Error("Not allowed by CORS"));
    }
  },
  methods: ["GET", "POST", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization"],
  credentials: true,
};

app.use(cors(corsOptions)); // Enable CORS globally
app.options("*", cors(corsOptions)); // Enable preflight for all routes

// --- ✅ Middleware ---
app.use(bodyParser.json()); // Parse JSON request bodies

// --- ✅ POST Endpoint ---
app.post("/run-audit", (req, res) => {
  const formData = req.body;
  console.log("Received data from React:", formData);

  // Spawn Python script
  const pythonProcess = spawn("python", ["audittimecalculator.py"]);

  // Send data to Python
  pythonProcess.stdin.write(JSON.stringify(formData));
  pythonProcess.stdin.end();

  let resultData = "";

  // Read data from Python stdout
  pythonProcess.stdout.on("data", (data) => {
    resultData += data.toString();
  });

  // When Python script ends
  pythonProcess.on("close", (code) => {
    if (code !== 0) {
      console.error(`Python process exited with code ${code}`);
      return res.status(500).send("Error in Python script");
    }
    console.log("Python process completed");
    try {
      const resultJSON = JSON.parse(resultData);
      res.json({ auditResults: resultJSON });
    } catch (err) {
      console.error("Failed to parse Python output:", err);
      res.status(500).send("Invalid output from Python script");
    }
  });

  pythonProcess.stderr.on("data", (data) => {
    console.error(`Python error: ${data.toString()}`);
  });
});

// --- ✅ Start Server ---
const PORT = 5000;
app.listen(PORT, "0.0.0.0", () => {
  console.log(`Server running on port ${PORT}`);
});
