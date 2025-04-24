import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import { spawn } from "child_process";

const app = express();

// CORS configuration
app.use(cors({
  origin: "https://helpful-treacle-7f8ee1.netlify.app", // your frontend URL
  methods: ["GET", "POST", "OPTIONS"], // Allow these HTTP methods
  allowedHeaders: ["Content-Type"], // Allow these headers
}));

app.use(bodyParser.json()); // Parse JSON body

// Handle OPTIONS preflight request
app.options("*", cors());

// API Endpoint to Handle React's Request
app.post("/run-audit", (req, res) => {
  const formData = req.body;
  console.log("Received data from React:", formData);

  // Spawn a Python process
  const pythonProcess = spawn("python", ["audittimecalculator.py"]);

  // Send formData to Python process as JSON
  pythonProcess.stdin.write(JSON.stringify(formData));
  pythonProcess.stdin.end();

  let resultData = "";
  // Capture output from Python script
  pythonProcess.stdout.on("data", (data) => {
    resultData += data.toString();
  });

  // Send results back to React when Python finishes
  pythonProcess.on("close", (code) => {
    if (code !== 0) {
      console.error(`Python process exited with code ${code}`);
      return res.status(500).send("Error in Python script");
    }
    console.log("Python process completed");
    res.json({ auditResults: JSON.parse(resultData) });
  });

  pythonProcess.stderr.on("data", (data) => {
    console.error(`Python error: ${data.toString()}`);
  });
});

// Start server on port from environment or 5000
app.listen(process.env.PORT || 5000, "0.0.0.0", () => {
  console.log("Server running...");
});
