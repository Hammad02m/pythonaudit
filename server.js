import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import { spawn } from "child_process";

const app = express();
app.use(cors());
app.use(bodyParser.json()); // Parse JSON body

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
    pythonProcess.on("close", () => {
        console.log("Python process completed");
        res.json({ auditResults: JSON.parse(resultData) });
    });

    pythonProcess.stderr.on("data", (data) => {
        console.error(`Python error: ${data}`);
    });
});

// Start server on port 5000
app.listen(5000, "0.0.0.0", () => console.log("Server running on port 5000"));

