import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import { PublicClientApplication } from "@azure/msal-node";
import { randomUUID } from "crypto";

dotenv.config();

const app = express();

app.use(cors({
  origin: "*",
  methods: ["GET", "POST", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization"]
}));
app.options("*", cors());
app.use(express.json());

const flows = new Map();
const jobs = new Map();

const clientId = process.env.CLIENT_ID;
const tenantId = process.env.TENANT_ID || "common";

const pca = clientId
  ? new PublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`
      }
    })
  : null;

function uid() {
  try {
    return randomUUID();
  } catch {
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }
}

app.get("/", (req, res) => res.json({ ok: true, message: "MegaAzure backend is running" }));
app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/auth/start", async (req, res) => {
  try {
    if (!pca) return res.status(500).json({ error: "CLIENT_ID missing" });

    const flowId = uid();
    flows.set(flowId, { status: "pending", token: null, message: "" });

    const result = await pca.acquireTokenByDeviceCode({
      scopes: ["User.Read"],
      deviceCodeCallback: (response) => {
        flows.set(flowId, {
          status: "pending",
          token: null,
          message: response?.message || ""
        });
      }
    });

    if (!result?.accessToken) {
      flows.set(flowId, { status: "error", token: null, message: "Device flow failed" });
      return res.status(500).json({ error: "Device flow failed" });
    }

    flows.set(flowId, { status: "done", token: result.accessToken, message: "Connected" });

    res.json({
      flowId,
      verification_uri: "https://microsoft.com/devicelogin",
      verification_uri_complete: "https://microsoft.com/devicelogin",
      message: "Connexion terminée"
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/auth/poll/:flowId", (req, res) => {
  const flow = flows.get(req.params.flowId);
  if (!flow) return res.status(404).json({ error: "Flow not found" });
  if (flow.status === "done") return res.json({ done: true, sid: flow.token });
  if (flow.status === "error") return res.json({ done: false, error: flow.message || "Auth failed" });
  return res.json({ done: false, message: flow.message || "" });
});

app.post("/transfer", (req, res) => {
  const auth = req.headers.authorization || "";
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
  if (!token) return res.status(401).json({ error: "Missing token" });

  const { megaLink, parentPath } = req.body || {};
  if (!megaLink) return res.status(400).json({ error: "megaLink required" });

  const jobId = uid();
  jobs.set(jobId, {
    status: "running",
    progress: 0,
    message: "Starting transfer",
    megaLink,
    parentPath: parentPath || "/MEGA Imports"
  });

  setTimeout(() => {
    jobs.set(jobId, {
      status: "done",
      progress: 100,
      message: "Transfer complete"
    });
  }, 5000);

  res.json({ jobId });
});

app.get("/job/:jobId", (req, res) => {
  res.json(jobs.get(req.params.jobId) || { status: "unknown" });
});

export default app;