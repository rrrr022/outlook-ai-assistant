import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import https from 'https';
import http from 'http';
import fs from 'fs';
import path from 'path';
import net from 'net';

// Load environment variables from the project root
const envPath = path.resolve(__dirname, '../../.env');
console.log(`📁 Loading .env from: ${envPath}`);
const result = dotenv.config({ path: envPath });
if (result.error) {
  console.error('❌ Error loading .env file:', result.error);
} else {
  console.log('✅ .env file loaded successfully');
  console.log(`   GITHUB_TOKEN: ${process.env.GITHUB_TOKEN ? 'Found' : 'NOT FOUND'}`);
}

import aiRoutes from './routes/aiRoutes';
import automationRoutes from './routes/automationRoutes';

const app = express();
const DEFAULT_PORT = parseInt(process.env.PORT || '3001', 10);

/**
 * Find an available port starting from the given port
 */
async function findAvailablePort(startPort: number): Promise<number> {
  return new Promise((resolve) => {
    const server = net.createServer();
    server.listen(startPort, () => {
      const port = (server.address() as net.AddressInfo).port;
      server.close(() => resolve(port));
    });
    server.on('error', () => {
      // Port in use, try next one
      resolve(findAvailablePort(startPort + 1));
    });
  });
}

// Middleware
app.use(cors({
  origin: true, // Allow all origins in development
  credentials: true,
}));
app.use(express.json());

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// API Routes
app.use('/api/ai', aiRoutes);
app.use('/api/automation', automationRoutes);

// Error handling middleware
app.use((err: Error, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error('Error:', err.message);
  res.status(500).json({ 
    success: false, 
    error: 'Internal server error',
    message: process.env.NODE_ENV === 'development' ? err.message : undefined,
  });
});

// Start server
const startServer = async () => {
  try {
    const port = await findAvailablePort(DEFAULT_PORT);
    
    // For development, create self-signed certificates if they don't exist
    const certPath = path.join(__dirname, '../../certs');
    
    if (fs.existsSync(path.join(certPath, 'server.key')) && 
        fs.existsSync(path.join(certPath, 'server.crt'))) {
      // Use HTTPS with certificates
      const httpsOptions = {
        key: fs.readFileSync(path.join(certPath, 'server.key')),
        cert: fs.readFileSync(path.join(certPath, 'server.crt')),
      };
      
      https.createServer(httpsOptions, app).listen(port, () => {
        console.log(`🚀 Server running on https://localhost:${port}`);
        console.log(`📊 Health check: https://localhost:${port}/health`);
        // Write port to file for frontend to read
        fs.writeFileSync(path.join(__dirname, '../../.server-port'), port.toString());
      });
    } else {
      // Fall back to HTTP for development
      http.createServer(app).listen(port, () => {
        console.log(`🚀 Server running on http://localhost:${port}`);
        console.log(`📊 Health check: http://localhost:${port}/health`);
        if (port !== DEFAULT_PORT) {
          console.log(`ℹ️  Note: Port ${DEFAULT_PORT} was in use, using ${port} instead`);
        }
        console.log(`⚠️  Note: For Outlook Add-in, HTTPS is required. Generate certificates in /certs folder.`);
        // Write port to file for frontend to read
        fs.writeFileSync(path.join(__dirname, '../../.server-port'), port.toString());
      });
    }
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
};

startServer();
