// progress.ts

// @ts-nocheck
// @ts-ignore

import { NextApiRequest, NextApiResponse } from 'next';

let clients = [];

export default function handler(_req: NextApiRequest, _res: NextApiResponse) {
  _res.setHeader('Content-Type', 'text/event-stream');
  _res.setHeader('Cache-Control', 'no-cache');
  _res.setHeader('Connection', 'keep-alive');

  const data = `data: ${JSON.stringify({ progress: 0, consoleLog: 'Connection established' })}\n\n`;
  _res.write(data);

  clients.push(_res);

  console.log('New client connected'); // Add console log for new client connection

  _req.on('close', () => {
    clients = clients.filter(_client => _client !== _res);
    console.log('Client disconnected'); // Add console log for client disconnection
  });

  _req.on('aborted', () => {
    clients = clients.filter(_client => _client !== _res);
    _res.end();
    console.log('Client aborted connection'); // Add console log for client connection abortion
  });
}

export function sendProgressUpdate(_progress, _consoleLog = '') {
  clients.forEach(_res => {
    _res.write(`data: ${JSON.stringify({ progress: _progress, consoleLog: _consoleLog })}\n\n`);
  });
  console.log(`Progress update sent: progress=${_progress}, consoleLog=${_consoleLog}`); // Add console log for progress update
}
