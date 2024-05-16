// progress.ts

// @ts-nocheck
// @ts-ignore

import { NextApiRequest, NextApiResponse } from 'next';

let clients = [];

export default function handler(req: NextApiRequest, res: NextApiResponse) {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const data = `data: ${JSON.stringify({ progress: 0, consoleLog: 'Connection established' })}\n\n`;
  res.write(data);

  clients.push(res);

  console.log('New client connected'); // Add console log for new client connection

  req.on('close', () => {
    clients = clients.filter(client => client !== res);
    console.log('Client disconnected'); // Add console log for client disconnection
  });

  req.on('aborted', () => {
    clients = clients.filter(client => client !== res);
    res.end();
    console.log('Client aborted connection'); // Add console log for client connection abortion
  });
}

export function sendProgressUpdate(progress, consoleLog = '') {
  clients.forEach(res => {
    res.write(`data: ${JSON.stringify({ progress, consoleLog })}\n\n`);
  });
  console.log(`Progress update sent: progress=${progress}, consoleLog=${consoleLog}`); // Add console log for progress update
}
