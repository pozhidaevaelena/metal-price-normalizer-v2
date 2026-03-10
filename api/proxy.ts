import type { VercelRequest, VercelResponse } from '@vercel/node';
import fetch from 'node-fetch';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  const targetUrl = req.query.url as string;
  
  if (!targetUrl) {
    return res.status(400).json({ error: 'URL is required' });
  }

  try {
    const response = await fetch(targetUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://mc.ru/prices',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
      },
      redirect: 'follow'
    });

    if (!response.ok) {
      return res.status(response.status).json({ error: `Target returned ${response.status} (${response.statusText})` });
    }

    const contentType = response.headers.get('content-type');
    if (contentType) {
      res.setHeader('Content-Type', contentType);
    }

    // For Vercel Serverless Functions, we might need to return the buffer instead of streaming
    // depending on the size and runtime, but streaming should work if supported.
    const buffer = await response.buffer();
    res.send(buffer);
  } catch (error: any) {
    console.error('Proxy error:', error);
    res.status(500).json({ error: `Failed to fetch target URL: ${error.message}` });
  }
}
