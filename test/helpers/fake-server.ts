import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';

export interface Request { method: string; path: string; headers: IncomingMessage['headers']; body: string }
export type Reply = (request: Request, response: ServerResponse) => void | Promise<void>;
export interface Endpoint {
  url: string;
  requests: Request[];
  setReply(reply: Reply): void;
  reset(): void;
  close(): Promise<void>;
}
export interface Servers { graph: Endpoint; login: Endpoint; close(): Promise<void> }

async function startEndpoint(): Promise<Endpoint> {
  const requests: Request[] = [];
  let reply: Reply = (_request, response) => { response.writeHead(404).end(); };
  const server: Server = createServer(async (request, response) => {
    let body = '';
    for await (const chunk of request) body += chunk.toString();
    const recorded = { method: request.method ?? 'GET', path: request.url ?? '/', headers: request.headers, body };
    requests.push(recorded);
    await reply(recorded, response);
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('Fake server did not bind');
  return {
    url: `http://127.0.0.1:${address.port}`, requests,
    setReply: (next) => { reply = next; },
    reset: () => { requests.length = 0; },
    close: () => new Promise((resolve, reject) => {
      server.close((error) => error ? reject(error) : resolve());
      server.closeAllConnections();
    }),
  };
}

export async function startServers(): Promise<Servers> {
  const graph = await startEndpoint();
  const login = await startEndpoint();
  return { graph, login, close: async () => { await Promise.all([graph.close(), login.close()]); } };
}
