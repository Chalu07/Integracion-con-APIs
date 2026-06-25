const rateLimitMap = new Map<string, { count: number; lastReset: number }>();

export function rateLimit(
  ip: string,
  limit: number = 5,
  windowMs: number = 60000
): boolean {
  const now = Date.now();
  const record = rateLimitMap.get(ip);

  if (!record || now - record.lastReset > windowMs) {
    rateLimitMap.set(ip, { count: 1, lastReset: now });
    return true;
  }

  if (record.count >= limit) {
    return false;
  }

  record.count++;
  return true;
}
