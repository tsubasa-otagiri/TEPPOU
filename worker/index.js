/**
 * TEPPOU – Cloudflare Worker
 * - IP 制限
 * - KV 読み書き（GET / POST）
 * - Cache API によるサーバーサイドキャッシュ（TTL 5分）
 * - 差分チェック（変更なし時は KV.put をスキップ）
 */

const ALLOWED_IPS = new Set([
  "210.172.143.39", // 主回線
  "210.172.143.37", // Mac用VPN回線
  "210.172.130.69", // 追加回線
]);

const RESOURCES = new Set(["records", "settings"]);

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function cors(body, init = {}) {
  return new Response(body, { ...init, headers: { ...CORS, ...(init.headers ?? {}) } });
}

export default {
  async fetch(request, env) {
    // ── CORS preflight ──────────────────────────────────────────────────────────
    if (request.method === "OPTIONS") return cors(null, { status: 204 });

    // ── IP 制限 ─────────────────────────────────────────────────────────────────
    const ip = request.headers.get("CF-Connecting-IP") ?? "";
    if (!ALLOWED_IPS.has(ip)) {
      return cors(JSON.stringify({ error: "Forbidden", ip }), {
        status: 403,
        headers: { "Content-Type": "application/json" },
      });
    }

    // ── ルーティング ────────────────────────────────────────────────────────────
    const url      = new URL(request.url);
    const resource = url.pathname.replace(/^\/api\//, "").replace(/\/$/, "");

    if (!RESOURCES.has(resource)) {
      return cors("Not Found", { status: 404 });
    }

    // ── GET（キャッシュ優先） ───────────────────────────────────────────────────
    if (request.method === "GET") {
      const cacheKey = new Request(`https://kv-cache.internal/${resource}`);
      const cache    = caches.default;

      const cached = await cache.match(cacheKey);
      if (cached) {
        const body = await cached.text();
        return cors(body, { headers: { "Content-Type": "application/json", "X-Cache": "HIT" } });
      }

      const value = (await env.TEPPOU_KV.get(resource)) ?? "[]";
      const resp  = cors(value, {
        headers: { "Content-Type": "application/json", "X-Cache": "MISS", "Cache-Control": "s-maxage=300" },
      });
      await cache.put(cacheKey, resp.clone());
      return resp;
    }

    // ── POST（差分チェック → KV.put → キャッシュ無効化） ───────────────────────
    if (request.method === "POST") {
      const newValue = await request.text();
      const current  = await env.TEPPOU_KV.get(resource);

      if (current !== newValue) {
        await env.TEPPOU_KV.put(resource, newValue);
        const cacheKey = new Request(`https://kv-cache.internal/${resource}`);
        await caches.default.delete(cacheKey);
      }

      return cors("OK");
    }

    return cors("Method Not Allowed", { status: 405 });
  },
};
