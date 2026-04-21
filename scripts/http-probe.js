import http from "node:http";
import https from "node:https";

const DEFAULT_TIMEOUT_MS = 1500;

const getClient = (url) => (url.protocol === "https:" ? https : http);

export const requestStatus = (targetUrl, timeoutMs = DEFAULT_TIMEOUT_MS) =>
  new Promise((resolveRequest) => {
    let settled = false;
    const url = typeof targetUrl === "string" ? new URL(targetUrl) : targetUrl;
    const finalize = (result) => {
      if (settled) {
        return;
      }
      settled = true;
      resolveRequest(result);
    };

    const request = getClient(url).request(url, {
      method: "GET",
      headers: {
        "Cache-Control": "no-cache",
      },
    }, (response) => {
      response.resume();
      finalize({
        ok: (response.statusCode || 0) >= 200 && (response.statusCode || 0) < 300,
        status: response.statusCode || 0,
      });
    });

    request.setTimeout(timeoutMs, () => {
      request.destroy();
      finalize({ ok: false, status: 0 });
    });

    request.on("error", () => {
      finalize({ ok: false, status: 0 });
    });

    request.end();
  });
