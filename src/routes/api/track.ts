import { createFileRoute } from "@tanstack/react-router";
import { getStore } from "@netlify/blobs";

// A real, valid 1x1 transparent GIF (43 bytes), base64-encoded.
const TRANSPARENT_GIF = Buffer.from(
  "R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==",
  "base64",
);

function pixelResponse(): Response {
  return new Response(TRANSPARENT_GIF, {
    status: 200,
    headers: {
      "Content-Type": "image/gif",
      "Content-Length": String(TRANSPARENT_GIF.length),
      // Never let the recipient's mail client (or any proxy) cache this,
      // or a second open from the same person won't re-fire the request.
      "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
      Pragma: "no-cache",
    },
  });
}

export const Route = createFileRoute("/api/track")({
  server: {
    handlers: {
      // GET /api/track?id=<rowIndex>:<email> — called by the invisible
      // <img> tag inserted into outgoing HTML-mode emails when "Track
      // opens" is enabled. We always respond with the pixel even if
      // storage fails, so a broken-image icon never shows up in the
      // recipient's inbox.
      GET: async ({ request }) => {
        try {
          const url = new URL(request.url);
          const id = url.searchParams.get("id");
          if (id) {
            const store = getStore({ name: "outreach-opens" });
            // Keep the EARLIEST open timestamp per id rather than
            // overwriting on every re-fetch (image proxies/clients can
            // refetch the same pixel more than once).
            let existing: string | null = null;
            try {
              existing = await store.get(id, { type: "text" });
            } catch {
              existing = null;
            }
            if (!existing) {
              await store.set(id, new Date().toISOString());
            }
          }
        } catch (err) {
          // Storage hiccups should never break pixel delivery.
          console.error("track pixel storage error", err);
        }
        return pixelResponse();
      },
    },
  },
});
